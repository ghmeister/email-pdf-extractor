import json
import os
import threading
import time
import imaplib
import email
import logging
import tempfile
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import msal
import requests
from sqlalchemy import event
from sqlalchemy.engine import Engine
from sqlalchemy.pool import NullPool
from flask import Flask, render_template, request, redirect, url_for, flash, Response, session
from flask_sqlalchemy import SQLAlchemy
from PyPDF2 import PdfReader
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret")
_APP_VERSION = os.environ.get("APP_VERSION", "dev")


@app.context_processor
def inject_globals():
    return {"app_version": _APP_VERSION}
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", "sqlite:////app/data/app.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "poolclass": NullPool,
    "connect_args": {"check_same_thread": False},
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)


@event.listens_for(Engine, "connect")
def _set_sqlite_pragmas(dbapi_conn, _):
    if isinstance(dbapi_conn, sqlite3.Connection):
        dbapi_conn.execute("PRAGMA journal_mode=WAL")
        dbapi_conn.execute("PRAGMA busy_timeout=10000")


db = SQLAlchemy(app)


class Rule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    enabled = db.Column(db.Boolean, default=True)
    sender_contains = db.Column(db.String(255), default="")
    subject_contains = db.Column(db.String(255), default="")
    filename_contains = db.Column(db.String(255), default="")
    body_contains = db.Column(db.String(255), default="")
    pdf_text_contains = db.Column(db.String(255), default="")


class LogEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    level = db.Column(db.String(10), default="INFO")
    message = db.Column(db.Text, nullable=False)


class ProcessedEmail(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    message_id = db.Column(db.Text, unique=True, nullable=False)  # Text: Message-IDs can exceed 255 chars
    processed_at = db.Column(db.DateTime, default=datetime.utcnow)


with app.app_context():
    db.create_all()
    if Rule.query.count() == 0:
        db.session.add(Rule(
            name="Invoice (EN)",
            enabled=True,
            pdf_text_contains="invoice|total due|amount due|payment due|bill to|invoice no|invoice number|invoice date",
        ))
        db.session.add(Rule(
            name="Rechnung (DE)",
            enabled=True,
            pdf_text_contains="rechnung|gesamtbetrag|zahlbar bis|rechnungsbetrag|mwst|mehrwertsteuer|rechnungsnummer|rechnungsdatum",
        ))
        db.session.commit()


# --- Auth ---

_UI_USER = os.environ.get("UI_USER", "admin")
_UI_PASSWORD = os.environ.get("UI_PASSWORD", "")


@app.before_request
def _require_auth():
    if not _UI_PASSWORD:
        return  # auth not configured — allow all (dev/local)
    if request.endpoint in ("login", "logout", "health", "static"):
        return
    if not session.get("logged_in"):
        return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if (request.form.get("username") == _UI_USER
                and request.form.get("password") == _UI_PASSWORD):
            session.permanent = True
            session["logged_in"] = True
            return redirect(request.args.get("next") or "/")
        error = "Invalid credentials"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# --- Polling state ---

_poll_thread: threading.Thread | None = None
_poll_last_at: datetime | None = None

# --- MSAL / OneDrive auth ---

_ONEDRIVE_SCOPES = ["Files.ReadWrite"]
_TOKEN_CACHE_PATH = Path(os.environ.get("TOKEN_CACHE_PATH", "/app/data/.token_cache.json"))
_msal_app: msal.PublicClientApplication | None = None
_msal_cache: msal.SerializableTokenCache | None = None


# --- Helpers ---

def log_message(message, level="INFO"):
    try:
        entry = LogEntry(level=level, message=message)
        db.session.add(entry)
        db.session.commit()
    except Exception:
        db.session.rollback()
    if level == "ERROR":
        logger.error(message)
    else:
        logger.info(message)


def load_rules():
    return Rule.query.filter_by(enabled=True).all()


def _condition_matches(condition_str, target):
    """Match if any pipe-separated term appears in target (OR logic)."""
    terms = [t.strip() for t in condition_str.split("|") if t.strip()]
    return any(t.lower() in target.lower() for t in terms)


def match_rule(message, attachment_name, pdf_text, email_body=""):
    for rule in load_rules():
        if rule.sender_contains and not _condition_matches(rule.sender_contains, message.get("From", "")):
            continue
        if rule.subject_contains and not _condition_matches(rule.subject_contains, message.get("Subject", "")):
            continue
        if rule.filename_contains and not _condition_matches(rule.filename_contains, attachment_name):
            continue
        if rule.body_contains and not _condition_matches(rule.body_contains, email_body):
            continue
        if rule.pdf_text_contains and not _condition_matches(rule.pdf_text_contains, pdf_text):
            continue
        return True
    return False


def extract_email_body(message):
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/plain" and not part.get("Content-Disposition"):
                return part.get_payload(decode=True).decode(errors="ignore")
    elif message.get_content_type() == "text/plain":
        return message.get_payload(decode=True).decode(errors="ignore")
    return ""


def extract_pdf_text(buffer):
    try:
        reader = PdfReader(buffer)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        log_message(f"Failed to extract PDF text: {exc}", "ERROR")
        return ""


def _get_msal_app() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    global _msal_app, _msal_cache
    if _msal_app is None:
        client_id = os.environ.get("ONEDRIVE_CLIENT_ID")
        if not client_id:
            raise ValueError("ONEDRIVE_CLIENT_ID is required.")
        _msal_cache = msal.SerializableTokenCache()
        if _TOKEN_CACHE_PATH.exists():
            _msal_cache.deserialize(_TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
        _msal_app = msal.PublicClientApplication(
            client_id,
            authority="https://login.microsoftonline.com/consumers",
            token_cache=_msal_cache,
        )
    return _msal_app, _msal_cache


def _persist_msal_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        _TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        _TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")


def get_access_token() -> str:
    msal_app, cache = _get_msal_app()

    accounts = msal_app.get_accounts()
    if accounts:
        result = msal_app.acquire_token_silent(_ONEDRIVE_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _persist_msal_cache(cache)
            return result["access_token"]

    # No cached token — device code flow (blocks until user authenticates)
    flow = msal_app.initiate_device_flow(scopes=_ONEDRIVE_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Could not start device flow: {flow}")

    log_message(f"ONE-TIME LOGIN REQUIRED — {flow['message']}", "INFO")
    result = msal_app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Device flow failed: {result.get('error_description')}")

    _persist_msal_cache(cache)
    return result["access_token"]


def upload_to_onedrive(filename, content_bytes, content_type="application/pdf"):
    access_token = get_access_token()
    folder_path = os.environ.get("ONEDRIVE_FOLDER_PATH", "/Documents/Invoices").strip("/")
    safe_folder = "/".join(requests.utils.quote(p, safe="") for p in folder_path.split("/"))
    safe_filename = requests.utils.quote(filename, safe="")
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{safe_folder}/{safe_filename}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": content_type,
    }
    response = requests.put(upload_url, headers=headers, data=content_bytes)
    response.raise_for_status()
    if content_type == "application/pdf":
        log_message(f"Uploaded '{filename}' to OneDrive folder '{folder_path}'.")


def _upload_sidecar(filename, message):
    """Upload a .meta.json sidecar alongside a PDF so pdf-renamer can read email context."""
    meta = {
        "source": "email",
        "from": message.get("From", ""),
        "subject": message.get("Subject", ""),
        "date": message.get("Date", ""),
        "message_id": message.get("Message-ID", "").strip(),
    }
    try:
        upload_to_onedrive(
            f"{filename}.meta.json",
            json.dumps(meta, ensure_ascii=False).encode("utf-8"),
            content_type="application/json",
        )
    except Exception as exc:
        log_message(f"Failed to upload sidecar for '{filename}': {exc}", "ERROR")


def _purge_old_records():
    now = datetime.utcnow()
    log_cutoff = now - timedelta(days=int(os.environ.get("LOG_RETENTION_DAYS", 30)))
    email_cutoff = now - timedelta(days=int(os.environ.get("PROCESSED_EMAIL_RETENTION_DAYS", 365)))
    deleted_logs = LogEntry.query.filter(LogEntry.created_at < log_cutoff).delete()
    deleted_emails = ProcessedEmail.query.filter(ProcessedEmail.processed_at < email_cutoff).delete()
    db.session.commit()
    if deleted_logs or deleted_emails:
        logger.info("Purged %d log entries and %d processed email records.", deleted_logs, deleted_emails)


def _process_message(message) -> bool:
    """Process all PDF attachments in one email. Returns True if no uploads failed."""
    email_body = extract_email_body(message)
    upload_failed = False
    for part in message.walk():
        is_pdf = part.get_content_type() == "application/pdf"
        is_attachment = "attachment" in part.get("Content-Disposition", "").lower()
        if not (is_pdf or is_attachment):
            continue
        filename = part.get_filename() or "attachment.pdf"
        if not filename.lower().endswith(".pdf"):
            continue
        content = part.get_payload(decode=True)
        with tempfile.SpooledTemporaryFile(mode="w+b", buffering=0) as tmp:
            tmp.write(content)
            tmp.seek(0)
            pdf_text = extract_pdf_text(tmp)
        if match_rule(message, filename, pdf_text, email_body):
            try:
                upload_to_onedrive(filename, content)
                _upload_sidecar(filename, message)
            except Exception as exc:
                log_message(f"Upload failed for '{filename}': {exc}", "ERROR")
                upload_failed = True
        else:
            log_message(f"Attachment '{filename}' did not match any rules.")
    return not upload_failed


def poll_inbox():
    global _poll_last_at
    host = os.environ.get("IMAP_HOST")
    port = int(os.environ.get("IMAP_PORT", 993))
    username = os.environ.get("IMAP_USER")
    password = os.environ.get("IMAP_PASS")
    folder = os.environ.get("IMAP_FOLDER", "INBOX")
    interval = int(os.environ.get("POLL_INTERVAL", 120))
    lookback_days = int(os.environ.get("POLL_LOOKBACK_DAYS", 90))

    with app.app_context():
        if not all([host, username, password]):
            log_message("IMAP connection settings are incomplete.", "ERROR")
            return

        # Eagerly authenticate with OneDrive so device flow runs at startup, not mid-poll
        try:
            get_access_token()
        except Exception as exc:
            log_message(f"OneDrive auth error at startup: {exc}", "ERROR")

        while True:
            try:
                _poll_last_at = datetime.utcnow()
                since_date = (_poll_last_at - timedelta(days=lookback_days)).strftime("%d-%b-%Y")
                with imaplib.IMAP4_SSL(host, port, timeout=30) as mail:
                    mail.login(username, password)
                    mail.select(folder)
                    status, data = mail.uid("search", None, f"UNSEEN SINCE {since_date}")
                    if status != "OK":
                        log_message("Failed to search inbox.", "ERROR")
                    else:
                        for uid in data[0].split():
                            status, msg_data = mail.uid("fetch", uid, "(BODY.PEEK[])")
                            if status != "OK":
                                continue
                            message = email.message_from_bytes(msg_data[0][1])
                            msg_id = message.get("Message-ID", "").strip()
                            if not msg_id:
                                continue
                            if ProcessedEmail.query.filter_by(message_id=msg_id).first():
                                continue
                            if _process_message(message):
                                try:
                                    db.session.add(ProcessedEmail(message_id=msg_id))
                                    db.session.commit()
                                except Exception:
                                    db.session.rollback()
                _purge_old_records()
            except Exception as exc:
                log_message(f"Email polling error: {exc}", "ERROR")
                db.session.rollback()

            time.sleep(interval)


# --- Routes ---

@app.route("/")
def index():
    rules = Rule.query.order_by(Rule.id).all()
    logs = LogEntry.query.order_by(LogEntry.created_at.desc()).limit(10).all()
    return render_template("index.html", rules=rules, logs=logs)


@app.route("/rules", methods=["GET", "POST"])
def rules():
    if request.method == "POST":
        form = request.form
        rule = Rule(
            name=form.get("name", "Unnamed rule"),
            enabled=bool(form.get("enabled")),
            sender_contains=form.get("sender_contains", ""),
            subject_contains=form.get("subject_contains", ""),
            filename_contains=form.get("filename_contains", ""),
            body_contains=form.get("body_contains", ""),
            pdf_text_contains=form.get("pdf_text_contains", ""),
        )
        db.session.add(rule)
        db.session.commit()
        flash("Rule saved.", "success")
        return redirect(url_for("rules"))

    rules_list = Rule.query.order_by(Rule.id).all()
    return render_template("rules.html", rules=rules_list)


@app.route("/rules/edit/<int:rule_id>", methods=["GET", "POST"])
def edit_rule(rule_id):
    rule = db.get_or_404(Rule, rule_id)
    if request.method == "POST":
        form = request.form
        rule.name = form.get("name", rule.name)
        rule.enabled = bool(form.get("enabled"))
        rule.sender_contains = form.get("sender_contains", "")
        rule.subject_contains = form.get("subject_contains", "")
        rule.filename_contains = form.get("filename_contains", "")
        rule.body_contains = form.get("body_contains", "")
        rule.pdf_text_contains = form.get("pdf_text_contains", "")
        db.session.commit()
        flash("Rule updated.", "success")
        return redirect(url_for("rules"))
    return render_template("edit_rule.html", rule=rule)


@app.route("/rules/delete/<int:rule_id>", methods=["POST"])
def delete_rule(rule_id):
    rule = db.get_or_404(Rule, rule_id)
    db.session.delete(rule)
    db.session.commit()
    flash("Rule deleted.", "success")
    return redirect(url_for("rules"))


@app.route("/logs")
def logs():
    page = int(request.args.get("page", 1))
    per_page = 25
    entries = LogEntry.query.order_by(LogEntry.created_at.desc()).paginate(
        page=page, per_page=per_page, error_out=False
    )
    return render_template("logs.html", entries=entries)


@app.route("/health")
def health():
    thread_alive = _poll_thread is not None and _poll_thread.is_alive()
    return {
        "status": "ok" if thread_alive else "degraded",
        "polling": "running" if thread_alive else "stopped",
        "last_poll": _poll_last_at.isoformat() if _poll_last_at else None,
    }, 200 if thread_alive else 503


def start_email_polling():
    global _poll_thread
    _poll_thread = threading.Thread(target=poll_inbox, daemon=True)
    _poll_thread.start()


start_email_polling()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
