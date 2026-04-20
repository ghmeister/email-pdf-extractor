import os
import threading
import time
import imaplib
import email
import logging
import tempfile
from datetime import datetime, timedelta
from functools import wraps

import requests
from flask import Flask, render_template, request, redirect, url_for, flash, Response
from flask_sqlalchemy import SQLAlchemy
from PyPDF2 import PdfReader
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret")
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get("DATABASE_URL", "sqlite:///data/app.db")
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

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


with app.app_context():
    db.create_all()
    if Rule.query.count() == 0:
        db.session.add(Rule(
            name="Invoices",
            enabled=True,
            subject_contains="invoice",
            pdf_text_contains="invoice",
        ))
        db.session.commit()


# --- Auth ---

def _check_auth(username, password):
    ui_user = os.environ.get("UI_USER", "admin")
    ui_password = os.environ.get("UI_PASSWORD", "")
    return bool(ui_password) and username == ui_user and password == ui_password


def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not _check_auth(auth.username, auth.password):
            return Response(
                "Authentication required.",
                401,
                {"WWW-Authenticate": 'Basic realm="Email PDF Extractor"'},
            )
        return f(*args, **kwargs)
    return decorated


# --- Polling state ---

_poll_thread: threading.Thread | None = None
_poll_last_at: datetime | None = None

# --- OAuth token cache ---

_token_cache: dict = {"token": None, "expires_at": 0.0}


# --- Helpers ---

def log_message(message, level="INFO"):
    entry = LogEntry(level=level, message=message)
    db.session.add(entry)
    db.session.commit()
    if level == "ERROR":
        logger.error(message)
    else:
        logger.info(message)


def get_env(name, default=None):
    return os.environ.get(name, default)


def load_rules():
    return Rule.query.filter_by(enabled=True).all()


def _condition_matches(condition_str, target):
    """Match if any pipe-separated term appears in target (OR logic)."""
    terms = [t.strip() for t in condition_str.split("|") if t.strip()]
    return any(t.lower() in target.lower() for t in terms)


def match_rule(message, attachment_name, pdf_text):
    for rule in load_rules():
        if rule.sender_contains and not _condition_matches(rule.sender_contains, message.get("From", "")):
            continue
        if rule.subject_contains and not _condition_matches(rule.subject_contains, message.get("Subject", "")):
            continue
        if rule.filename_contains and not _condition_matches(rule.filename_contains, attachment_name):
            continue
        if rule.body_contains:
            body = extract_email_body(message)
            if not _condition_matches(rule.body_contains, body):
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


def get_access_token():
    now = time.time()
    if _token_cache["token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["token"]

    refresh_token = get_env("ONEDRIVE_REFRESH_TOKEN")
    client_id = get_env("ONEDRIVE_CLIENT_ID")
    client_secret = get_env("ONEDRIVE_CLIENT_SECRET")
    tenant_id = get_env("ONEDRIVE_TENANT_ID")
    if not all([refresh_token, client_id, client_secret, tenant_id]):
        raise ValueError("Missing OneDrive OAuth configuration.")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    response = requests.post(token_url, data={
        "grant_type": "refresh_token",
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "scope": "offline_access Files.ReadWrite.All openid profile email",
    })
    response.raise_for_status()
    data = response.json()
    _token_cache["token"] = data["access_token"]
    _token_cache["expires_at"] = now + data.get("expires_in", 3600)
    return _token_cache["token"]


def upload_to_onedrive(filename, content_bytes):
    access_token = get_access_token()
    folder_path = get_env("ONEDRIVE_FOLDER_PATH", "/Documents/Invoices").strip("/")
    safe_folder = "/".join(requests.utils.quote(p, safe="") for p in folder_path.split("/"))
    safe_filename = requests.utils.quote(filename, safe="")
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{safe_folder}/{safe_filename}:/content"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/pdf",
    }
    response = requests.put(upload_url, headers=headers, data=content_bytes)
    response.raise_for_status()
    log_message(f"Uploaded '{filename}' to OneDrive folder '{folder_path}'.")


def _purge_old_logs():
    retention_days = int(get_env("LOG_RETENTION_DAYS", 30))
    cutoff = datetime.utcnow() - timedelta(days=retention_days)
    deleted = LogEntry.query.filter(LogEntry.created_at < cutoff).delete()
    if deleted:
        db.session.commit()
        logger.info("Purged %d log entries older than %d days.", deleted, retention_days)


def poll_inbox():
    global _poll_last_at
    host = get_env("IMAP_HOST")
    port = int(get_env("IMAP_PORT", 993))
    username = get_env("IMAP_USER")
    password = get_env("IMAP_PASS")
    folder = get_env("IMAP_FOLDER", "INBOX")
    interval = int(get_env("POLL_INTERVAL", 120))

    with app.app_context():
        if not all([host, username, password]):
            log_message("IMAP connection settings are incomplete.", "ERROR")
            return

        while True:
            try:
                _poll_last_at = datetime.utcnow()
                with imaplib.IMAP4_SSL(host, port) as mail:
                    mail.login(username, password)
                    mail.select(folder)
                    status, messages = mail.search(None, "UNSEEN")
                    if status != "OK":
                        log_message("Failed to search inbox.", "ERROR")
                    else:
                        for num in messages[0].split():
                            status, data = mail.fetch(num, "RFC822")
                            if status != "OK":
                                continue
                            message = email.message_from_bytes(data[0][1])
                            upload_failed = False
                            for part in message.walk():
                                content_disposition = part.get("Content-Disposition", "")
                                if part.get_content_type() == "application/pdf" or "attachment" in content_disposition.lower():
                                    filename = part.get_filename() or "attachment.pdf"
                                    if filename.lower().endswith(".pdf"):
                                        content = part.get_payload(decode=True)
                                        with tempfile.SpooledTemporaryFile(mode="w+b", buffering=0) as tmp:
                                            tmp.write(content)
                                            tmp.seek(0)
                                            pdf_text = extract_pdf_text(tmp)
                                        if match_rule(message, filename, pdf_text):
                                            try:
                                                upload_to_onedrive(filename, content)
                                            except Exception as exc:
                                                log_message(f"Upload failed for '{filename}': {exc}", "ERROR")
                                                upload_failed = True
                                        else:
                                            log_message(f"Attachment '{filename}' did not match any rules.")
                            if not upload_failed:
                                mail.store(num, "+FLAGS", "\\Seen")
                _purge_old_logs()
            except Exception as exc:
                log_message(f"Email polling error: {exc}", "ERROR")

            time.sleep(interval)


# --- Routes ---

@app.route("/")
@require_auth
def index():
    rules = Rule.query.order_by(Rule.id).all()
    logs = LogEntry.query.order_by(LogEntry.created_at.desc()).limit(10).all()
    return render_template("index.html", rules=rules, logs=logs)


@app.route("/rules", methods=["GET", "POST"])
@require_auth
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
@require_auth
def edit_rule(rule_id):
    rule = Rule.query.get_or_404(rule_id)
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
@require_auth
def delete_rule(rule_id):
    rule = Rule.query.get_or_404(rule_id)
    db.session.delete(rule)
    db.session.commit()
    flash("Rule deleted.", "success")
    return redirect(url_for("rules"))


@app.route("/logs")
@require_auth
def logs():
    page = int(request.args.get("page", 1))
    per_page = 25
    entries = LogEntry.query.order_by(LogEntry.created_at.desc()).paginate(page=page, per_page=per_page, error_out=False)
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
