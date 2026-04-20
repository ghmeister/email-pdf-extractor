import os
import threading
import time
import imaplib
import email
import logging
import tempfile
from datetime import datetime

import requests
from flask import Flask, render_template, request, redirect, url_for, flash
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
        default_rule = Rule(
            name="Invoices",
            enabled=True,
            subject_contains="invoice",
            pdf_text_contains="invoice",
        )
        db.session.add(default_rule)
        db.session.commit()

def log_message(message, level="INFO"):
    entry = LogEntry(level=level, message=message)
    db.session.add(entry)
    db.session.commit()
    if level == "ERROR":
        logger.error(message)
    else:
        logger.info(message)

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

    rules = Rule.query.order_by(Rule.id).all()
    return render_template("rules.html", rules=rules)

@app.route("/rules/delete/<int:rule_id>", methods=["POST"])
def delete_rule(rule_id):
    rule = Rule.query.get_or_404(rule_id)
    db.session.delete(rule)
    db.session.commit()
    flash("Rule deleted.", "success")
    return redirect(url_for("rules"))

@app.route("/logs")
def logs():
    page = int(request.args.get("page", 1))
    per_page = 25
    entries = LogEntry.query.order_by(LogEntry.created_at.desc()).paginate(page=page, per_page=per_page, error_out=False)
    return render_template("logs.html", entries=entries)

@app.route("/health")
def health():
    return {"status": "ok"}

def get_env(name, default=None):
    return os.environ.get(name, default)

def load_rules():
    return Rule.query.filter_by(enabled=True).all()

def match_rule(message, attachment_name, pdf_text):
    for rule in load_rules():
        if rule.sender_contains and rule.sender_contains.lower() not in message.get("From", "").lower():
            continue
        if rule.subject_contains and rule.subject_contains.lower() not in message.get("Subject", "").lower():
            continue
        if rule.filename_contains and rule.filename_contains.lower() not in attachment_name.lower():
            continue
        if rule.body_contains:
            body = extract_email_body(message).lower()
            if rule.body_contains.lower() not in body:
                continue
        if rule.pdf_text_contains and rule.pdf_text_contains.lower() not in pdf_text.lower():
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
        text = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text.append(page_text)
        return "\n".join(text)
    except Exception as exc:
        log_message(f"Failed to extract PDF text: {exc}", "ERROR")
        return ""

def get_access_token():
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
    return response.json()["access_token"]

def upload_to_onedrive(filename, content_bytes):
    access_token = get_access_token()
    folder_path = get_env("ONEDRIVE_FOLDER_PATH", "/Documents/Invoices").strip("/")
    upload_url = (
        f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}/{filename}:/content"
    )
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/pdf",
    }
    response = requests.put(upload_url, headers=headers, data=content_bytes)
    response.raise_for_status()
    log_message(f"Uploaded '{filename}' to OneDrive folder '{folder_path}'.")

def poll_inbox():
    host = get_env("IMAP_HOST")
    port = int(get_env("IMAP_PORT", 993))
    username = get_env("IMAP_USER")
    password = get_env("IMAP_PASS")
    folder = get_env("IMAP_FOLDER", "INBOX")
    interval = int(get_env("POLL_INTERVAL", 120))

    if not all([host, username, password]):
        log_message("IMAP connection settings are incomplete.", "ERROR")
        return

    while True:
        try:
            with imaplib.IMAP4_SSL(host, port) as mail:
                mail.login(username, password)
                mail.select(folder)
                status, messages = mail.search(None, "UNSEEN")
                if status != "OK":
                    log_message("Failed to search inbox.", "ERROR")
                    continue

                for num in messages[0].split():
                    status, data = mail.fetch(num, "RFC822")
                    if status != "OK":
                        continue
                    message = email.message_from_bytes(data[0][1])
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
                                        log_message(f"Upload failed for {filename}: {exc}", "ERROR")
                                else:
                                    log_message(f"Attachment '{filename}' did not match any rules.")
                    mail.store(num, "+FLAGS", "\\Seen")
        except Exception as exc:
            log_message(f"Email polling error: {exc}", "ERROR")

        time.sleep(interval)

def start_email_polling():
    thread = threading.Thread(target=poll_inbox, daemon=True)
    thread.start()

if __name__ == "__main__":
    start_email_polling()
    app.run(host="0.0.0.0", port=5000)
