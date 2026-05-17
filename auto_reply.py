"""
auto_reply.py — Module 3: Auto-Reply Automation
================================================
• Polls IMAP inbox (replies@yourdomain.com) every 60 seconds for UNSEEN messages
• Matches In-Reply-To / From header against ceo_data.xlsx
• Dispatches personalised auto-reply within 5 seconds of detection
• Injects {{first_name}}, {{company}}, {{meeting_link}} into template
• Logs every reply to Sheet 3 'Replies Log' in ceo_data.xlsx
"""

import os
import email
import email.utils
import imaplib
import logging
import smtplib
import threading
import time
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from dotenv import load_dotenv

load_dotenv()

# ── IMAP (incoming) ───────────────────────────────────────────────────────
IMAP_HOST  = os.getenv("IMAP_HOST", "imap.gmail.com")
IMAP_PORT  = int(os.getenv("IMAP_PORT", "993"))
EMAIL_USER = os.getenv("IMAP_USER", "replies@yourdomain.com")
EMAIL_PASS = os.getenv("IMAP_PASSWORD", "")

# ── SMTP (outgoing) ───────────────────────────────────────────────────────
SMTP_HOST  = os.getenv("SMTP_HOST", "smtp.sendgrid.net")
SMTP_PORT  = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER  = os.getenv("SMTP_USER", "apikey")
SMTP_PASS  = os.getenv("SMTP_PASSWORD", "")

# ── Sender identity ───────────────────────────────────────────────────────
SENDER_NAME    = os.getenv("SENDER_NAME",    "Your Name")
SENDER_TITLE   = os.getenv("SENDER_TITLE",   "Head of Partnerships")
SENDER_COMPANY = os.getenv("SENDER_COMPANY", "Your Company")
SENDER_PHONE   = os.getenv("SENDER_PHONE",   "+1 000 000 0000")
SENDER_WEBSITE = os.getenv("SENDER_WEBSITE", "https://yourdomain.com")
MEETING_LINK   = os.getenv("CALENDAR_LINK",  "https://calendly.com/yourlink")
UNSUBSCRIBE_BASE = os.getenv("UNSUBSCRIBE_BASE", "https://yourdomain.com/unsubscribe")

# ── Files ─────────────────────────────────────────────────────────────────
EXCEL_FILE    = "ceo_data.xlsx"
TEMPLATE_FILE = "auto_reply_template.txt"
LOG_FILE      = "listener_log.txt"
LOG_SHEET     = "Replies Log"

POLL_INTERVAL = 60  # seconds between inbox checks

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)
_excel_lock = threading.Lock()

# ── Helpers ───────────────────────────────────────────────────────────────

def load_ceo_df() -> pd.DataFrame:
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name="CEO Master List")
    except Exception as exc:
        log.error("Cannot read CEO data: %s", exc)
        return pd.DataFrame()

def lookup_sender(from_email: str, df: pd.DataFrame) -> dict:
    """Return CEO row dict or sensible defaults."""
    if df.empty:
        return {}
    hit = df[df["Email Address"].str.lower() == from_email.lower()]
    return hit.iloc[0].to_dict() if not hit.empty else {}

def get_auto_reply(sender_email: str, sender_name: str) -> str:
    """Build personalised auto-reply body from template."""
    df  = load_ceo_df()
    ceo = lookup_sender(sender_email, df)

    first_name = str(ceo.get("Full Name", sender_name)).split()[0] if ceo else sender_name.split()[0]
    company    = str(ceo.get("Company Name", "your company")) if ceo else "your company"

    template = Path(TEMPLATE_FILE).read_text(encoding="utf-8")
    return (
        template
        .replace("{{first_name}}",    first_name)
        .replace("{{company}}",       company)
        .replace("{{meeting_link}}",  MEETING_LINK)
        .replace("{{sender_name}}",   SENDER_NAME)
        .replace("{{sender_title}}",  SENDER_TITLE)
        .replace("{{sender_company}}",SENDER_COMPANY)
        .replace("{{sender_phone}}",  SENDER_PHONE)
        .replace("{{sender_website}}",SENDER_WEBSITE)
        .replace("{{unsubscribe_link}}", f"{UNSUBSCRIBE_BASE}?email={sender_email}")
    )

def send_auto_reply(to_email: str, to_name: str, original_subject: str, body: str) -> bool:
    try:
        msg = MIMEMultipart()
        msg["From"]    = EMAIL_USER
        msg["To"]      = to_email
        msg["Subject"] = f"Re: {original_subject} [Auto-Reply]"
        msg.attach(MIMEText(body, "plain", "utf-8"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.ehlo()
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.sendmail(EMAIL_USER, to_email, msg.as_string())
        log.info("[%s] Auto-replied to %s <%s>", datetime.now().strftime("%H:%M:%S"), to_name, to_email)
        return True
    except Exception as exc:
        log.error("Auto-reply FAILED for %s: %s", to_email, exc)
        return False

def append_replies_log(entry: dict):
    """Thread-safe append to 'Replies Log' sheet."""
    with _excel_lock:
        wb = load_workbook(EXCEL_FILE)

        if LOG_SHEET not in wb.sheetnames:
            ws = wb.create_sheet(LOG_SHEET)
            headers = [
                "Timestamp (UTC)", "From Email", "Sender Name",
                "CEO Name", "Company", "Original Subject",
                "Reply Sent At", "Status", "Reply Preview",
            ]
            ws.append(headers)
            for cell in ws[1]:
                cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
                cell.fill      = PatternFill("solid", start_color="1F4E79")
                cell.alignment = Alignment(horizontal="center")
        else:
            ws = wb[LOG_SHEET]

        ws.append([
            entry["timestamp"],
            entry["from_email"],
            entry["sender_name"],
            entry["ceo_name"],
            entry["company"],
            entry["subject"],
            entry["reply_sent_at"],
            entry["status"],
            entry["reply_preview"],
        ])
        wb.save(EXCEL_FILE)

def process_message(raw_bytes: bytes, imap, uid: bytes):
    """Parse one email and dispatch auto-reply within 5 seconds."""
    t_detect = time.time()

    msg         = email.message_from_bytes(raw_bytes)
    from_raw    = msg.get("From", "")
    sender_name, sender_email = email.utils.parseaddr(from_raw)
    subject     = msg.get("Subject", "(no subject)")
    in_reply_to = msg.get("In-Reply-To", "")

    log.info("New reply from %s <%s> | Subject: %s", sender_name, sender_email, subject)

    # Cross-reference against CEO data
    df  = load_ceo_df()
    ceo = lookup_sender(sender_email, df)
    
    # ── THE GATEKEEPER FIX ────────────────────────────────────────────────
    # If the email address is NOT in the Excel sheet, ignore it!
    if not ceo:
        log.info("   ↳ Ignored: %s is not in our CEO Master List.", sender_email)
        imap.store(uid, "+FLAGS", "\\Seen") # Mark as seen so we don't check it again
        return
    # ──────────────────────────────────────────────────────────────────────

    ceo_name = str(ceo.get("Full Name", sender_name))
    company  = str(ceo.get("Company Name", "Unknown"))

    # Build + dispatch reply
    body = get_auto_reply(sender_email, sender_name)
    ok   = send_auto_reply(sender_email, ceo_name, subject, body)

    elapsed = time.time() - t_detect
    log.info("   ↳ Reply dispatched in %.1f seconds.", elapsed)

    # Log to Excel
    append_replies_log({
        "timestamp":     datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "from_email":    sender_email,
        "sender_name":   sender_name,
        "ceo_name":      ceo_name,
        "company":       company,
        "subject":       subject,
        "reply_sent_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        "status":        "Sent" if ok else "Failed",
        "reply_preview": body[:120].replace("\n", " ") + "…",
    })

    # Mark as Seen
    imap.store(uid, "+FLAGS", "\\Seen")

def listen_inbox():
    """Main polling loop — runs indefinitely."""
    seen_uids: set = set()
    log.info("Auto-reply listener started. Polling every %ds.", POLL_INTERVAL)

    while True:
        try:
            mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
            mail.login(EMAIL_USER, EMAIL_PASS)
            mail.select("inbox")

            _, data = mail.search(None, "UNSEEN")
            uids = data[0].split()

            if uids:
                log.info("Found %d unseen message(s).", len(uids))

            for uid in uids:
                if uid in seen_uids:
                    continue
                seen_uids.add(uid)
                _, raw = mail.fetch(uid, "(RFC822)")
                process_message(raw[0][1], mail, uid)

            mail.logout()

        except imaplib.IMAP4.error as exc:
            log.error("IMAP auth/connection error: %s", exc)
        except Exception as exc:
            log.error("Unexpected error: %s", exc)

        log.info("Sleeping %ds …", POLL_INTERVAL)
        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    if not EMAIL_PASS:
        log.error("IMAP_PASSWORD not set in .env — aborting.")
    else:
        listen_inbox()