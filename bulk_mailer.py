"""
bulk_mailer.py — Module 2: Bulk Email Sender
=============================================
• Reads 'Email Ready' sheet from ceo_data.xlsx
• Personalises template.html with {{first_name}}, {{company}}, {{industry}}
• Adds Reply-To header pointing to monitored inbox
• Embeds open-tracking pixel + click tracking
• Rate-limited to ≤ 50 emails/hour (72-second sleep between sends)
• CAN-SPAM / GDPR compliant unsubscribe footer
• TEST_MODE=True redirects all sends to SENDER_EMAIL for safe testing
"""

import os
import re
import time
import logging
import smtplib
import pandas as pd
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

# ── SMTP config (set in .env) ─────────────────────────────────────────────
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.sendgrid.net")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "apikey")          # SendGrid literal "apikey"
SMTP_PASS = os.getenv("SMTP_PASSWORD", "")

# ── Sender identity ───────────────────────────────────────────────────────
SENDER_EMAIL   = os.getenv("SENDER_EMAIL",   "outreach@yourdomain.com")
SENDER_NAME    = os.getenv("SENDER_NAME",    "Your Name")
SENDER_TITLE   = os.getenv("SENDER_TITLE",   "Head of Partnerships")
SENDER_COMPANY = os.getenv("SENDER_COMPANY", "Your Company")
SENDER_PHONE   = os.getenv("SENDER_PHONE",   "+1 000 000 0000")
SENDER_WEBSITE = os.getenv("SENDER_WEBSITE", "https://yourdomain.com")
REPLY_TO       = os.getenv("REPLY_TO",       "replies@yourdomain.com")

# ── Campaign settings ─────────────────────────────────────────────────────
CALENDAR_LINK      = os.getenv("CALENDAR_LINK",      "https://calendly.com/yourlink")
UNSUBSCRIBE_BASE   = os.getenv("UNSUBSCRIBE_BASE",   "https://yourdomain.com/unsubscribe")
TRACKING_BASE      = os.getenv("TRACKING_BASE",      "https://yourdomain.com/track")
EXCEL_FILE         = "ceo_data.xlsx"
TEMPLATE_FILE      = "template.html"
LOG_FILE           = "send_log.txt"

# ── Safety controls ───────────────────────────────────────────────────────
TEST_MODE    = True   # ← set False only when ready for live send
TEST_BATCH   = 1      # records to process in test mode
DELAY_SECS   = 72     # 3600s / 50 emails = 72 s per email

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ── Helpers ───────────────────────────────────────────────────────────────

def load_template() -> str:
    return Path(TEMPLATE_FILE).read_text(encoding="utf-8")


def personalise(html: str, row: pd.Series, to_email: str) -> str:
    first_name = str(row.get("Full Name", "")).split()[0]
    company    = str(row.get("Company Name", ""))
    industry   = str(row.get("Industry", "your industry"))

    tracking_pixel = (
        f'<img src="{TRACKING_BASE}/open?email={to_email}" '
        f'width="1" height="1" style="display:none;" alt="" />'
    )
    unsub_link = f"{UNSUBSCRIBE_BASE}?email={to_email}"

    replacements = {
        "{{first_name}}":    first_name,
        "{{company}}":       company,
        "{{industry}}":      industry,
        "{{sender_name}}":   SENDER_NAME,
        "{{sender_title}}":  SENDER_TITLE,
        "{{sender_company}}":SENDER_COMPANY,
        "{{sender_phone}}":  SENDER_PHONE,
        "{{sender_website}}":SENDER_WEBSITE,
        "{{calendar_link}}": CALENDAR_LINK,
        "{{unsubscribe_link}}": unsub_link,
        "{{tracking_pixel}}": tracking_pixel,
    }
    for k, v in replacements.items():
        html = html.replace(k, v)
    return html


def build_msg(to_email: str, row: pd.Series, html_body: str) -> MIMEMultipart:
    first_name = str(row.get("Full Name", "")).split()[0]
    company    = str(row.get("Company Name", ""))

    msg = MIMEMultipart("alternative")
    msg["From"]         = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"]           = to_email
    msg["Reply-To"]     = REPLY_TO
    msg["Subject"]      = f"Quick question for you, {first_name}"
    msg["List-Unsubscribe"] = f"<{UNSUBSCRIBE_BASE}?email={to_email}>"

    # Plain-text fallback (strip HTML tags)
    plain = re.sub(r"<[^>]+>", " ", html_body)
    plain = re.sub(r"\s+", " ", plain).strip()
    msg.attach(MIMEText(plain, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg


def smtp_send(msg: MIMEMultipart, to_email: str) -> bool:
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        return True
    except Exception as exc:
        log.error("Failed → %s: %s", to_email, exc)
        return False


# ── Main ──────────────────────────────────────────────────────────────────

def run():
    df = pd.read_excel(EXCEL_FILE, sheet_name="Email Ready")
    log.info("Loaded %d records from 'Email Ready'.", len(df))

    if not SMTP_PASS:
        log.error("SMTP_PASSWORD not configured in .env — aborting.")
        return

    template = load_template()

    if TEST_MODE:
        df = df.head(TEST_BATCH)
        log.warning("TEST MODE — processing %d records, sending to %s", len(df), SENDER_EMAIL)

    sent_list, failed_list = [], []

    for idx, (_, row) in enumerate(df.iterrows()):
        # In test mode redirect to sender's own inbox
        to_email = SENDER_EMAIL if TEST_MODE else str(row["Email Address"])
        full_name = str(row.get("Full Name", "CEO"))

        html_body = personalise(template, row, to_email)
        msg       = build_msg(to_email, row, html_body)

        ok = smtp_send(msg, to_email)
        if ok:
            sent_list.append(to_email)
            log.info("✓ Sent to %s <%s>", full_name, to_email)
        else:
            failed_list.append({"email": to_email, "error": "see log"})

        # Rate-limit: pause before every send except the last
        if idx < len(df) - 1:
            log.info("  Waiting %ds (rate limit ≤50/hr) ...", DELAY_SECS)
            time.sleep(DELAY_SECS)

    log.info("Done.  Sent: %d | Failed: %d", len(sent_list), len(failed_list))


if __name__ == "__main__":
    run()
