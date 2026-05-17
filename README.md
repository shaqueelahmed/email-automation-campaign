# 📧 Email Automation Campaign

> **End-to-end Python automation system** — CEO data extraction, personalised bulk emailing, and intelligent auto-reply.  
> Built as a 3-day assignment. All 3 modules fully implemented and tested.

---

## 📁 Project Structure

```
email_automation_campaign/
├── scraper.py               # Module 1 — CEO data extraction & Excel export
├── bulk_mailer.py           # Module 2 — personalised bulk email sender
├── auto_reply.py            # Module 3 — IMAP inbox listener & auto-reply
├── email_verifier.py        # Free email finder & verifier (no API key needed)
├── template.html            # HTML outreach email template
├── auto_reply_template.txt  # Plain-text auto-reply template
├── ceo_data.xlsx            # Output Excel — 60 CEOs, 3 sheets
├── .env.example             # Environment variable template (copy → .env)
├── requirements.txt         # Python dependencies
└── README.md                # This file
```

---

## ⚙️ Prerequisites

- Python 3.10 or higher
- A [SendGrid](https://sendgrid.com) free account (100 emails/day)
- A Gmail account with 2-Step Verification enabled
- Optional: [Calendly](https://calendly.com) free account for meeting link

---

## 🚀 Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment

```bash
cp .env.example .env
```

Open `.env` and fill in your credentials (see [Environment Variables](#-environment-variables) below).

### 3. Run Module 1 — Data Extraction

```bash
python scraper.py
```

**Output:** `ceo_data.xlsx` with:
- **Sheet 1 "CEO Master List"** — 60 CEOs, all 10 fields, bold headers, sorted by revenue ↓
- **Sheet 2 "Email Ready"** — valid emails only, green/yellow conditional formatting

### 4. Run Module 2 — Bulk Email Sender

> ⚠️ Always test first. `TEST_MODE = True` sends all emails only to your own inbox.

```bash
python bulk_mailer.py
```

Check your inbox — you should receive a personalised email.  
When ready for live send, open `bulk_mailer.py` and set `TEST_MODE = False`.

### 5. Run Module 3 — Auto-Reply Listener

```bash
python auto_reply.py
```

Keeps running. Reply to the test email from Step 4 — within 60 seconds you'll see the auto-reply arrive in your inbox and a new "Replies Log" sheet appear in `ceo_data.xlsx`.

---

## 🔑 Environment Variables

Copy `.env.example` to `.env` and fill in:

```env
# ── SendGrid SMTP ──────────────────────────────────────
SMTP_HOST=smtp.sendgrid.net
SMTP_PORT=587
SMTP_USER=apikey                          # always literally "apikey"
SMTP_PASSWORD=SG.your_sendgrid_key_here  # from SendGrid → Settings → API Keys

# ── Your Sender Identity ───────────────────────────────
SENDER_EMAIL=yourname@gmail.com          # must match SendGrid verified sender
SENDER_NAME=Your Full Name
SENDER_TITLE=Head of Partnerships
SENDER_COMPANY=Your Company Name
SENDER_PHONE=+91 00000 00000
SENDER_WEBSITE=https://yourwebsite.com
REPLY_TO=yourname@gmail.com

# ── Gmail IMAP (for Module 3) ──────────────────────────
IMAP_HOST=imap.gmail.com
IMAP_PORT=993
IMAP_USER=yourname@gmail.com
IMAP_PASSWORD=xxxxxxxxxxxx               # Gmail App Password (16 chars, no spaces)

# ── Campaign Links ─────────────────────────────────────
CALENDAR_LINK=https://calendly.com/your-link
UNSUBSCRIBE_BASE=https://docs.google.com/forms/your-form-id
TRACKING_BASE=https://yourwebsite.com/track

# ── Optional Free Email Finder APIs ───────────────────
APOLLO_API_KEY=your_apollo_key           # apollo.io — 50 free/month
SNOV_CLIENT_ID=your_snov_client_id       # snov.io — 50 free/month
SNOV_CLIENT_SECRET=your_snov_secret
SKRAPP_API_KEY=your_skrapp_key           # skrapp.io — 50 free/month
```

### How to get Gmail App Password

1. Go to **myaccount.google.com → Security**
2. Enable **2-Step Verification**
3. Search **"App Passwords"** → create one named `AutoReply`
4. Copy the 16-character password → paste as `IMAP_PASSWORD` (no spaces)

---

## 📋 Excel Output Schema

### Sheet 1 — CEO Master List (60 records)

| # | Column | Description |
|---|--------|-------------|
| 1 | Full Name | CEO's complete legal name |
| 2 | Company Name | Current employer |
| 3 | Industry | Sector (Tech, Finance, FMCG…) |
| 4 | Country | HQ country — normalised |
| 5 | Email Address | Corporate email — 🟢 green if valid, 🟡 yellow if unverified |
| 6 | Mobile / Contact | Publicly listed phone number |
| 7 | LinkedIn URL | Public profile URL |
| 8 | Net Worth (USD) | Forbes estimate |
| 9 | Company Revenue | Annual revenue in USD billions |
| 10 | Data Source URL | Where the data was sourced |

### Sheet 2 — Email Ready

Filtered view — only rows with validated email addresses.  
Columns: Full Name, Company Name, Industry, Email Address, Company Revenue, LinkedIn URL.

### Sheet 3 — Replies Log *(auto-created by Module 3)*

| Column | Description |
|--------|-------------|
| Timestamp (UTC) | When the reply was detected |
| From Email | Sender's email address |
| Sender Name | Parsed from email header |
| CEO Name | Matched from CEO Master List |
| Company | Matched company name |
| Original Subject | Subject line of the incoming reply |
| Reply Sent At | Timestamp when auto-reply was dispatched |
| Status | "Sent" or "Failed" |
| Reply Preview | First 120 characters of auto-reply body |

---

## 📧 Email Verification (email_verifier.py)

A custom free verification system — **no paid API key required:**

| Method | How It Works | Confidence |
|--------|-------------|------------|
| Pattern Generation | Builds all common formats: `firstname.lastname@`, `f.lastname@`, etc. | Medium |
| MX Record Check | DNS lookup confirms domain accepts email | Medium |
| SMTP RCPT Probe | Connects to mail server, tests address without sending email | High |
| Apollo.io API | 50 free lookups/month — sign up with Gmail | High |
| Snov.io API | 50 free credits/month | High |
| Skrapp.io API | 50 free searches/month | High |

---

## 🔒 Security Best Practices

- ✅ All credentials stored in `.env` — never hardcoded in scripts
- ✅ `.env` excluded from ZIP submission — use `.env.example` instead
- ✅ Gmail App Password used (not main Gmail password)
- ✅ SendGrid API key scoped appropriately
- ✅ `TEST_MODE = True` default prevents accidental live sends

---

## ✅ Compliance

| Requirement | Implementation |
|-------------|----------------|
| CAN-SPAM | Unsubscribe link in every email footer |
| CAN-SPAM | `List-Unsubscribe` header on all outbound emails |
| GDPR | Only publicly available business contact data used |
| Rate Limiting | 72-second delay = max 50 emails/hour |
| Data Source | Forbes, Bloomberg, LinkedIn public profiles only |

> **Note on sending domain:** For this assignment, a Gmail address verified through SendGrid's Single Sender Verification was used for testing. In a production environment, a dedicated domain (e.g. `outreach@yourdomain.com`) with full SPF/DKIM/DMARC authentication via SendGrid's Domain Authentication would be used.

---

## 🛠 Tech Stack

| Library | Purpose |
|---------|---------|
| `pandas` | Data cleaning, deduplication, sorting |
| `openpyxl` | Excel formatting, conditional colours, freeze panes |
| `requests` + `beautifulsoup4` | HTML scraping |
| `selenium` | Browser automation for JS-rendered pages (ready to use) |
| `smtplib` | SMTP email sending |
| `imaplib` | IMAP inbox polling |
| `dnspython` | MX record lookups for email verification |
| `python-dotenv` | `.env` credential loading |

---

## 📦 Deliverables Checklist

- [x] `scraper.py` — working data pipeline with comments
- [x] `bulk_mailer.py` — personalised, rate-limited bulk sender
- [x] `auto_reply.py` — IMAP listener + auto-reply dispatcher
- [x] `email_verifier.py` — free email finder, no API key needed
- [x] `template.html` — HTML email with all variables + tracking
- [x] `auto_reply_template.txt` — auto-reply template with meeting link
- [x] `ceo_data.xlsx` — 60 CEOs, 3 sheets, professionally formatted
- [x] `.env.example` — all credentials documented
- [x] `requirements.txt` — all dependencies
- [x] `README.md` — complete setup and run guide

---

## 📞 Support

For any issues running the modules, check:
1. `send_log.txt` — Module 2 send history and errors
2. `listener_log.txt` — Module 3 polling and reply log
3. Make sure `.env` is in the same folder as the scripts
4. On Windows, if `.env` isn't loading, temporarily hardcode credentials for testing only
