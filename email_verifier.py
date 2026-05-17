"""
email_verifier.py — Free corporate email finder & verifier
===========================================================
No API key required. Uses three techniques in order:

  1. Pattern generation   — builds common corporate email formats
  2. MX record check      — confirms the domain actually receives mail
  3. SMTP RCPT probe      — connects to mail server and tests if the
                            address is accepted WITHOUT sending any email
                            (some servers block this — gracefully falls back)
"""

import re
import dns.resolver
import smtplib
import socket
import time
import logging
import os
import requests
from typing import Optional

log = logging.getLogger(__name__)

# ── Optional free-tier API keys (set in .env) ─────────────────────────────
SKRAPP_KEY  = os.getenv("SKRAPP_API_KEY", "")
SNOV_CLIENT = os.getenv("SNOV_CLIENT_ID", "")
SNOV_SECRET = os.getenv("SNOV_CLIENT_SECRET", "")
APOLLO_KEY  = os.getenv("APOLLO_API_KEY", "")

EMAIL_RE = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

# Most common corporate email patterns (ordered by frequency)
PATTERNS = [
    "{first}.{last}",
    "{first}",
    "{f}{last}",
    "{first}{last}",
    "{first}_{last}",
    "{first}-{last}",
    "{f}.{last}",
    "{last}.{first}",
    "{last}{f}",
]

# ── 1. Pattern generator ───────────────────────────────────────────────────

def generate_candidates(first: str, last: str, domain: str) -> list[str]:
    """Return all pattern-based email candidates for a person + domain."""
    first = first.lower().strip()
    last  = last.lower().strip()
    f     = first[0] if first else ""
    l     = last[0]  if last  else ""

    candidates = []
    for pattern in PATTERNS:
        local = (
            pattern
            .replace("{first}", first)
            .replace("{last}",  last)
            .replace("{f}",     f)
            .replace("{l}",     l)
        )
        email = f"{local}@{domain}"
        if EMAIL_RE.match(email):
            candidates.append(email)
    return candidates

# ── 2. MX record check ────────────────────────────────────────────────────

def get_mx_server(domain: str) -> Optional[str]:
    """Return the highest-priority MX host for a domain, or None."""
    try:
        records = dns.resolver.resolve(domain, "MX", lifetime=5)
        sorted_records = sorted(records, key=lambda r: r.preference)
        return str(sorted_records[0].exchange).rstrip(".")
    except Exception:
        return None

def domain_has_mx(domain: str) -> bool:
    return get_mx_server(domain) is not None

# ── 3. SMTP RCPT probe ────────────────────────────────────────────────────

def smtp_verify(email_addr: str, from_addr: str = "verify@checker.io",
                timeout: int = 10) -> Optional[bool]:
    """
    Connect to the domain's MX server and issue RCPT TO.
    Returns:
      True  — address accepted (250)
      False — address rejected (550/551)
      None  — server blocked the check (greylisting / catch-all / timeout)
    """
    domain = email_addr.split("@")[1]
    mx     = get_mx_server(domain)
    if not mx:
        return None
    try:
        with smtplib.SMTP(timeout=timeout) as smtp:
            smtp.connect(mx, 25)
            smtp.helo(socket.getfqdn())
            smtp.mail(from_addr)
            code, _ = smtp.rcpt(email_addr)
            smtp.quit()
            if code == 250:
                return True
            if code in (550, 551, 553):
                return False
            return None          # Ambiguous (e.g. 452 greylisting)
    except (smtplib.SMTPConnectError, socket.timeout, OSError):
        return None              # Port 25 blocked (common on cloud VMs)
    except Exception:
        return None

# ── Free-tier API integrations ────────────────────────────────────────────

def skrapp_find(first: str, last: str, company_domain: str) -> Optional[str]:
    if not SKRAPP_KEY: return None
    try:
        r = requests.post(
            "https://app.skrapp.io/api/v2/find",
            json={"firstName": first, "lastName": last, "domain": company_domain},
            headers={"X-Access-Key": SKRAPP_KEY, "Content-Type": "application/json"},
            timeout=10,
        )
        return r.json().get("email") or None
    except Exception: return None

def snov_get_token() -> Optional[str]:
    if not SNOV_CLIENT or not SNOV_SECRET: return None
    try:
        r = requests.post(
            "https://api.snov.io/v1/oauth/access_token",
            data={"grant_type": "client_credentials",
                  "client_id": SNOV_CLIENT, "client_secret": SNOV_SECRET},
            timeout=10,
        )
        return r.json().get("access_token")
    except Exception: return None

def snov_find(first: str, last: str, domain: str) -> Optional[str]:
    token = snov_get_token()
    if not token: return None
    try:
        r = requests.post(
            "https://api.snov.io/v1/get-emails-from-names",
            data={"access_token": token, "firstName": first, "lastName": last, "domain": domain},
            timeout=10,
        )
        emails = r.json().get("data", {}).get("items", [])
        if emails: return emails[0].get("email")
    except Exception: pass
    return None

def apollo_find(first: str, last: str, company_name: str) -> Optional[str]:
    if not APOLLO_KEY: return None
    try:
        r = requests.post(
            "https://api.apollo.io/v1/people/match",
            json={"first_name": first, "last_name": last,
                  "organization_name": company_name, "reveal_personal_emails": False},
            headers={"Content-Type": "application/json", "Cache-Control": "no-cache",
                     "X-Api-Key": APOLLO_KEY},
            timeout=10,
        )
        return (r.json().get("person") or {}).get("email")
    except Exception: return None

# ── Master finder ─────────────────────────────────────────────────────────

def find_and_verify_email(first: str, last: str, company_domain: str, company_name: str = "", use_smtp: bool = True) -> dict:
    result = {"email": None, "valid": False, "method": "none", "confidence": "unverified"}

    for api_name, fn, args in [
        ("Apollo.io",  apollo_find,  (first, last, company_name)),
        ("Snov.io",    snov_find,    (first, last, company_domain)),
        ("Skrapp.io",  skrapp_find,  (first, last, company_domain)),
    ]:
        found = fn(*args)
        if found and EMAIL_RE.match(found):
            result.update(email=found, valid=True, method=api_name, confidence="high")
            log.info("  [%s] Found: %s", api_name, found)
            return result

    if not domain_has_mx(company_domain):
        log.warning("  No MX record for %s — skipping", company_domain)
        return result

    candidates = generate_candidates(first, last, company_domain)
    for candidate in candidates:
        if use_smtp:
            verified = smtp_verify(candidate)
            if verified is True:
                result.update(email=candidate, valid=True, method="SMTP probe", confidence="high")
                log.info("  [SMTP] Verified: %s", candidate)
                return result
            if verified is False:
                continue

        if result["email"] is None:
            result.update(email=candidates[0], valid=True, method="Pattern (unverified)", confidence="medium")

    if result["email"]:
        log.info("  [Pattern] Best guess: %s  (%s)", result["email"], result["confidence"])
    else:
        log.warning("  Could not determine email for %s %s @ %s", first, last, company_domain)

    return result

# ── Batch helper ─────────────────────────────────────────────────────────

def enrich_dataframe(df, first_col="Full Name", company_col="Company Name",
                     email_col="Email Address", valid_col="Email Valid"):
    import pandas as pd
    for i, row in df.iterrows():
        if str(row.get(valid_col, False)).lower() == "true":
            continue

        full_name = str(row.get(first_col, ""))
        parts     = full_name.strip().split()
        first     = parts[0] if parts else ""
        last      = parts[-1] if len(parts) > 1 else ""

        # FIX: Extract domain from existing incomplete email (e.g. @apple.com)
        existing_email = str(row.get(email_col, ""))
        if "@" in existing_email:
            domain = existing_email.split("@")[1].strip()
        else:
            continue

        company_name = str(row.get(company_col, ""))
        res = find_and_verify_email(first, last, domain, company_name)

        if res["email"]:
            df.at[i, email_col]  = res["email"]
            df.at[i, valid_col]  = res["valid"]

        time.sleep(0.3)

    return df