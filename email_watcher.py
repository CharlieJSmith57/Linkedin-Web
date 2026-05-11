"""
email_watcher.py — Automated LinkedIn export email watcher

Waits for LinkedIn's first installment email:
  Subject: "The first installment of your LinkedIn data archive is ready!"
  From:    messages-noreply@linkedin.com

This email contains Connections.csv and arrives within ~10 minutes
of requesting the archive on LinkedIn.

Recency check: only accepts emails received AFTER this script started,
using the exact parsed Date header — not just the calendar day.
"""

from __future__ import annotations

import email
import email.header
import email.utils
import imaplib
import logging
import re
import time
import zipfile
from datetime import datetime, timezone, timedelta
from html.parser import HTMLParser
from pathlib import Path

import requests
import config

log = logging.getLogger(__name__)

GMAIL_IMAP_HOST  = "imap.gmail.com"
GMAIL_IMAP_PORT  = 993
POLL_INTERVAL    = 60    # seconds between inbox checks
MAX_WAIT         = 3600  # 1 hour timeout

LINKEDIN_SENDER  = "messages-noreply@linkedin.com"

# Subject must contain this string (case-insensitive)
TARGET_SUBJECT   = "first installment of your linkedin data archive is ready"

# Record exactly when this script started (UTC, timezone-aware)
# Subtract a small buffer for clock skew between Gmail and local machine
SCRIPT_START_UTC = datetime.now(timezone.utc) - timedelta(minutes=2)


# ── HTML link extractor ───────────────────────────────────────────────────────

class LinkExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            for name, value in attrs:
                if name == "href" and value:
                    self.links.append(value)


def extract_links_from_html(html: str) -> list[str]:
    parser = LinkExtractor()
    parser.feed(html)
    return parser.links


# ── Gmail ─────────────────────────────────────────────────────────────────────

def connect_gmail() -> imaplib.IMAP4_SSL:
    log.info("Connecting to Gmail as %s...", config.GMAIL_ADDRESS)
    mail = imaplib.IMAP4_SSL(GMAIL_IMAP_HOST, GMAIL_IMAP_PORT)
    mail.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
    mail.select("inbox")
    log.info("Connected.")
    return mail


def decode_subject(raw: str) -> str:
    parts = email.header.decode_header(raw)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="ignore"))
        else:
            result.append(part)
    return "".join(result)


def parse_email_date(date_str: str) -> datetime | None:
    """Parse an email Date header into a timezone-aware datetime."""
    try:
        parsed = email.utils.parsedate_to_datetime(date_str)
        # Ensure timezone-aware
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed
    except Exception:
        return None


def find_linkedin_email(mail: imaplib.IMAP4_SSL) -> str | None:
    """
    Search for the first-installment LinkedIn email received after
    SCRIPT_START_UTC. Returns download URL or None.
    """
    since_str = SCRIPT_START_UTC.strftime("%d-%b-%Y")
    _, data = mail.search(None, f'(FROM "{LINKEDIN_SENDER}" SINCE "{since_str}")')
    msg_ids = data[0].split()

    if not msg_ids:
        return None

    for msg_id in reversed(msg_ids):
        _, msg_data = mail.fetch(msg_id, "(RFC822)")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        subject  = decode_subject(msg.get("Subject", ""))
        date_str = msg.get("Date", "")
        received = parse_email_date(date_str)

        # ── Recency check: exact timestamp, not just calendar day ──────────
        if received is None:
            log.debug("Could not parse date for email '%s' — skipping", subject)
            continue

        if received < SCRIPT_START_UTC:
            log.debug(
                "Skipping old email from %s: '%s' (received %s, script started %s)",
                date_str, subject,
                received.strftime("%Y-%m-%d %H:%M:%S UTC"),
                SCRIPT_START_UTC.strftime("%Y-%m-%d %H:%M:%S UTC"),
            )
            continue

        # ── Subject check ──────────────────────────────────────────────────
        if TARGET_SUBJECT not in subject.lower():
            log.debug("Subject doesn't match: '%s'", subject)
            continue

        log.info("✓ Found matching email: '%s' (received %s)",
                 subject, received.strftime("%Y-%m-%d %H:%M UTC"))

        url = extract_download_url(msg)
        if url:
            return url
        else:
            log.warning("Email matched but no download link found — will retry.")

    return None


def extract_download_url(msg: email.message.Message) -> str | None:
    """Extract the LinkedIn archive download URL from the email body."""
    for part in msg.walk():
        content_type = part.get_content_type()
        try:
            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
        except Exception:
            continue

        if content_type == "text/html":
            for link in extract_links_from_html(body):
                if _is_linkedin_download_link(link):
                    log.info("Download URL found: %s...", link[:80])
                    return link

        elif content_type == "text/plain":
            for m in re.findall(r'https?://[^\s<>"]+', body):
                if _is_linkedin_download_link(m):
                    log.info("Download URL found (plain text): %s...", m[:80])
                    return m

    return None


def _is_linkedin_download_link(url: str) -> bool:
    if not url:
        return False
    u = url.lower()
    return "linkedin.com" in u and any(
        k in u for k in ("ambry", "dms/ambry", "checkpoint", "download", "archive")
    )


# ── Download + extract ────────────────────────────────────────────────────────

def download_and_extract(url: str) -> Path | None:
    inbox = Path(config.CSV_INBOX_DIR)
    inbox.mkdir(parents=True, exist_ok=True)
    zip_path = inbox / "linkedin_export.zip"

    log.info("Downloading archive ZIP...")
    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        }
        r = requests.get(url, headers=headers, stream=True, timeout=120)
        r.raise_for_status()

        with open(zip_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=16384):
                f.write(chunk)

        size_kb = zip_path.stat().st_size / 1024
        log.info("Downloaded: %.1f KB", size_kb)

        if size_kb < 1:
            log.error("ZIP too small — LinkedIn may require a fresh browser session to download.")
            return None

    except Exception as e:
        log.error("Download failed: %s", e)
        return None

    # Extract Connections.csv
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            names = zf.namelist()
            log.info("ZIP contents: %s", names)

            csv_names = [
                n for n in names
                if "connection" in n.lower() and n.endswith(".csv")
            ]

            if not csv_names:
                log.error("No Connections CSV found in ZIP. Contents: %s", names)
                return None

            extracted = zf.extract(csv_names[0], inbox)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            final_path = inbox / f"Connections_{ts}.csv"
            Path(extracted).rename(final_path)
            log.info("Extracted: %s", final_path.name)
            return final_path

    except zipfile.BadZipFile:
        log.error("Not a valid ZIP — the download link may have expired.")
        return None
    except Exception as e:
        log.error("Extraction error: %s", e)
        return None
    finally:
        if zip_path.exists():
            zip_path.unlink()


# ── Main watcher loop ─────────────────────────────────────────────────────────

def wait_for_export_email() -> Path | None:
    """
    Poll Gmail until the first-installment LinkedIn email arrives.
    Only accepts emails received after this script started.
    Returns path to Connections.csv, or None on failure/timeout.
    """
    log.info("=" * 60)
    log.info("Email watcher started at %s UTC",
             SCRIPT_START_UTC.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("Watching for: '%s'", TARGET_SUBJECT)
    log.info("Only emails received AFTER script start will be accepted.")
    log.info("Expected arrival: ~10 minutes after requesting archive.")
    log.info("=" * 60)

    deadline = time.time() + MAX_WAIT
    mail = None

    try:
        mail = connect_gmail()
    except imaplib.IMAP4.error as e:
        log.error("Gmail login failed: %s", e)
        log.error(
            "\nCheck:\n"
            "  1. GMAIL_ADDRESS in config.py is correct\n"
            "  2. GMAIL_APP_PASSWORD is the 16-char App Password\n"
            "  3. 2-Step Verification is enabled\n"
            "  4. IMAP is enabled in Gmail settings\n"
        )
        return None

    try:
        while time.time() < deadline:
            remaining = int(deadline - time.time())
            log.info("Checking inbox... (%dm %ds remaining)",
                     remaining // 60, remaining % 60)

            try:
                mail.select("inbox")
                url = find_linkedin_email(mail)
            except imaplib.IMAP4.abort:
                log.warning("IMAP connection dropped, reconnecting...")
                try:
                    mail.logout()
                except Exception:
                    pass
                time.sleep(5)
                mail = connect_gmail()
                continue

            if url:
                return download_and_extract(url)

            log.info("Not found yet — next check in %ds.", POLL_INTERVAL)
            time.sleep(POLL_INTERVAL)

    finally:
        try:
            mail.logout()
        except Exception:
            pass

    log.error("Timed out after 1 hour. No matching email received.")
    log.error("You can still continue manually: drop Connections.csv into %s",
              config.CSV_INBOX_DIR)
    return None


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [EMAIL] %(levelname)s %(message)s",
    )
    result = wait_for_export_email()
    print("\nResult:", result or "FAILED")
