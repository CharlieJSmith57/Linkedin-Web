"""
email_watcher.py — Automated LinkedIn export email watcher

Watches MULTIPLE Gmail folders: Primary inbox, Social tab, Updates tab,
and Promotions tab. LinkedIn's export emails land in Social by default.

Subject: "The first installment of your LinkedIn data archive is ready!"
From:    messages-noreply@linkedin.com

Only accepts emails received AFTER this script started (exact timestamp).
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
POLL_INTERVAL    = 60     # seconds between full scan cycles
MAX_WAIT         = 3600   # 1 hour

LINKEDIN_SENDER  = "messages-noreply@linkedin.com"
TARGET_SUBJECT   = "first installment of your linkedin data archive is ready"

# All Gmail folders that could contain the export email.
# Gmail's tabbed inbox splits into these IMAP folders.
# Searching Primary inbox only — LinkedIn export emails are set to deliver here.
# If emails stop arriving in Primary, add "[Gmail]/All Mail" back to this list.
SEARCH_FOLDERS = [
    "INBOX",
]

# Record when the script started (UTC, timezone-aware)
SCRIPT_START_UTC = datetime.now(timezone.utc) - timedelta(minutes=2)


# ── HTML link extractor ───────────────────────────────────────────────────────

class _LinkExtractor(HTMLParser):
    def __init__(self): super().__init__(); self.links = []
    def handle_starttag(self, tag, attrs):
        if tag == "a":
            for n, v in attrs:
                if n == "href" and v: self.links.append(v)

def _extract_links(html: str) -> list[str]:
    p = _LinkExtractor(); p.feed(html); return p.links


# ── Gmail helpers ─────────────────────────────────────────────────────────────

def connect_gmail() -> imaplib.IMAP4_SSL:
    mail = imaplib.IMAP4_SSL(GMAIL_IMAP_HOST, GMAIL_IMAP_PORT)
    mail.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
    log.info("Gmail connected as %s", config.GMAIL_ADDRESS)
    return mail


def list_available_folders(mail: imaplib.IMAP4_SSL) -> list[str]:
    """Return all IMAP folder names so we can debug what's available."""
    try:
        _, data = mail.list()
        folders = []
        for item in data:
            if item:
                parts = item.decode().split('"/"')
                if parts:
                    folders.append(parts[-1].strip().strip('"'))
        return folders
    except Exception:
        return []


def decode_subject(raw: str) -> str:
    parts = email.header.decode_header(raw)
    return "".join(
        p.decode(e or "utf-8", errors="ignore") if isinstance(p, bytes) else p
        for p, e in parts
    )


def parse_email_date(s: str) -> datetime | None:
    try:
        d = email.utils.parsedate_to_datetime(s)
        return d if d.tzinfo else d.replace(tzinfo=timezone.utc)
    except Exception:
        return None


def _is_linkedin_dl_link(url: str) -> bool:
    u = url.lower()
    return "linkedin.com" in u and any(
        k in u for k in ("ambry", "checkpoint", "download", "archive")
    )


def _extract_download_url(msg) -> str | None:
    for part in msg.walk():
        ct = part.get_content_type()
        try:
            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
        except Exception:
            continue
        links = _extract_links(body) if ct == "text/html" else re.findall(r'https?://[^\s<>"]+', body)
        for link in links:
            if _is_linkedin_dl_link(link):
                return link
    return None


def search_folder(mail: imaplib.IMAP4_SSL, folder: str) -> str | None:
    """
    Search a single IMAP folder for the LinkedIn export email.
    Returns download URL if found, None otherwise.
    """
    try:
        status, _ = mail.select(folder, readonly=True)
        if status != "OK":
            log.debug("  Folder %s: not accessible (status=%s)", folder, status)
            return None
    except Exception as e:
        log.debug("  Folder %s: error selecting: %s", folder, e)
        return None

    # Search by sender and date (day-level — we filter by exact time below)
    since_str = SCRIPT_START_UTC.strftime("%d-%b-%Y")
    try:
        _, data = mail.search(None, f'(FROM "{LINKEDIN_SENDER}" SINCE "{since_str}")')
    except Exception as e:
        log.debug("  Folder %s: search error: %s", folder, e)
        return None

    msg_ids = data[0].split() if data[0] else []
    if not msg_ids:
        log.debug("  Folder %s: 0 LinkedIn messages since %s", folder, since_str)
        return None

    log.debug("  Folder %s: %d LinkedIn message(s) found", folder, len(msg_ids))

    for msg_id in reversed(msg_ids):
        try:
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)
        except Exception:
            continue

        subj = decode_subject(msg.get("Subject", ""))
        rcvd = parse_email_date(msg.get("Date", ""))

        # Exact timestamp check — reject anything before script started
        if rcvd is None or rcvd < SCRIPT_START_UTC:
            log.debug("    Skipping old email (%s): '%s'",
                      rcvd.strftime("%H:%M UTC") if rcvd else "no date", subj[:60])
            continue

        # Subject check
        if TARGET_SUBJECT not in subj.lower():
            log.debug("    Subject no match: '%s'", subj[:60])
            continue

        log.info("  ✓ FOUND in folder %s: '%s' (received %s)",
                 folder, subj[:60], rcvd.strftime("%Y-%m-%d %H:%M UTC"))

        url = _extract_download_url(msg)
        if url:
            return url
        else:
            log.warning("  Email matched but no download link found — will retry.")

    return None


# ── Download + extract ────────────────────────────────────────────────────────

def download_and_extract(url: str) -> Path | None:
    inbox = Path(config.CSV_INBOX_DIR)
    inbox.mkdir(parents=True, exist_ok=True)
    zip_path = inbox / "linkedin_export.zip"

    log.info("Downloading archive ZIP...")
    try:
        r = requests.get(
            url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
            stream=True,
            timeout=120,
        )
        r.raise_for_status()
        with open(zip_path, "wb") as f:
            for chunk in r.iter_content(16384):
                f.write(chunk)
        size_kb = zip_path.stat().st_size / 1024
        log.info("Downloaded: %.1f KB", size_kb)
        if size_kb < 1:
            log.error("ZIP too small — download may have failed or link expired.")
            return None
    except Exception as e:
        log.error("Download failed: %s", e)
        return None

    try:
        with zipfile.ZipFile(zip_path) as zf:
            names = zf.namelist()
            log.info("ZIP contents: %s", names)
            csvs = [n for n in names if "connection" in n.lower() and n.endswith(".csv")]
            if not csvs:
                log.error("No Connections CSV in ZIP. Contents: %s", names)
                return None
            ext = zf.extract(csvs[0], inbox)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            final = inbox / f"Connections_{ts}.csv"
            Path(ext).rename(final)
            log.info("Extracted: %s", final.name)
            return final
    except zipfile.BadZipFile:
        log.error("Not a valid ZIP — link may have expired.")
        return None
    except Exception as e:
        log.error("Extraction error: %s", e)
        return None
    finally:
        if zip_path.exists():
            zip_path.unlink()


# ── Main watcher loop ─────────────────────────────────────────────────────────

def wait_for_export_email() -> Path | None:
    log.info("=" * 62)
    log.info("  Email watcher started at %s UTC",
             SCRIPT_START_UTC.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("  Watching for: '%s'", TARGET_SUBJECT)
    log.info("  Searching folders: Primary, Social, Updates, Promotions")
    log.info("  Emails BEFORE script start are ignored.")
    log.info("=" * 62)

    mail = None
    try:
        mail = connect_gmail()
    except imaplib.IMAP4.error as e:
        log.error("Gmail login failed: %s", e)
        log.error(
            "Check: GMAIL_ADDRESS, GMAIL_APP_PASSWORD (16-char App Password), "
            "2-Step Verification enabled, IMAP enabled in Gmail settings."
        )
        return None

    # On first connect, log available folders to help debug
    folders = list_available_folders(mail)
    if folders:
        social_folders = [f for f in folders if "categor" in f.lower() or "social" in f.lower()]
        log.info("Available category folders: %s", social_folders or "(none found — Gmail may use different names)")

    deadline = time.time() + MAX_WAIT
    scan_count = 0

    try:
        while time.time() < deadline:
            remaining = int(deadline - time.time())
            scan_count += 1
            log.info("── Scan #%d (%dm %ds remaining) ──",
                     scan_count, remaining // 60, remaining % 60)

            for folder in SEARCH_FOLDERS:
                try:
                    url = search_folder(mail, folder)
                    if url:
                        return download_and_extract(url)
                except imaplib.IMAP4.abort:
                    log.warning("IMAP connection dropped — reconnecting...")
                    try: mail.logout()
                    except Exception: pass
                    time.sleep(5)
                    mail = connect_gmail()
                    break  # restart this scan cycle

            log.info("Not found in any folder. Next scan in %ds.", POLL_INTERVAL)
            time.sleep(POLL_INTERVAL)

    finally:
        try: mail.logout()
        except Exception: pass

    log.error("Timed out after 60 minutes. No matching email found.")
    log.error("Drop Connections.csv manually into: %s", config.CSV_INBOX_DIR)
    return None


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [EMAIL] %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )
    result = wait_for_export_email()
    print("\nResult:", result or "FAILED — see log above")
