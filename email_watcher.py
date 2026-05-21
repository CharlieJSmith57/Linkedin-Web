"""
email_watcher.py — Watches Gmail for LinkedIn first-installment export email.
Subject: "The first installment of your LinkedIn data archive is ready!"
Only accepts emails received AFTER this script started.
"""
from __future__ import annotations
import email, email.header, email.utils, imaplib, logging, re, time, zipfile
from datetime import datetime, timezone, timedelta
from html.parser import HTMLParser
from pathlib import Path
import requests, config

log = logging.getLogger(__name__)
GMAIL_IMAP_HOST  = "imap.gmail.com"
POLL_INTERVAL    = 60
MAX_WAIT         = 3600
LINKEDIN_SENDER  = "messages-noreply@linkedin.com"
TARGET_SUBJECT   = "first installment of your linkedin data archive is ready"
SCRIPT_START_UTC = datetime.now(timezone.utc) - timedelta(minutes=2)

class LinkExtractor(HTMLParser):
    def __init__(self): super().__init__(); self.links = []
    def handle_starttag(self, tag, attrs):
        if tag == "a":
            for n, v in attrs:
                if n == "href" and v: self.links.append(v)

def extract_links(html):
    p = LinkExtractor(); p.feed(html); return p.links

def connect_gmail():
    mail = imaplib.IMAP4_SSL(GMAIL_IMAP_HOST, 993)
    mail.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
    mail.select("inbox"); return mail

def decode_subject(raw):
    parts = email.header.decode_header(raw)
    return "".join(p.decode(e or "utf-8") if isinstance(p, bytes) else p for p, e in parts)

def parse_email_date(s):
    try:
        d = email.utils.parsedate_to_datetime(s)
        return d if d.tzinfo else d.replace(tzinfo=timezone.utc)
    except: return None

def is_linkedin_dl_link(url):
    u = url.lower()
    return "linkedin.com" in u and any(k in u for k in ("ambry","checkpoint","download","archive"))

def find_email(mail):
    since = SCRIPT_START_UTC.strftime("%d-%b-%Y")
    _, data = mail.search(None, f'(FROM "{LINKEDIN_SENDER}" SINCE "{since}")')
    for mid in reversed(data[0].split()):
        _, md = mail.fetch(mid, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        subj = decode_subject(msg.get("Subject",""))
        rcvd = parse_email_date(msg.get("Date",""))
        if not rcvd or rcvd < SCRIPT_START_UTC: continue
        if TARGET_SUBJECT not in subj.lower(): continue
        log.info("Found: '%s'", subj)
        for part in msg.walk():
            ct = part.get_content_type()
            try: body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
            except: continue
            links = extract_links(body) if ct=="text/html" else re.findall(r'https?://[^\s<>"]+', body)
            for link in links:
                if is_linkedin_dl_link(link): return link
    return None

def download_and_extract(url):
    inbox = Path(config.CSV_INBOX_DIR)
    zip_path = inbox / "linkedin_export.zip"
    try:
        r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, stream=True, timeout=120)
        r.raise_for_status()
        with open(zip_path,"wb") as f:
            for chunk in r.iter_content(16384): f.write(chunk)
        if zip_path.stat().st_size < 1024: return None
        with zipfile.ZipFile(zip_path) as zf:
            csvs = [n for n in zf.namelist() if "connection" in n.lower() and n.endswith(".csv")]
            if not csvs: return None
            ext = zf.extract(csvs[0], inbox)
            final = inbox / f"Connections_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            Path(ext).rename(final); return final
    except Exception as e: log.error("Download failed: %s", e); return None
    finally:
        if zip_path.exists(): zip_path.unlink()

def wait_for_export_email():
    log.info("Email watcher started — waiting for LinkedIn export email")
    deadline = time.time() + MAX_WAIT
    try: mail = connect_gmail()
    except imaplib.IMAP4.error as e: log.error("Gmail login failed: %s", e); return None
    try:
        while time.time() < deadline:
            try:
                mail.select("inbox")
                url = find_email(mail)
            except imaplib.IMAP4.abort:
                try: mail.logout()
                except: pass
                time.sleep(5); mail = connect_gmail(); continue
            if url:
                csv = download_and_extract(url)
                return csv
            log.info("Not found — checking again in %ds", POLL_INTERVAL)
            time.sleep(POLL_INTERVAL)
    finally:
        try: mail.logout()
        except: pass
    return None

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [EMAIL] %(message)s")
    print(wait_for_export_email() or "FAILED")
