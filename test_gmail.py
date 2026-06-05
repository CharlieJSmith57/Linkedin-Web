"""
test_gmail.py — Gmail connection verifier

Run this before triggering a LinkedIn export to confirm
the email watcher can access your inbox.

    python test_gmail.py

What it checks:
  1. Can connect to Gmail IMAP with your credentials
  2. Can select the inbox
  3. Shows how many total messages are in the inbox
  4. Searches for any past LinkedIn export emails (as a sanity check)
  5. Confirms the exact subject line it will watch for
"""

import email
import email.header
import email.utils
import imaplib
import sys

try:
    import config
except Exception as e:
    print(f"ERROR: Could not load config.py — {e}")
    sys.exit(1)

LINKEDIN_SENDER = "messages-noreply@linkedin.com"
TARGET_SUBJECT  = "first installment of your linkedin data archive is ready"

W = 60

def hdr(msg):
    print("\n" + "─" * W)
    print(f"  {msg}")
    print("─" * W)

def ok(msg):   print(f"  [  OK  ]  {msg}")
def fail(msg): print(f"  [ FAIL ]  {msg}")
def info(msg): print(f"  [ INFO ]  {msg}")

def decode_subject(raw):
    parts = email.header.decode_header(raw)
    return "".join(
        p.decode(e or "utf-8", errors="ignore") if isinstance(p, bytes) else p
        for p, e in parts
    )

def run():
    print("=" * W)
    print("  Gmail connection test")
    print("=" * W)

    # ── Step 1: Credentials present ──────────────────────────────────────────
    hdr("1. Checking config.py credentials")
    if not config.GMAIL_ADDRESS or config.GMAIL_ADDRESS == "your_gmail@gmail.com":
        fail("GMAIL_ADDRESS is not set in config.py"); sys.exit(1)
    if not config.GMAIL_APP_PASSWORD or config.GMAIL_APP_PASSWORD == "xxxx xxxx xxxx xxxx":
        fail("GMAIL_APP_PASSWORD is not set in config.py"); sys.exit(1)
    ok(f"GMAIL_ADDRESS    = {config.GMAIL_ADDRESS}")
    ok(f"GMAIL_APP_PASSWORD = {'*' * 16}")

    # ── Step 2: Connect ───────────────────────────────────────────────────────
    hdr("2. Connecting to Gmail IMAP")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", 993)
        ok("TCP connection established")
    except Exception as e:
        fail(f"Could not reach imap.gmail.com: {e}")
        sys.exit(1)

    try:
        mail.login(config.GMAIL_ADDRESS, config.GMAIL_APP_PASSWORD)
        ok("Login successful")
    except imaplib.IMAP4.error as e:
        fail(f"Login failed: {e}")
        print()
        print("  Troubleshooting:")
        print("  • GMAIL_APP_PASSWORD must be the 16-char App Password,")
        print("    NOT your Google account password.")
        print("  • Generate one at: myaccount.google.com/apppasswords")
        print("  • 2-Step Verification must be enabled.")
        print("  • IMAP must be on: Gmail → Settings → See all settings")
        print("    → Forwarding and POP/IMAP → Enable IMAP")
        sys.exit(1)

    # ── Step 3: Select inbox ──────────────────────────────────────────────────
    hdr("3. Selecting inbox")
    try:
        status, data = mail.select("INBOX")
        if status != "OK":
            fail(f"Could not select INBOX: {data}")
            sys.exit(1)
        msg_count = int(data[0]) if data[0] else 0
        ok(f"Inbox selected — {msg_count:,} messages total")
    except Exception as e:
        fail(f"Error selecting inbox: {e}")
        sys.exit(1)

    # ── Step 4: Search for past LinkedIn export emails ────────────────────────
    hdr("4. Searching for LinkedIn export emails (any date)")
    try:
        _, data = mail.search(None, f'(FROM "{LINKEDIN_SENDER}")')
        msg_ids = data[0].split() if data[0] else []
        info(f"Found {len(msg_ids)} email(s) from {LINKEDIN_SENDER}")

        if msg_ids:
            # Show the 3 most recent
            for mid in reversed(msg_ids[-3:]):
                _, md = mail.fetch(mid, "(RFC822.HEADER)")
                msg = email.message_from_bytes(md[0][1])
                subj = decode_subject(msg.get("Subject", "(no subject)"))
                date = msg.get("Date", "(no date)")
                matched = "✓" if TARGET_SUBJECT in subj.lower() else " "
                print(f"  [{matched}] {date[:22]}  {subj[:55]}")
            if len(msg_ids) > 3:
                print(f"      ... and {len(msg_ids)-3} older message(s)")
        else:
            info("No past LinkedIn emails found — that's fine if you haven't exported yet.")

    except Exception as e:
        fail(f"Search error: {e}")

    # ── Step 5: Summary ───────────────────────────────────────────────────────
    hdr("5. What the watcher will look for")
    info(f"From:    {LINKEDIN_SENDER}")
    info(f"Subject: (contains) '{TARGET_SUBJECT}'")
    info("Folder:  INBOX (Primary)")
    info(f"Timing:  only emails received AFTER you run pipeline.py")
    info(f"Poll:    every 60 seconds for up to 60 minutes")

    mail.logout()

    print()
    print("=" * W)
    print("  All checks passed. Gmail access is working.")
    print("  You can now run pipeline.py and request a LinkedIn export.")
    print("=" * W)
    print()

if __name__ == "__main__":
    run()
