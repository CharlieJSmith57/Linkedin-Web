"""
pipeline.py — Daily pipeline orchestrator.

Run manually:         python pipeline.py
Run on a schedule:    python scheduler.py   (uses Windows Task Scheduler)

Full daily cycle:
  1. Open LinkedIn data export page in browser   (you manually download CSV)
  2. Watch csv_inbox/ for the new CSV file
  3. Parse CSV → detect changes vs. stored data
  4. Write changes to Excel workbook
  5. Export Excel → docs/data/network.json
  6. git commit + push → GitHub Pages updates

The slow enricher (enricher.py) runs on its own separate schedule
and writes back into the same Excel file.
"""

from __future__ import annotations

import logging
import subprocess
import sys
import time
import webbrowser
from pathlib import Path
from datetime import date, datetime

import config

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [PIPELINE] %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# LinkedIn data export URL — lands you directly on the right page
LINKEDIN_EXPORT_URL = (
    "https://www.linkedin.com/mypreferences/d/download-my-data"
)

# How long to wait for the CSV to appear in the inbox folder (seconds)
CSV_WAIT_TIMEOUT = 3600  # 1 hour max — you have time to download and drop it
CSV_POLL_INTERVAL = 10   # check every 10 seconds


def open_linkedin_export_page():
    """Open LinkedIn's data export page so it's right in front of you."""
    log.info("Opening LinkedIn data export page...")
    webbrowser.open(LINKEDIN_EXPORT_URL)
    log.info(
        "ACTION REQUIRED:\n"
        "  1. On LinkedIn: select 'Download larger data archive'\n"
        "  2. Click 'Request archive'\n"
        "  3. Wait for the email (~10 min), download the ZIP\n"
        "  4. Open the ZIP, copy Connections.csv into: %s\n"
        "  Pipeline will detect the file and continue automatically.",
        config.CSV_INBOX_DIR,
    )


def wait_for_csv() -> Path | None:
    """
    Poll the csv_inbox directory for a new CSV file.
    Returns the path to the file when found, or None on timeout.
    Ignores files older than today so stale files don't trigger a re-run.
    """
    inbox = Path(config.CSV_INBOX_DIR)
    deadline = time.time() + CSV_WAIT_TIMEOUT
    seen_before = {f for f in inbox.glob("*.csv")}

    log.info("Watching %s for new CSV... (timeout: 1 hour)", inbox)

    while time.time() < deadline:
        current = set(inbox.glob("*.csv"))
        new_files = current - seen_before
        if new_files:
            csv_path = sorted(new_files)[0]  # take the first if multiple dropped
            log.info("Detected new CSV: %s", csv_path.name)
            return csv_path
        time.sleep(CSV_POLL_INTERVAL)

    log.error("Timed out waiting for CSV. Re-run pipeline.py after dropping the file.")
    return None


def run_module(script_name: str, label: str):
    """Run a pipeline module as a subprocess and stream its output."""
    log.info("── Starting: %s ──", label)
    result = subprocess.run(
        [sys.executable, script_name],
        cwd=str(config.REPO_ROOT),
    )
    if result.returncode != 0:
        log.error("%s exited with code %d", label, result.returncode)
        raise RuntimeError(f"{label} failed")
    log.info("── Done: %s ──", label)


def run():
    log.info("════════════════════════════════════════")
    log.info("  LinkedIn Pipeline — %s", date.today().isoformat())
    log.info("════════════════════════════════════════")

    # Step 1: Open LinkedIn export page
    open_linkedin_export_page()

    # Step 2: Watch Gmail for the export-ready email and auto-download
    log.info("Starting email watcher — go to LinkedIn and click 'Request archive' now.")
    from email_watcher import wait_for_export_email
    csv_path = wait_for_export_email()

    # Fallback: if email watcher fails, wait for manual CSV drop
    if not csv_path:
        log.warning("Email watcher did not find the CSV. Falling back to manual mode.")
        log.warning("Drop Connections.csv into: %s", config.CSV_INBOX_DIR)
        csv_path = wait_for_csv()

    if not csv_path:
        log.error("No CSV found after 1 hour. Exiting.")
        sys.exit(1)

    # Small pause to ensure the file is fully written before reading
    time.sleep(2)

    # Step 3 + 4: Parse CSV and update Excel
    import os
    os.environ["PIPELINE_CSV_PATH"] = str(csv_path)

    run_module("change_detector.py", "Change detector")

    # Step 5: Export Excel → network.json
    run_module("export_to_json.py", "JSON exporter")

    # Step 6: Commit + push
    run_module("git_push.py", "Git push")

    # Archive the processed CSV so it's not picked up again
    archive_dir = Path(config.CSV_INBOX_DIR) / "processed"
    archive_dir.mkdir(exist_ok=True)
    archived = archive_dir / f"{date.today().isoformat()}_{csv_path.name}"
    csv_path.rename(archived)
    log.info("CSV archived to: %s", archived)

    log.info("════════════════════════════════════════")
    log.info("  Pipeline complete. GitHub Pages will update in ~30s.")
    log.info("════════════════════════════════════════")


if __name__ == "__main__":
    run()
