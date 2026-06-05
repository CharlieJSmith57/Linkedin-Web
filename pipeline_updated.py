"""
pipeline.py  —  Main orchestrator for the LinkedIn network pipeline
B:\\linkedin-pipeline\\pipeline.py

RUN ORDER (per execution)
─────────────────────────
1. [NEW] Trigger LinkedIn CSV export request  ← starts the 15-20 min download clock
2. Run the slow enricher  (8-12 profiles, randomized, ~10-15 min)
3. After enricher finishes, check email for the LinkedIn export notification
4. If CSV found → run change_detector → export_to_json → git_push
5. Every 4th enricher run → also run the full daily CSV pipeline
   (regardless of whether a new CSV arrived via email)

SCHEDULING
──────────
Windows Task Scheduler calls:  python pipeline.py
Recommended: every 2 hours, or twice daily.

The "every 4th run" counter persists in run_counter.json in the data dir
so it survives restarts.

FILE LOCATIONS
──────────────
Pipeline code:  B:\\linkedin-pipeline\\
Persistent data: B:\\linkedin-data\\
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import time
import webbrowser
from datetime import datetime, date
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────

PIPELINE_DIR  = Path(__file__).parent
DATA_DIR      = Path(os.environ.get("LINKEDIN_DATA_DIR", r"B:\linkedin-data"))
CSV_INBOX     = DATA_DIR / "csv_inbox"
COUNTER_FILE  = DATA_DIR / "run_counter.json"
LOG_FILE      = DATA_DIR / "logs" / f"pipeline_{date.today()}.log"

# LinkedIn data export URL — opens directly to the export request page
LINKEDIN_EXPORT_URL = (
    "https://www.linkedin.com/mypreferences/d/download-my-data"
)

# ── Logging ───────────────────────────────────────────────────────────────────

LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("pipeline")


# ── Run counter ───────────────────────────────────────────────────────────────

def _load_counter() -> int:
    try:
        return json.loads(COUNTER_FILE.read_text())["enricher_runs"]
    except Exception:
        return 0


def _save_counter(n: int) -> None:
    COUNTER_FILE.write_text(json.dumps({"enricher_runs": n}))


# ── Step helpers ──────────────────────────────────────────────────────────────

def step_open_csv_export() -> None:
    """
    Open the LinkedIn data-export page in the default browser.
    This starts the 15-20 minute download clock.  The user clicks
    'Request archive' once — or has already set up auto-export.
    We open it silently; if already requested today we can skip.
    """
    log.info("STEP 1 — Opening LinkedIn export page to start download clock")
    try:
        webbrowser.open(LINKEDIN_EXPORT_URL)
        log.info("  Browser opened: %s", LINKEDIN_EXPORT_URL)
    except Exception as exc:
        log.warning("  Could not open browser: %s", exc)


def step_run_enricher() -> bool:
    """
    Run ingest_enriched.py (the slow LinkedIn API scraper).
    Returns True if it completed without error.
    """
    log.info("STEP 2 — Running enricher (8-12 profiles)")
    script = PIPELINE_DIR / "ingest_enriched.py"
    if not script.exists():
        log.warning("  ingest_enriched.py not found — skipping")
        return False
    result = subprocess.run(
        [sys.executable, str(script)],
        capture_output=False,   # let it print to stdout/stderr live
    )
    if result.returncode != 0:
        log.error("  Enricher exited with code %d", result.returncode)
        return False
    log.info("  Enricher completed successfully")
    return True


def step_check_email_for_csv() -> Path | None:
    """
    Run email_watcher.py in one-shot mode to check for the LinkedIn export email.
    If found, the watcher downloads the ZIP, extracts Connections.csv,
    and drops it into csv_inbox/.
    Returns the path to the CSV if found, else None.
    """
    log.info("STEP 3 — Checking email for LinkedIn export notification")
    script = PIPELINE_DIR / "email_watcher.py"
    if not script.exists():
        log.warning("  email_watcher.py not found — skipping email check")
        return _find_csv_in_inbox()

    result = subprocess.run(
        [sys.executable, str(script), "--once"],
        capture_output=False,
    )
    if result.returncode != 0:
        log.warning("  email_watcher returned code %d", result.returncode)

    return _find_csv_in_inbox()


def _find_csv_in_inbox() -> Path | None:
    """Return the newest CSV in csv_inbox/, or None."""
    CSV_INBOX.mkdir(parents=True, exist_ok=True)
    csvs = sorted(CSV_INBOX.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)
    if csvs:
        log.info("  Found CSV: %s", csvs[0].name)
        return csvs[0]
    log.info("  No CSV in inbox yet")
    return None


def step_run_change_detector(csv_path: Path) -> bool:
    """Run change_detector.py on the given CSV."""
    log.info("STEP 4a — Running change detector on %s", csv_path.name)
    script = PIPELINE_DIR / "change_detector.py"
    if not script.exists():
        log.warning("  change_detector.py not found")
        return False
    result = subprocess.run(
        [sys.executable, str(script), str(csv_path)],
        capture_output=False,
    )
    if result.returncode != 0:
        log.error("  change_detector exited %d", result.returncode)
        return False
    return True


def step_export_to_json() -> bool:
    """Run export_to_json.py to rebuild network.json."""
    log.info("STEP 4b — Exporting to network.json")
    script = PIPELINE_DIR / "export_to_json.py"
    if not script.exists():
        log.warning("  export_to_json.py not found")
        return False
    result = subprocess.run([sys.executable, str(script)], capture_output=False)
    if result.returncode != 0:
        log.error("  export_to_json exited %d", result.returncode)
        return False
    return True


def step_git_push() -> bool:
    """Run git_push.py to commit and push to GitHub Pages."""
    log.info("STEP 4c — Git push")
    script = PIPELINE_DIR / "git_push.py"
    if not script.exists():
        log.warning("  git_push.py not found")
        return False
    result = subprocess.run([sys.executable, str(script)], capture_output=False)
    if result.returncode != 0:
        log.error("  git_push exited %d", result.returncode)
        return False
    return True


def step_run_full_daily_pipeline() -> None:
    """
    Every 4th enricher run: run the complete daily CSV pipeline
    (change_detector → export_to_json → git_push) regardless of email.
    This ensures the front-end stays current even if no new CSV arrived.
    """
    log.info("STEP 5 — 4th enricher run: running full daily pipeline")
    csv = _find_csv_in_inbox()
    if csv:
        if step_run_change_detector(csv):
            step_export_to_json()
            step_git_push()
    else:
        # No CSV but still push to pick up any enricher-only changes
        step_export_to_json()
        step_git_push()


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    log.info("=" * 60)
    log.info("LinkedIn Pipeline starting — %s", datetime.now().isoformat())

    # ── Step 1: Open export page FIRST to start the download clock ────────────
    step_open_csv_export()

    # Brief pause so browser can load before enricher hammers the CPU
    time.sleep(3)

    # ── Step 2: Run enricher (takes ~10-15 min) ───────────────────────────────
    enricher_ok = step_run_enricher()

    # ── Update run counter ─────────────────────────────────────────────────────
    run_count = _load_counter() + 1
    _save_counter(run_count)
    log.info("Enricher run count: %d", run_count)

    # ── Step 3: Check email for CSV (export should be ready by now) ───────────
    csv_path = step_check_email_for_csv()

    # ── Step 4: If CSV arrived, run the change detection pipeline ─────────────
    if csv_path:
        ok = step_run_change_detector(csv_path)
        if ok:
            step_export_to_json()
            step_git_push()
    else:
        log.info("No new CSV — skipping change detection this run")

    # ── Step 5: Every 4th enricher run → also run the full daily pipeline ─────
    if run_count % 4 == 0:
        log.info("4th run milestone — triggering full daily pipeline")
        step_run_full_daily_pipeline()
    else:
        log.info(
            "Next full daily pipeline in %d enricher run(s)",
            4 - (run_count % 4),
        )

    log.info("Pipeline complete — %s", datetime.now().isoformat())
    log.info("=" * 60)


if __name__ == "__main__":
    main()
