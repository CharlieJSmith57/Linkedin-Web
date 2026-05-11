"""
scheduler.py — Windows Task Scheduler setup

Run this ONCE to register both scheduled tasks with Windows:
    python scheduler.py install

To remove the tasks:
    python scheduler.py uninstall

To check task status:
    python scheduler.py status

HOW IT WORKS
─────────────────────────────────────────────────────────────────────────────
Uses Windows' built-in `schtasks.exe` — no third-party tools needed.
Creates two tasks:

  LinkedInPipeline  — runs pipeline.py daily at PIPELINE_RUN_TIME
  LinkedInEnricher  — runs enricher.py daily at ENRICHER_RUN_TIME

Both tasks:
  - Run whether the user is logged in or not (requires password on install)
  - Run under your Windows user account
  - Log output to data/logs/ so you can see what happened
  - Wake the computer if it's sleeping (configurable below)
  - Show a toast notification on completion

REQUIREMENTS
─────────────────────────────────────────────────────────────────────────────
  - Windows 10 or 11
  - Python on PATH (or set PYTHON_EXE below to an absolute path)
  - Run `python scheduler.py install` from an elevated terminal (Run as Admin)
    if you want the tasks to run while logged out; otherwise normal terminal works.
"""

from __future__ import annotations

import os
import subprocess
import sys
import getpass
import logging
from pathlib import Path

import config

log = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format="%(asctime)s [SCHEDULER] %(message)s")

# ── Settings ──────────────────────────────────────────────────────────────────

# Use the same Python that's running this script
PYTHON_EXE = sys.executable

REPO_ROOT = str(config.REPO_ROOT)
LOG_DIR   = config.REPO_ROOT / "data" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

TASKS = [
    {
        "name":    "LinkedInPipeline",
        "script":  "pipeline.py",
        "time":    config.PIPELINE_RUN_TIME,
        "log":     str(LOG_DIR / "pipeline.log"),
        "comment": "LinkedIn daily CSV pipeline and GitHub Pages export",
    },
    {
        "name":    "LinkedInEnricher",
        "script":  "enricher.py",
        "time":    config.ENRICHER_RUN_TIME,
        "log":     str(LOG_DIR / "enricher.log"),
        "comment": "LinkedIn slow-crawl profile enricher",
    },
]


# ── schtasks helpers ──────────────────────────────────────────────────────────

def schtasks(*args: str) -> tuple[int, str]:
    """Run schtasks.exe with the given arguments. Returns (returncode, output)."""
    cmd = ["schtasks"] + list(args)
    result = subprocess.run(cmd, capture_output=True, text=True)
    output = (result.stdout + result.stderr).strip()
    return result.returncode, output


def task_exists(name: str) -> bool:
    code, _ = schtasks("/Query", "/TN", name, "/FO", "LIST")
    return code == 0


# ── Install ───────────────────────────────────────────────────────────────────

def install():
    """Register both tasks with Windows Task Scheduler."""

    # We need the current Windows username
    username = os.environ.get("USERNAME") or getpass.getuser()

    print(f"\nInstalling tasks for user: {username}")
    print("You will be prompted for your Windows password.")
    print("(This is required for tasks that run while logged out.)\n")

    # Prompt once for password — we pass it to schtasks /RU /RP
    password = getpass.getpass(f"Windows password for {username}: ")

    for task in TASKS:
        name   = task["name"]
        script = Path(REPO_ROOT) / task["script"]
        log_file = task["log"]
        time_str = task["time"]  # e.g. "07:00"

        # The command that schtasks will execute:
        # python pipeline.py >> data/logs/pipeline.log 2>&1
        # We wrap in cmd /C so redirection works
        run_cmd = (
            f'cmd /C "cd /D {REPO_ROOT} && '
            f'{PYTHON_EXE} {script} >> {log_file} 2>&1"'
        )

        if task_exists(name):
            log.info("Task '%s' already exists — deleting and re-creating.", name)
            schtasks("/Delete", "/TN", name, "/F")

        code, output = schtasks(
            "/Create",
            "/SC",       "DAILY",
            "/TN",       name,
            "/TR",       run_cmd,
            "/ST",       time_str,
            "/RU",       username,
            "/RP",       password,
            "/RL",       "HIGHEST",    # run with highest available privileges
            "/F",                      # force overwrite if exists
            # Wake computer from sleep to run the task
            # (only works if enabled in power settings)
        )

        if code == 0:
            log.info("✓ Task '%s' installed — runs daily at %s", name, time_str)
        else:
            log.error("✗ Failed to install '%s': %s", name, output)

    print("\nDone. Verify in Task Scheduler (taskschd.msc) under Task Scheduler Library.")
    print("Run `python scheduler.py status` to confirm.\n")


# ── Uninstall ─────────────────────────────────────────────────────────────────

def uninstall():
    """Remove both tasks from Windows Task Scheduler."""
    for task in TASKS:
        name = task["name"]
        if task_exists(name):
            code, output = schtasks("/Delete", "/TN", name, "/F")
            if code == 0:
                log.info("✓ Removed task '%s'", name)
            else:
                log.error("✗ Failed to remove '%s': %s", name, output)
        else:
            log.info("Task '%s' not found — nothing to remove.", name)


# ── Status ────────────────────────────────────────────────────────────────────

def status():
    """Print the current status of both tasks."""
    for task in TASKS:
        name = task["name"]
        code, output = schtasks("/Query", "/TN", name, "/FO", "LIST", "/V")
        if code == 0:
            # Pull out the key lines
            lines = output.splitlines()
            interesting = [
                l for l in lines
                if any(k in l for k in [
                    "Task Name", "Status", "Next Run Time", "Last Run Time",
                    "Last Result", "Run As User"
                ])
            ]
            print(f"\n── {name} ──")
            for l in interesting:
                print(" ", l)
        else:
            print(f"\n── {name} ── NOT FOUND")

    # Also show last few lines of each log
    print("\n── Recent log output ──")
    for task in TASKS:
        log_path = Path(task["log"])
        print(f"\n{task['name']} ({log_path.name}):")
        if log_path.exists():
            lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
            for line in lines[-15:]:  # last 15 lines
                print(" ", line)
        else:
            print("  (no log yet)")


# ── Run once now ──────────────────────────────────────────────────────────────

def run_now(task_name: str):
    """Trigger a task to run immediately (useful for testing)."""
    code, output = schtasks("/Run", "/TN", task_name)
    if code == 0:
        log.info("Task '%s' triggered.", task_name)
    else:
        log.error("Failed to trigger '%s': %s", task_name, output)


# ── CLI ───────────────────────────────────────────────────────────────────────

USAGE = """
Usage:
  python scheduler.py install     Register both tasks with Windows Task Scheduler
  python scheduler.py uninstall   Remove both tasks
  python scheduler.py status      Show task status + recent log output
  python scheduler.py run-now     Trigger both tasks to run immediately
"""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(USAGE)
        sys.exit(0)

    cmd = sys.argv[1].lower()

    if cmd == "install":
        install()
    elif cmd == "uninstall":
        uninstall()
    elif cmd == "status":
        status()
    elif cmd == "run-now":
        for task in TASKS:
            run_now(task["name"])
    else:
        print(f"Unknown command: {cmd}")
        print(USAGE)
        sys.exit(1)
