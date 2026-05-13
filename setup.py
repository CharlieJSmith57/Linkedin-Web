"""
setup.py — First-time setup and migration tool

Run this ONCE after installing or updating the project:
    python setup.py

What it does:
  1. Creates B:\linkedin-data\ and all required subdirectories
  2. Migrates any existing data from the old B:\linkedin-network\data\ location
  3. Copies starter JSON config files if they don't already exist in the data root
  4. Verifies config.py credentials are filled in
  5. Prints a clear summary of what's where

Safe to re-run at any time — never overwrites existing data files.
"""

import shutil
import sys
from pathlib import Path

# ── Load config ───────────────────────────────────────────────────────────────
try:
    import config
except Exception as e:
    print(f"ERROR: Could not load config.py: {e}")
    print("Make sure config.py exists in the same folder as setup.py.")
    sys.exit(1)

W = 56  # print width

def header(msg):
    print("=" * W)
    print(f"  {msg}")
    print("=" * W)

def status(tag, msg):
    print(f"  [{tag:<8}]  {msg}")

def warn(msg):
    print(f"  [WARNING]  {msg}")

def ok(msg):
    print(f"  [OK]       {msg}")


# ── Starter JSON files that should live in DATA_ROOT ─────────────────────────
# These are copied from the project folder if they don't already exist.
# The project folder copies are kept as templates only.

STARTER_FILES = {
    "industry_map.json":    config.REPO_ROOT / "data" / "industry_map.json",
    "overrides.json":       config.REPO_ROOT / "data" / "overrides.json",
    "pending_changes.json": config.REPO_ROOT / "data" / "pending_changes.json",
    "custom_taxonomy.json": config.REPO_ROOT / "data" / "custom_taxonomy.json",
}


def run():
    header("LinkedIn Network — Setup")
    print()
    print(f"  Project code : {config.REPO_ROOT}")
    print(f"  Data root    : {config.DATA_ROOT}")
    print()

    # ── 1. Create data directories ────────────────────────────────────────────
    header("Creating data directories")
    dirs = [
        config.DATA_ROOT,
        config.SNAPSHOT_DIR,
        config.CSV_INBOX_DIR,
        config.LOG_DIR,
    ]
    for d in dirs:
        d.mkdir(parents=True, exist_ok=True)
        tag = "EXISTS" if d.exists() else "CREATED"
        status(tag, str(d))
    print()

    # ── 2. Migrate existing data from old location ───────────────────────────
    old_data = config.REPO_ROOT / "data"
    migrated = 0

    if old_data.exists() and old_data != config.DATA_ROOT:
        header("Checking for data to migrate")

        # Migrate Excel workbook
        old_xlsx = old_data / "network_master.xlsx"
        new_xlsx = config.EXCEL_PATH
        if old_xlsx.exists() and not new_xlsx.exists():
            shutil.copy2(old_xlsx, new_xlsx)
            status("MIGRATED", f"network_master.xlsx → {new_xlsx}")
            migrated += 1
        elif old_xlsx.exists() and new_xlsx.exists():
            status("SKIPPED", f"network_master.xlsx already exists at destination")

        # Migrate snapshots
        old_snaps = old_data / "snapshots"
        if old_snaps.exists():
            for f in old_snaps.glob("*.json"):
                dest = config.SNAPSHOT_DIR / f.name
                if not dest.exists():
                    shutil.copy2(f, dest)
                    migrated += 1
            if list(old_snaps.glob("*.json")):
                status("MIGRATED", f"snapshots → {config.SNAPSHOT_DIR}")

        if migrated == 0:
            ok("No old data to migrate.")
        else:
            ok(f"{migrated} file(s) migrated to {config.DATA_ROOT}")
        print()

    # ── 3. Copy starter JSON configs if missing ───────────────────────────────
    header("Config files")
    for filename, src in STARTER_FILES.items():
        dest = config.DATA_ROOT / filename
        if dest.exists():
            status("EXISTS", f"{filename}")
        elif src.exists():
            shutil.copy2(src, dest)
            status("CREATED", f"{filename} (copied from project template)")
        else:
            warn(f"{filename} template not found in project — skipping")
    print()

    # ── 4. Verify credentials ─────────────────────────────────────────────────
    header("Credential check")
    creds_ok = True
    checks = [
        ("LINKEDIN_EMAIL",    config.LINKEDIN_EMAIL,    "your_linkedin_email@gmail.com"),
        ("LINKEDIN_PASSWORD", config.LINKEDIN_PASSWORD, "your_linkedin_password"),
        ("GMAIL_ADDRESS",     config.GMAIL_ADDRESS,     "your_gmail@gmail.com"),
        ("GMAIL_APP_PASSWORD",config.GMAIL_APP_PASSWORD,"xxxx xxxx xxxx xxxx"),
    ]
    for name, val, placeholder in checks:
        if val == placeholder or not val:
            warn(f"{name} is not set in config.py")
            creds_ok = False
        else:
            ok(f"{name} is set")
    print()

    # ── 5. Summary ────────────────────────────────────────────────────────────
    header("Setup complete")
    print()
    if not creds_ok:
        print("  ⚠ Fill in your credentials in config.py before running pipeline.py")
    else:
        print("  ✓ All credentials set")
    print()
    print("  Your data lives at:")
    print(f"    {config.DATA_ROOT}")
    print()
    print("  This folder is SEPARATE from the project and will NOT be")
    print("  affected when you extract future project updates.")
    print()
    print("  Next steps:")
    print("    1. python pipeline.py     — run the daily pipeline")
    print("    2. python scheduler.py install  — set up auto-scheduling")
    print("=" * W)


if __name__ == "__main__":
    run()
