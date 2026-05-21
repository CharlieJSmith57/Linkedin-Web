"""setup.py — First-time setup and migration. Run once after installing or updating."""
import shutil, sys
from pathlib import Path
try:
    import config
except Exception as e:
    print(f"ERROR: {e}"); sys.exit(1)

W = 56
def hdr(m): print("="*W); print(f"  {m}"); print("="*W)
def st(t,m): print(f"  [{t:<8}]  {m}")

def run():
    hdr("LinkedIn Network — Setup")
    print(f"\n  Project code : {config.REPO_ROOT}")
    print(f"  Data root    : {config.DATA_ROOT}\n")

    hdr("Creating data directories")
    for d in [config.DATA_ROOT, config.SNAPSHOT_DIR, config.CSV_INBOX_DIR, config.LOG_DIR]:
        d.mkdir(parents=True, exist_ok=True)
        st("EXISTS" if d.exists() else "CREATED", str(d))
    print()

    hdr("Checking for data to migrate")
    old_xlsx = config.REPO_ROOT / "data" / "network_master.xlsx"
    if old_xlsx.exists() and not config.EXCEL_PATH.exists():
        shutil.copy2(old_xlsx, config.EXCEL_PATH)
        st("MIGRATED", "network_master.xlsx")
    elif config.EXCEL_PATH.exists():
        st("SKIPPED", "network_master.xlsx already exists at destination")
    else:
        st("OK", "No old data to migrate.")
    print()

    hdr("Config files")
    starters = ["industry_map.json","overrides.json","pending_changes.json","custom_taxonomy.json"]
    for fn in starters:
        src  = config.REPO_ROOT / "data" / fn
        dest = config.DATA_ROOT / fn
        if dest.exists(): st("EXISTS", fn)
        elif src.exists(): shutil.copy2(src, dest); st("CREATED", fn)
        else: print(f"  [WARNING]  {fn} template not found")
    print()

    hdr("Credential check")
    ok = True
    for name, val, ph in [
        ("LINKEDIN_EMAIL",    config.LINKEDIN_EMAIL,    "your_linkedin_email@gmail.com"),
        ("LINKEDIN_PASSWORD", config.LINKEDIN_PASSWORD, "your_linkedin_password"),
        ("GMAIL_ADDRESS",     config.GMAIL_ADDRESS,     "your_gmail@gmail.com"),
        ("GMAIL_APP_PASSWORD",config.GMAIL_APP_PASSWORD,"xxxx xxxx xxxx xxxx"),
    ]:
        if val == ph or not val: print(f"  [WARNING]  {name} not set"); ok = False
        else: st("OK", f"{name} is set")
    print()

    hdr("Setup complete")
    print()
    if not ok: print("  Set your credentials in config.py before running pipeline.py")
    else: print("  All credentials set")
    print(f"\n  Your data lives at:\n    {config.DATA_ROOT}")
    print("\n  Next steps:")
    print("    1. python pipeline.py")
    print("    2. python scheduler.py install")
    print("="*W)

if __name__ == "__main__": run()
