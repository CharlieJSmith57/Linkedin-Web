"""git_push.py — Auto-commit and push network.json to GitHub."""
from __future__ import annotations
import logging, subprocess, sys
from datetime import date
from pathlib import Path
import config

log = logging.getLogger(__name__)
JSON_REL_PATH = Path("docs") / "data" / "network.json"

def run(cmd, cwd):
    r = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True)
    return r.returncode, r.stdout.strip(), r.stderr.strip()

def push():
    repo = config.REPO_ROOT
    json_path = repo / JSON_REL_PATH
    if not json_path.exists():
        log.error("network.json not found — run export_to_json.py first.")
        sys.exit(1)
    code, stdout, _ = run(["git", "status", "--porcelain", str(JSON_REL_PATH)], repo)
    if not stdout:
        log.info("network.json unchanged since last push — nothing to commit.")
        return
    run(["git", "add", str(JSON_REL_PATH)], repo)
    message = config.GIT_COMMIT_MESSAGE.format(date=date.today().isoformat())
    code, _, err = run(["git", "commit", "-m", message], repo)
    if code != 0:
        log.error("git commit failed: %s", err); sys.exit(1)
    log.info("Committed: %s", message)
    code, _, err = run(["git", "push", config.GIT_REMOTE, config.GIT_BRANCH], repo)
    if code != 0:
        log.error("git push failed: %s", err); sys.exit(1)
    log.info("Pushed to %s/%s", config.GIT_REMOTE, config.GIT_BRANCH)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [GIT] %(levelname)s %(message)s")
    push()
