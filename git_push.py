"""
git_push.py — Auto-commit and push the updated network.json to GitHub.

Called automatically by the pipeline after export_to_json.py runs.
Can also be run manually:

    python git_push.py

Requires git to be installed and the repo to have a remote already configured
(i.e., you've already done `git remote add origin https://github.com/you/repo.git`).
"""

from __future__ import annotations

import logging
import subprocess
import sys
from datetime import date
from pathlib import Path

import config

log = logging.getLogger(__name__)

# The file we're committing — relative to repo root so git sees it cleanly
JSON_REL_PATH = Path("docs") / "data" / "network.json"


def run(cmd: list[str], cwd: Path) -> tuple[int, str, str]:
    """Run a shell command, return (returncode, stdout, stderr)."""
    result = subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
    )
    return result.returncode, result.stdout.strip(), result.stderr.strip()


def push():
    repo = config.REPO_ROOT
    json_path = repo / JSON_REL_PATH

    if not json_path.exists():
        log.error("network.json not found at %s — run export_to_json.py first.", json_path)
        sys.exit(1)

    # Check if there are actually changes to commit
    code, stdout, _ = run(["git", "status", "--porcelain", str(JSON_REL_PATH)], repo)
    if not stdout:
        log.info("network.json unchanged since last push — nothing to commit.")
        return

    # Stage the file
    code, _, err = run(["git", "add", str(JSON_REL_PATH)], repo)
    if code != 0:
        log.error("git add failed: %s", err)
        sys.exit(1)

    # Commit
    message = config.GIT_COMMIT_MESSAGE.format(date=date.today().isoformat())
    code, _, err = run(["git", "commit", "-m", message], repo)
    if code != 0:
        log.error("git commit failed: %s", err)
        sys.exit(1)
    log.info("Committed: %s", message)

    # Push
    code, _, err = run(
        ["git", "push", config.GIT_REMOTE, config.GIT_BRANCH],
        repo,
    )
    if code != 0:
        log.error("git push failed: %s", err)
        log.error("You may need to pull first: git pull --rebase origin main")
        sys.exit(1)

    log.info("Pushed to %s/%s — GitHub Pages will update in ~30 seconds.",
             config.GIT_REMOTE, config.GIT_BRANCH)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s [GIT] %(levelname)s %(message)s")
    push()
