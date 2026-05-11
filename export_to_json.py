"""
export_to_json.py — Static data exporter
Run this after each pipeline/enricher cycle:

    python export_to_json.py

Reads the master Excel workbook, applies the full title taxonomy to every
connection, and writes docs/data/network.json — the single file the
GitHub Pages front-end fetches on load.

After running, commit and push:
    git add docs/data/network.json
    git commit -m "chore: refresh network data $(date +%Y-%m-%d)"
    git push

Or use git_push.py to do that automatically.

OUTPUT SHAPE
────────────
{
  "generated":   "2026-05-01T14:32:00",
  "count":       312,
  "connections": [
    {
      "id":                   "ACoAAB...",
      "first":                "Sarah",
      "last":                 "Chen",
      "headline":             "Senior Protection Engineer @ Pacific Gas & Electric",
      "location":             "San Francisco, CA",
      "photo_url":            "https://...",
      "connected_on":         "2024-09-15",
      "current_company":      "Pacific Gas & Electric",
      "current_title":        "Senior Protection & Relay Engineer",
      "current_start":        "2024-03-01",
      "discipline":           "Electrical — Protection & Relay",
      "discipline_family":    "Electrical",
      "discipline_specialty": "Protection & Relay",
      "seniority":            "Senior",
      "school":               "MIT",
      "industry":             "Energy / Utilities",
      "skills":               ["Power Systems", "Relay Protection", "PSCAD"],
      "summary":              "...",
      "enriched_on":          "2026-04-30",
      "history": [
        {
          "company":   "Pacific Gas & Electric",
          "title":     "Senior Protection & Relay Engineer",
          "start":     "2024-03-01",
          "end":       null,
          "current":   true,
          "discipline": "Electrical — Protection & Relay",
          "seniority":  "Senior"
        },
        ...
      ],
      "education": [
        {
          "school":     "MIT",
          "degree":     "B.S.",
          "field":      "Electrical Engineering",
          "start_year": 2016,
          "end_year":   2020
        }
      ]
    }
  ]
}
"""

from __future__ import annotations

import json
import logging
import re
import sys
from datetime import datetime, date
from pathlib import Path

import openpyxl

import config
from title_taxonomy import classify

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [EXPORTER] %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

EXCEL_PATH  = Path(config.EXCEL_PATH)
OUTPUT_PATH = Path(config.REPO_ROOT) / "docs" / "data" / "network.json"


# ── Excel helpers ─────────────────────────────────────────────────────────────

def sheet_to_dicts(ws: openpyxl.worksheet.worksheet.Worksheet) -> list[dict]:
    """Convert a worksheet to a list of dicts keyed by the header row."""
    rows = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h is not None else f"col_{i}"
               for i, h in enumerate(next(rows))]
    return [
        {headers[i]: cell for i, cell in enumerate(row)}
        for row in rows
        if any(cell is not None for cell in row)  # skip blank rows
    ]


def safe_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def safe_date(val) -> str | None:
    """Return ISO date string or None for any date-like value."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.date().isoformat() if isinstance(val, datetime) else val.isoformat()
    s = str(val).strip()
    if not s or s.lower() in ("none", "null", ""):
        return None
    # Try parsing common formats
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s  # return raw if unparseable — front-end will handle gracefully


def split_semicolon(val) -> list[str]:
    """Split a semicolon-delimited cell value into a clean list."""
    if not val:
        return []
    return [x.strip() for x in str(val).split(";") if x.strip()]


# ── Education string → list ────────────────────────────────────────────────────
# The enricher stores education as:
#   "MIT (B.S. Electrical Engineering, 2016–2020); Stanford (M.S. AI, 2020–2022)"
# We re-parse it into structured dicts here.

EDU_RE = re.compile(
    r"^(?P<school>[^(]+?)"                  # school name (up to first paren or end)
    r"(?:\s*\((?P<detail>[^)]+)\))?"        # optional (degree, field, years)
    r"\s*$"
)
YEAR_RANGE_RE = re.compile(r"(\d{4})\s*[–\-]\s*(\d{4})")
SINGLE_YEAR_RE = re.compile(r"graduated\s+(\d{4})")


def parse_edu_string(edu_str: str) -> list[dict]:
    """Convert the semicolon-separated education string back into structured dicts."""
    entries = []
    for chunk in split_semicolon(edu_str):
        m = EDU_RE.match(chunk)
        if not m:
            entries.append({"school": chunk, "degree": "", "field": "", "start_year": None, "end_year": None})
            continue

        school = m.group("school").strip()
        detail = m.group("detail") or ""

        # Extract years
        yr = YEAR_RANGE_RE.search(detail)
        start_year = int(yr.group(1)) if yr else None
        end_year   = int(yr.group(2)) if yr else None
        if not end_year:
            sy = SINGLE_YEAR_RE.search(detail)
            if sy:
                end_year = int(sy.group(1))

        # Remove year portion from detail to get degree/field
        detail_clean = YEAR_RANGE_RE.sub("", detail)
        detail_clean = SINGLE_YEAR_RE.sub("", detail_clean).strip(" ,")

        # Heuristic: first comma-separated token is degree, rest is field
        parts = [p.strip() for p in detail_clean.split(",", 1) if p.strip()]
        degree = parts[0] if parts else ""
        field  = parts[1] if len(parts) > 1 else ""

        entries.append({
            "school":     school,
            "degree":     degree,
            "field":      field,
            "start_year": start_year,
            "end_year":   end_year,
        })
    return entries


# ── Build history from Employment History tab ──────────────────────────────────

def build_history_map(history_rows: list[dict]) -> dict[str, list[dict]]:
    """
    Group Employment History rows by connection_id.
    Applies taxonomy to each historical title.
    Returns {connection_id: [job_entry, ...]} sorted newest-first.
    """
    history_map: dict[str, list[dict]] = {}

    for row in history_rows:
        conn_id = safe_str(row.get("connection_id"))
        if not conn_id:
            continue

        title   = safe_str(row.get("title") or row.get("job_title") or "")
        company = safe_str(row.get("company") or row.get("company_name") or "")
        start   = safe_date(row.get("start") or row.get("start_date"))
        end     = safe_date(row.get("end")   or row.get("end_date"))
        current = str(row.get("is_current") or row.get("current") or "").upper() in ("YES", "TRUE", "1", "CURRENT")
        if end is None and current:
            end = None  # keep null for current roles

        taxonomy = classify(title)

        entry = {
            "company":              company,
            "title":                title,
            "start":                start,
            "end":                  end,
            "current":              current,
            "discipline":           taxonomy["discipline"],
            "discipline_family":    taxonomy["discipline_family"],
            "discipline_specialty": taxonomy["discipline_specialty"],
            "seniority":            taxonomy["seniority"],
        }

        history_map.setdefault(conn_id, []).append(entry)

    # Sort each person's history: current jobs first, then by start date desc
    for conn_id in history_map:
        history_map[conn_id].sort(
            key=lambda e: (
                0 if e["current"] else 1,
                e["start"] or "0000-00-00",
            ),
            reverse=False,
        )
        # Flip so current is [0], oldest is [-1]
        history_map[conn_id] = list(reversed(history_map[conn_id]))

    return history_map


# ── Main export ────────────────────────────────────────────────────────────────

def export():
    log.info("Loading workbook: %s", EXCEL_PATH)
    if not EXCEL_PATH.exists():
        log.error("Workbook not found. Run pipeline.py first.")
        sys.exit(1)

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)

    if "Current Snapshot" not in wb.sheetnames:
        log.error("'Current Snapshot' sheet not found in workbook.")
        sys.exit(1)

    snapshot_rows = sheet_to_dicts(wb["Current Snapshot"])
    log.info("Snapshot rows loaded: %d", len(snapshot_rows))

    history_rows = []
    if "Employment History" in wb.sheetnames:
        history_rows = sheet_to_dicts(wb["Employment History"])
        log.info("Employment history rows loaded: %d", len(history_rows))
    else:
        log.warning("'Employment History' sheet not found — history will be empty.")

    wb.close()

    history_map = build_history_map(history_rows)

    connections = []
    skipped = 0

    for row in snapshot_rows:
        conn_id = safe_str(row.get("connection_id") or row.get("linkedin_id") or "")
        first   = safe_str(row.get("first_name") or row.get("first") or "")
        last    = safe_str(row.get("last_name")  or row.get("last")  or "")

        if not conn_id or (not first and not last):
            skipped += 1
            continue

        current_title = safe_str(
            row.get("current_title") or row.get("title") or ""
        )
        taxonomy = classify(current_title)

        # Skills: stored as semicolon string in Excel
        skills_raw = safe_str(row.get("skills") or "")
        skills = split_semicolon(skills_raw)

        # Education: stored as semicolon string in Excel
        edu_raw = safe_str(row.get("education") or "")
        education = parse_edu_string(edu_raw) if edu_raw else []

        conn = {
            # Identity
            "id":                   conn_id,
            "first":                first,
            "last":                 last,
            "headline":             safe_str(row.get("headline") or ""),
            "location":             safe_str(row.get("location") or ""),
            "photo_url":            safe_str(row.get("profile_photo_url") or row.get("photo_url") or ""),
            "connected_on":         safe_date(row.get("connected_on") or row.get("data_since")),

            # Current role
            "current_company":      safe_str(row.get("current_company") or ""),
            "current_title":        current_title,
            "current_start":        safe_date(row.get("current_start_date") or row.get("current_start")),

            # Taxonomy (from current title)
            "discipline":           taxonomy["discipline"],
            "discipline_family":    taxonomy["discipline_family"],
            "discipline_specialty": taxonomy["discipline_specialty"],
            "seniority":            taxonomy["seniority"],

            # Enriched fields
            "school":               safe_str(row.get("school") or ""),
            "industry":             safe_str(row.get("industry") or "Unknown"),
            "skills":               skills,
            "summary":              safe_str(row.get("summary") or ""),
            "enriched_on":          safe_date(row.get("enriched_on")),

            # Nested data
            "education":            education,
            "history":              history_map.get(conn_id, []),
        }
        connections.append(conn)

    # Sort alphabetically for stable diffs in git
    connections.sort(key=lambda c: (c["last"].lower(), c["first"].lower()))

    payload = {
        "generated":   datetime.now().isoformat(timespec="seconds"),
        "count":       len(connections),
        "connections": connections,
    }

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    log.info(
        "Exported %d connections → %s  (skipped %d malformed rows)",
        len(connections), OUTPUT_PATH, skipped,
    )

    # Print a quick discipline breakdown for sanity-checking
    from collections import Counter
    families = Counter(c["discipline_family"] for c in connections)
    log.info("Discipline breakdown:")
    for family, count in families.most_common():
        log.info("  %-30s %d", family, count)


if __name__ == "__main__":
    export()
