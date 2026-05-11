"""
change_detector.py — CSV diff engine

Reads the incoming Connections.csv from data/csv_inbox/,
compares it against the master Excel workbook, and writes
change events into the Employment History tab.

Change types written to history:
  INITIAL   — first time we see this connection
  ARRIVAL   — they changed to a new company (opens new entry, today as start)
  DEPARTURE — previous company closed (yesterday as end date)
  UPDATE    — title, headline, or location changed (no date change)

Called by pipeline.py with env var PIPELINE_CSV_PATH set to the CSV file.
Can also be run standalone for testing:
    python change_detector.py path/to/Connections.csv
"""

from __future__ import annotations

import csv
import logging
import os
import sys
from datetime import date, timedelta
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import config
from title_taxonomy import classify

log = logging.getLogger(__name__)

EXCEL_PATH = Path(config.EXCEL_PATH)

# ── Column definitions ────────────────────────────────────────────────────────
# Current Snapshot sheet columns (in order)
SNAP_COLS = [
    "connection_id",        # derived from LinkedIn URL slug
    "linkedin_public_id",   # the slug itself (e.g. "jane-smith-123")
    "first_name",
    "last_name",
    "headline",
    "location",
    "current_company",
    "current_title",
    "current_start_date",
    "profile_photo_url",
    "school",
    "discipline",
    "discipline_family",
    "discipline_specialty",
    "seniority",
    "industry",
    "skills",
    "education",
    "summary",
    "enriched_on",
    "connected_on",
    "data_since",
    "last_updated",
    "change_log",
]

# Employment History sheet columns (in order)
HIST_COLS = [
    "connection_id",
    "first_name",
    "last_name",
    "company",
    "title",
    "discipline",
    "discipline_family",
    "discipline_specialty",
    "seniority",
    "start",
    "end",
    "is_current",
    "change_type",          # INITIAL / ARRIVAL / DEPARTURE / UPDATE
    "recorded_on",          # date this row was written
]

LINKEDIN_BLUE = "0A66C2"
HEADER_FILL   = PatternFill("solid", fgColor=LINKEDIN_BLUE)
HEADER_FONT   = Font(color="FFFFFF", bold=True, size=10)
ALT_FILL      = PatternFill("solid", fgColor="F3F6FA")


# ── CSV parsing ───────────────────────────────────────────────────────────────

def parse_csv(csv_path: Path) -> list[dict]:
    """
    Parse a LinkedIn Connections.csv export.
    Returns a list of dicts with normalised keys.
    Handles both the "larger archive" and "specific files" CSV formats.
    """
    records = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        # LinkedIn sometimes puts notes at the top — skip until we find the header
        lines = f.readlines()

    # Find the header row (contains "First Name" or "FirstName")
    header_idx = 0
    for i, line in enumerate(lines):
        if "first" in line.lower() and ("name" in line.lower() or "Name" in line):
            header_idx = i
            break

    content = "".join(lines[header_idx:])
    reader = csv.DictReader(content.splitlines())

    for row in reader:
        # Normalise key names across different LinkedIn export formats
        norm = {k.strip().lower().replace(" ", "_"): (v or "").strip()
                for k, v in row.items()}

        first = norm.get("first_name") or norm.get("firstname") or ""
        last  = norm.get("last_name")  or norm.get("lastname")  or ""
        url   = norm.get("url") or norm.get("profile_url") or norm.get("linkedin_url") or ""
        company = norm.get("company") or norm.get("current_company") or ""
        position = norm.get("position") or norm.get("title") or norm.get("current_position") or ""
        connected_on = norm.get("connected_on") or norm.get("connection_date") or ""

        if not first and not last:
            continue  # skip blank rows

        # Derive a stable ID from the profile URL slug
        public_id = _slug_from_url(url)
        conn_id   = "csv_" + public_id if public_id else f"csv_{first}_{last}".lower().replace(" ", "_")

        records.append({
            "connection_id":      conn_id,
            "linkedin_public_id": public_id,
            "first_name":         first,
            "last_name":          last,
            "current_company":    company,
            "current_title":      position,
            "connected_on":       connected_on,
            "url":                url,
        })

    log.info("CSV parsed: %d connections found", len(records))
    return records


def _slug_from_url(url: str) -> str:
    """Extract the profile slug from a LinkedIn URL."""
    url = url.strip().rstrip("/")
    if "/in/" in url:
        return url.split("/in/")[-1].split("?")[0].strip("/")
    return url.replace("https://", "").replace("http://", "").replace("www.linkedin.com/in/", "").strip("/")


# ── Workbook management ────────────────────────────────────────────────────────

def load_or_create_workbook() -> tuple[Workbook, dict, list[dict]]:
    """
    Load the master workbook if it exists, or create a fresh one.
    Returns (workbook, snap_by_id, history_rows).
    snap_by_id: {connection_id: {col: value}} for quick lookup.
    """
    if EXCEL_PATH.exists():
        wb = openpyxl.load_workbook(EXCEL_PATH)
    else:
        log.info("Creating new workbook at %s", EXCEL_PATH)
        wb = _create_workbook()

    # Ensure both sheets exist
    if "Current Snapshot" not in wb.sheetnames:
        _create_snapshot_sheet(wb)
    if "Employment History" not in wb.sheetnames:
        _create_history_sheet(wb)

    snap_by_id = _load_snapshot(wb)
    history_rows = _load_history(wb)

    return wb, snap_by_id, history_rows


def _create_workbook() -> Workbook:
    wb = Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    _create_snapshot_sheet(wb)
    _create_history_sheet(wb)
    return wb


def _create_snapshot_sheet(wb: Workbook):
    ws = wb.create_sheet("Current Snapshot")
    _write_header(ws, SNAP_COLS)


def _create_history_sheet(wb: Workbook):
    ws = wb.create_sheet("Employment History")
    _write_header(ws, HIST_COLS)


def _write_header(ws, cols: list[str]):
    for i, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i, value=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def _load_snapshot(wb: Workbook) -> dict:
    ws = wb["Current Snapshot"]
    headers = [c.value for c in ws[1]]
    snap = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        d = {headers[i]: row[i] for i in range(len(headers)) if i < len(row)}
        cid = d.get("connection_id")
        if cid:
            snap[cid] = d
    return snap


def _load_history(wb: Workbook) -> list[dict]:
    ws = wb["Employment History"]
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        rows.append({headers[i]: row[i] for i in range(len(headers)) if i < len(row)})
    return rows


# ── Core diff logic ────────────────────────────────────────────────────────────

def process_changes(
    csv_records: list[dict],
    snap_by_id: dict,
    wb: Workbook,
) -> int:
    """
    Diff CSV records against stored snapshot.
    Writes changes directly into the workbook sheets.
    Returns count of changed records.
    """
    ws_snap = wb["Current Snapshot"]
    ws_hist = wb["Employment History"]

    today     = date.today().isoformat()
    yesterday = (date.today() - timedelta(days=1)).isoformat()

    snap_headers = [c.value for c in ws_snap[1]]
    hist_headers = [c.value for c in ws_hist[1]]

    # Build row-number index for snapshot (connection_id → row number)
    snap_row_idx = {}
    for row in ws_snap.iter_rows(min_row=2):
        cid = row[0].value
        if cid:
            snap_row_idx[cid] = row[0].row

    changed = 0

    for rec in csv_records:
        cid   = rec["connection_id"]
        first = rec["first_name"]
        last  = rec["last_name"]
        new_co    = rec["current_company"]
        new_title = rec["current_title"]

        taxonomy = classify(new_title)

        if cid not in snap_by_id:
            # ── NEW CONNECTION ────────────────────────────────────────────
            log.info("INITIAL: %s %s @ %s", first, last, new_co)

            snap_row = _blank_snap(rec, taxonomy, today)
            _append_row(ws_snap, snap_headers, snap_row)
            snap_row_idx[cid] = ws_snap.max_row

            if new_co:
                _append_row(ws_hist, hist_headers, {
                    "connection_id":      cid,
                    "first_name":         first,
                    "last_name":          last,
                    "company":            new_co,
                    "title":              new_title,
                    "discipline":         taxonomy["discipline"],
                    "discipline_family":  taxonomy["discipline_family"],
                    "discipline_specialty": taxonomy["discipline_specialty"],
                    "seniority":          taxonomy["seniority"],
                    "start":              today,
                    "end":                None,
                    "is_current":         "YES",
                    "change_type":        "INITIAL",
                    "recorded_on":        today,
                })
            changed += 1

        else:
            # ── EXISTING CONNECTION — check for changes ───────────────────
            stored = snap_by_id[cid]
            old_co    = stored.get("current_company") or ""
            old_title = stored.get("current_title")   or ""

            updates = {}
            log_parts = []

            if new_co and new_co != old_co:
                # Company changed — close old entry, open new one
                log.info("COMPANY CHANGE: %s %s  %s → %s", first, last, old_co, new_co)

                # Close old history entry
                _close_current_history(ws_hist, hist_headers, cid, yesterday)

                # Open new history entry
                _append_row(ws_hist, hist_headers, {
                    "connection_id":        cid,
                    "first_name":           first,
                    "last_name":            last,
                    "company":              new_co,
                    "title":                new_title,
                    "discipline":           taxonomy["discipline"],
                    "discipline_family":    taxonomy["discipline_family"],
                    "discipline_specialty": taxonomy["discipline_specialty"],
                    "seniority":            taxonomy["seniority"],
                    "start":                today,
                    "end":                  None,
                    "is_current":           "YES",
                    "change_type":          "ARRIVAL",
                    "recorded_on":          today,
                })

                updates["current_company"]    = new_co
                updates["current_title"]      = new_title
                updates["current_start_date"] = today
                updates["discipline"]           = taxonomy["discipline"]
                updates["discipline_family"]    = taxonomy["discipline_family"]
                updates["discipline_specialty"] = taxonomy["discipline_specialty"]
                updates["seniority"]            = taxonomy["seniority"]
                log_parts.append(f"company: {old_co} → {new_co}")
                changed += 1

            elif new_title and new_title != old_title:
                # Title changed within same company
                log.info("TITLE CHANGE: %s %s  '%s' → '%s'", first, last, old_title, new_title)
                _append_row(ws_hist, hist_headers, {
                    "connection_id":        cid,
                    "first_name":           first,
                    "last_name":            last,
                    "company":              new_co,
                    "title":                new_title,
                    "discipline":           taxonomy["discipline"],
                    "discipline_family":    taxonomy["discipline_family"],
                    "discipline_specialty": taxonomy["discipline_specialty"],
                    "seniority":            taxonomy["seniority"],
                    "start":                today,
                    "end":                  None,
                    "is_current":           "YES",
                    "change_type":          "UPDATE",
                    "recorded_on":          today,
                })
                updates["current_title"]        = new_title
                updates["discipline"]           = taxonomy["discipline"]
                updates["discipline_family"]    = taxonomy["discipline_family"]
                updates["discipline_specialty"] = taxonomy["discipline_specialty"]
                updates["seniority"]            = taxonomy["seniority"]
                log_parts.append(f"title: {old_title!r} → {new_title!r}")
                changed += 1

            # Always keep connected_on if we have it and didn't before
            if rec.get("connected_on") and not stored.get("connected_on"):
                updates["connected_on"] = rec["connected_on"]

            if updates:
                updates["last_updated"] = today
                if log_parts:
                    existing_log = stored.get("change_log") or ""
                    new_entry = f"{today}: " + "; ".join(log_parts)
                    updates["change_log"] = (existing_log + "\n" + new_entry).strip()
                _update_snap_row(ws_snap, snap_headers, snap_row_idx[cid], updates)

    log.info("Change detection complete: %d records changed / added", changed)
    return changed


def _blank_snap(rec: dict, taxonomy: dict, today: str) -> dict:
    return {
        "connection_id":        rec["connection_id"],
        "linkedin_public_id":   rec["linkedin_public_id"],
        "first_name":           rec["first_name"],
        "last_name":            rec["last_name"],
        "headline":             "",
        "location":             "",
        "current_company":      rec["current_company"],
        "current_title":        rec["current_title"],
        "current_start_date":   today,
        "profile_photo_url":    "",
        "school":               "",
        "discipline":           taxonomy["discipline"],
        "discipline_family":    taxonomy["discipline_family"],
        "discipline_specialty": taxonomy["discipline_specialty"],
        "seniority":            taxonomy["seniority"],
        "industry":             "Unknown",
        "skills":               "",
        "education":            "",
        "summary":              "",
        "enriched_on":          None,
        "connected_on":         rec.get("connected_on", ""),
        "data_since":           today,
        "last_updated":         today,
        "change_log":           f"{today}: INITIAL",
    }


def _append_row(ws, headers: list[str], data: dict):
    row_num = ws.max_row + 1
    for i, col in enumerate(headers, 1):
        val = data.get(col)
        cell = ws.cell(row=row_num, column=i, value=val)
        # Alternate row shading
        if row_num % 2 == 0:
            cell.fill = ALT_FILL


def _update_snap_row(ws, headers: list[str], row_num: int, updates: dict):
    col_idx = {h: i + 1 for i, h in enumerate(headers)}
    for field, value in updates.items():
        if field in col_idx:
            ws.cell(row=row_num, column=col_idx[field], value=value)


def _close_current_history(ws, headers: list[str], conn_id: str, end_date: str):
    """Set end date and is_current=NO on the open history entry for this person."""
    col_idx = {h: i + 1 for i, h in enumerate(headers)}
    cid_col     = col_idx.get("connection_id", 1)
    current_col = col_idx.get("is_current")
    end_col     = col_idx.get("end")

    for row in ws.iter_rows(min_row=2):
        if row[cid_col - 1].value == conn_id:
            if current_col and row[current_col - 1].value == "YES":
                row[current_col - 1].value = "NO"
                if end_col:
                    row[end_col - 1].value = end_date


# ── Entry point ────────────────────────────────────────────────────────────────

def run(csv_path: Path | None = None):
    if csv_path is None:
        # Check env var set by pipeline.py
        env_path = os.environ.get("PIPELINE_CSV_PATH")
        if env_path:
            csv_path = Path(env_path)
        elif len(sys.argv) > 1:
            csv_path = Path(sys.argv[1])
        else:
            log.error("No CSV path provided. Usage: python change_detector.py path/to/Connections.csv")
            sys.exit(1)

    if not csv_path.exists():
        log.error("CSV not found: %s", csv_path)
        sys.exit(1)

    csv_records = parse_csv(csv_path)
    if not csv_records:
        log.warning("No records parsed from CSV — check the file format.")
        return

    wb, snap_by_id, history_rows = load_or_create_workbook()
    process_changes(csv_records, snap_by_id, wb)

    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(EXCEL_PATH)
    log.info("Workbook saved: %s", EXCEL_PATH)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [CHANGE_DETECTOR] %(levelname)s %(message)s",
    )
    run()
