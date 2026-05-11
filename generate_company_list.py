"""
generate_company_list.py — One-shot utility

Reads the master Excel workbook and prints every unique company name
found anywhere in the Employment History tab.

Usage:
    python generate_company_list.py

Then copy the printed list and send it to Claude with:
    "Categorize these companies by industry: [paste list]"

Claude will return a JSON block you paste into data/industry_map.json.
The enricher reads that file automatically on its next run.
"""

from pathlib import Path
import openpyxl
import config

EXCEL_PATH = Path(config.EXCEL_PATH)

def main():
    if not EXCEL_PATH.exists():
        print(f"ERROR: Workbook not found at {EXCEL_PATH}")
        print("Run pipeline.py at least once first.")
        return

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)

    companies = set()

    # Pull from Employment History tab (column: 'company')
    if "Employment History" in wb.sheetnames:
        ws = wb["Employment History"]
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        try:
            co_col = headers.index("company")
        except ValueError:
            co_col = None

        if co_col is not None:
            for row in ws.iter_rows(min_row=2, values_only=True):
                val = row[co_col]
                if val and str(val).strip():
                    companies.add(str(val).strip())

    # Also pull current_company from Snapshot tab as a cross-check
    if "Current Snapshot" in wb.sheetnames:
        ws2 = wb["Current Snapshot"]
        headers2 = [c.value for c in next(ws2.iter_rows(min_row=1, max_row=1))]
        try:
            co_col2 = headers2.index("current_company")
        except ValueError:
            co_col2 = None

        if co_col2 is not None:
            for row in ws2.iter_rows(min_row=2, values_only=True):
                val = row[co_col2]
                if val and str(val).strip():
                    companies.add(str(val).strip())

    wb.close()

    sorted_cos = sorted(companies)
    print(f"\nFound {len(sorted_cos)} unique companies:\n")
    print(", ".join(sorted_cos))
    print("\n─────────────────────────────────────────────────────")
    print("Send the list above to Claude with this message:")
    print()
    print('  "Categorize these companies by industry and return')
    print("   a JSON object in this format:")
    print('   {\"Company Name\": \"Industry\", ...}')
    print('   Use broad industry labels like: Technology, Finance,')
    print('   Engineering, Consulting, Healthcare, Energy, Real Estate,')
    print('   Government, Education, Media, Manufacturing, Retail."')
    print()
    print("Then save Claude's JSON response to: data/industry_map.json")


if __name__ == "__main__":
    main()
