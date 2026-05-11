# LinkedIn Network Graph

A self-hosted network visualization of your LinkedIn connections.
Data lives in a private Excel workbook on your machine.
The front-end is a static GitHub Pages site fed by `network.json`.

---

## Repo structure

```
linkedin-pipeline/
│
├── docs/                        ← GitHub Pages root
│   ├── index.html               ← The visualization web app
│   └── data/
│       └── network.json         ← Auto-generated; committed by git_push.py
│
├── data/                        ← Local only (.gitignored)
│   ├── network_master.xlsx      ← Master workbook (two tabs)
│   ├── industry_map.json        ← Company → industry mapping (you build this)
│   ├── csv_inbox/               ← Drop LinkedIn CSVs here
│   ├── snapshots/               ← Daily JSON backups
│   └── logs/                    ← Scheduler logs
│
├── config.py                    ← Your credentials + paths (.gitignored)
├── title_taxonomy.py            ← Granular engineering title classifier
├── pipeline.py                  ← Daily orchestrator (CSV → Excel → JSON → push)
├── enricher.py                  ← Slow background profile enricher
├── export_to_json.py            ← Excel → network.json
├── change_detector.py           ← CSV diff logic
├── exporter.py                  ← Excel writer
├── git_push.py                  ← Auto-commit + push
├── scheduler.py                 ← Windows Task Scheduler setup
├── generate_company_list.py     ← Print all companies for industry categorization
└── requirements.txt
```

---

## First-time setup (Mac — development)

```bash
# 1. Clone your repo
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME

# 2. Create virtualenv
python3 -m venv .venv
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Edit config.py
#    Set LINKEDIN_EMAIL, LINKEDIN_PASSWORD, and verify all paths

# 5. Run the pipeline once manually
python pipeline.py
#    → Opens LinkedIn export page
#    → You download the CSV and drop it in data/csv_inbox/
#    → Pipeline processes it, exports JSON, pushes to GitHub

# 6. Test the taxonomy classifier
python title_taxonomy.py
```

---

## Production setup (Windows desktop)

```powershell
# 1. Clone repo
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME

# 2. Install dependencies (no venv needed for Task Scheduler simplicity)
pip install -r requirements.txt

# 3. Edit config.py with your credentials

# 4. Install scheduled tasks (run once, as Administrator)
python scheduler.py install

# 5. Verify tasks are registered
python scheduler.py status

# 6. Trigger a test run immediately
python scheduler.py run-now
```

---

## GitHub Pages setup

1. Push this repo to GitHub (private repo is fine)
2. In repo Settings → Pages:
   - Source: **Deploy from a branch**
   - Branch: `main`, folder: `/docs`
3. Your site will be at `https://YOUR_USERNAME.github.io/YOUR_REPO_NAME/`
4. The `docs/data/network.json` file is what the site fetches — it updates every time `export_to_json.py` runs and pushes

---

## Daily workflow

**Automatic** (after Task Scheduler is set up):
- 7:00 AM — `pipeline.py` starts, opens LinkedIn export page
- You download CSV within the hour and drop it in `data/csv_inbox/`
- Pipeline detects it, updates Excel, exports JSON, pushes to GitHub
- 8:00 AM — `enricher.py` fetches 8–12 profiles (runs quietly in background)

**Manual** (any time):
```bash
python pipeline.py        # full daily cycle
python enricher.py        # run one enrichment batch
python export_to_json.py  # just re-export without fetching anything
python git_push.py        # just push current network.json
```

---

## Industry categorization

```bash
# 1. Generate company list
python generate_company_list.py
# → Prints all company names

# 2. Send list to Claude:
#    "Categorize these companies by industry and return
#     {"Company Name": "Industry"} JSON."

# 3. Paste Claude's JSON response into data/industry_map.json

# 4. Re-export (enricher picks up new industries on next run too)
python export_to_json.py && python git_push.py
```

---

## Title taxonomy

The `title_taxonomy.py` module classifies every job title into two fields:

| Field | Example |
|---|---|
| `discipline` | `Electrical — Protection & Relay` |
| `discipline_family` | `Electrical` |
| `discipline_specialty` | `Protection & Relay` |
| `seniority` | `Senior` |

Covered disciplines: Electrical (10 specialties), Civil (8), Mechanical (7),
Software/Controls (12), Environmental (6), Chemical/Oil & Gas (3),
Commissioning (5), Robotics (3), Defense (5), Manufacturing (3),
Project Management (5), Consulting, Finance, Business.

To test a title: `python title_taxonomy.py`

---

## Troubleshooting

**"Workbook not found"** → Run `pipeline.py` at least once first to create it.

**"No public_id found"** → The enricher needs `linkedin_public_id` in the snapshot sheet.
This is populated when `change_detector.py` parses the CSV (the URL column).

**LinkedIn blocks the enricher** → Increase `ENRICHER_DELAY_MIN`/`MAX` in config.py.
Reduce `ENRICHER_DAILY_MAX` to 5–6 for a while.

**Git push fails** → Run `git pull --rebase origin main` in the repo folder first.

**GitHub Pages shows old data** → Hard-refresh the browser (Ctrl+Shift+R).
The front-end adds `?v=timestamp` to bust the cache, but browsers sometimes
hold on anyway.
