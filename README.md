# Netikot XView Dashboard

A lightweight Flask web dashboard for monitoring **Netikot / XView** connectivity check results. It reads a SQLite results database (read-only), summarizes the health of every monitored IP and site, highlights failure trends, and exports a formatted weekly Excel report with charts.

---

## Features

- **Overview dashboard** – total checks, success rate, unique IPs, healthy vs. failing targets, and stats for the most recent execution run.
- **Last-run detection** – groups results into runs automatically (a gap of 30+ minutes between results starts a new run).
- **Failure analysis**
  - Top failure types from each IP's latest result
  - 30-day failure-type trend chart (one result per IP per day)
  - Per-IP error count since its last successful check, with a distribution chart
- **Focused views** – incorrect passwords, closed ports, latest failures, latest successes, and per-site coverage.
- **Search** – look up any IP address, site name, or detail text across all results.
- **Weekly Excel report** – one-click `.xlsx` download with a failure-type line chart, an IP errors bar chart, and supporting tables.
- **Light / dark theme** – toggle saved in the browser.
- **Fast and safe** – the database is opened in read-only, immutable mode, and the main dashboard is cached for 5 minutes.

---

## Tech stack

| Component | Technology |
|-----------|------------|
| Backend   | Python 3.12, Flask 3 |
| Database  | SQLite (read-only) |
| Reports   | openpyxl |
| Frontend  | Jinja2 templates, CSS, inline SVG charts |
| Deployment | Docker |

---

## Project structure

```
Netikot-XView-Dashboard/
├── app.py                 # Flask app: routes, SQL queries, chart data, Excel export
├── requirements.txt       # Python dependencies
├── Dockerfile             # Container image definition
├── .dockerignore
├── static/
│   └── style.css          # Dashboard styling (light & dark themes)
└── templates/
    ├── base.html          # Layout, sidebar navigation, theme toggle
    ├── index.html         # Main dashboard
    ├── failure_reasons.html
    ├── failure_chart.html
    ├── failure_counts.html
    ├── ip_errors_since_success.html
    ├── ip_error_column_chart.html
    ├── results.html       # Generic results table page
    └── table.html
```

---

## Database requirements

The app expects a SQLite database at:

```
/data/NetikotDB.db
```

It must contain a table named `tbl_Results` with these columns:

| Column       | Description |
|--------------|-------------|
| `ID`         | Unique row ID (increasing) |
| `IP_Address` | Target IP address |
| `Port`       | Target port |
| `Site_Name`  | Site the target belongs to |
| `Is_Success` | `1` = success, `0` = failure |
| `Details`    | Result / error message (e.g. `Port is closed`, `timeout`, `password is incorrect`) |
| `Date`       | Check timestamp (SQLite-compatible datetime) |

An optional view named **`vew_LastMonthRun`** is used for the monthly charts. If it doesn't exist, the app falls back to the last 30 days of `tbl_Results` automatically.

> To use a different path, change `DB_PATH` at the top of `app.py`.

---

## Getting started

### Run with Docker (recommended)

```bash
# Build the image
docker build -t netikot-xview-dashboard .

# Run it, mounting the folder that contains NetikotDB.db
docker run -d \
  --name netikot-dashboard \
  -p 5000:5000 \
  -v /path/to/db-folder:/data:ro \
  netikot-xview-dashboard
```

Open **http://localhost:5000** in your browser.

### Run locally

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# Make sure the database exists at /data/NetikotDB.db (or edit DB_PATH in app.py)
python app.py
```

The server listens on `0.0.0.0:5000`.

---

## Pages and routes

| Route | Description |
|-------|-------------|
| `/` | Main dashboard with summary cards and charts |
| `/monthly_failure_types` | 30-day failure-type trend and daily breakdown (alias: `/failure_reasons`) |
| `/ip_errors_since_success` | Errors per IP since its last successful check |
| `/latest_failures` | Failed results from each IP's latest check |
| `/latest_successes` | Successful results from each IP's latest check |
| `/incorrect_passwords` | Targets currently failing with an incorrect password |
| `/port_closed` | Targets currently failing with `Port is closed` |
| `/site_summary` | Success / failure totals per site |
| `/search_ip?q=<text>` | Search by IP, site name, or detail text |
| `/download_report` | Download the weekly Excel report |

List pages show up to **500** rows (`ROW_LIMIT` in `app.py`).

---

## Weekly Excel report

`/download_report` generates `weekly_report-YYYYMMDD-HHMMSS.xlsx` covering the last 7 days of data. It contains:

1. **Weekly Report** – a failure-types-per-day line chart and a top-IP-errors bar chart
2. **Weekly Failure Types** – daily failure counts grouped by type
3. **IP Errors Graph** – bar chart of the IPs with the most errors since their last success
4. **IP Errors Table** – full filterable table (IP, site, errors, last error type, last success date)

---

## Configuration

These constants at the top of `app.py` control behavior:

| Setting | Default | Purpose |
|---------|---------|---------|
| `DB_PATH` | `/data/NetikotDB.db` | Path to the SQLite database |
| `ROW_LIMIT` | `500` | Maximum rows shown on list pages |
| `DASHBOARD_CACHE_TTL_SECONDS` | `300` | How long the main dashboard is cached |

---

## Notes

- The dashboard never writes to the database; it's opened with `mode=ro&immutable=1`.
- Because of the 5-minute cache, new results may take up to 5 minutes to appear on the main page. Other pages always query live data.
- The Docker container runs as a non-root user (`appuser`).
