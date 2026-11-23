# Netikot XView Dashboard

A lightweight Flask web dashboard for monitoring **Netikot / XView** connectivity check results. It reads a SQLite results database (read-only), summarizes the health of every monitored IP and site, highlights failure trends, and exports a formatted weekly Excel report with charts.

---

## Features

- **Overview dashboard** – total checks, success rate, unique IPs, healthy vs. failing targets, and stats for the most recent execution run.
- **Weekly Status card** – shows the share of current targets that are "updated": their latest check succeeded, or they had a successful check within the 7 days before their latest failure.
- **Last-run detection** – groups results into runs automatically (a gap of 30+ minutes between results starts a new run). The current target list is taken from the first run on the most recent day of data, so IPs that are no longer checked drop out of the stats.
- **Failure analysis**
  - Top failure types from each IP's latest result
  - 30-day failure-type trend chart (one result per IP per day)
  - Per-IP error count since its last successful check, with a distribution chart
- **Focused views** – incorrect passwords, closed ports, latest failures, latest successes, and per-site coverage.
- **Search** – look up any IP address, site name, or detail text across all results.
- **Weekly Excel report** – one-click `.xlsx` download with a failure-type line chart, an IP errors bar chart, and supporting tables.
- **Light / dark theme** – toggle saved in the browser.
- **Fast and safe** – the database is opened in read-only, immutable mode, and every page plus the Excel report is pre-built and refreshed in the background every 10 minutes.

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

> To use a different path, set the `NETIKOT_DB_PATH` environment variable (for example `-e NETIKOT_DB_PATH=/data/other.db`).

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

# Make sure the database exists at /data/NetikotDB.db (or set NETIKOT_DB_PATH)
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

These settings at the top of `app.py` control behavior:

| Setting | Default | Purpose |
|---------|---------|---------|
| `DB_PATH` | `/data/NetikotDB.db` | Path to the SQLite database (override with the `NETIKOT_DB_PATH` env var) |
| `ROW_LIMIT` | `500` | Maximum rows shown on list pages |
| `PAGE_CACHE_REFRESH_SECONDS` | `600` | How often the background thread rebuilds all cached pages and the report |

---

## Notes

- The dashboard never writes to the database; it's opened with `mode=ro&immutable=1`.
- All pages (except search) and the Excel report are served from a cache that a background thread rebuilds every 10 minutes, so new results can take up to 10 minutes to appear. If a refresh fails, the previous pages stay in place. Search always queries live data.
- Open pages reload themselves in the browser every 10 minutes.
- The background refresher starts when the app is launched with `python app.py` (as the Docker image does).
- The Docker container runs as a non-root user (`appuser`).
