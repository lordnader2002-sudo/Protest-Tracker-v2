# Protest Tracker

> Real-time protest & demonstration monitoring for your property portfolio — powered by Mobilize.us

---

## What It Does

Protest Tracker automatically scans the [Mobilize.us](https://www.mobilize.us) public event API for upcoming protests, rallies, town halls, and demonstrations near your properties. Every run produces:

- **An Excel workbook** (`protest_tracker_report.xlsx`) with color-coded, grouped results ready to share with stakeholders
- **A live web dashboard** (`index.html`) with a Summary overview, proximity charts, and fully sortable/filterable event tables

---

## Dashboard

Open `index.html` in any modern browser. No server required — it reads `dashboard_data.json` from the same directory.

### Tabs

| Tab | What it shows |
|-----|---------------|
| **Summary** | KPI tiles, proximity breakdown charts, top exposed properties, event type frequency |
| **3-Day Events** | All protest-type events within 3 miles, starting today through +3 days |
| **No Kings (30-Day)** | Events matching "No Kings" keywords within 3 miles, today through +30 days |

### Reading the color bands

| Color | Distance |
|-------|----------|
| 🔴 Red | < 1 mile from property |
| 🟡 Amber | 1 – 2 miles |
| 🟢 Green | > 2 miles (within search radius) |

### Flag columns

| Flag | Meaning |
|------|---------|
| **New** | Event ID was not present in the previous run's data — appeared since last check |
| **Dup** | The same event appears near 2 or more distinct properties |
| **Recur** | The same event has multiple distinct timeslots |

Rows marked **New** are displayed in **bold** for quick scanning.

### Summary page

The Summary tab loads automatically on open and shows:

- **Proximity Alert banner** — threat level (LOW / MODERATE / HIGH / CRITICAL) based on the number of events within 1 mile
- **Five KPI tiles** — total counts for both windows plus New, Duplicate, and Recurring flags
- **Proximity Breakdown** — animated bar charts showing the red/amber/green distribution for each reporting window
- **Top Exposed Properties** — ranked by event count for both windows
- **Event Types** — frequency breakdown across all tracked events

---

## Running the Tracker

### Prerequisites

```bash
pip install requests openpyxl tqdm zipcodes
```

### Basic run

```bash
python3 protest_tracker.py
```

Outputs:
- `protest_tracker_report.xlsx` — Excel workbook
- `dashboard_data.json` — data file consumed by the dashboard

### Options

```
--output FILENAME     Excel output filename (default: protest_tracker_report.xlsx)
--csv    FILENAME     Properties CSV to use (default: properties.csv)
--api-key KEY         Mobilize.us API key (or set MOBILIZE_API_KEY env var)
```

### API key (recommended)

An API key gives you a dedicated rate-limit bucket and avoids 429 errors on large portfolios. Set it as an environment variable:

```bash
export MOBILIZE_API_KEY="your_key_here"
python3 protest_tracker.py
```

---

## Properties File

`properties.csv` defines the locations to monitor. Each row is one property.

### Required columns

| Column | Description |
|--------|-------------|
| `property_id` | Unique identifier |
| `name` | Display name used in the dashboard and Excel report |
| `address` | Street address |
| `postal_code` | 5-digit ZIP code |
| `lat` | Latitude (decimal degrees) |
| `lon` | Longitude (decimal degrees) |

### Example

```csv
property_id,name,address,postal_code,lat,lon
P001,Downtown Tower,100 Main St,10001,40.7484,-73.9967
P002,Midtown Plaza,500 5th Ave,10036,40.7549,-73.9840
```

---

## Configuration

Key constants at the top of `protest_tracker.py`:

| Constant | Default | Description |
|----------|---------|-------------|
| `SEARCH_RADIUS_MI` | `3` | Miles from property to search |
| `DAYS_GENERAL` | `3` | Days ahead for the general window |
| `DAYS_NO_KINGS` | `30` | Days ahead for the No Kings window |
| `NO_KINGS_KEYWORDS` | see file | Keywords that trigger the No Kings filter |

---

## Excel Workbook

The generated `.xlsx` file contains:

- **Summary sheet** — run metadata, counts, and configuration snapshot
- **3-Day Events sheet** — color-banded, grouped, sortable event table
- **No Kings sheet** — same format for the 30-day No Kings window

Distance bands use the same red / amber / green color coding as the dashboard. Groups are separated by thick borders.

---

## How "New" Events Are Detected

On each run the tracker saves a cache of event IDs. On the next run, any event ID not present in the previous cache is flagged as **New**. This requires the cache file to exist from a prior run — on the very first run all events will appear as New.

---

## Data Source

All event data is pulled live from the [Mobilize.us public API](https://api.mobilize.us/v1/events). No data is stored beyond the generated output files. Results reflect what is publicly listed on Mobilize.us at the time of the run.
