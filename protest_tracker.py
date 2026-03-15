#!/usr/bin/env python3
"""
Protest Tracker
===============
Queries the Mobilize.us public API for events near every property in
properties.csv and produces a professional Excel workbook.

Two reporting windows
---------------------
  • 3-Day Events  : Any public events (protest / rally / town-hall …)
                    within 3 miles of a property, starting today through +3 days.
  • No Kings (30d): Events whose title or description contains a "No Kings"
                    keyword, within 3 miles, today through +30 days.

Usage
-----
  python3 protest_tracker.py [--output FILENAME] [--csv PROPERTIES_CSV]
"""

import argparse
import csv
import math
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Configuration ─────────────────────────────────────────────────────────────

PROPERTIES_CSV  = "properties.csv"
MOBILIZE_API    = "https://api.mobilize.us/v1/events"

# Optional API key — set via env var MOBILIZE_API_KEY or --api-key argument.
# Authenticated requests have their own rate-limit bucket (avoids shared-IP 429s).
MOBILIZE_API_KEY = os.environ.get("MOBILIZE_API_KEY", "")

SEARCH_RADIUS_MI  = 3    # hard filter: must be within this many miles of property
CLUSTER_RADIUS_MI = 35   # properties within this distance share one API query
DAYS_GENERAL      = 3    # 3-day event window
DAYS_NO_KINGS     = 30   # No Kings event window

NO_KINGS_KEYWORDS = ["no kings", "nokings", "no_kings", "#nokings", "50501"]

# Event types to EXCLUDE from the 3-day sheet (pure campaign/admin work).
# Everything else — including PROTEST, RALLY, MEETING, COMMUNITY, etc. — is kept.
EXCLUDE_TYPES = {"PHONE_BANK", "TEXT_BANK", "AUTOMATED_PHONE_BANK", "LETTER_WRITING",
                 "VOTER_REG", "FUNDRAISER", "TRAINING", "FRIEND_TO_FRIEND_OUTREACH"}

REQUEST_DELAY  = 2.0   # seconds between successful API requests
MAX_RETRIES    = 3     # max retries on transient errors (NOT 429)
RETRY_BACKOFF  = 2.0   # exponential backoff base (seconds): 2, 4, 8
PER_PAGE       = 200   # max results per page

# ── Excel colour palette ───────────────────────────────────────────────────────

HDR_NAVY   = PatternFill("solid", fgColor="1F3864")   # dark navy  (title rows)
HDR_BLUE   = PatternFill("solid", fgColor="2E75B6")   # medium blue (column headers)
HDR_PURPLE = PatternFill("solid", fgColor="5B2C8D")   # purple      (No Kings title)
COL_PURPLE = PatternFill("solid", fgColor="7B5EA7")   # lighter purple (No Kings col hdrs)
ALT_BLUE   = PatternFill("solid", fgColor="DCE6F1")   # pale blue   (even data rows)
ALT_PURPLE = PatternFill("solid", fgColor="E8DEFF")   # pale purple (even No Kings rows)
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")

THIN = Border(
    left   = Side(style="thin", color="B8CCE4"),
    right  = Side(style="thin", color="B8CCE4"),
    top    = Side(style="thin", color="B8CCE4"),
    bottom = Side(style="thin", color="B8CCE4"),
)

# ── Sheet column definitions ───────────────────────────────────────────────────

COLUMNS = [
    # (header label,            row-dict key,      col width)
    ("Property ID",             "property_id",      14),
    ("Property Name",           "property_name",    28),
    ("Property Address",        "property_addr",    36),
    ("Event Title",             "event_title",      42),
    ("Event Type",              "event_type",       16),
    ("Date & Time (UTC)",       "event_date",       22),
    ("Event Location",          "event_location",   44),
    ("Distance (mi)",           "distance_mi",      14),
    ("Event URL",               "event_url",        50),
]

# ── Helpers ────────────────────────────────────────────────────────────────────

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Great-circle distance in miles between two lat/lon points."""
    R = 3958.8
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2) ** 2
    return R * 2 * math.asin(math.sqrt(a))


def pad_zip(raw: str) -> str:
    """Zero-pad a zipcode to 5 digits (handles New England codes stored as ints)."""
    return str(raw).strip().zfill(5)


def load_properties(path: str) -> list[dict]:
    props = []
    with open(path, newline="", encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            try:
                props.append({
                    "id":      row["property_id"].strip(),
                    "name":    row["name"].strip(),
                    "address": row["address"].strip(),
                    "zip":     pad_zip(row["postal_code"]),
                    "lat":     float(row["lat"]),
                    "lon":     float(row["lon"]),
                })
            except (ValueError, KeyError):
                pass   # skip malformed rows
    return props


class RateLimitError(Exception):
    """Raised when Mobilize.us returns HTTP 429 so the caller can retry later."""


def _build_headers() -> dict:
    headers = {"User-Agent": "ProtestTracker/1.0 (internal security monitoring tool)"}
    if MOBILIZE_API_KEY:
        headers["Authorization"] = f"Bearer {MOBILIZE_API_KEY}"
    return headers


def fetch_events_for_zip(zipcode: str, max_dist: int,
                         start_ts: int, end_ts: int) -> list[dict]:
    """Return all public Mobilize events for a zip within the given time window.

    On HTTP 429 (rate-limit): skips the zip immediately — no hanging retries.
    On transient errors: retries up to MAX_RETRIES times with exponential backoff.
    """
    params: dict = {
        "zipcode":        zipcode,
        "max_dist":       max_dist,
        "timeslot_start": f"gte_{start_ts}",
        "timeslot_end":   f"lte_{end_ts}",
        "per_page":       PER_PAGE,
        "visibility":     "PUBLIC",
    }
    events: list[dict] = []
    url: str | None = MOBILIZE_API
    headers = _build_headers()

    while url:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = requests.get(
                    url,
                    params=params if url == MOBILIZE_API else None,
                    headers=headers,
                    timeout=20,
                )

                if resp.status_code == 429:
                    # Signal caller to retry this cluster later.
                    raise RateLimitError(zipcode)

                if resp.status_code == 400:
                    # Bad request — invalid zip or malformed params. Not retryable.
                    print(f"    [400] Bad request for zip={zipcode} — skipping cluster.",
                          file=sys.stderr)
                    return events

                resp.raise_for_status()
                payload = resp.json()
                events.extend(payload.get("data", []))
                url = payload.get("next")
                params = None
                time.sleep(REQUEST_DELAY)
                break   # success — move to next page

            except requests.RequestException as exc:
                wait = RETRY_BACKOFF ** attempt
                print(f"    [WARN] API error zip={zipcode} (attempt {attempt}/{MAX_RETRIES}): "
                      f"{exc}. Retrying in {wait:.0f}s …", file=sys.stderr)
                time.sleep(wait)
        else:
            # All retries exhausted for this page (transient errors only)
            print(f"    [ERROR] Giving up on zip={zipcode} after {MAX_RETRIES} attempts.",
                  file=sys.stderr)
            url = None

    return events


def event_location_str(loc: dict) -> str:
    parts = []
    if loc.get("venue"):
        parts.append(loc["venue"])
    addr = ", ".join(a for a in loc.get("address_lines", []) if a)
    if addr:
        parts.append(addr)
    city_state = ", ".join(filter(None, [loc.get("locality"), loc.get("region")]))
    if city_state:
        parts.append(city_state)
    if loc.get("postal_code"):
        parts.append(loc["postal_code"])
    return ", ".join(parts) if parts else "Virtual / TBD"


def expand_event(event: dict, prop: dict) -> list[dict]:
    """
    Expand a Mobilize event into one row per timeslot, filtered to
    SEARCH_RADIUS_MI from the property.  Returns [] if outside radius
    or location is unavailable.
    """
    loc = event.get("location") or {}
    elat = loc.get("lat")
    elon = loc.get("lon")
    if elat is None or elon is None:
        return []   # virtual / no coords

    dist = haversine(prop["lat"], prop["lon"], elat, elon)
    if dist > SEARCH_RADIUS_MI:
        return []

    title      = (event.get("title") or "").strip()
    etype      = (event.get("event_type") or "OTHER").replace("_", " ").title()
    browser_url = event.get("browser_url", "")
    loc_str    = event_location_str(loc)

    rows = []
    for ts in event.get("timeslots") or []:
        start_unix = ts.get("start_date")
        if not start_unix:
            continue
        dt_utc = datetime.fromtimestamp(start_unix, tz=timezone.utc)
        rows.append({
            "property_id":    prop["id"],
            "property_name":  prop["name"],
            "property_addr":  prop["address"],
            "event_title":    title,
            "event_type":     etype,
            "event_date":     dt_utc.strftime("%b %d, %Y  %I:%M %p UTC"),
            "event_dt_sort":  dt_utc,
            "event_location": loc_str,
            "distance_mi":    round(dist, 2),
            "event_url":      browser_url,
        })
    return rows


def is_no_kings(event: dict) -> bool:
    text = (
        (event.get("title") or "") + " " + (event.get("description") or "")
    ).lower()
    return any(kw in text for kw in NO_KINGS_KEYWORDS)


# ── Geographic clustering ──────────────────────────────────────────────────────

def cluster_properties(properties: list[dict]) -> list[list[dict]]:
    """
    Group properties into geographic clusters so nearby properties share a
    single API query instead of one query per zip code.

    Algorithm: greedy single-linkage — seed each cluster from the first
    unassigned property and absorb any unassigned property within
    CLUSTER_RADIUS_MI of that seed.  Returns a list of clusters, where
    each cluster is a list of property dicts.
    """
    assigned: set[int] = set()
    clusters: list[list[dict]] = []

    for i, seed in enumerate(properties):
        if i in assigned:
            continue
        cluster = [seed]
        assigned.add(i)
        for j, candidate in enumerate(properties):
            if j in assigned:
                continue
            if haversine(seed["lat"], seed["lon"],
                         candidate["lat"], candidate["lon"]) <= CLUSTER_RADIUS_MI:
                cluster.append(candidate)
                assigned.add(j)
        clusters.append(cluster)

    return clusters


def cluster_query_params(cluster: list[dict]) -> tuple[str, int]:
    """
    Return (zip_code, query_radius_miles) for a cluster.

    zip  – taken from the property nearest to the cluster's geographic centroid
    radius – distance from centroid to farthest member + SEARCH_RADIUS_MI + 1 mile buffer
    """
    centroid_lat = sum(p["lat"] for p in cluster) / len(cluster)
    centroid_lon = sum(p["lon"] for p in cluster) / len(cluster)

    nearest = min(cluster,
                  key=lambda p: haversine(centroid_lat, centroid_lon, p["lat"], p["lon"]))
    max_dist = max(haversine(centroid_lat, centroid_lon, p["lat"], p["lon"])
                   for p in cluster)
    query_radius = math.ceil(max_dist + SEARCH_RADIUS_MI + 1)

    return nearest["zip"], query_radius


# ── Core collection logic ──────────────────────────────────────────────────────

def collect_events(
    properties: list[dict],
) -> tuple[list[dict], list[dict]]:
    """
    Cluster properties geographically, then fire one Mobilize.us API query
    per cluster instead of one per zip code.  Typical reduction: ~200 calls
    down to ~40-60 calls.

    Returns:
      general_rows  – protest-type events within DAYS_GENERAL days
      no_kings_rows – No Kings events within DAYS_NO_KINGS days
    Both lists are sorted by event start time.
    """
    now_ts  = int(datetime.now(tz=timezone.utc).timestamp())
    end_30d = now_ts + DAYS_NO_KINGS * 86_400
    end_3d  = now_ts + DAYS_GENERAL  * 86_400

    clusters = cluster_properties(properties)
    print(f"  Clustered {len(properties)} properties into {len(clusters)} geographic groups.\n")

    general_rows:  list[dict] = []
    no_kings_rows: list[dict] = []
    seen_general:  set[tuple] = set()
    seen_no_kings: set[tuple] = set()

    def process_events(events: list[dict], cluster: list[dict]) -> None:
        for ev in events:
            eid       = ev.get("id")
            etype_raw = (ev.get("event_type") or "OTHER").upper()
            no_kings  = is_no_kings(ev)
            for prop in cluster:
                rows = expand_event(ev, prop)
                if not rows:
                    continue
                for row in rows:
                    ts_key = (eid, prop["id"], row["event_dt_sort"])
                    if (etype_raw not in EXCLUDE_TYPES
                            and row["event_dt_sort"].timestamp() <= end_3d
                            and ts_key not in seen_general):
                        seen_general.add(ts_key)
                        general_rows.append(row)
                    if no_kings and ts_key not in seen_no_kings:
                        seen_no_kings.add(ts_key)
                        no_kings_rows.append(row)

    # ── First pass ────────────────────────────────────────────────────────────
    retry_queue: list[list[dict]] = []   # clusters that hit 429 on first pass

    for idx, cluster in enumerate(clusters, 1):
        zip_code, query_radius = cluster_query_params(cluster)
        names = ", ".join(p["name"] for p in cluster[:3])
        suffix = f" +{len(cluster)-3} more" if len(cluster) > 3 else ""
        print(f"  [{idx:>3}/{len(clusters)}] zip={zip_code}  r={query_radius}mi  "
              f"({names}{suffix})")

        try:
            events = fetch_events_for_zip(zip_code, query_radius, now_ts, end_30d)
            print(f"           → {len(events)} event(s) returned")
            process_events(events, cluster)
        except RateLimitError:
            print(f"    [429] Rate-limited — queued for retry pass.", file=sys.stderr)
            retry_queue.append(cluster)

    # ── Retry pass (60 s head-start, then 5 s between attempts) ─────────────
    if retry_queue:
        print(f"\n  {len(retry_queue)} cluster(s) rate-limited. "
              f"Waiting 60s before retry pass …")
        time.sleep(60)

        still_failed: list[str] = []
        for attempt_num in range(1, MAX_RETRIES + 1):
            next_retry: list[list[dict]] = []

            for cluster in retry_queue:
                zip_code, query_radius = cluster_query_params(cluster)
                names = cluster[0]["name"]
                print(f"  [retry {attempt_num}/{MAX_RETRIES}] zip={zip_code}  ({names})")
                try:
                    events = fetch_events_for_zip(zip_code, query_radius, now_ts, end_30d)
                    print(f"           → {len(events)} event(s) returned")
                    process_events(events, cluster)
                    time.sleep(5)
                except RateLimitError:
                    next_retry.append(cluster)

            retry_queue = next_retry
            if not retry_queue:
                print("  All retries succeeded.")
                break

            if attempt_num < MAX_RETRIES:
                wait = 60 * attempt_num
                print(f"  {len(retry_queue)} still rate-limited. "
                      f"Waiting {wait}s before next retry …")
                time.sleep(wait)

        if retry_queue:
            for cluster in retry_queue:
                zip_code, _ = cluster_query_params(cluster)
                still_failed.append(zip_code)
                for prop in cluster:
                    print(f"  [FAILED] {prop['name']} (zip={prop['zip']}) — "
                          f"no data after {MAX_RETRIES} retries.", file=sys.stderr)

    general_rows.sort(key=lambda r: r["event_dt_sort"])
    no_kings_rows.sort(key=lambda r: r["event_dt_sort"])
    return general_rows, no_kings_rows


# ── Excel helpers ──────────────────────────────────────────────────────────────

def _cell(ws, row: int, col: int, value,
          font=None, fill=None, align=None, border=None, hyperlink=None):
    c = ws.cell(row=row, column=col, value=value)
    if font:    c.font      = font
    if fill:    c.fill      = fill
    if align:   c.alignment = align
    if border:  c.border    = border
    if hyperlink:
        c.hyperlink = hyperlink
    return c


def write_data_sheet(ws, rows: list[dict],
                     title_fill: PatternFill,
                     col_hdr_fill: PatternFill,
                     alt_fill: PatternFill,
                     sheet_title: str,
                     subtitle: str) -> None:

    ws.sheet_view.showGridLines = False
    num_cols = len(COLUMNS)
    last_col = get_column_letter(num_cols)

    # ── Row 1: Title ──
    ws.merge_cells(f"A1:{last_col}1")
    _cell(ws, 1, 1, sheet_title,
          font   = Font(name="Calibri", bold=True, size=16, color="FFFFFF"),
          fill   = title_fill,
          align  = Alignment(horizontal="center", vertical="center"),
          border = THIN)
    ws.row_dimensions[1].height = 34

    # ── Row 2: Subtitle ──
    ws.merge_cells(f"A2:{last_col}2")
    _cell(ws, 2, 1, subtitle,
          font   = Font(name="Calibri", italic=True, size=10, color="FFFFFF"),
          fill   = title_fill,
          align  = Alignment(horizontal="center", vertical="center"),
          border = THIN)
    ws.row_dimensions[2].height = 16

    # ── Row 3: Column headers ──
    for col_idx, (label, _, _) in enumerate(COLUMNS, 1):
        _cell(ws, 3, col_idx, label,
              font   = Font(name="Calibri", bold=True, size=11, color="FFFFFF"),
              fill   = col_hdr_fill,
              align  = Alignment(horizontal="center", vertical="center", wrap_text=True),
              border = THIN)
    ws.row_dimensions[3].height = 22
    ws.freeze_panes = "A4"

    # ── Data rows ──
    if not rows:
        ws.merge_cells(f"A4:{last_col}4")
        _cell(ws, 4, 1, "No events found for this reporting period.",
              font  = Font(name="Calibri", italic=True, size=11, color="808080"),
              align = Alignment(horizontal="center", vertical="center"))
        ws.row_dimensions[4].height = 24
    else:
        for row_idx, row in enumerate(rows, start=4):
            is_even = (row_idx % 2 == 0)
            row_fill = alt_fill if is_even else WHITE_FILL
            ws.row_dimensions[row_idx].height = 26

            for col_idx, (_, key, _) in enumerate(COLUMNS, 1):
                val = row.get(key, "")

                is_url  = (key == "event_url")
                is_wrap = (key in ("event_title", "event_location", "event_url"))
                is_cent = (key == "distance_mi")

                font = Font(name="Calibri", size=10,
                            color="0563C1" if is_url and val else "000000",
                            underline="single" if is_url and val else None)
                align = Alignment(
                    vertical   = "center",
                    horizontal = "center" if is_cent else "left",
                    wrap_text  = is_wrap,
                )
                c = _cell(ws, row_idx, col_idx, val,
                          font=font, fill=row_fill, align=align, border=THIN)
                if is_url and val:
                    c.hyperlink = val

    # ── Column widths ──
    for col_idx, (_, _, width) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def build_summary_sheet(ws, general_rows: list[dict],
                         no_kings_rows: list[dict]) -> None:
    ws.sheet_view.showGridLines = False
    now_str = datetime.now().strftime("%B %d, %Y at %I:%M %p")

    # Title
    ws.merge_cells("A1:C1")
    _cell(ws, 1, 1, "PROTEST TRACKER — EXECUTIVE SUMMARY",
          font   = Font(name="Calibri", bold=True, size=18, color="FFFFFF"),
          fill   = HDR_NAVY,
          align  = Alignment(horizontal="center", vertical="center"),
          border = THIN)
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:C2")
    _cell(ws, 2, 1, f"Report generated: {now_str}",
          font   = Font(name="Calibri", italic=True, size=11, color="FFFFFF"),
          fill   = HDR_NAVY,
          align  = Alignment(horizontal="center", vertical="center"),
          border = THIN)
    ws.row_dimensions[2].height = 20

    # Section header
    _cell(ws, 3, 1, "Metric",
          font  = Font(name="Calibri", bold=True, size=11, color="FFFFFF"),
          fill  = HDR_BLUE,
          align = Alignment(horizontal="center", vertical="center"),
          border = THIN)
    _cell(ws, 3, 2, "Value",
          font  = Font(name="Calibri", bold=True, size=11, color="FFFFFF"),
          fill  = HDR_BLUE,
          align = Alignment(horizontal="center", vertical="center"),
          border = THIN)
    ws.merge_cells("B3:C3")
    ws.row_dimensions[3].height = 20

    # Unique property IDs with hits
    gen_props  = {r["property_id"] for r in general_rows}
    nk_props   = {r["property_id"] for r in no_kings_rows}
    gen_events = {(r["event_title"], r["event_date"]) for r in general_rows}
    nk_events  = {(r["event_title"], r["event_date"]) for r in no_kings_rows}

    end_3d  = (datetime.now() + timedelta(days=DAYS_GENERAL)).strftime("%b %d, %Y")
    end_30d = (datetime.now() + timedelta(days=DAYS_NO_KINGS)).strftime("%b %d, %Y")

    stats = [
        ("Properties monitored",                  "205"),
        ("Search radius",                          f"{SEARCH_RADIUS_MI} miles"),
        ("─── 3-Day Window ───",                  ""),
        (f"  Window end date",                     end_3d),
        (f"  Properties with events",              str(len(gen_props))),
        (f"  Unique events found",                 str(len(gen_events))),
        (f"  Total property-event pairs",          str(len(general_rows))),
        ("─── No Kings 30-Day Window ───",        ""),
        (f"  Window end date",                     end_30d),
        (f"  Properties with No Kings events",     str(len(nk_props))),
        (f"  Unique No Kings events found",        str(len(nk_events))),
        (f"  Total property-event pairs",          str(len(no_kings_rows))),
        ("─── Data Sources ───",                  ""),
        ("  Primary source",                       "Mobilize.us Public API"),
        ("  Keywords (No Kings filter)",           ", ".join(NO_KINGS_KEYWORDS)),
    ]

    for i, (metric, value) in enumerate(stats, start=4):
        is_even   = (i % 2 == 0)
        is_hdr    = metric.startswith("───")
        row_fill  = (PatternFill("solid", fgColor="D9E2F3") if is_hdr
                     else (ALT_BLUE if is_even else WHITE_FILL))
        bold      = is_hdr

        _cell(ws, i, 1, metric,
              font  = Font(name="Calibri", bold=bold, size=11,
                           color="1F3864" if is_hdr else "000000"),
              fill  = row_fill,
              align = Alignment(vertical="center"),
              border = THIN)
        ws.merge_cells(f"B{i}:C{i}")
        _cell(ws, i, 2, value,
              font  = Font(name="Calibri", bold=bold, size=11),
              fill  = row_fill,
              align = Alignment(vertical="center"),
              border = THIN)
        ws.row_dimensions[i].height = 22

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 5


# ── Build workbook ─────────────────────────────────────────────────────────────

def build_excel(general_rows: list[dict], no_kings_rows: list[dict],
                output_path: str) -> None:
    wb = Workbook()

    # Summary sheet
    ws_sum = wb.active
    ws_sum.title = "Summary"
    build_summary_sheet(ws_sum, general_rows, no_kings_rows)

    # 3-day events sheet
    ws_gen = wb.create_sheet("3-Day Events")
    now_str = datetime.now().strftime("%B %d, %Y")
    end_3d  = (datetime.now() + timedelta(days=DAYS_GENERAL)).strftime("%B %d, %Y")
    write_data_sheet(
        ws_gen, general_rows,
        title_fill   = HDR_NAVY,
        col_hdr_fill = HDR_BLUE,
        alt_fill     = ALT_BLUE,
        sheet_title  = "Protest / Rally Events Near Tracked Properties — 3-Day Window",
        subtitle     = (f"{now_str}  through  {end_3d}"
                        f"  •  Within {SEARCH_RADIUS_MI} miles of property"
                        f"  •  Source: Mobilize.us"),
    )

    # No Kings sheet
    ws_nk = wb.create_sheet("No Kings Events (30-Day)")
    end_30d = (datetime.now() + timedelta(days=DAYS_NO_KINGS)).strftime("%B %d, %Y")
    write_data_sheet(
        ws_nk, no_kings_rows,
        title_fill   = HDR_PURPLE,
        col_hdr_fill = COL_PURPLE,
        alt_fill     = ALT_PURPLE,
        sheet_title  = "No Kings Protest Events — 30-Day Window",
        subtitle     = (f"{now_str}  through  {end_30d}"
                        f"  •  Within {SEARCH_RADIUS_MI} miles of property"
                        f"  •  Source: Mobilize.us"),
    )

    wb.save(output_path)
    print(f"\n  ✓ Excel workbook saved → {output_path}")


# ── Entry point ────────────────────────────────────────────────────────────────

def main() -> None:
    ap = argparse.ArgumentParser(
        description="Protest Tracker — generates an Excel report of events near tracked properties",
    )
    ap.add_argument(
        "--output", "-o",
        default="protest_tracker_report.xlsx",
        help="Output Excel file name (default: protest_tracker_report.xlsx)",
    )
    ap.add_argument(
        "--csv", "-c",
        default=PROPERTIES_CSV,
        help=f"Path to properties CSV (default: {PROPERTIES_CSV})",
    )
    ap.add_argument(
        "--api-key", "-k",
        default="",
        help="Mobilize.us API key (overrides MOBILIZE_API_KEY env var)",
    )
    args = ap.parse_args()

    # Allow CLI flag to override env var
    if args.api_key:
        global MOBILIZE_API_KEY
        MOBILIZE_API_KEY = args.api_key

    print("=" * 64)
    print("  PROTEST TRACKER")
    print("=" * 64)
    print(f"  Properties CSV : {args.csv}")
    print(f"  Output file    : {args.output}")
    print(f"  Authenticated  : {'Yes' if MOBILIZE_API_KEY else 'No (set MOBILIZE_API_KEY for better rate limits)'}")
    print(f"  Search radius  : {SEARCH_RADIUS_MI} miles")
    print(f"  Cluster radius : {CLUSTER_RADIUS_MI} miles (properties grouped for fewer API calls)")
    print(f"  General window : today + {DAYS_GENERAL} days")
    print(f"  No Kings window: today + {DAYS_NO_KINGS} days")
    print("=" * 64)

    properties = load_properties(args.csv)
    print(f"\nLoaded {len(properties)} properties.\n")
    print("Querying Mobilize.us public API …\n")

    general_rows, no_kings_rows = collect_events(properties)

    print(f"\n{'─'*64}")
    print(f"  3-Day events found        : {len(general_rows)} property-event rows")
    print(f"  No Kings events found     : {len(no_kings_rows)} property-event rows")
    print(f"{'─'*64}")

    print("\nBuilding Excel workbook …")
    build_excel(general_rows, no_kings_rows, args.output)
    print("Done.\n")


if __name__ == "__main__":
    main()
