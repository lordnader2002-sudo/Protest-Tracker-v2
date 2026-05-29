#!/usr/bin/env python3
"""Build detention_centers.csv — the anchor set for the ICE Detention protest stream.

This is a one-time / occasional helper. It is NOT run by the GitHub Actions
workflow; the resulting CSV is committed to the repo and read at runtime by
protest_tracker.py (which loads it with the same loader as properties.csv).

Sources (both fetched from GitHub raw CDN — no API key, no geocoder needed):

  1. Vera Institute "ICE Detention Trends" — FY2026 daily population file.
     Authoritative list of CURRENTLY-ACTIVE facilities (ICE facility code +
     name + state). We keep facilities with any population in FY2026.
       https://github.com/vera-institute/ice-detention-trends

  2. The Marshall Project "DHS Immigration Detention" — locations.csv.
     ICE facility code (DETLOC) -> street address + geocoded lat/lon
     (Google-geocoded, manually verified).
       https://github.com/themarshallproject/dhs_immigration_detention

The two share ICE's official facility code, so we join Vera's current set to
Marshall's coordinates. Facilities without a coordinate match are skipped (and
reported), since the runtime needs lat/lon to measure protest proximity.

Output columns match properties.csv exactly:
    property_id,name,address,postal_code,lat,lon

Usage:  python build_detention_csv.py [-o detention_centers.csv]
"""
from __future__ import annotations

import argparse
import csv
import io
import sys
import urllib.request

VERA_FY2026 = ("https://raw.githubusercontent.com/vera-institute/"
               "ice-detention-trends/main/facilities/by_fiscal_year/FY2026.csv")
MARSHALL_LOCATIONS = ("https://raw.githubusercontent.com/themarshallproject/"
                      "dhs_immigration_detention/master/locations.csv")


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": "ProtestTracker-detention-builder/1.0"})
    with urllib.request.urlopen(req, timeout=90) as resp:
        # Marshall's file uses latin-1 + bare \r line endings; decode permissively.
        return resp.read().decode("latin-1")


def parse_csv(text: str) -> list[dict]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return list(csv.DictReader(io.StringIO(text)))


def valid_coord(lat: float, lon: float) -> bool:
    if lat == 0 and lon == 0:
        return False
    # Continental US + AK/HI/PR/GU envelope.
    return 13.0 <= lat <= 72.0 and -180.0 <= lon <= -64.0


def title_case(name: str) -> str:
    """Title-case a SHOUTING facility name while keeping short tokens sensible."""
    small = {"of", "and", "the", "at", "in", "for"}
    out = []
    for i, w in enumerate(name.split()):
        lw = w.lower()
        if w.isupper() and len(w) <= 3 and lw not in small and not w.isalpha():
            out.append(w)            # keep acronyms / codes like USP, ICE
        elif lw in small and i != 0:
            out.append(lw)
        else:
            out.append(w.capitalize())
    return " ".join(out)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--output", default="detention_centers.csv")
    args = ap.parse_args()

    print("Fetching Vera FY2026 active-facility list …")
    vera = parse_csv(fetch(VERA_FY2026))
    print("Fetching Marshall Project geocoded locations …")
    marshall = parse_csv(fetch(MARSHALL_LOCATIONS))

    # Marshall: code -> location record
    mp = {r["DETLOC"].strip(): r for r in marshall if r.get("DETLOC", "").strip()}

    # Vera: currently-active facilities (any population recorded in FY2026)
    active: dict[str, dict] = {}
    for r in vera:
        code = (r.get("detention_facility_code") or "").strip()
        if not code:
            continue
        try:
            pop = float(r.get("midnight_pop") or 0) + float(r.get("daily_pop") or 0)
        except ValueError:
            pop = 0.0
        a = active.setdefault(code, {"name": r.get("detention_facility_name", ""),
                                     "state": r.get("state", ""), "maxpop": 0.0})
        a["maxpop"] = max(a["maxpop"], pop)
    active = {c: v for c, v in active.items() if v["maxpop"] > 0}
    print(f"  Vera: {len(active)} currently-active facilities (population > 0 in FY2026)")

    rows, skipped = [], []
    for code, info in sorted(active.items()):
        loc = mp.get(code)
        if not loc:
            skipped.append((code, info["name"], "no coordinate match"))
            continue
        try:
            lat = float(loc.get("Lat") or "")
            lon = float(loc.get("Lng") or "")
        except ValueError:
            skipped.append((code, info["name"], "blank coords"))
            continue
        if not valid_coord(lat, lon):
            skipped.append((code, info["name"], f"implausible coords {lat},{lon}"))
            continue

        name = title_case((loc.get("Name") or info["name"]).strip())
        street = (loc.get("Address") or "").strip()
        city = (loc.get("City") or "").strip()
        state = (loc.get("State") or info["state"] or "").strip()
        zip_ = (loc.get("Zip") or "").strip().zfill(5)
        address = " ".join(" ".join(p for p in [street, city, state, zip_] if p).split())
        rows.append({"property_id": code, "name": name, "address": address,
                     "postal_code": zip_, "lat": round(lat, 6), "lon": round(lon, 6)})

    with open(args.output, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["property_id", "name", "address",
                                           "postal_code", "lat", "lon"])
        w.writeheader()
        w.writerows(rows)

    print(f"\n  ✓ Wrote {len(rows)} detention centers → {args.output}")
    print(f"  ⚠ Skipped {len(skipped)} active facilities with no usable coordinates.")
    for code, name, why in skipped[:10]:
        print(f"      - {code:10} {name[:34]:34} ({why})", file=sys.stderr)
    if len(skipped) > 10:
        print(f"      … and {len(skipped) - 10} more", file=sys.stderr)


if __name__ == "__main__":
    main()
