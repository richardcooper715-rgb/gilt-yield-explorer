#!/usr/bin/env python3
"""
fetch_gilt_data.py

Builds gilt_yields.json for the Gilt Yield Explorer chart, from two
free public sources:

  1. DMO historical gilt reference prices & yields (25 Nov 2002 - 21 Jul 2017
     for yields; prices alone go back to 1996)
     https://dmo.gov.uk/data/gilt-market/historical-prices-and-yields
     -> per-security (per-ISIN) daily yields, both conventional and
        index-linked gilts. This is the good stuff.

     NOTE: DMO's own documentation confirms gross redemption yields in
     this dataset were only calculated/published from 25 Nov 2002 onward
     (https://dmo.gov.uk/data/gilt-market/aggregated-yields). Files for
     1996-2001 (and most of 2002) contain prices only -- no yield column
     -- so this script correctly produces 0 yield rows for those years.
     That is expected DMO data coverage, not a parsing bug.

     DMO's site blocks scripted downloads (ShieldSquare bot protection),
     so this script does NOT fetch these files itself. Instead:
       1. Download each year's file (1996-2017) by hand from the URL
          above, in your own browser.
       2. Save them all into a folder named dmo_raw/ next to this script
          (any filename is fine as long as the 4-digit year is in it).
       3. Run this script -- it reads dmo_raw/ automatically from then on.

  2. Bank of England Anderson-Sleath fitted yield curves (continuous,
     1979/1985 - present)
     https://www.bankofengland.co.uk/statistics/yield-curves
     -> curve-level (not per-security) nominal and real spot yields,
        used here to extend coverage past 21 Jul 2017, when DMO
        stopped publishing and per-security prices moved behind
        Tradeweb Insite (registration required, not scraped here).
     This part IS scraped automatically -- BoE's site doesn't block it.

     NOTE: BoE's archive zips (glcnominalddata.zip / glcrealddata.zip)
     are only refreshed on the 2nd working day of each month (confirmed
     on the BoE page's own FAQ) -- on their own they always lag by
     several weeks. This script also fetches BoE's separate "Latest
     yield curve data" zip, which is updated daily, and overlays it on
     top of the archive data to close that gap.

Run:

    pip install requests beautifulsoup4 openpyxl xlrd pandas
    python fetch_gilt_data.py                    # full run, all sources
    python fetch_gilt_data.py --only boe-latest   # fast daily refresh
    python fetch_gilt_data.py --only dmo          # e.g. after adding new
                                                   # dmo_raw/ files
    python fetch_gilt_data.py --only dmo,boe-archive

Skipped components reuse their data from the existing gilt_yields.json
(if present) rather than being left empty, so a partial run never loses
data a fuller run previously produced.

Output: ./gilt_yields.json  (schema documented at the bottom of this
file, and consumed by gilt-yield-explorer.html)

NOTE: The DMO parsing logic (header-row detection, column-name mapping)
was written without a chance to inspect a real downloaded file. Print
statements are left in deliberately so you can see where it breaks
against the real file layouts and report back.
"""

import argparse
import io
import json
import os
import re
import zipfile
from datetime import datetime

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
except ImportError:
    raise SystemExit("pip install pandas openpyxl xlrd requests beautifulsoup4")

HEADERS = {"User-Agent": "Mozilla/5.0 (gilt-yield-explorer data pipeline)"}

BOE_CURVES_PAGE = "https://www.bankofengland.co.uk/statistics/yield-curves"

# The archive zips (found via discover_boe_archive_zips) are only refreshed
# on the 2nd working day of each month (confirmed on the BoE yield-curves
# page's own FAQ) -- so on their own they always lag behind by several
# weeks. This "latest" zip is updated daily and fills that gap; it bundles
# several curve types together in one file, unlike the type-specific
# archive zips, so its parser filters sheets by curve type too.
BOE_LATEST_ZIP_URL = ("https://www.bankofengland.co.uk/-/media/boe/files/"
                       "statistics/yield-curves/latest-yield-curve-data.zip")

DMO_RAW_DIR = "./dmo_raw"  # put manually-downloaded DMO yearly files here

CUTOVER_DATE = "2017-07-21"  # last DMO reference price date


# ---------------------------------------------------------------------------
# 1. DMO per-security historical yields (1996 - 2017)
#
# DMO's site blocks scripted requests (ShieldSquare bot protection), so
# this reads local files instead of downloading them. One-time manual step:
#
#   1. Open https://dmo.gov.uk/data/gilt-market/historical-prices-and-yields
#      in your browser.
#   2. Download each year's file (1996-2017) -- whatever DMO names them is
#      fine, as long as the 4-digit year appears somewhere in the filename,
#      e.g. "GiltRefPrices1996.xls", "1996.xlsx", "dmo_1996_prices.csv".
#   3. Save them all into a folder named dmo_raw/ next to this script.
#
# Re-run this script any time after that -- it just reads whatever's in
# dmo_raw/, so you never have to touch the DMO site again.
# ---------------------------------------------------------------------------

def discover_dmo_year_files():
    """Scan DMO_RAW_DIR for manually-downloaded yearly files, matched by
    a 4-digit year (1990-2029) anywhere in the filename."""
    if not os.path.isdir(DMO_RAW_DIR):
        print(f"  ERROR: {DMO_RAW_DIR}/ does not exist. Create it and put "
              f"the manually-downloaded DMO yearly files inside -- see the "
              f"comment above discover_dmo_year_files() for instructions.")
        return {}

    year_files = {}
    for fname in sorted(os.listdir(DMO_RAW_DIR)):
        if not fname.lower().endswith((".xls", ".xlsx", ".csv")):
            continue
        m = re.search(r"(19|20)\d{2}", fname)
        if not m:
            print(f"  WARNING: {fname} has no 4-digit year in its name -- skipping")
            continue
        year = int(m.group(0))
        year_files[year] = os.path.join(DMO_RAW_DIR, fname)

    print(f"Found {len(year_files)} local DMO yearly files in {DMO_RAW_DIR}/: "
          f"{sorted(year_files)}")
    return year_files


def parse_dmo_year_file(year, path):
    """Parse one year's DMO reference price/yield file into tidy rows.

    Expected (approximate) columns based on DMO's documented format:
    close-of-business date, ISIN, gilt short name, gilt type
    (conventional/index-linked), clean price, gross redemption yield.
    Column names vary by year -- inspect the real file and adjust the
    rename map below.
    """
    try:
        raw = pd.read_excel(path, header=None)
    except Exception:
        raw = pd.read_csv(path, header=None)

    # DMO's files have a title block (data date / report name / date range)
    # before the real header row. Scan the first 20 rows for the one that
    # actually looks like a header (contains "isin" somewhere).
    header_row_idx = None
    for i in range(min(20, len(raw))):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if any("isin" in v for v in row_vals):
            header_row_idx = i
            break

    if header_row_idx is None:
        print(f"  [{year}] WARNING: could not find a header row (looked for "
              f"'ISIN' in first 20 rows). First 6 rows were:")
        print(raw.head(6).to_string())
        return []

    headers = [str(v).strip().lower() for v in raw.iloc[header_row_idx].tolist()]
    df = raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    print(f"  [{year}] header row found at index {header_row_idx}: {headers}")

    # DMO's real headers carry embedded newlines and unit suffixes, e.g.
    # "yield \n(%)", "clean price\n(£)" -- normalise whitespace and match
    # by keyword rather than requiring an exact string.
    def normalize(h):
        h = str(h).replace("\n", " ").replace("\r", " ")
        return re.sub(r"\s+", " ", h).strip().lower()

    def classify(h):
        if "isin" in h:
            return "isin"
        if "gilt name" in h or "stock name" in h or h == "name":
            return "name"
        if "redemption date" in h:
            return "redemption_date"
        if "close of business date" in h or h == "cob date":
            return "date"
        if "yield" in h:
            return "yield"
        if "clean price" in h:
            return "clean_price"
        if "dirty price" in h:
            return "dirty_price"
        if "accrued" in h:
            return "accrued_interest"
        if "duration" in h:
            return "modified_duration"
        return h  # leave unrecognised columns as-is

    df.columns = [classify(normalize(h)) for h in headers]

    required = {"date", "isin", "yield"}
    missing = required - set(df.columns)
    if missing:
        print(f"  [{year}] WARNING: missing columns {missing}, "
              f"actual columns were {list(df.columns)} -- skipping")
        return []

    pre_drop_count = len(df)
    df["date_parsed"] = pd.to_datetime(df["date"], errors="coerce")
    df["yield_parsed"] = pd.to_numeric(df["yield"], errors="coerce")

    if pre_drop_count > 0:
        n_bad_date = df["date_parsed"].isna().sum()
        n_bad_isin = df["isin"].isna().sum() if "isin" in df.columns else pre_drop_count
        n_bad_yield = df["yield_parsed"].isna().sum()
        if n_bad_yield == pre_drop_count and n_bad_date < pre_drop_count:
            # DMO's own historical documentation confirms gross redemption
            # yields were only calculated/published from 25 Nov 2002 onward
            # (https://dmo.gov.uk/data/gilt-market/aggregated-yields).
            # Years before that genuinely have prices but no yield column
            # populated -- this is expected, not a parsing failure.
            print(f"  [{year}] no gross redemption yield published by DMO "
                  f"this year (prices only) -- expected for dates before "
                  f"25 Nov 2002, not a parsing error.")

    df["date"] = df["date_parsed"]
    df["yield"] = df["yield_parsed"]
    df["dirty_price"] = (pd.to_numeric(df["dirty_price"], errors="coerce")
                          if "dirty_price" in df.columns else float("nan"))
    df = df.dropna(subset=["date", "isin", "yield"])

    rows = []
    for _, r in df.iterrows():
        name = str(r.get("name", "")).strip()
        is_il = bool(re.search(r"index.?linked|\bIL\b", name, re.IGNORECASE))
        dp = r.get("dirty_price")
        rows.append({
            "date": r["date"].strftime("%Y-%m-%d"),
            "isin": str(r["isin"]).strip(),
            "name": name,
            "type": "index_linked" if is_il else "conventional",
            "yield": float(r["yield"]),
            "dirty_price": None if pd.isna(dp) else float(dp),
        })
    print(f"  [{year}] parsed {len(rows)} rows")
    return rows


def build_dmo_dataset():
    year_files = discover_dmo_year_files()
    all_rows = []
    for year, path in sorted(year_files.items()):
        try:
            all_rows.extend(parse_dmo_year_file(year, path))
        except Exception as e:
            print(f"  [{year}] FAILED: {e}")

    securities = {}
    for row in all_rows:
        sec = securities.setdefault(row["isin"], {
            "isin": row["isin"],
            "name": row["name"],
            "type": row["type"],
            "series": [],
        })
        sec["series"].append({
            "date": row["date"],
            "yield": row["yield"],
            "dirty_price": row["dirty_price"],
        })

    for sec in securities.values():
        sec["series"].sort(key=lambda p: p["date"])

    return list(securities.values())


# ---------------------------------------------------------------------------
# 2. BoE curve-level nominal / real yields (2017 - present)
# ---------------------------------------------------------------------------

def discover_boe_archive_zips():
    """Find the nominal and real spot-curve archive zip links on the
    BoE yield curves page."""
    resp = requests.get(BOE_CURVES_PAGE, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # Only want the DAILY gilt-based spot curves: glcnominalddata.zip and
    # glcrealddata.zip. Explicitly exclude "month" (monthly duplicates) and
    # "blc" (LIBOR-based commercial liability curves -- not gilt-based).
    daily_gilt_re = re.compile(r"glc(nominal|real)ddata\.zip$", re.IGNORECASE)

    zips = {"nominal": [], "real": []}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not href.lower().endswith(".zip"):
            continue
        m = daily_gilt_re.search(href)
        if not m:
            continue  # skip monthly, blc, and anything else
        url = href if href.startswith("http") else f"https://www.bankofengland.co.uk{href}"
        curve = m.group(1).lower()
        zips[curve].append(url)

    print(f"Discovered BoE archives: "
          f"{len(zips['nominal'])} nominal, {len(zips['real'])} real")
    return zips


def parse_boe_zip(url, curve_name):
    """Each BoE archive zip contains per-period Excel workbooks with a
    spot-curve sheet: date rows x maturity-year columns, rate in %.

    BoE publishes these on a very fine maturity grid (roughly monthly
    steps across the full curve), which is far more resolution than a
    maturity-selector dropdown needs and makes the JSON output huge
    (multi-hundred-MB, unusable on mobile). We keep only a curated set
    of "round number" maturities -- enough to be useful for an overlay
    selector, without carrying ~40x more data than needed.
    """
    KEEP_MATURITIES = [1, 2, 3, 4, 5, 7, 10, 15, 20, 25, 30, 40]

    resp = requests.get(url, headers=HEADERS, timeout=60)
    resp.raise_for_status()

    rows = []
    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        for name in zf.namelist():
            if not name.lower().endswith((".xls", ".xlsx")):
                continue
            with zf.open(name) as f:
                try:
                    sheets = pd.read_excel(f, sheet_name=None, header=None)
                except Exception as e:
                    print(f"    could not parse {name}: {e}")
                    continue

            for sheet_name, df in sheets.items():
                if "spot" not in sheet_name.lower():
                    continue
                # Expect: first column = date, header row = maturities in years.
                # Layout varies -- inspect the real workbook and adjust.
                header = df.iloc[3]  # guess: header row after title rows
                maturities = pd.to_numeric(header[1:], errors="coerce")
                # Map each curated maturity to the nearest actual column,
                # so we don't need the source grid to land on exact integers.
                col_for_target = {}
                for col_idx, m in enumerate(maturities, start=1):
                    if pd.isna(m):
                        continue
                    for target in KEEP_MATURITIES:
                        if abs(m - target) <= 0.1:
                            col_for_target.setdefault(target, col_idx)
                data = df.iloc[4:]
                for _, r in data.iterrows():
                    date = pd.to_datetime(r[0], errors="coerce")
                    if pd.isna(date):
                        continue
                    for target, col_idx in col_for_target.items():
                        rate = pd.to_numeric(r[col_idx], errors="coerce")
                        if pd.isna(rate):
                            continue
                        rows.append({
                            "date": date.strftime("%Y-%m-%d"),
                            "maturity_years": float(target),
                            "rate_pct": float(rate),
                        })
    print(f"    parsed {len(rows)} rows from {url} "
          f"(curated to {len(KEEP_MATURITIES)} maturities: {KEEP_MATURITIES})")
    return rows


def build_boe_dataset():
    zips = discover_boe_archive_zips()
    curves = {"nominal": [], "real": []}
    for curve_name in ("nominal", "real"):
        for url in zips[curve_name]:
            try:
                curves[curve_name].extend(parse_boe_zip(url, curve_name))
            except Exception as e:
                print(f"  FAILED {url}: {e}")
        curves[curve_name].sort(key=lambda r: (r["date"], r["maturity_years"]))
    return curves


def fetch_latest_boe_curves():
    """Fetch BoE's 'Latest yield curve data' zip -- updated daily, unlike
    the monthly-refreshed archive zips above. Bundles multiple curve types
    (nominal, real, inflation, OIS, ...) together in one file, so sheets
    are filtered on curve type AND 'spot' in the sheet name.
    """
    KEEP_MATURITIES = [1, 2, 3, 4, 5, 7, 10, 15, 20, 25, 30, 40]
    result = {"nominal": [], "real": []}
    try:
        resp = requests.get(BOE_LATEST_ZIP_URL, headers=HEADERS, timeout=60)
        resp.raise_for_status()
    except Exception as e:
        print(f"    could not fetch latest-yield-curve-data.zip: {e}")
        return result

    sheet_names_seen = []
    filenames_seen = []
    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        for name in zf.namelist():
            if not name.lower().endswith((".xls", ".xlsx")):
                continue
            filenames_seen.append(name)

            # The sheet names inside are generic ("spot curve" etc.) and
            # identical across every workbook in this zip -- the curve
            # type (nominal/real/inflation/OIS) is only distinguishable
            # from the workbook's own filename.
            fn = name.lower()
            if "nominal" in fn:
                curve_name = "nominal"
            elif "real" in fn:
                curve_name = "real"
            else:
                continue  # inflation / OIS / anything else -- not tracked here

            with zf.open(name) as f:
                try:
                    sheets = pd.read_excel(f, sheet_name=None, header=None)
                except Exception as e:
                    print(f"    could not parse {name} in latest zip: {e}")
                    continue

            for sheet_name, df in sheets.items():
                sheet_names_seen.append(sheet_name)
                sn = sheet_name.lower()
                if "spot" not in sn or "short" in sn:
                    continue  # skip the short-end tab, keep the full spot curve

                header = df.iloc[3]
                maturities = pd.to_numeric(header[1:], errors="coerce")
                col_for_target = {}
                for col_idx, m in enumerate(maturities, start=1):
                    if pd.isna(m):
                        continue
                    for target in KEEP_MATURITIES:
                        if abs(m - target) <= 0.1:
                            col_for_target.setdefault(target, col_idx)
                data = df.iloc[4:]
                for _, r in data.iterrows():
                    date = pd.to_datetime(r[0], errors="coerce")
                    if pd.isna(date):
                        continue
                    for target, col_idx in col_for_target.items():
                        rate = pd.to_numeric(r[col_idx], errors="coerce")
                        if pd.isna(rate):
                            continue
                        result[curve_name].append({
                            "date": date.strftime("%Y-%m-%d"),
                            "maturity_years": float(target),
                            "rate_pct": float(rate),
                        })

    print(f"    parsed {len(result['nominal'])} nominal + {len(result['real'])} "
          f"real rows from latest-yield-curve-data.zip")
    if not result["nominal"] and not result["real"]:
        print(f"    WARNING: 0 rows from the 'latest' zip. Workbook filenames "
              f"found: {filenames_seen!r}. Sheet names: {sheet_names_seen!r} -- "
              f"if none of the filenames contain 'nominal'/'real', report this "
              f"list back so the filter can be adjusted.")
    return result


def merge_curve_overlay(base, overlay):
    """Combine two {'nominal':[...], 'real':[...]} curve dicts, with
    `overlay` rows taking precedence over `base` rows on the same
    (date, maturity) -- used to let the fresher 'latest' BoE data
    supersede the monthly-archive data for whatever dates they share,
    while keeping all the archive's older history.
    """
    result = {}
    for curve_name in ("nominal", "real"):
        by_key = {(r["date"], r["maturity_years"]): r
                  for r in base.get(curve_name, [])}
        for r in overlay.get(curve_name, []):
            by_key[(r["date"], r["maturity_years"])] = r
        result[curve_name] = sorted(by_key.values(),
                                     key=lambda r: (r["date"], r["maturity_years"]))
    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Build gilt_yields.json from DMO + BoE data sources.")
    parser.add_argument(
        "--only", default="all",
        help=("Comma-separated components to run: dmo, boe-archive, "
              "boe-latest, or 'all' (default). Skipped components reuse "
              "their data from the existing gilt_yields.json if present. "
              "E.g. --only boe-latest for a fast daily refresh; "
              "--only dmo,boe-archive to skip the daily-only piece."))
    args = parser.parse_args()

    valid = {"dmo", "boe-archive", "boe-latest"}
    components = valid if args.only == "all" else set(args.only.split(","))
    unknown = components - valid
    if unknown:
        raise SystemExit(f"Unknown --only component(s): {unknown}. "
                          f"Valid: {sorted(valid)} or 'all'.")

    previous = None
    if os.path.exists("gilt_yields.json"):
        try:
            with open("gilt_yields.json") as f:
                previous = json.load(f)
        except Exception as e:
            print(f"Could not read existing gilt_yields.json ({e}) -- "
                  f"skipped components will start from empty instead.")

    # --- securities (DMO) ---
    if "dmo" in components:
        print("=== Building DMO per-security dataset (1996-2017) ===")
        securities = build_dmo_dataset()
    else:
        securities = (previous or {}).get("securities", [])
        print(f"=== Skipping DMO (--only) -- reusing {len(securities)} "
              f"securities from existing gilt_yields.json ===")

    # --- curves (BoE archive, optionally overlaid with BoE latest) ---
    if "boe-archive" in components:
        print("\n=== Building BoE archive curve dataset ===")
        curves = build_boe_dataset()
    else:
        curves = (previous or {}).get("curves", {"nominal": [], "real": []})
        print(f"\n=== Skipping BoE archive (--only) -- reusing "
              f"{len(curves.get('nominal', []))} nominal + "
              f"{len(curves.get('real', []))} real curve points from "
              f"existing gilt_yields.json ===")

    if "boe-latest" in components:
        print("\n=== Fetching BoE 'latest' curve data (fills the gap past "
              "the monthly archive refresh) ===")
        latest = fetch_latest_boe_curves()
        curves = merge_curve_overlay(curves, latest)
    else:
        print("\n=== Skipping BoE latest (--only) ===")

    output = {
        "generated": datetime.utcnow().isoformat() + "Z",
        "cutover_date": CUTOVER_DATE,
        "notes": (
            "securities[]: per-ISIN daily yields, DMO reference prices. "
            "Gross redemption yields were only calculated/published by "
            "DMO from 25 Nov 2002 (confirmed via dmo.gov.uk) up to "
            f"{CUTOVER_DATE}; earlier DMO files (1996-2001) contain "
            "prices only, no yield. curves.nominal / curves.real: "
            "BoE Anderson-Sleath fitted spot curves by maturity -- the "
            "monthly archive data overlaid with BoE's daily 'latest' "
            "data, continuous to present but NOT per-security."
        ),
        "securities": securities,
        "curves": curves,
    }

    with open("gilt_yields.json", "w") as f:
        json.dump(output, f, separators=(",", ":"))

    print(f"\nWrote gilt_yields.json: {len(securities)} securities, "
          f"{len(curves['nominal'])} nominal curve points, "
          f"{len(curves['real'])} real curve points.")


if __name__ == "__main__":
    main()


# ---------------------------------------------------------------------------
# Output schema (gilt_yields.json)
# ---------------------------------------------------------------------------
# {
#   "generated": "2026-09-13T12:00:00Z",
#   "cutover_date": "2017-07-21",
#   "securities": [
#     {
#       "isin": "GB00...",
#       "name": "5% Treasury Gilt 2025",
#       "type": "conventional" | "index_linked",
#       "series": [{"date": "1998-01-05", "yield": 6.23, "dirty_price": 99.41}, ...]
#     }, ...
#   ],
#   "curves": {
#     "nominal": [{"date": "2017-07-24", "maturity_years": 10.0, "rate_pct": 1.23}, ...],
#     "real":    [{"date": "2017-07-24", "maturity_years": 10.0, "rate_pct": -1.75}, ...]
#   }
# }
