#!/usr/bin/env python3
"""
fetch_gilt_data.py

Builds gilt_yields.json for the Gilt Yield Explorer chart, from three
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
     -> curve-level (not per-security) nominal, real, and implied-
        inflation (BEI) curves, in both spot (zero coupon) and forward
        form, used here to extend coverage past 21 Jul 2017, when DMO
        stopped publishing and per-security prices moved behind
        Tradeweb Insite (registration required, not scraped here).
     This part IS scraped automatically -- BoE's site doesn't block it.

     NOTE on par vs zero coupon: BoE does NOT publish a par-yield curve
     at all -- confirmed on their FAQ, which documents only "spot" and
     "forward" as the two measures they produce. What this script pulls
     from the "spot curve" sheet IS the zero-coupon curve already. Par
     yields, if wanted, are bootstrapped from the spot curve client-side
     in the app (see gilt-yield-explorer.html), not fetched from BoE.

     NOTE on BEI: BoE publishes the implied inflation curve directly
     (glcinflationddata.zip), which is a better BEI series than a naive
     nominal-minus-real spread, so this script fetches it as its own
     curve type rather than deriving it.

     NOTE on refresh cadence: BoE's archive zips are only refreshed on
     the 2nd working day of each month (confirmed on the BoE page's own
     FAQ) -- on their own they always lag by several weeks. This script
     also fetches BoE's separate "Latest yield curve data" zip, which is
     updated daily, and overlays it on top of the archive data to close
     that gap.

  3. DMO's published UK RPI history (report D4O, back to June 1980)
     https://www.dmo.gov.uk/data/gilt-market/index-linked-gilts
     -> feeds the Portfolio Cashflows tab, which needs to compute each
        index-linked gilt's Reference Index / Index Ratio itself. DMO's
        own reference list of gilts in issue (which would give the base
        RPI and lag type directly) is a PDF, not scraped here -- instead
        each security's first_price_date (already available, no new
        fetch needed) is used as a proxy for its first issue date, from
        which the app derives both the lag type (3-month from 2005-06
        onward, 8-month before) and the base Reference Index, following
        DMO's own published formula (igcalc.pdf). RPI history only
        covers actual published prints; cashflow dates beyond the last
        known print are projected using the BEI curve above.
     This part is NOT scraped automatically -- DMO blocks scripted
     requests to this data endpoint with ShieldSquare bot-detection
     (confirmed; same as the historical gilt-price page below). Like
     DMO_RAW_DIR, this needs a manually-downloaded export in RPI_RAW_DIR
     (see fetch_rpi_history's docstring for exact steps).

Run:

    pip install requests beautifulsoup4 openpyxl xlrd pandas
    python fetch_gilt_data.py                    # full run, all sources
    python fetch_gilt_data.py --only boe-latest   # fast daily refresh
    python fetch_gilt_data.py --only dmo          # e.g. after adding new
                                                   # dmo_raw/ files
    python fetch_gilt_data.py --only rpi          # after adding/updating
                                                   # a file in rpi_raw/
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
RPI_RAW_DIR = "./rpi_raw"  # put a manually-downloaded DMO RPI Data export here

CUTOVER_DATE = "2017-07-21"  # last DMO reference price date

# BoE publishes gilt-based nominal, real, and implied-inflation (BEI) curves.
# There is no "par yield" curve published anywhere -- confirmed via BoE's own
# FAQ, which only documents spot and forward measures. Par yields, if wanted,
# have to be bootstrapped from the spot curve (done client-side in the app).
BOE_CURVE_TYPES = ("nominal", "real", "inflation")

# AXIS_MATURITIES: the original curated "round number" set -- kept purely
# as a reference for what the app uses as fixed axis tick labels and the
# Time Series maturity-overlay dropdown (see gilt-yield-explorer.html's
# own AXIS_MATURITIES constant, which must match this one).
AXIS_MATURITIES = [1, 2, 3, 4, 5, 7, 10, 15, 20, 25, 30, 40]

# KEEP_MATURITIES: every integer year 1-40. Used for the actual fetched
# curve data. This used to match AXIS_MATURITIES exactly (12 points), but
# was widened so the forward curve (and now spot/par too, for free) is
# sampled at every year rather than only at the round-number points --
# the app still labels its axes only at AXIS_MATURITIES, this just gives
# it more real data to draw a smoother/more accurate line through.
KEEP_MATURITIES = list(range(1, 41))

# BoE's daily archive goes back decades, so keeping all 40 maturities for
# every single day since inception made gilt_yields.json too big to push
# to GitHub (126MB, over GitHub's 100MB single-file push limit). This
# trims maturity resolution for older dates, where the extra detail
# matters less, while keeping full year-by-year granularity for the
# recent period this app is mostly used to analyse:
#   - before GRANULARITY_CUTOVER_DATE: only AXIS_MATURITIES (12 points/day)
#   - from GRANULARITY_CUTOVER_DATE:   all of KEEP_MATURITIES (40 points/day)
GRANULARITY_CUTOVER_DATE = "2016-01-01"


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


FRACTION_MAP = {
    "\u215b": 0.125, "\u00bc": 0.25, "\u215c": 0.375, "\u00bd": 0.5,
    "\u215d": 0.625, "\u00be": 0.75, "\u215e": 0.875,
    "\u2153": 1 / 3, "\u2154": 2 / 3,
}


def parse_coupon_pct(name):
    """UK gilt names always lead with the coupon rate, e.g. '4\u00bd% Treasury
    Gilt 2028' or '0\u215b% Index-linked Treasury Gilt 2029'. Handles both
    the unicode-fraction style (older/most DMO naming) and plain decimals
    (e.g. '0.125%'). Returns None for names that don't parse (e.g. the
    handful of old floating-rate gilts, which don't have a fixed coupon).
    """
    m = re.match(r"^\s*(\d+)?([\u215b\u00bc\u215c\u00bd\u215d\u00be\u215e"
                 r"\u2153\u2154])?\s*%", name)
    if m and (m.group(1) or m.group(2)):
        whole = float(m.group(1)) if m.group(1) else 0.0
        frac = FRACTION_MAP.get(m.group(2), 0.0) if m.group(2) else 0.0
        return whole + frac
    m2 = re.match(r"^\s*(\d+(?:\.\d+)?)\s*%", name)
    if m2:
        return float(m2.group(1))
    # Tradeweb-style naming, e.g. "Treasury Gilt 4.25 12/46" or
    # "Treasury Gilt IL 0.125 03/29" -- coupon as a plain decimal after
    # "Gilt" (and an optional "IL" marker), no leading/trailing %. Seen
    # in the user's manually-added post-2017 Tradeweb file, which uses a
    # different naming convention to DMO's own historical files.
    m3 = re.search(r"\bGilt\b\s+(?:IL\s+)?(\d+(?:\.\d+)?)\b", name, re.IGNORECASE)
    if m3:
        return float(m3.group(1))
    return None


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
    df["redemption_date_parsed"] = (
        pd.to_datetime(df["redemption_date"], errors="coerce")
        if "redemption_date" in df.columns else pd.NaT)
    df = df.dropna(subset=["date", "isin", "yield"])

    rows = []
    for _, r in df.iterrows():
        name = str(r.get("name", "")).strip()
        is_il = bool(re.search(r"index.?linked|\bIL\b", name, re.IGNORECASE))
        dp = r.get("dirty_price")
        rd = r.get("redemption_date_parsed")
        rows.append({
            "date": r["date"].strftime("%Y-%m-%d"),
            "isin": str(r["isin"]).strip(),
            "name": name,
            "type": "index_linked" if is_il else "conventional",
            "yield": float(r["yield"]),
            "dirty_price": None if pd.isna(dp) else float(dp),
            "redemption_date": (rd.strftime("%Y-%m-%d")
                                 if rd is not None and not pd.isna(rd) else None),
            "coupon_pct": parse_coupon_pct(name),
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
            "redemption_date": row["redemption_date"],
            "coupon_pct": row["coupon_pct"],
            "series": [],
        })
        # A security's name/redemption_date/coupon are constant, but rows
        # from different years could disagree on redemption_date if DMO
        # ever corrected it -- keep the latest non-null value seen.
        if row["redemption_date"]:
            sec["redemption_date"] = row["redemption_date"]
        if row["coupon_pct"] is not None:
            sec["coupon_pct"] = row["coupon_pct"]
        sec["series"].append({
            "date": row["date"],
            "yield": row["yield"],
            "dirty_price": row["dirty_price"],
        })

    for sec in securities.values():
        sec["series"].sort(key=lambda p: p["date"])
        # Proxy for first issue date -- used client-side to infer each
        # index-linked gilt's indexation lag (3-month for gilts first
        # issued from 2005-06 onward, 8-month before that) and as the
        # base date for its index ratio calculation, since DMO's own
        # "gilts in issue" reference list (a PDF) isn't scraped here.
        sec["first_price_date"] = sec["series"][0]["date"] if sec["series"] else None

    return list(securities.values())


# ---------------------------------------------------------------------------
# 2. BoE curve-level nominal / real yields (2017 - present)
# ---------------------------------------------------------------------------

def discover_boe_archive_zips():
    """Find the nominal, real, and inflation daily archive zip links on
    the BoE yield curves page."""
    resp = requests.get(BOE_CURVES_PAGE, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # Only want the DAILY gilt-based curves: glcnominalddata.zip,
    # glcrealddata.zip, glcinflationddata.zip. Explicitly exclude "month"
    # (monthly duplicates) and "blc"/"ois" (not gilt-based).
    daily_gilt_re = re.compile(r"glc(nominal|real|inflation)ddata\.zip$",
                                re.IGNORECASE)

    zips = {c: [] for c in BOE_CURVE_TYPES}
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not href.lower().endswith(".zip"):
            continue
        m = daily_gilt_re.search(href)
        if not m:
            continue  # skip monthly, blc, ois, and anything else
        url = href if href.startswith("http") else f"https://www.bankofengland.co.uk{href}"
        curve = m.group(1).lower()
        zips[curve].append(url)

    print("Discovered BoE archives: " +
          ", ".join(f"{len(zips[c])} {c}" for c in BOE_CURVE_TYPES))
    return zips


def extract_curve_measures(sheets):
    """Given {sheet_name: DataFrame} for one workbook, pull out the spot
    and forward curve tabs (skipping the separately-fitted "short end"
    tabs, which use a different maturity grid and would double-count
    points already in the full curve -- see BoE's own FAQ on this).
    Returns {"spot": [rows], "forward": [rows]}.

    Maturity resolution depends on date -- see GRANULARITY_CUTOVER_DATE
    above: coarse (AXIS_MATURITIES) before it, full (KEEP_MATURITIES)
    from it onward. This keeps file size manageable while giving the
    recent period (what this app is mostly used to analyse) full detail.
    """
    out = {"spot": [], "forward": []}
    for sheet_name, df in sheets.items():
        sn = sheet_name.lower()
        if "short" in sn:
            continue
        if "spot" in sn:
            measure = "spot"
        elif "fwd" in sn or "forward" in sn:
            measure = "forward"
        else:
            continue

        header = df.iloc[3]  # guess: header row after title rows
        maturities = pd.to_numeric(header[1:], errors="coerce")
        # Map both the full and coarse target sets to actual columns once
        # per sheet -- which one applies is decided per-row, by date.
        col_for_target_full = {}
        for col_idx, m in enumerate(maturities, start=1):
            if pd.isna(m):
                continue
            for target in KEEP_MATURITIES:
                if abs(m - target) <= 0.1:
                    col_for_target_full.setdefault(target, col_idx)
        col_for_target_coarse = {t: col_for_target_full[t]
                                  for t in AXIS_MATURITIES
                                  if t in col_for_target_full}

        data = df.iloc[4:]
        for _, r in data.iterrows():
            date = pd.to_datetime(r[0], errors="coerce")
            if pd.isna(date):
                continue
            date_str = date.strftime("%Y-%m-%d")
            col_for_target = (col_for_target_full
                               if date_str >= GRANULARITY_CUTOVER_DATE
                               else col_for_target_coarse)
            for target, col_idx in col_for_target.items():
                rate = pd.to_numeric(r[col_idx], errors="coerce")
                if pd.isna(rate):
                    continue
                out[measure].append({
                    "date": date_str,
                    "maturity_years": float(target),
                    "rate_pct": float(rate),
                })
    return out


def parse_boe_zip(url):
    """Each BoE archive zip contains per-period Excel workbooks with
    spot and forward curve sheets: date rows x maturity-year columns,
    rate in %. Returns {"spot": [rows], "forward": [rows]}.

    BoE publishes these on a very fine maturity grid (roughly monthly
    steps across the full curve), which is far more resolution than a
    maturity-selector dropdown needs and makes the JSON output huge
    (multi-hundred-MB, unusable on mobile). We keep only a curated set
    of "round number" maturities (KEEP_MATURITIES) -- enough to be
    useful for a selector, without carrying ~40x more data than needed.
    """
    resp = requests.get(url, headers=HEADERS, timeout=60)
    resp.raise_for_status()

    result = {"spot": [], "forward": []}
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
            measures = extract_curve_measures(sheets)
            result["spot"].extend(measures["spot"])
            result["forward"].extend(measures["forward"])

    print(f"    parsed {len(result['spot'])} spot + {len(result['forward'])} "
          f"forward rows from {url} (curated to {len(KEEP_MATURITIES)} "
          f"maturities: {KEEP_MATURITIES})")
    return result


def build_boe_dataset():
    zips = discover_boe_archive_zips()
    curves = {c: {"spot": [], "forward": []} for c in BOE_CURVE_TYPES}
    for curve_name in BOE_CURVE_TYPES:
        for url in zips[curve_name]:
            try:
                parsed = parse_boe_zip(url)
                curves[curve_name]["spot"].extend(parsed["spot"])
                curves[curve_name]["forward"].extend(parsed["forward"])
            except Exception as e:
                print(f"  FAILED {url}: {e}")
        for measure in ("spot", "forward"):
            curves[curve_name][measure].sort(
                key=lambda r: (r["date"], r["maturity_years"]))
    return curves


def fetch_latest_boe_curves():
    """Fetch BoE's 'Latest yield curve data' zip -- updated daily, unlike
    the monthly-refreshed archive zips above. Bundles multiple curve types
    (nominal, real, inflation, OIS, ...) together in one file, so the
    curve type is read from the workbook's filename (sheet names are
    generic and identical across every workbook in this zip).
    """
    result = {c: {"spot": [], "forward": []} for c in BOE_CURVE_TYPES}
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

            fn = name.lower()
            curve_name = None
            for c in BOE_CURVE_TYPES:
                if c in fn:
                    curve_name = c
                    break
            if curve_name is None:
                continue  # OIS / BLC / anything else -- not tracked here

            with zf.open(name) as f:
                try:
                    sheets = pd.read_excel(f, sheet_name=None, header=None)
                except Exception as e:
                    print(f"    could not parse {name} in latest zip: {e}")
                    continue
            sheet_names_seen.extend(sheets.keys())
            measures = extract_curve_measures(sheets)
            result[curve_name]["spot"].extend(measures["spot"])
            result[curve_name]["forward"].extend(measures["forward"])

    totals = {c: len(result[c]["spot"]) + len(result[c]["forward"])
              for c in BOE_CURVE_TYPES}
    print(f"    parsed from latest-yield-curve-data.zip: " +
          ", ".join(f"{c}={totals[c]}" for c in BOE_CURVE_TYPES) +
          " (spot+forward rows combined)")
    if sum(totals.values()) == 0:
        print(f"    WARNING: 0 rows from the 'latest' zip. Workbook filenames "
              f"found: {filenames_seen!r}. Sheet names: {sheet_names_seen!r} -- "
              f"if none of the filenames contain 'nominal'/'real'/'inflation', "
              f"report this list back so the filter can be adjusted.")
    return result


def merge_curve_overlay(base, overlay):
    """Combine two {curve_type: {measure: [...]}} curve dicts, with
    `overlay` rows taking precedence over `base` rows on the same
    (date, maturity) -- used to let the fresher 'latest' BoE data
    supersede the monthly-archive data for whatever dates they share,
    while keeping all the archive's older history.
    """
    result = {}
    for curve_name in BOE_CURVE_TYPES:
        result[curve_name] = {}
        for measure in ("spot", "forward"):
            by_key = {(r["date"], r["maturity_years"]): r
                      for r in base.get(curve_name, {}).get(measure, [])}
            for r in overlay.get(curve_name, {}).get(measure, []):
                by_key[(r["date"], r["maturity_years"])] = r
            result[curve_name][measure] = sorted(
                by_key.values(), key=lambda r: (r["date"], r["maturity_years"]))
    return result


# ---------------------------------------------------------------------------
# RPI history (for index-linked gilt cashflow projections)
# ---------------------------------------------------------------------------

DMO_RPI_REPORT_URL = "https://www.dmo.gov.uk/data/XmlDataReport?reportCode=D4O"


def _try_parse_rpi_response(raw, content_type):
    """Attempt to parse a response body as RPI data in XML/Excel/CSV.
    Returns a DataFrame or None if nothing recognisable came back."""
    df = None
    if "xml" in content_type or raw.lstrip()[:1] == b"<":
        try:
            df = pd.read_xml(io.BytesIO(raw), parser="etree")
        except Exception:
            pass
    if df is None:
        try:
            df = pd.read_excel(io.BytesIO(raw))
        except Exception:
            pass
    if df is None:
        try:
            candidate = pd.read_csv(io.BytesIO(raw))
            # A 1-column "DataFrame" usually means we CSV-parsed an HTML
            # page by accident (one long "line" per row) -- reject it.
            if candidate.shape[1] > 1:
                df = candidate
        except Exception:
            pass
    return df


def discover_rpi_raw_files():
    """Files the user has manually downloaded into RPI_RAW_DIR (DMO's RPI
    Data report, exported via a real browser -- see fetch_rpi_history's
    docstring for why this is necessary rather than fetched directly)."""
    if not os.path.isdir(RPI_RAW_DIR):
        return []
    return sorted(
        os.path.join(RPI_RAW_DIR, fn) for fn in os.listdir(RPI_RAW_DIR)
        if fn.lower().endswith((".xls", ".xlsx", ".csv")))


def parse_rpi_raw_file(path):
    try:
        df = pd.read_excel(path)
    except Exception:
        try:
            df = pd.read_csv(path)
        except Exception as e:
            print(f"  could not read {path}: {e}")
            return []
    rows = _rows_from_rpi_dataframe(df)
    print(f"  parsed {len(rows)} months from {path}")
    return rows


def fetch_rpi_history():
    """Get DMO's published UK RPI history (report D4O), which goes back
    to June 1980. Used by the app to compute index-linked gilts' Reference
    Index / Index Ratio itself (per DMO's own published formula -- see
    igcalc.pdf), rather than needing DMO's separate PDF reference list of
    gilts in issue.

    CONFIRMED: DMO blocks scripted requests to this data endpoint with
    ShieldSquare bot-detection (same thing that blocked the historical
    gilt-price page early in this project) -- both the interactive
    report page and its supposed XML feed return a "ShieldSquare Block"
    page rather than data, regardless of URL. So, like the DMO yearly
    price files (DMO_RAW_DIR), this now expects a manually-downloaded
    copy in RPI_RAW_DIR:

        1. Visit https://www.dmo.gov.uk/data/gilt-market/index-linked-gilts
           in a real browser and open the "RPI data" report.
        2. Export/download it (CSV or Excel).
        3. Save the file into ./rpi_raw/ next to this script (any
           filename works, as long as it ends .csv/.xls/.xlsx).

    The web-fetch attempt below is kept as a fallback in case DMO ever
    lifts the block, but given it's confirmed blocked, don't expect it
    to work -- the manual-file path is the real one.

    Returns a list of {"month": "YYYY-MM-01", "rpi": float}, sorted by
    month, or an empty list if no data could be found (in which case
    the app falls back to BEI-only projection with no historical RPI
    anchor -- print output will make this failure obvious).
    """
    raw_files = discover_rpi_raw_files()
    if raw_files:
        print(f"  found {len(raw_files)} manually-downloaded RPI file(s) "
              f"in {RPI_RAW_DIR}: {raw_files}")
        rows = []
        for path in raw_files:
            rows.extend(parse_rpi_raw_file(path))
        # de-dupe by month (in case of overlapping manual exports), keep
        # the last-seen value for any repeated month.
        by_month = {r["month"]: r["rpi"] for r in rows}
        out = sorted(({"month": m, "rpi": v} for m, v in by_month.items()),
                      key=lambda r: r["month"])
        if out:
            return out
        print(f"  WARNING: found files in {RPI_RAW_DIR} but couldn't "
              f"parse any RPI rows from them -- falling back to the "
              f"(likely blocked) web fetch below for diagnostics.")

    url_variants = [
        DMO_RPI_REPORT_URL,
        # Fallback only -- this is DMO's interactive report page, not a
        # data file, but kept in case the XML endpoint above ever moves.
        "https://www.dmo.gov.uk/data/ExportReport?reportCode=D4O",
    ]

    last_html_resp = None
    last_html_url = None
    for url in url_variants:
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"  could not fetch {url}: {e}")
            continue

        content_type = resp.headers.get("Content-Type", "").lower()
        if "html" in content_type:
            last_html_resp = resp  # keep the page itself for link-scanning below
            last_html_url = url
            continue

        df = _try_parse_rpi_response(resp.content, content_type)
        if df is not None and not df.empty:
            print(f"  RPI history: {url} worked ({content_type})")
            return _rows_from_rpi_dataframe(df)

    # Nothing worked directly -- scan the HTML report page for a real
    # download link, so we know what to try next rather than guessing blind.
    if last_html_resp is not None:
        soup = BeautifulSoup(last_html_resp.text, "html.parser")
        title = soup.title.string.strip() if soup.title and soup.title.string else "(no title)"
        candidates = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            hl = href.lower()
            if any(k in hl for k in (".csv", ".xlsx", ".xls", ".xml",
                                      "download", "export", "/media/")):
                candidates.append(href)

        # Same bot-detection block that hit DMO's historical-prices page
        # earlier in this project (ShieldSquare) -- check for it here too,
        # since "a data endpoint returns HTML" is exactly what that looks
        # like, and a candidate-link scan can't fix a block page anyway.
        page_text = soup.get_text(" ", strip=True).lower()
        block_markers = ["shieldsquare", "captcha", "access denied",
                         "request unsuccessful", "blocked"]
        hit_markers = [m for m in block_markers if m in page_text]

        print(f"  WARNING: RPI report URL(s) returned HTML, not data. "
              f"Last tried: {last_html_url} -- page title: {title!r}. "
              f"Found {len(candidates)} candidate link(s) that might be "
              f"the real download: {candidates!r}.")
        if hit_markers:
            print(f"  This looks like a bot-detection block page "
                  f"(matched: {hit_markers!r}) -- same kind of thing that "
                  f"blocked DMO's historical-prices page earlier in this "
                  f"project, not a wrong URL. Likely needs a different "
                  f"approach (e.g. manual download) rather than a URL fix.")
        else:
            print(f"  Not an obvious bot-block page -- if one of the "
                  f"candidate links above looks right, or the title/first "
                  f"part of the page suggests what's actually going on, "
                  f"report it back and I'll adjust.")
    else:
        print("  WARNING: could not fetch any RPI report URL variant.")
    return []


def _rows_from_rpi_dataframe(df):
    # Column names are unconfirmed -- match by keyword like the DMO price
    # files, rather than assuming exact names.
    cols = {str(c).strip().lower(): c for c in df.columns}
    date_col = next((cols[c] for c in cols if "date" in c or "month" in c
                      or "period" in c), None)
    rpi_col = next((cols[c] for c in cols if "rpi" in c or "index" in c),
                    None)
    if date_col is None or rpi_col is None:
        print(f"  WARNING: RPI history parsed ({len(df)} rows) but "
              f"couldn't identify date/RPI columns. Columns found: "
              f"{list(df.columns)!r}. First 3 rows:\n{df.head(3)}")
        return []

    out = []
    for _, r in df.iterrows():
        d = pd.to_datetime(r[date_col], errors="coerce")
        v = pd.to_numeric(r[rpi_col], errors="coerce")
        if pd.isna(d) or pd.isna(v):
            continue
        out.append({"month": d.strftime("%Y-%m-01"), "rpi": float(v)})
    out.sort(key=lambda r: r["month"])
    print(f"  parsed {len(out)} months of RPI history "
          f"({out[0]['month'] if out else '?'} to "
          f"{out[-1]['month'] if out else '?'})")
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Build gilt_yields.json from DMO + BoE data sources.")
    parser.add_argument(
        "--only", default="all",
        help=("Comma-separated components to run: dmo, boe-archive, "
              "boe-latest, rpi, or 'all' (default). Skipped components "
              "reuse their data from the existing gilt_yields.json if "
              "present. E.g. --only boe-latest for a fast daily refresh; "
              "--only dmo,boe-archive to skip the daily-only piece."))
    args = parser.parse_args()

    valid = {"dmo", "boe-archive", "boe-latest", "rpi"}
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
    empty_curves = {c: {"spot": [], "forward": []} for c in BOE_CURVE_TYPES}
    if "boe-archive" in components:
        print("\n=== Building BoE archive curve dataset "
              "(nominal, real, inflation -- spot + forward) ===")
        curves = build_boe_dataset()
    else:
        curves = (previous or {}).get("curves", empty_curves)
        counts = ", ".join(
            f"{c}={len(curves.get(c, {}).get('spot', []))}sp+"
            f"{len(curves.get(c, {}).get('forward', []))}fwd"
            for c in BOE_CURVE_TYPES)
        print(f"\n=== Skipping BoE archive (--only) -- reusing {counts} "
              f"curve points from existing gilt_yields.json ===")

    if "boe-latest" in components:
        print("\n=== Fetching BoE 'latest' curve data (fills the gap past "
              "the monthly archive refresh) ===")
        latest = fetch_latest_boe_curves()
        curves = merge_curve_overlay(curves, latest)
    else:
        print("\n=== Skipping BoE latest (--only) ===")

    # --- RPI history (for index-linked gilt cashflow projections) ---
    if "rpi" in components:
        print("\n=== Fetching DMO RPI history ===")
        rpi_history = fetch_rpi_history()
        if not rpi_history and previous and previous.get("rpi_history"):
            rpi_history = previous["rpi_history"]
            print(f"  fetch returned nothing (likely blocked, or no file "
                  f"in {RPI_RAW_DIR} on this machine) -- keeping the "
                  f"{len(rpi_history)} months already in gilt_yields.json "
                  f"rather than overwriting them with an empty result.")
    else:
        rpi_history = (previous or {}).get("rpi_history", [])
        print(f"\n=== Skipping RPI history (--only) -- reusing "
              f"{len(rpi_history)} months from existing gilt_yields.json ===")

    output = {
        "generated": datetime.utcnow().isoformat() + "Z",
        "cutover_date": CUTOVER_DATE,
        "notes": (
            "securities[]: per-ISIN daily yields + dirty prices, DMO "
            "reference prices. Gross redemption yields were only "
            "calculated/published by DMO from 25 Nov 2002 (confirmed via "
            f"dmo.gov.uk) up to {CUTOVER_DATE}; earlier DMO files "
            "(1996-2001) contain prices only, no yield. "
            "curves.{nominal,real,inflation}.{spot,forward}: BoE "
            "Anderson-Sleath fitted curves by maturity (continuously "
            "compounded, per BoE's own FAQ) -- monthly archive data "
            "overlaid with BoE's daily 'latest' data. 'spot' = zero "
            "coupon yields; 'forward' = instantaneous forward rates; "
            "'inflation' = implied breakeven inflation (BEI), published "
            "directly by BoE rather than derived as nominal-minus-real. "
            "There is no published par-yield curve -- BoE's FAQ confirms "
            "only spot and forward are produced; par yields are "
            "bootstrapped from the spot curve client-side in the app. "
            f"Maturity resolution is date-dependent (kept file size "
            f"under GitHub's 100MB push limit): only "
            f"{AXIS_MATURITIES} years before {GRANULARITY_CUTOVER_DATE}, "
            f"all of {KEEP_MATURITIES[0]}-{KEEP_MATURITIES[-1]} (every "
            f"year) from {GRANULARITY_CUTOVER_DATE} onward. "
            "securities[].redemption_date, .coupon_pct (parsed from the "
            "name), and .first_price_date (proxy for first issue date, "
            "used to infer each index-linked gilt's indexation lag and "
            "as the base date for its index ratio -- DMO's own 'gilts in "
            "issue' reference list is a PDF and isn't scraped here) are "
            "for the app's Portfolio Cashflows tab. rpi_history[]: "
            "DMO's published UK RPI series (report D4O), used the same "
            "way, per DMO's published Reference Index formula (see "
            "igcalc.pdf) -- computed client-side, not by this script."
        ),
        "securities": securities,
        "curves": curves,
        "rpi_history": rpi_history,
    }

    with open("gilt_yields.json", "w") as f:
        json.dump(output, f, separators=(",", ":"))

    curve_summary = ", ".join(
        f"{c}={len(curves.get(c, {}).get('spot', []))}sp+"
        f"{len(curves.get(c, {}).get('forward', []))}fwd"
        for c in BOE_CURVE_TYPES)
    print(f"\nWrote gilt_yields.json: {len(securities)} securities, "
          f"curves: {curve_summary}, rpi_history: {len(rpi_history)} months.")


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
#       "redemption_date": "2025-03-07" | null,
#       "coupon_pct": 5.0 | null,           # parsed from name; null for floaters
#       "first_price_date": "2003-01-06",   # proxy for first issue date
#       "series": [{"date": "1998-01-05", "yield": 6.23, "dirty_price": 99.41}, ...]
#     }, ...
#   ],
#   "curves": {
#     "nominal":   {"spot": [{"date":"2017-07-24","maturity_years":10.0,"rate_pct":1.23}, ...],
#                   "forward": [...same shape...]},
#     "real":      {"spot": [...], "forward": [...]},
#     "inflation": {"spot": [...], "forward": [...]}   # implied BEI, published directly by BoE
#   },
#   "rpi_history": [{"month": "1987-01-01", "rpi": 100.0}, ...]   # DMO report D4O
#   # No "par" key -- BoE doesn't publish par yields; the app derives them
#   # client-side from curves.<type>.spot via a standard bootstrap.
#   # Maturity coverage is date-dependent: only AXIS_MATURITIES (12
#   # points) before GRANULARITY_CUTOVER_DATE, all of KEEP_MATURITIES
#   # (every year 1-40) from that date onward -- see the constants above.
# }
