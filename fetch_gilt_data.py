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

  4. Hargreaves Lansdown's public daily gilt-price pages (no login)
     https://www.hl.co.uk/shares/corporate-bonds-gilts/bond-prices/uk-gilts
     https://www.hl.co.uk/shares/corporate-bonds-gilts/bond-prices/uk-index-linked-gilts
     -> fills the daily security-level gap left once DMO stopped
        publishing per-security prices in 2017 -- previously only
        available via slow, one-by-one manual Tradeweb downloads. Plain
        server-rendered HTML tables, confirmed scrapable with no login.
        HL shows clean price; this script adds accrued interest
        (Actual/Actual, index-ratio-adjusted for linkers) to get the
        dirty_price the rest of the app expects. Data is from NetBuilder,
        a market data vendor -- solid for this app's purposes, but not
        as authoritative as DMO/BoE/ONS, and HL could change their page
        layout at any time (this parser fails loudly, with diagnostics,
        rather than silently, if that happens).

Run:

    pip install requests beautifulsoup4 openpyxl xlrd pandas
    python fetch_gilt_data.py                    # full run, all sources
    python fetch_gilt_data.py --only boe-latest,hl-prices  # daily refresh
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
import calendar
from datetime import date, datetime

import requests
from bs4 import BeautifulSoup

try:
    import pandas as pd
except ImportError:
    raise SystemExit("pip install pandas openpyxl xlrd requests beautifulsoup4")

HEADERS = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")}

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
IL_GILTS_RAW_DIR = "./il_gilts_raw"  # put a manually-downloaded DMO "Index-linked
                                     # Gilts in Issue" export here (see
                                     # fetch_il_index_ratio_anchors docstring)

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


def merge_dmo_into_previous(dmo_securities, previous_securities):
    """build_dmo_dataset() only knows about what's currently in dmo_raw/,
    so used on its own it would wholesale REPLACE the securities list --
    silently losing any dates that exist only because a later source
    (e.g. HL's daily automated prices) added them and aren't also
    present in a dmo_raw file. This merges instead: for each ISIN, the
    freshly re-parsed dmo_raw dates take priority (so a new/corrected
    Tradeweb or DMO file properly supersedes existing values for the
    dates it covers), but every other date and every security not
    touched by this dmo_raw pass is preserved from the previous run
    rather than dropped.
    """
    prev_by_isin = {s["isin"]: s for s in previous_securities}
    merged = []
    seen_isins = set()

    for dsec in dmo_securities:
        isin = dsec["isin"]
        seen_isins.add(isin)
        prev = prev_by_isin.get(isin)
        if prev is None:
            merged.append(dsec)
            continue
        by_date = {p["date"]: p for p in prev.get("series", [])}
        for p in dsec["series"]:
            by_date[p["date"]] = p  # fresh dmo_raw data wins for this date
        merged_sec = dict(prev)
        merged_sec.update({
            "name": dsec["name"],
            "type": dsec["type"],
            "redemption_date": dsec.get("redemption_date") or prev.get("redemption_date"),
            "coupon_pct": (dsec.get("coupon_pct") if dsec.get("coupon_pct") is not None
                           else prev.get("coupon_pct")),
            "series": sorted(by_date.values(), key=lambda p: p["date"]),
        })
        earliest_dates = [d for d in
                          (merged_sec["series"][0]["date"] if merged_sec["series"] else None,
                           prev.get("first_price_date")) if d]
        merged_sec["first_price_date"] = min(earliest_dates) if earliest_dates else None
        merged.append(merged_sec)

    # Securities the previous run knew about but this dmo_raw pass never
    # touched at all (e.g. a gilt HL discovered that isn't in any
    # dmo_raw file) -- keep them completely as-is.
    for isin, prev in prev_by_isin.items():
        if isin not in seen_isins:
            merged.append(prev)

    return merged


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
# Index-linked gilt index-ratio anchors (from DMO's "Index-linked Gilts in
# Issue" report) -- fixes systematic error in deriving the index ratio
# purely from first_price_date (see fetch_il_index_ratio_anchors below).
# ---------------------------------------------------------------------------

def discover_il_gilts_raw_files():
    if not os.path.isdir(IL_GILTS_RAW_DIR):
        return []
    return sorted(
        os.path.join(IL_GILTS_RAW_DIR, fn) for fn in os.listdir(IL_GILTS_RAW_DIR)
        if fn.lower().endswith((".xls", ".xlsx", ".csv")))


def parse_il_gilts_raw_file(path):
    """Parse one manually-downloaded copy of DMO's "Index-linked Gilts in
    Issue" report ("Download"-style sheet: Gilt Name / ISIN Code / "Index
    Ratio for settlement on <date>" / ..., with rows grouped under
    "3-month Indexation Lag" and "8-month Indexation Lag" section labels).

    Returns {"settlement_date": "YYYY-MM-DD", "ratios": {isin: ratio},
    "lag_months": {isin: 3|8}}, or None if the file doesn't look like
    this report (e.g. wrong sheet, or DMO changed its layout).
    """
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None)
    except Exception:
        try:
            sheets = {"": pd.read_csv(path, header=None)}
        except Exception as e:
            print(f"  could not read {path}: {e}")
            return None

    for sheet_name, df in sheets.items():
        header_row_idx = None
        settlement_date = None
        isin_col = ratio_col = None
        for i in range(min(20, len(df))):
            row_vals = [str(v) for v in df.iloc[i].tolist()]
            if (any("isin" in v.lower() for v in row_vals)
                    and any("index ratio" in v.lower() for v in row_vals)):
                header_row_idx = i
                for col_idx, v in enumerate(row_vals):
                    vl = v.lower()
                    if "isin" in vl:
                        isin_col = col_idx
                    m = re.search(r"index ratio for settlement on\s+"
                                  r"([\w\-]+)", v, re.IGNORECASE)
                    if m:
                        ratio_col = col_idx
                        settlement_date = pd.to_datetime(m.group(1), errors="coerce")
                break
        if header_row_idx is None or isin_col is None or ratio_col is None:
            continue  # not this sheet -- try the next one (e.g. "Download" vs others)
        if settlement_date is None or pd.isna(settlement_date):
            print(f"  WARNING: found the header row in {path} ({sheet_name}) "
                  f"but couldn't parse the settlement date from it -- "
                  f"can't use these ratios without an anchor date.")
            continue

        ratios, lag_months = {}, {}
        current_lag = None
        for i in range(header_row_idx + 1, len(df)):
            first_cell = df.iat[i, 0]
            if isinstance(first_cell, str) and "indexation lag" in first_cell.lower():
                current_lag = 8 if "8-month" in first_cell else (
                    3 if "3-month" in first_cell else current_lag)
                continue
            isin = df.iat[i, isin_col]
            ratio = df.iat[i, ratio_col]
            if pd.isna(isin) or pd.isna(ratio):
                continue
            isin = str(isin).strip()
            if not re.match(r"^GB00[0-9A-Z]{7}\d$", isin):
                continue
            try:
                ratios[isin] = float(ratio)
            except (TypeError, ValueError):
                continue
            if current_lag:
                lag_months[isin] = current_lag

        if ratios:
            print(f"  parsed {len(ratios)} index ratios from {path} "
                  f"({sheet_name}), anchored to settlement date "
                  f"{settlement_date.strftime('%Y-%m-%d')}")
            return {"settlement_date": settlement_date.strftime("%Y-%m-%d"),
                    "ratios": ratios, "lag_months": lag_months}

    print(f"  WARNING: {path} didn't match the expected 'Index-linked "
          f"Gilts in Issue' layout (Gilt Name / ISIN Code / Index Ratio "
          f"columns with 3-month/8-month section labels) in any sheet.")
    return None


def fetch_il_index_ratio_anchors():
    """Get real, DMO-published index ratios (and lag type) per currently-
    issued index-linked gilt, to anchor the app's own Reference Index
    formula to rather than deriving everything from first_price_date --
    that proxy turned out to be materially wrong for gilts whose earliest
    available price data doesn't reach back to their actual issue date
    (confirmed: error scaled almost exactly with each gilt's true
    accumulated inflation uplift, from a user cross-check against this
    report).

    DMO's own reference list ("gilts in issue") is normally a PDF and
    isn't scraped automatically here -- like RPI history, this expects a
    manually-downloaded copy (exportable as Excel from a real browser
    session) in IL_GILTS_RAW_DIR:

        1. Visit https://www.dmo.gov.uk/data/gilt-market/index-linked-gilts
        2. Open/export the "Index-linked gilts in issue" report as Excel.
        3. Save it into ./il_gilts_raw/ next to this script.

    Returns {"settlement_date": ..., "ratios": {isin: ratio},
    "lag_months": {isin: 3|8}}, or an empty dict if no usable file was
    found (in which case the app falls back to the first_price_date
    proxy method for any gilt without an anchor).
    """
    files = discover_il_gilts_raw_files()
    if not files:
        print(f"  no files found in {IL_GILTS_RAW_DIR} -- see "
              f"fetch_il_index_ratio_anchors docstring for how to get "
              f"DMO's 'Index-linked Gilts in Issue' report.")
        return {}
    parsed = [r for r in (parse_il_gilts_raw_file(p) for p in files) if r]
    if not parsed:
        return {}
    return max(parsed, key=lambda r: r["settlement_date"])


# ---------------------------------------------------------------------------
# Hargreaves Lansdown daily gilt prices (fills the daily security-level gap
# now that DMO no longer publishes per-security prices post-cutover, and
# Tradeweb requires slow one-by-one manual downloads).
# ---------------------------------------------------------------------------
#
# HL's public bond-price pages (no login required) are plain server-rendered
# HTML tables -- confirmed by fetching them directly. Two pages, matching
# our own conventional/index_linked split:
HL_CONVENTIONAL_URL = "https://www.hl.co.uk/shares/corporate-bonds-gilts/bond-prices/uk-gilts"
HL_INDEX_LINKED_URL = "https://www.hl.co.uk/shares/corporate-bonds-gilts/bond-prices/uk-index-linked-gilts"

ISIN_RE = re.compile(r"\bGB00[0-9A-Z]{7}\d\b")


def parse_hl_gilt_table(html, is_index_linked):
    """Parse one HL bond-price page into {isin: {...}}. Returns {} if the
    page doesn't match the expected layout. Whenever the result is empty,
    ALWAYS prints one diagnostic line covering every stage (table found?
    which one, out of how many candidates? headers? row/cell counts?
    sample row content?) -- earlier versions of this diagnostic were
    split across several separate conditional branches, any one of which
    could individually fail to fire and leave a bare "parsed 0" with no
    explanation at all, which is what actually happened in practice.
    This version can't have that gap: everything funnels through one
    unconditional check at the end.
    """
    soup = BeautifulSoup(html, "html.parser")
    all_tables = soup.find_all("table")
    candidates = []
    for t in all_tables:
        header_row = t.find("tr")
        if not header_row:
            continue
        hdrs = [c.get_text(strip=True).lower() for c in header_row.find_all(["th", "td"])]
        if any("issuer" in h for h in hdrs) and any("coupon" in h for h in hdrs):
            candidates.append((t, hdrs))
    table, headers = candidates[0] if candidates else (None, None)

    idx_issuer = idx_coupon = idx_maturity = idx_price = idx_ytm0 = None
    rows = []
    rows_with_cells = 0
    result = {}

    if table is not None:
        def col(*keywords):
            for i, h in enumerate(headers):
                if all(k in h for k in keywords):
                    return i
            return None

        idx_issuer = col("issuer")
        idx_coupon = col("coupon")
        idx_maturity = col("maturity")
        idx_price = col("price")
        idx_ytm0 = col("ytm", "0")  # "YTM 0% tax" -- effectively the gross yield

        rows = table.find_all("tr")[1:]  # skip header row
        for r in rows:
            # Some tables mark the row-identifying cell (issuer, here) as
            # <th scope="row"> rather than <td> -- accept either, unlike
            # a td-only search which would silently misalign every column.
            cells = r.find_all(["td", "th"])
            if not cells or idx_issuer is None or idx_issuer >= len(cells):
                continue
            rows_with_cells += 1
            issuer_text = cells[idx_issuer].get_text(" ", strip=True)
            m = ISIN_RE.search(issuer_text)
            if not m:
                continue
            isin = m.group(0)

            def cell_num(idx):
                if idx is None or idx >= len(cells):
                    return None
                txt = cells[idx].get_text(strip=True).replace(",", "")
                try:
                    return float(txt)
                except ValueError:
                    return None

            price = cell_num(idx_price)
            if price is None:
                continue
            coupon_pct = cell_num(idx_coupon)
            maturity_date = None
            if idx_maturity is not None and idx_maturity < len(cells):
                md = pd.to_datetime(cells[idx_maturity].get_text(strip=True),
                                     errors="coerce", dayfirst=True)
                if md is not None and not pd.isna(md):
                    maturity_date = md.strftime("%Y-%m-%d")

            result[isin] = {
                "name": cells[idx_issuer].get_text(" ", strip=True).split(" GBP")[0].strip(),
                "coupon_pct": coupon_pct,
                "maturity_date": maturity_date,
                "clean_price": price,
                "yield_pct": cell_num(idx_ytm0),
                "type": "index_linked" if is_index_linked else "conventional",
            }

    if not result:
        page_title = (soup.title.string.strip()
                       if soup.title and soup.title.string else "?")
        sample_texts = None
        if rows:
            sample = rows[0].find_all(["td", "th"])
            sample_texts = [c.get_text(strip=True) for c in sample]
        print(f"  WARNING: extracted 0 gilts from this page. Page title: "
              f"{page_title!r}. <table> elements on page: {len(all_tables)}, "
              f"matching 'issuer'+'coupon' headers: {len(candidates)}. "
              f"{'Using the first match.' if candidates else 'No match at all.'} "
              f"Headers used: {headers!r}. Column indices -- issuer: "
              f"{idx_issuer}, coupon: {idx_coupon}, maturity: {idx_maturity}, "
              f"price: {idx_price}, ytm0: {idx_ytm0}. Data <tr> rows: "
              f"{len(rows)}, rows with a usable issuer cell: {rows_with_cells}. "
              f"First row's cells: {sample_texts!r}.")
    return result


def fetch_hl_gilt_prices():
    """Fetch and parse both HL bond-price pages. Returns {isin: {...}},
    or {} for a page that fails (with diagnostics printed), so a problem
    with one page doesn't lose data from the other.
    """
    combined = {}
    for url, is_il in ((HL_CONVENTIONAL_URL, False), (HL_INDEX_LINKED_URL, True)):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"  could not fetch {url}: {e}")
            continue
        parsed = parse_hl_gilt_table(resp.text, is_il)
        print(f"  parsed {len(parsed)} {'index-linked' if is_il else 'conventional'} "
              f"gilt prices from {url}")
        combined.update(parsed)
    return combined


# ---- accrued interest / dirty price -----------------------------------

def _add_months(d, n):
    month = d.month - 1 + n
    year = d.year + month // 12
    month = month % 12 + 1
    day = min(d.day, calendar.monthrange(year, month)[1])
    return date(year, month, day)


def _coupon_window(redemption_date_str, today):
    """Last and next semi-annual coupon dates (day/month anchored on the
    redemption date) bracketing `today`. None if already matured."""
    redemption = date.fromisoformat(redemption_date_str)
    if today >= redemption:
        return None
    d = redemption
    guard = 0
    while d > today and guard < 200:
        d = _add_months(d, -6)
        guard += 1
    return d, _add_months(d, 6)


def compute_dirty_price(clean_price, coupon_pct, redemption_date_str, today,
                         index_ratio=None):
    """Dirty price = clean price + accrued interest, using the standard
    Actual/Actual semi-annual convention. `index_ratio` (if given)
    inflation-adjusts the accrued interest for index-linked gilts, using
    the closest thing we have -- the DMO-anchored index ratio -- rather
    than a full Reference Index recomputation, which is a reasonable
    approximation given how small accrued interest is relative to price.
    Does NOT model the UK ex-dividend period (a further, smaller
    simplification, consistent with the rest of this app's approach to
    index-linked gilts).
    """
    if clean_price is None or coupon_pct is None or not redemption_date_str:
        return clean_price
    window = _coupon_window(redemption_date_str, today)
    if window is None:
        return clean_price  # matured -- no accrual to add
    last_coupon, next_coupon = window
    days_since = (today - last_coupon).days
    days_in_period = (next_coupon - last_coupon).days
    if days_in_period <= 0:
        return clean_price
    ai = (coupon_pct / 2) * (days_since / days_in_period)
    if index_ratio:
        ai *= index_ratio
    return round(clean_price + ai, 4)


def apply_hl_prices(securities, hl_data, il_index_ratio_anchors):
    """Merge today's HL prices into `securities` (mutates and returns it):
    updates today's series point for matching ISINs (adding one if not
    already present for today), and creates a new minimal security entry
    for any ISIN HL has that we don't -- e.g. a gilt issued after our
    last DMO/Tradeweb data. Dirty price is computed from HL's clean price
    via compute_dirty_price(); yield comes straight from HL for
    conventional gilts (HL doesn't publish one for linkers).
    """
    today_str = datetime.utcnow().strftime("%Y-%m-%d")
    today = date.fromisoformat(today_str)
    by_isin = {s["isin"]: s for s in securities}
    ratios = (il_index_ratio_anchors or {}).get("ratios", {})
    updated, added = 0, 0

    for isin, hl in hl_data.items():
        sec = by_isin.get(isin)
        if sec is None:
            sec = {
                "isin": isin,
                "name": hl["name"],
                "type": hl["type"],
                "redemption_date": hl["maturity_date"],
                "coupon_pct": hl["coupon_pct"],
                "series": [],
                "first_price_date": today_str,
            }
            securities.append(sec)
            by_isin[isin] = sec
            added += 1

        redemption_date = sec.get("redemption_date") or hl["maturity_date"]
        index_ratio = ratios.get(isin) if sec["type"] == "index_linked" else None
        dirty_price = compute_dirty_price(hl["clean_price"], sec.get("coupon_pct"),
                                           redemption_date, today, index_ratio)
        point = {
            "date": today_str,
            "yield": hl["yield_pct"],  # None for index-linked -- HL doesn't publish one
            "dirty_price": dirty_price,
        }
        if sec["series"] and sec["series"][-1]["date"] == today_str:
            sec["series"][-1] = point
        else:
            sec["series"].append(point)
        updated += 1

    print(f"  applied HL prices to {updated} securities ({added} newly "
          f"added, not previously in the dataset)")
    return securities


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Build gilt_yields.json from DMO + BoE data sources.")
    parser.add_argument(
        "--only", default="all",
        help=("Comma-separated components to run: dmo, boe-archive, "
              "boe-latest, rpi, il-ratios, hl-prices, or 'all' (default). "
              "Skipped components reuse their data from the existing "
              "gilt_yields.json if present. E.g. --only boe-latest,"
              "hl-prices for a fast daily refresh; --only dmo,boe-archive "
              "to skip the daily-only pieces."))
    args = parser.parse_args()

    valid = {"dmo", "boe-archive", "boe-latest", "rpi", "il-ratios", "hl-prices"}
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
        dmo_securities = build_dmo_dataset()
        prev_securities = (previous or {}).get("securities", [])
        if prev_securities:
            securities = merge_dmo_into_previous(dmo_securities, prev_securities)
            print(f"  merged onto {len(prev_securities)} previously-known "
                  f"securities -- dates covered by the current dmo_raw "
                  f"files take priority, everything else (e.g. HL-only "
                  f"daily prices) is preserved")
        else:
            securities = dmo_securities
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

    # --- IL gilt index-ratio anchors (fixes the first_price_date proxy) ---
    if "il-ratios" in components:
        print("\n=== Fetching DMO index-linked gilt index-ratio anchors ===")
        il_index_ratio_anchors = fetch_il_index_ratio_anchors()
        if not il_index_ratio_anchors and previous and previous.get("il_index_ratio_anchors"):
            il_index_ratio_anchors = previous["il_index_ratio_anchors"]
            print(f"  fetch returned nothing (no file in {IL_GILTS_RAW_DIR} "
                  f"on this machine) -- keeping the anchors already in "
                  f"gilt_yields.json rather than overwriting them with an "
                  f"empty result.")
    else:
        il_index_ratio_anchors = (previous or {}).get("il_index_ratio_anchors", {})
        print(f"\n=== Skipping IL index-ratio anchors (--only) -- reusing "
              f"existing gilt_yields.json data ===")

    # --- HL daily gilt prices (security-level daily update, no manual
    # Tradeweb download needed) ---
    if "hl-prices" in components:
        print("\n=== Fetching Hargreaves Lansdown daily gilt prices ===")
        hl_data = fetch_hl_gilt_prices()
        if hl_data:
            securities = apply_hl_prices(securities, hl_data, il_index_ratio_anchors)
        else:
            print("  no HL prices fetched -- leaving securities unchanged "
                  "for today.")
    else:
        print("\n=== Skipping HL daily prices (--only) ===")

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
            "igcalc.pdf) -- computed client-side, not by this script. "
            "Each security's most recent series point may come from "
            "Hargreaves Lansdown's public daily gilt-price pages (no DMO/"
            "Tradeweb equivalent exists post-cutover) rather than DMO -- "
            "dirty_price is computed from HL's clean price plus accrued "
            "interest (Actual/Actual, index-ratio-adjusted for linkers "
            "using the DMO anchor above; the UK ex-dividend period isn't "
            "modelled). yield is HL's gross YTM for conventional gilts, "
            "null for linkers (HL doesn't publish one)."
        ),
        "securities": securities,
        "curves": curves,
        "rpi_history": rpi_history,
        "il_index_ratio_anchors": il_index_ratio_anchors,
    }

    with open("gilt_yields.json", "w") as f:
        json.dump(output, f, separators=(",", ":"))

    curve_summary = ", ".join(
        f"{c}={len(curves.get(c, {}).get('spot', []))}sp+"
        f"{len(curves.get(c, {}).get('forward', []))}fwd"
        for c in BOE_CURVE_TYPES)
    print(f"\nWrote gilt_yields.json: {len(securities)} securities, "
          f"curves: {curve_summary}, rpi_history: {len(rpi_history)} months, "
          f"il_index_ratio_anchors: {len(il_index_ratio_anchors.get('ratios', {}))} ISINs.")


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
