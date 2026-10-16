#!/usr/bin/env python3
"""
fetch_sonia.py

Keeps sonia.json up to date for the Gilt Holdings app (gilt-holdings.html),
which uses the Bank of England's SONIA Compounded Index (series IUDZOS2) as
a "what if the money had sat in cash at SONIA instead" comparator.

Source: Bank of England Interactive Statistical Database, which serves a
plain CSV for any series code without a login:
    https://www.bankofengland.co.uk/boeapps/database/

The script merges what it downloads into the existing sonia.json (new
values win), so a failed or partial download never loses history. If the
BoE request fails for any reason it prints why and exits 0, leaving
sonia.json untouched -- the rest of the daily workflow carries on.

Usage:
    python fetch_sonia.py
"""

import csv
import io
import json
import os
import sys
from datetime import datetime, timezone

import requests

OUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sonia.json")
SERIES = "IUDZOS2"
# The index starts at 100 on 23 Apr 2018.
URL = ("https://www.bankofengland.co.uk/boeapps/database/_iadb-fromshowcolumns.asp"
       "?csv.x=yes&Datefrom=23/Apr/2018&Dateto=now&SeriesCodes=" + SERIES +
       "&CSVF=TN&UsingCodes=Y&VPD=Y&VFD=N")
HEADERS = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
           "Accept": "text/csv,text/plain,*/*"}


def load_existing():
    try:
        with open(OUT_FILE, encoding="utf-8") as f:
            return {d: v for d, v in json.load(f).get("data", [])}
    except (OSError, ValueError):
        return {}


def parse_boe_csv(text):
    """BoE's CSV is 'DATE,IUDZOS2' then rows like '23 Apr 2018,100.0000'."""
    rows = {}
    reader = csv.reader(io.StringIO(text))
    header = next(reader, None)
    if not header or SERIES not in [h.strip().upper() for h in header]:
        raise ValueError(f"unexpected header: {header!r} (first 200 chars: {text[:200]!r})")
    col = [h.strip().upper() for h in header].index(SERIES)
    for r in reader:
        if len(r) <= col or not r[0].strip() or not r[col].strip():
            continue
        d = None
        for fmt in ("%d %b %Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                d = datetime.strptime(r[0].strip(), fmt).strftime("%Y-%m-%d")
                break
            except ValueError:
                pass
        if d is None:
            continue
        rows[d] = round(float(r[col]), 8)
    return rows


def main():
    existing = load_existing()
    try:
        resp = requests.get(URL, headers=HEADERS, timeout=60)
        resp.raise_for_status()
        fresh = parse_boe_csv(resp.text)
    except Exception as e:  # noqa: BLE001 -- never break the workflow over this
        print(f"SONIA: download failed, sonia.json left unchanged ({e})")
        return 0
    if not fresh:
        print("SONIA: BoE returned no rows, sonia.json left unchanged")
        return 0
    merged = {**existing, **fresh}
    data = sorted(merged.items())
    out = {
        "generated": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "series_code": SERIES,
        "name": "SONIA Compounded Index",
        "source": "Bank of England Database (https://www.bankofengland.co.uk/boeapps/database/), series IUDZOS2",
        "data": [[d, v] for d, v in data],
    }
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, separators=(",", ":"))
    print(f"SONIA: {len(fresh)} rows downloaded, {len(data)} in sonia.json, latest {data[-1][0]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
