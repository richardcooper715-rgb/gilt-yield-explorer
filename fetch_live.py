#!/usr/bin/env python3
"""
fetch_live.py

Snapshot of intraday benchmark gilt yields for gilt-live.html, written to
live_gilts.json. Run every 15 minutes during UK market hours by
.github/workflows/live-quotes.yml, which publishes the file to the
`live-data` branch (so the frequent updates don't clutter main's history or
trigger GitHub Pages rebuilds). The page reads it from raw.githubusercontent.com.

Sources (both unofficial public quote feeds, so either may change or stop
without notice -- the script records whatever it could get and carries on):

  1. MarketWatch benchmark gilts (Tullett Prebon), 1-50 years:
     the JSON quote service behind https://www.marketwatch.com/investing/bond/tmbmkgb-10y?countryCode=BX
     MarketWatch's own web pages refuse scripted requests, but the quote
     service doesn't. NOTE: its instrument details (ISIN, coupon, maturity)
     are stale -- e.g. "10 Year" is labelled 1.125% 2039 while the price and
     yield are clearly the current 10-year benchmark's. So the script works
     out which gilt each quote really is: the one (from gilt_latest.json)
     whose own price/yield maths turns MarketWatch's yield into MarketWatch's
     clean price. That pins it down to within 0.001 of price.

  2. CNBC benchmark gilts (2, 5, 10, 30, 40, 50 years), with the actual
     benchmark gilt's coupon and maturity:
     the quote service behind https://www.cnbc.com/quotes/UK5Y
     The page also calls this directly from the browser (it allows that), so
     the copy here is only a fallback.

Usage:
    python fetch_live.py [output_path]
"""

import calendar
import json
import os
import sys
from datetime import date, datetime, timedelta, timezone

import requests

HEADERS = {"User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
           "Accept": "application/json"}

MW_TENORS = ["01Y", "02Y", "03Y", "04Y", "05Y", "06Y", "07Y", "08Y", "09Y", "10Y",
             "12Y", "15Y", "20Y", "25Y", "30Y", "40Y", "50Y"]
# The public read-only entitlement token MarketWatch's own pages use.
MW_URL = ("https://api.wsj.net/api/dylan/quotes/v2/comp/quoteByDialect?dialect=official"
          "&needed=CompositeTrading|BondSpecific&MaxInstrumentMatches=1&accept=application/json"
          "&EntitlementToken=cecc4267a0194af89ca343805a3e57af&ckey=cecc4267a0&dialects=Charting&id=")

CNBC_SYMBOLS = ["UK1Y", "UK2Y", "UK3Y", "UK5Y", "UK7Y", "UK10Y", "UK15Y", "UK20Y", "UK30Y", "UK40Y", "UK50Y"]
CNBC_URL = ("https://quote.cnbc.com/quote-html-webservice/restQuote/symbolType/symbol?symbols={}"
            "&requestMethod=itv&noform=1&partnerId=2&fund=1&exthrs=1&output=json")


def _val(d, *path):
    for p in path:
        if not isinstance(d, dict):
            return None
        d = d.get(p)
    return d


def fetch_marketwatch():
    ids = ",".join(f"Bond-BX-TMBMKGB-{t}" for t in MW_TENORS)
    resp = requests.get(MW_URL + ids, headers={**HEADERS, "Referer": "https://www.marketwatch.com/"}, timeout=30)
    resp.raise_for_status()
    out = []
    for r in resp.json().get("InstrumentResponses", []):
        matches = r.get("Matches") or []
        if not matches:
            continue
        m = matches[0]
        tenor = r["RequestId"].rsplit("-", 1)[-1]
        y = _val(m, "CompositeTrading", "Last", "Price", "Value")
        if y is None:
            continue
        chg = _val(m, "CompositeTrading", "NetChange", "Value")
        out.append({
            "tenor": tenor,
            "years": int(tenor[:-1]) if tenor.endswith("Y") else int(tenor[:-1]) / 12,
            "name": _val(m, "Instrument", "CommonName"),
            "yield": y,
            "change": chg,
            "prev_close": (y - chg) if chg is not None else None,
            "open": _val(m, "CompositeTrading", "Open", "Value"),
            "high": _val(m, "CompositeTrading", "High", "Value"),
            "low": _val(m, "CompositeTrading", "Low", "Value"),
            "price": _val(m, "BondSpecific", "TradePrice", "Value"),
            "price_change": _val(m, "BondSpecific", "TradeNetChange", "Value"),
            "time_uk": _val(m, "CompositeTrading", "Last", "Time"),   # UK local time, no zone given
        })
    return out


def _add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y += d.year
    return date(y, m + 1, min(d.day, calendar.monthrange(y, m + 1)[1]))


def clean_price_from_yield(coupon, maturity, gross_yield, trade_date):
    """Clean price per 100 of a conventional gilt (semi-annual, DMO-style) at a gross redemption yield."""
    settle = trade_date + timedelta(days=1)
    nxt = maturity
    while _add_months(nxt, -6) > settle:
        nxt = _add_months(nxt, -6)
    last = _add_months(nxt, -6)
    frac = (nxt - settle).days / (nxt - last).days
    v = 1 / (1 + gross_yield / 200)
    pv, k = 0.0, 0
    while _add_months(nxt, 6 * k) <= maturity:
        pv += coupon / 2 * v ** (frac + k)
        k += 1
    pv += 100 * v ** (frac + k - 1)
    return pv - coupon / 2 * (settle - last).days / (nxt - last).days


def identify_marketwatch(rows, latest_path="gilt_latest.json"):
    """Add the actual gilt (name, coupon, maturity, ISIN) to each MarketWatch row where it can be pinned down."""
    try:
        with open(latest_path, encoding="utf-8") as f:
            secs = json.load(f).get("securities", [])
    except (OSError, ValueError):
        return
    today = datetime.now(timezone.utc).date()
    gilts = []
    for s in secs:
        if s.get("type") != "conventional" or s.get("coupon_pct") is None or "strip" in (s.get("name") or "").lower():
            continue
        mat = date.fromisoformat(s["redemption_date"])
        if mat > today + timedelta(days=30):
            gilts.append((s, mat))
    for r in rows:
        if r.get("price") is None or r.get("yield") is None:
            continue
        # only gilts with roughly the right maturity (a benchmark is never far from its tenor)
        lo, hi = (0.2, 2.5) if r["years"] <= 1 else (0.4 * r["years"], 1.8 * r["years"])
        near = [(s, mat) for s, mat in gilts if lo <= (mat - today).days / 365.25 <= hi]
        scored = sorted(((abs(clean_price_from_yield(s["coupon_pct"], mat, r["yield"], today) - r["price"]), s["isin"], s, mat)
                         for s, mat in near), key=lambda t: t[0])
        scored = [(e, s, mat) for e, _, s, mat in scored]
        if scored and scored[0][0] < 0.02 and (len(scored) == 1 or scored[1][0] > 0.03):
            _, s, mat = scored[0]
            r.update({"gilt": s["name"], "isin": s["isin"], "coupon": s["coupon_pct"], "maturity_date": mat.isoformat()})


def _pct(s):
    try:
        return float(str(s).replace("%", "").replace("+", ""))
    except (TypeError, ValueError):
        return None


def fetch_cnbc():
    resp = requests.get(CNBC_URL.format("|".join(CNBC_SYMBOLS)), headers=HEADERS, timeout=30)
    resp.raise_for_status()
    out = []
    for q in _val(resp.json(), "FormattedQuoteResult", "FormattedQuote") or []:
        if str(q.get("code")) != "0" or not q.get("last"):
            continue
        out.append({
            "symbol": q.get("symbol"),
            "name": q.get("name"),
            "yield": _pct(q.get("last")),
            "change": _pct(q.get("change")),
            "prev_close": _pct(q.get("previous_day_closing")),
            "open": _pct(q.get("open")), "high": _pct(q.get("high")), "low": _pct(q.get("low")),
            "price": _pct(q.get("bond_last_price")),
            "price_change": _pct(q.get("bond_change_price")),
            "coupon": _pct(q.get("coupon")),
            "maturity_date": q.get("maturity_date"),
            "time": q.get("last_time"),
        })
    return out


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else "live_gilts.json"
    result = {"fetched": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"), "sources": {}}
    here = os.path.dirname(os.path.abspath(__file__))
    for name, fn in (("marketwatch", fetch_marketwatch), ("cnbc", fetch_cnbc)):
        try:
            rows = fn()
            if name == "marketwatch":
                identify_marketwatch(rows, os.path.join(here, "gilt_latest.json"))
            result["sources"][name] = {"ok": True, "rows": rows}
            print(f"{name}: {len(rows)} benchmarks")
        except Exception as e:  # noqa: BLE001 -- keep going with whatever works
            result["sources"][name] = {"ok": False, "error": str(e)[:300], "rows": []}
            print(f"{name}: FAILED ({e})")
    if not any(s["rows"] for s in result["sources"].values()):
        print("No data from either source; not writing a file.")
        return 1
    with open(path, "w", encoding="utf-8") as f:
        json.dump(result, f, separators=(",", ":"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
