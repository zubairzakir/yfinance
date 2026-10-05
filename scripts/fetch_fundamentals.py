"""Download company profile, dividend, earnings and fund data for resolved symbols.

Run after build_universe.py. For every Yahoo symbol in <out>/_state/rows.json,
makes one quoteSummary request (profile, dividends, calendar, earnings,
estimates, fund details) and saves the raw response to
<out>/_state/quotesummary/<SYMBOL>.json. Existing files are skipped, so
re-running resumes.

quoteSummary needs a cookie + crumb; this gets them the way a browser does
(fc.yahoo.com, then /v1/test/getcrumb) and refreshes them on 401.

Usage:
    python scripts/fetch_fundamentals.py --out Tickers_out
"""

import argparse
import json
import logging
import os
import sys
import time

sys.path.insert(0, os.path.dirname(__file__))
from build_universe import Yahoo  # noqa: E402

logger = logging.getLogger("fetch_fundamentals")

QUOTE_SUMMARY_URL = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/{}"
MODULES = ",".join([
    "assetProfile", "summaryDetail", "calendarEvents", "earnings", "earningsHistory",
    "earningsTrend", "defaultKeyStatistics", "financialData", "price", "fundProfile", "quoteType",
])


class YahooWithCrumb(Yahoo):
    def __init__(self, min_interval):
        super().__init__(min_interval)
        self.crumb = None

    def refresh_crumb(self):
        self.s.cookies.clear()
        self.get("https://fc.yahoo.com", {})  # sets the A3 cookie (returns 404, that's fine)
        r = self.get("https://query1.finance.yahoo.com/v1/test/getcrumb", {})
        if r.status_code != 200 or not r.text or "<" in r.text:
            raise RuntimeError(f"could not get crumb: HTTP {r.status_code} {r.text[:80]}")
        self.crumb = r.text

    def quote_summary(self, symbol):
        if not self.crumb:
            self.refresh_crumb()
        for _ in range(2):
            r = self.get(QUOTE_SUMMARY_URL.format(symbol), {"modules": MODULES, "crumb": self.crumb})
            if r.status_code == 401:
                logger.info("crumb expired, refreshing")
                self.refresh_crumb()
                continue
            if r.status_code == 404:
                return None
            body = r.json().get("quoteSummary", {})
            return (body.get("result") or [None])[0]
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", required=True)
    ap.add_argument("--min-interval", type=float, default=2.0)
    args = ap.parse_args()

    qs_dir = os.path.join(args.out, "_state", "quotesummary")
    os.makedirs(qs_dir, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s",
                        handlers=[logging.StreamHandler(),
                                  logging.FileHandler(os.path.join(args.out, "_state", "fundamentals.log"))])

    rows = json.load(open(os.path.join(args.out, "_state", "rows.json")))
    symbols = sorted({r["Yahoo Symbol"] for r in rows.values() if r.get("Yahoo Symbol")})
    yh = YahooWithCrumb(args.min_interval)
    t0 = time.monotonic()
    for n, sym in enumerate(symbols, 1):
        path = os.path.join(qs_dir, f"{sym}.json")
        if os.path.exists(path):
            continue
        res = yh.quote_summary(sym)
        with open(path + ".tmp", "w") as f:
            json.dump(res, f)
        os.replace(path + ".tmp", path)
        logger.info("[%d/%d] %-14s %s  (%d requests, %.0fs)", n, len(symbols), sym,
                    ",".join(k for k in (res or {}) if k != "quoteType") or "NO DATA",
                    yh.requests, time.monotonic() - t0)
    print(f"done: {len(symbols)} symbols, {yh.requests} requests, {time.monotonic() - t0:.0f}s")


if __name__ == "__main__":
    sys.exit(main())
