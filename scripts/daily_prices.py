"""Append each completed trading day's prices to one CSV.

Reads the instruments from Tickers_4Oct26_updated.xlsx (rows with Status
"Listed") in the same folder, and appends one row per instrument per trading
day to daily_prices.csv: open, high, low, close, adjusted close, volume,
dividend, split, and change from the previous close.

Each run picks up where the last one stopped. For every instrument it adds
the trading days whose session closed after the last day already in the
file, up to 30 minutes before now (giving Yahoo time to finalise the day).
On the very first run, an instrument gets the sessions that closed in the
last 24 hours. So nothing is skipped or written twice, even if a run is
late, missed, or interrupted.

Usage (from the folder holding the spreadsheet):
    python daily_prices.py                  # normal daily run
    python daily_prices.py --since 2026-10-01   # first run: also backfill from a date
"""

import argparse
import csv
import datetime as dt
import os
import sys
import time

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
SPREADSHEET = os.path.join(HERE, "Tickers_4Oct26_updated.xlsx")
OUTPUT = os.path.join(HERE, "daily_prices.csv")
LOG = os.path.join(HERE, "daily_prices.log")
CHART_URL = "https://query2.finance.yahoo.com/v8/finance/chart/{}"
SETTLE = dt.timedelta(minutes=30)
FIRST_RUN_LOOKBACK = dt.timedelta(hours=24)

COLUMNS = ["Trading Date", "Yahoo Symbol", "Name", "Exchange", "Currency", "Price Unit",
           "Open", "High", "Low", "Close", "Adj Close", "Volume", "Previous Close", "Change", "Change %",
           "Dividend", "Stock Split", "Session Close (UTC)", "Added At"]


def log(msg):
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    print(line, flush=True)
    with open(LOG, "a") as f:
        f.write(line + "\n")


class Yahoo:
    def __init__(self, min_interval=1.0):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                                        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
        self.min_interval = min_interval
        self.last = 0.0
        self.rate_limited = False

    def daily(self, symbol, start):
        params = {"period1": int(start.timestamp()), "period2": int(time.time()),
                  "interval": "1d", "events": "div,splits", "includeAdjustedClose": "true"}
        backoff = 30
        for _ in range(4):
            wait = self.last + self.min_interval - time.monotonic()
            if wait > 0:
                time.sleep(wait)
            self.last = time.monotonic()
            try:
                r = self.s.get(CHART_URL.format(symbol), params=params, timeout=30)
            except requests.RequestException as e:
                log(f"  {symbol}: network error ({e}), retrying in {backoff}s")
                time.sleep(backoff)
                backoff *= 2
                continue
            if r.status_code == 429 or r.status_code >= 500:
                log(f"  Yahoo busy (HTTP {r.status_code}), waiting {backoff}s")
                time.sleep(backoff)
                backoff *= 2
                self.rate_limited = True
                continue
            self.rate_limited = False
            if r.status_code != 200:
                return None
            try:
                return r.json()["chart"]["result"][0]
            except (ValueError, KeyError, IndexError, TypeError):
                return None
        return None


def sessions(res):
    """Yield (session_close_utc, trading_date, row_values) for each daily bar."""
    meta = res.get("meta", {})
    ts = res.get("timestamp") or []
    q = (res.get("indicators", {}).get("quote") or [{}])[0]
    adj = (res.get("indicators", {}).get("adjclose") or [{}])[0].get("adjclose") or [None] * len(ts)
    reg = (meta.get("currentTradingPeriod") or {}).get("regular") or {}
    length = (reg.get("end", 0) - reg.get("start", 0)) or 8 * 3600
    offset = reg.get("gmtoffset", meta.get("gmtoffset", 0))
    events = res.get("events", {})
    divs = {int(k): v["amount"] for k, v in events.get("dividends", {}).items()}
    splits = {int(k): v["numerator"] / v["denominator"] for k, v in events.get("splits", {}).items()}
    for i, t in enumerate(ts):
        close = q.get("close", [None])[i]
        if close is None:
            continue
        start = dt.datetime.fromtimestamp(t, dt.timezone.utc)
        date = (start + dt.timedelta(seconds=offset)).date()
        # Dividend/split events are stamped at the session start or local midnight; match by date.
        div = sum(a for k, a in divs.items() if (dt.datetime.fromtimestamp(k, dt.timezone.utc) + dt.timedelta(seconds=offset)).date() == date)
        spl = next((r for k, r in splits.items() if (dt.datetime.fromtimestamp(k, dt.timezone.utc) + dt.timedelta(seconds=offset)).date() == date), None)
        yield start + dt.timedelta(seconds=length), date, {
            "Open": q.get("open", [None])[i], "High": q.get("high", [None])[i], "Low": q.get("low", [None])[i],
            "Close": close, "Adj Close": adj[i], "Volume": q.get("volume", [None])[i],
            "Dividend": div or None, "Stock Split": spl,
        }


def last_dates():
    """Latest trading date already in the CSV, per symbol."""
    if not os.path.exists(OUTPUT):
        return {}
    df = pd.read_csv(OUTPUT, usecols=["Trading Date", "Yahoo Symbol", "Close"])
    df["Trading Date"] = pd.to_datetime(df["Trading Date"]).dt.date
    g = df.sort_values("Trading Date").groupby("Yahoo Symbol").last()
    return {s: (r["Trading Date"], r["Close"]) for s, r in g.iterrows()}


def r6(x):
    return round(x, 6) if isinstance(x, float) else x


def main(argv=None, now=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--since", help="For instruments not yet in the file, start from this date (YYYY-MM-DD) "
                                    "instead of the last 24 hours")
    ap.add_argument("--min-interval", type=float, default=1.0, help="Seconds between requests (default 1)")
    args = ap.parse_args(argv)

    if not os.path.exists(SPREADSHEET):
        sys.exit(f"Can't find {SPREADSHEET} - put this script in the same folder as the spreadsheet.")
    inst = pd.read_excel(SPREADSHEET)
    inst = inst[inst["Status"] == "Listed"].drop_duplicates("Yahoo Symbol")

    now = now or dt.datetime.now(dt.timezone.utc)
    cutoff = now - SETTLE
    first_run_from = (dt.datetime.fromisoformat(args.since).replace(tzinfo=dt.timezone.utc)
                      if args.since else now - FIRST_RUN_LOOKBACK)
    have = last_dates()
    added_at = dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")
    log(f"Run started: {len(inst)} instruments, "
        f"{'continuing from previous runs' if have else 'first run'}, data up to {cutoff:%Y-%m-%d %H:%M} UTC")

    yh = Yahoo(args.min_interval)
    new_file = not os.path.exists(OUTPUT)
    n_rows, failed, blocked, t0 = 0, [], 0, time.monotonic()
    with open(OUTPUT, "a", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        if new_file:
            w.writeheader()
        for k, (_, row) in enumerate(inst.iterrows(), 1):
            sym = row["Yahoo Symbol"]
            last_date, prev_close = have.get(sym, (None, None))
            # Fetch a few extra days so the first new row has a previous close.
            since = (dt.datetime.combine(last_date, dt.time(), dt.timezone.utc) if last_date else first_run_from)
            res = yh.daily(sym, since - dt.timedelta(days=10))
            if res is None:
                failed.append(sym)
                blocked = blocked + 1 if yh.rate_limited else 0
                if blocked >= 3:
                    log("Yahoo is refusing requests from this connection. Stopping - nothing is lost; "
                        "run again in an hour or so and it will carry on from here.")
                    break
                continue
            blocked = 0
            for close_utc, date, v in sessions(res):
                if close_utc > cutoff:
                    continue  # session still open or just closed - next run gets it
                is_new = date > last_date if last_date else close_utc > first_run_from
                if is_new:
                    change = v["Close"] - prev_close if prev_close else None
                    w.writerow({
                        "Trading Date": date.isoformat(), "Yahoo Symbol": sym, "Name": row["Yahoo Name"],
                        "Exchange": row["Exchange"], "Currency": row["Currency"], "Price Unit": row["Price Unit"],
                        **{c: r6(v[c]) for c in ("Open", "High", "Low", "Close", "Adj Close", "Volume",
                                                 "Dividend", "Stock Split")},
                        "Previous Close": r6(prev_close), "Change": r6(change),
                        "Change %": r6(change / prev_close) if change is not None and prev_close else None,
                        "Session Close (UTC)": close_utc.strftime("%Y-%m-%d %H:%M"), "Added At": added_at,
                    })
                    n_rows += 1
                prev_close = v["Close"]
            f.flush()
            if k % 50 == 0:
                log(f"  {k}/{len(inst)} done, {n_rows} rows added")

    log(f"Run finished in {time.monotonic() - t0:.0f}s: {n_rows} rows added to {os.path.basename(OUTPUT)}")
    if failed:
        log(f"No data this run for {len(failed)}: {' '.join(failed)} (they'll be retried next run)")


if __name__ == "__main__":
    main()
