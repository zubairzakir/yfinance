"""Daily run: official daily prices plus 1-minute bars for every listed instrument.

Reads the instruments from Tickers_4Oct26_updated.xlsx (rows with Status
"Listed") in the same folder, and for each one appends:

  daily_prices.csv        One row per completed trading day: Yahoo's official
                          open, high, low, close, adjusted close and volume
                          (including auction volume), dividend, split, and
                          change from the previous close.
  minute_prices.csv       Every 1-minute bar for every instrument, including
                          pre/post-market, all in one file.

Each run continues from the last day / minute already saved for each
instrument, up to 30 minutes before now (Yahoo delays some exchanges). On the
first run an instrument gets the last 24 hours. So late, missed or interrupted
runs neither skip nor duplicate anything - but Yahoo only keeps about 30 days
of 1-minute data, so run at least every few weeks to keep minute history whole.

Usage (from the folder holding the spreadsheet):
    python daily_prices.py
    python daily_prices.py --since 2026-10-01   # first run: start from a date
"""

import argparse
import csv
import datetime as dt
import glob
import json
import os
import sys
import time
from zoneinfo import ZoneInfo

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
SPREADSHEET = os.path.join(HERE, "Tickers_4Oct26_updated.xlsx")
OUTPUT = os.path.join(HERE, "daily_prices.csv")
MINUTE_FILE = os.path.join(HERE, "minute_prices.csv")
PROGRESS = os.path.join(HERE, ".minute_progress.json")
OLD_MINUTE_DIR = os.path.join(HERE, "minute")  # earlier version: one file per stock
LOG = os.path.join(HERE, "daily_prices.log")
LOCK = os.path.join(HERE, ".daily_prices.lock")
CHART_URL = "https://query2.finance.yahoo.com/v8/finance/chart/{}"
SETTLE = dt.timedelta(minutes=30)
FIRST_RUN_LOOKBACK = dt.timedelta(hours=24)
MINUTE_HISTORY = dt.timedelta(days=29)   # Yahoo keeps ~30 days of 1-minute bars
MINUTE_CHUNK = dt.timedelta(days=7)      # and serves at most ~7 days per request
UTC = dt.timezone.utc

COLUMNS = ["Trading Date", "Yahoo Symbol", "Name", "Exchange", "Currency", "Price Unit",
           "Open", "High", "Low", "Close", "Adj Close", "Volume", "Previous Close", "Change", "Change %",
           "Dividend", "Stock Split", "Session Open (UTC)", "Session Close (UTC)", "Added At"]
MINUTE_COLUMNS = ["Yahoo Symbol", "Time (UTC)", "Exchange Time", "Open", "High", "Low", "Close", "Volume"]


def log(msg):
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    print(line, flush=True)
    with open(LOG, "a") as f:
        f.write(line + "\n")


class Yahoo:
    def __init__(self, min_interval=1.0):
        self.s = requests.Session()
        # A plain identifier: Yahoo refused a full browser string in testing.
        self.s.headers["User-Agent"] = "Mozilla/5.0"
        self.min_interval = min_interval
        self.last = 0.0
        self.rate_limited = False
        self.requests = 0

    def chart(self, symbol, start, end, interval):
        """Yahoo chart data, or None if unavailable. Sets rate_limited if Yahoo kept refusing."""
        params = {"period1": int(start.timestamp()), "period2": int(end.timestamp()), "interval": interval,
                  "events": "div,splits", "includeAdjustedClose": "true",
                  "includePrePost": "true" if interval != "1d" else "false"}
        backoff = 30
        for _ in range(4):
            wait = self.last + self.min_interval - time.monotonic()
            if wait > 0:
                time.sleep(wait)
            self.last = time.monotonic()
            self.requests += 1
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


# ---------------------------------------------------------------- daily bars

def sessions(res):
    """Yield (session_open_utc, session_close_utc, trading_date, row_values) for each daily bar."""
    meta = res.get("meta", {})
    ts = res.get("timestamp") or []
    q = (res.get("indicators", {}).get("quote") or [{}])[0]
    adj = (res.get("indicators", {}).get("adjclose") or [{}])[0].get("adjclose") or [None] * len(ts)
    reg = (meta.get("currentTradingPeriod") or {}).get("regular") or {}
    length = (reg.get("end", 0) - reg.get("start", 0)) or 8 * 3600
    offset = reg.get("gmtoffset", meta.get("gmtoffset", 0))

    def local_date(t):
        return (dt.datetime.fromtimestamp(int(t), UTC) + dt.timedelta(seconds=offset)).date()

    events = res.get("events", {})
    divs, splits = {}, {}
    for k, v in events.get("dividends", {}).items():
        divs[local_date(k)] = divs.get(local_date(k), 0) + v["amount"]
    for k, v in events.get("splits", {}).items():
        splits[local_date(k)] = v["numerator"] / v["denominator"]
    for i, t in enumerate(ts):
        close = q.get("close", [None])[i]
        if close is None:
            continue
        start = dt.datetime.fromtimestamp(t, UTC)
        date = local_date(t)
        yield start, start + dt.timedelta(seconds=length), date, {
            "Open": q.get("open", [None])[i], "High": q.get("high", [None])[i], "Low": q.get("low", [None])[i],
            "Close": close, "Adj Close": adj[i], "Volume": q.get("volume", [None])[i],
            "Dividend": divs.get(date), "Stock Split": splits.get(date),
        }


def last_dates():
    """Latest trading date and close already in daily_prices.csv, per symbol."""
    if not os.path.exists(OUTPUT):
        return {}
    df = pd.read_csv(OUTPUT, usecols=["Trading Date", "Yahoo Symbol", "Close"])
    df["Trading Date"] = pd.to_datetime(df["Trading Date"]).dt.date
    g = df.sort_values("Trading Date").groupby("Yahoo Symbol").last()
    return {s: (r["Trading Date"], r["Close"]) for s, r in g.iterrows()}


def r6(x):
    return round(x, 6) if isinstance(x, float) else x


def update_daily(yh, writer, row, have, cutoff, first_run_from, added_at):
    """Append new completed sessions for one instrument. Returns rows added, or None on failure."""
    sym = row["Yahoo Symbol"]
    last_date, prev_close = have.get(sym, (None, None))
    since = dt.datetime.combine(last_date, dt.time(), UTC) if last_date else first_run_from
    # A few extra days so the first new row has a previous close.
    res = yh.chart(sym, since - dt.timedelta(days=10), dt.datetime.now(UTC), "1d")
    if res is None:
        return None
    added = 0
    for open_utc, close_utc, date, v in sessions(res):
        if close_utc > cutoff:
            continue  # session still open or just closed - the next run gets it
        if (date > last_date) if last_date else (close_utc > first_run_from):
            change = v["Close"] - prev_close if prev_close else None
            writer.writerow({
                "Trading Date": date.isoformat(), "Yahoo Symbol": sym, "Name": row["Yahoo Name"],
                "Exchange": row["Exchange"], "Currency": row["Currency"], "Price Unit": row["Price Unit"],
                **{c: r6(v[c]) for c in ("Open", "High", "Low", "Close", "Adj Close", "Volume",
                                         "Dividend", "Stock Split")},
                "Previous Close": r6(prev_close), "Change": r6(change),
                "Change %": r6(change / prev_close) if change is not None and prev_close else None,
                "Session Open (UTC)": open_utc.strftime("%Y-%m-%d %H:%M"),
                "Session Close (UTC)": close_utc.strftime("%Y-%m-%d %H:%M"), "Added At": added_at,
            })
            added += 1
        prev_close = v["Close"]
    return added


# --------------------------------------------------------------- minute bars

def load_progress():
    """UTC time up to which each symbol's minute bars have been fetched."""
    try:
        with open(PROGRESS) as f:
            return {k: dt.datetime.fromisoformat(v) for k, v in json.load(f).items()}
    except (OSError, ValueError):
        pass
    if not os.path.exists(MINUTE_FILE):
        return {}
    # Progress file lost: rebuild it from the last bar per symbol in the minute file.
    df = pd.read_csv(MINUTE_FILE, usecols=["Yahoo Symbol", "Time (UTC)"])
    last = pd.to_datetime(df.groupby("Yahoo Symbol")["Time (UTC)"].max(), utc=True)
    return {s: t.to_pydatetime() + dt.timedelta(minutes=1) for s, t in last.items()}


def save_progress(progress):
    with open(PROGRESS + ".tmp", "w") as f:
        json.dump({k: v.isoformat() for k, v in progress.items()}, f, indent=0)
    os.replace(PROGRESS + ".tmp", PROGRESS)


def merge_old_minute_files():
    """Fold the per-stock files written by the earlier version into minute_prices.csv."""
    files = sorted(glob.glob(os.path.join(OLD_MINUTE_DIR, "*.csv")))
    if not files:
        return
    progress = {}
    old_progress = os.path.join(OLD_MINUTE_DIR, "_fetched_up_to.json")
    if os.path.exists(old_progress):
        with open(old_progress) as f:
            progress = {k: dt.datetime.fromisoformat(v) for k, v in json.load(f).items()}
    new_file = not os.path.exists(MINUTE_FILE)
    with open(MINUTE_FILE, "a", newline="") as out:
        w = csv.writer(out)
        if new_file:
            w.writerow(MINUTE_COLUMNS)
        for path in files:
            sym = os.path.basename(path)[:-4]
            with open(path, newline="") as f:
                rows = list(csv.reader(f))[1:]
            for r in rows:
                w.writerow([sym] + r)
            if sym not in progress and rows:
                progress[sym] = (dt.datetime.strptime(rows[-1][0], "%Y-%m-%d %H:%M").replace(tzinfo=UTC)
                                 + dt.timedelta(minutes=1))
    save_progress(progress)
    os.rename(OLD_MINUTE_DIR, OLD_MINUTE_DIR + "_old_merged_into_minute_prices")
    log(f"Merged {len(files)} per-stock minute files into {os.path.basename(MINUTE_FILE)} "
        f"(old folder renamed to minute_old_merged_into_minute_prices - safe to delete)")


def update_minutes(yh, w, sym, end, first_run_from, progress):
    """Append complete 1-minute bars up to `end` for one instrument via csv writer `w`.
    Returns bars added, or None if a request failed (the rest is retried next run)."""
    start = progress.get(sym) or first_run_from
    oldest = end - MINUTE_HISTORY
    if start < oldest:
        if start != first_run_from:
            log(f"  {sym}: minute data from {start:%Y-%m-%d %H:%M} to {oldest:%Y-%m-%d %H:%M} UTC "
                f"is no longer on Yahoo (over 30 days since the last run) - continuing from there")
        start = oldest
    added = 0
    chunk_start = start
    while chunk_start < end:
        chunk_end = min(chunk_start + MINUTE_CHUNK, end)
        res = yh.chart(sym, chunk_start, chunk_end, "1m")
        if res is None:
            return None
        tz = ZoneInfo(res.get("meta", {}).get("exchangeTimezoneName") or "UTC")
        q = (res.get("indicators", {}).get("quote") or [{}])[0]
        rows = {}
        for i, t in enumerate(res.get("timestamp") or []):
            bar = dt.datetime.fromtimestamp(t, UTC)
            close = q.get("close", [None])[i]
            # Only whole minutes inside this window; Yahoo can include a forming bar.
            if close is None or bar < chunk_start or bar + dt.timedelta(minutes=1) > chunk_end:
                continue
            rows[t] = [sym, bar.strftime("%Y-%m-%d %H:%M"), bar.astimezone(tz).strftime("%Y-%m-%d %H:%M"),
                       r6(q["open"][i]), r6(q["high"][i]), r6(q["low"][i]), r6(close), q["volume"][i]]
        for t in sorted(rows):
            w.writerow(rows[t])
        added += len(rows)
        progress[sym] = chunk_start = chunk_end
    return added


# ---------------------------------------------------------------------- main

def main(argv=None, now=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--since", help="For instruments not yet saved, start from this date (YYYY-MM-DD) "
                                    "instead of the last 24 hours")
    ap.add_argument("--min-interval", type=float, default=1.0, help="Seconds between requests (default 1)")
    args = ap.parse_args(argv)

    if not os.path.exists(SPREADSHEET):
        sys.exit(f"Can't find {SPREADSHEET} - put this script in the same folder as the spreadsheet.")
    if os.path.exists(LOCK) and time.time() - os.path.getmtime(LOCK) < 3 * 3600:
        sys.exit("Another run is still in progress (or crashed within the last 3 hours). "
                 f"If you're sure none is running, delete {LOCK} and try again.")
    open(LOCK, "w").close()
    try:
        run(args, now)
    finally:
        os.remove(LOCK)


def run(args, now):
    inst = pd.read_excel(SPREADSHEET)
    inst = inst[inst["Status"] == "Listed"].drop_duplicates("Yahoo Symbol")
    merge_old_minute_files()

    now = now or dt.datetime.now(UTC)
    cutoff = now - SETTLE
    minute_end = cutoff.replace(second=0, microsecond=0)
    first_run_from = (dt.datetime.fromisoformat(args.since).replace(tzinfo=UTC)
                      if args.since else now - FIRST_RUN_LOOKBACK)
    have = last_dates()
    progress = load_progress()
    added_at = dt.datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z")
    log(f"Run started: {len(inst)} instruments, {'continuing' if have else 'first run'}, "
        f"data up to {cutoff:%Y-%m-%d %H:%M} UTC")

    yh = Yahoo(args.min_interval)
    new_file = not os.path.exists(OUTPUT)
    new_minute_file = not os.path.exists(MINUTE_FILE)
    days = minutes = blocked = 0
    failed, t0 = [], time.monotonic()
    with open(OUTPUT, "a", newline="") as f, open(MINUTE_FILE, "a", newline="") as mf:
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        mw = csv.writer(mf)
        if new_file:
            w.writeheader()
        if new_minute_file:
            mw.writerow(MINUTE_COLUMNS)
        for k, (_, row) in enumerate(inst.iterrows(), 1):
            sym = row["Yahoo Symbol"]
            try:
                d = update_daily(yh, w, row, have, cutoff, first_run_from, added_at)
                f.flush()
                m = update_minutes(yh, mw, sym, minute_end, first_run_from, progress) if d is not None else None
                mf.flush()
                os.fsync(mf.fileno())
                save_progress(progress)  # only after the bars are safely on disk
            except Exception as e:  # one bad instrument must not stop the run
                log(f"  {sym}: unexpected error {type(e).__name__}: {e}")
                d = m = None
            days += d or 0
            minutes += m or 0
            if d is None or m is None:
                failed.append(sym)
                blocked = blocked + 1 if yh.rate_limited else 0
                if blocked >= 3:
                    log("Yahoo is refusing requests from this connection. Stopping - nothing is lost; "
                        "run again in an hour or so and it will carry on from here.")
                    break
            else:
                blocked = 0
            if k % 50 == 0:
                log(f"  {k}/{len(inst)} done: {days} daily rows, {minutes:,} minute bars")

    log(f"Run finished in {(time.monotonic() - t0) / 60:.0f} min ({yh.requests} requests): "
        f"{days} daily rows, {minutes:,} minute bars added")
    if failed:
        log(f"Incomplete this run for {len(failed)}: {' '.join(failed)} (they'll catch up next run)")


if __name__ == "__main__":
    main()
