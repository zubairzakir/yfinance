"""Download prices for every stock in symbols.csv since the last run.

Each run creates a folder named with today's date holding ONE Excel file:
    2026-10-05/prices_2026-10-05.xlsx
        Daily   - one row per stock per trading day: official open, high, low,
                  close, volume (incl. auction volume), previous close, change
        Minute  - every 1-minute bar (incl. pre/post-market) for every stock

The first run covers the last 24 hours. Each later run starts exactly where
the previous one ended. Yahoo keeps 1-minute data for ~30 days only.
"""

import datetime as dt
import os
import sys
import time
from zoneinfo import ZoneInfo

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
SYMBOLS = os.path.join(HERE, "symbols.csv")
LAST_RUN = os.path.join(HERE, ".last_run")
URL = "https://query2.finance.yahoo.com/v8/finance/chart/{}"
UTC = dt.timezone.utc
SETTLE = dt.timedelta(minutes=30)      # Yahoo delays some exchanges
MAX_MINUTE_AGE = dt.timedelta(days=29)
CHUNK = dt.timedelta(days=7)           # Yahoo serves max ~7 days of 1-minute bars per request
EXCEL_ROWS = 1_000_000

session = requests.Session()
session.headers["User-Agent"] = "Mozilla/5.0"
last_request = 0.0


def get(symbol, start, end, interval):
    """Yahoo chart data or None. Raises RuntimeError if Yahoo keeps refusing."""
    global last_request
    params = {"period1": int(start.timestamp()), "period2": int(end.timestamp()), "interval": interval,
              "events": "div,splits", "includePrePost": "true" if interval == "1m" else "false"}
    wait = 30
    for _ in range(4):
        time.sleep(max(0, last_request + 1.0 - time.monotonic()))
        last_request = time.monotonic()
        try:
            r = session.get(URL.format(symbol), params=params, timeout=30)
        except requests.RequestException:
            time.sleep(wait)
            wait *= 2
            continue
        if r.status_code == 429 or r.status_code >= 500:
            print(f"  Yahoo busy, waiting {wait}s")
            time.sleep(wait)
            wait *= 2
            continue
        try:
            return r.json()["chart"]["result"][0]
        except Exception:
            return None
    raise RuntimeError("Yahoo is refusing requests")


def daily_rows(sym, info, since, cutoff):
    res = get(sym, since - dt.timedelta(days=10), cutoff + SETTLE, "1d")
    if not res:
        return []
    meta, ts = res.get("meta", {}), res.get("timestamp") or []
    q = res["indicators"]["quote"][0]
    reg = (meta.get("currentTradingPeriod") or {}).get("regular") or {}
    length = dt.timedelta(seconds=(reg.get("end", 0) - reg.get("start", 0)) or 8 * 3600)
    tz = ZoneInfo(meta.get("exchangeTimezoneName") or "UTC")
    divs = {dt.datetime.fromtimestamp(int(k), tz).date(): v["amount"]
            for k, v in (res.get("events") or {}).get("dividends", {}).items()}
    rows, prev = [], None
    for i, t in enumerate(ts):
        close = q["close"][i]
        if close is None:
            continue
        opened = dt.datetime.fromtimestamp(t, UTC)
        closed = opened + length
        day = opened.astimezone(tz).date()
        if since < closed <= cutoff:
            rows.append({"Date": day, "Symbol": sym, **info,
                         "Open": q["open"][i], "High": q["high"][i], "Low": q["low"][i], "Close": close,
                         "Volume": q["volume"][i], "Previous Close": prev,
                         "Change %": (close / prev - 1) if prev else None, "Dividend": divs.get(day)})
        prev = close
    return rows


def minute_rows(sym, since, cutoff):
    rows, start = [], max(since, cutoff - MAX_MINUTE_AGE)
    while start < cutoff:
        end = min(start + CHUNK, cutoff)
        res = get(sym, start, end, "1m")
        if res and res.get("timestamp"):
            q = res["indicators"]["quote"][0]
            tz = ZoneInfo(res.get("meta", {}).get("exchangeTimezoneName") or "UTC")
            for i, t in enumerate(res["timestamp"]):
                bar = dt.datetime.fromtimestamp(t, UTC)
                if q["close"][i] is None or bar < start or bar + dt.timedelta(minutes=1) > end:
                    continue
                rows.append({"Symbol": sym, "Time (UTC)": bar.replace(tzinfo=None),
                             "Exchange Time": bar.astimezone(tz).replace(tzinfo=None),
                             "Open": q["open"][i], "High": q["high"][i], "Low": q["low"][i],
                             "Close": q["close"][i], "Volume": q["volume"][i]})
        start = end
    return rows


def main():
    stocks = pd.read_csv(SYMBOLS)
    now = dt.datetime.now(UTC)
    cutoff = (now - SETTLE).replace(second=0, microsecond=0)
    try:
        since = dt.datetime.fromisoformat(open(LAST_RUN).read().strip())
    except (OSError, ValueError):
        since = cutoff - dt.timedelta(hours=24)
    print(f"Getting prices from {since:%Y-%m-%d %H:%M} to {cutoff:%Y-%m-%d %H:%M} UTC for {len(stocks)} stocks")

    daily, minute, failed = [], [], []
    for n, s in enumerate(stocks.itertuples(index=False), 1):
        info = {"Name": s.Name, "Exchange": s.Exchange, "Currency": s.Currency, "Price Unit": s._4}
        try:
            daily += daily_rows(s._0, info, since, cutoff)
            minute += minute_rows(s._0, since, cutoff)
        except RuntimeError:
            sys.exit("Yahoo is refusing requests right now. Nothing was saved - try again in an hour.")
        except Exception as e:
            failed.append(s._0)
            print(f"  {s._0}: skipped ({e})")
        if n % 25 == 0:
            print(f"  {n}/{len(stocks)} stocks done")

    if not daily and not minute:
        print("Nothing new since the last run - no file created.")
        return
    today = dt.date.today().isoformat()
    folder = os.path.join(HERE, today)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"prices_{today}.xlsx")
    if os.path.exists(path):
        path = os.path.join(folder, f"prices_{today}_{dt.datetime.now():%H%M}.xlsx")
    print("Writing Excel file (takes a minute or two)...")
    minute_df = pd.DataFrame(minute)
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        pd.DataFrame(daily).to_excel(xl, sheet_name="Daily", index=False)
        for i in range(0, max(len(minute_df), 1), EXCEL_ROWS):
            name = "Minute" if i == 0 else f"Minute {i // EXCEL_ROWS + 1}"
            minute_df.iloc[i:i + EXCEL_ROWS].to_excel(xl, sheet_name=name, index=False)
    with open(LAST_RUN, "w") as f:
        f.write(cutoff.isoformat())
    print(f"\nDone: {len(daily)} daily rows, {len(minute):,} minute bars -> {os.path.relpath(path, HERE)}")
    if failed:
        print(f"Skipped {len(failed)}: {' '.join(failed)}")


if __name__ == "__main__":
    main()
