"""Build hourly / daily (or any bar size) open-high-low-close-volume from the 1-minute files.

Usage (from the Tickers_4Oct26 folder):
    python make_views.py NVDA 1h            # hourly bars for NVDA -> views/NVDA_1h.csv
    python make_views.py NVDA 1d            # daily bars
    python make_views.py all 1h             # every stock, one file each
    python make_views.py GSK.L 15min --from 2026-10-01 --to 2026-10-31
    python make_views.py NVDA 1h --regular  # regular session only (no pre/post-market)

Bar sizes: 5min, 15min, 30min, 1h, 1d, 1w. Bars line up with each exchange's
opening time (e.g. 09:30-10:30 in New York). --regular uses the official
session times recorded in daily_prices.csv. Daily volume from minutes misses
auction trading; daily_prices.csv has the official daily figures.
"""

import argparse
import glob
import os

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
FREQ = {"5min": "5min", "15min": "15min", "30min": "30min", "1h": "1h", "1d": "1D", "1w": "W-FRI"}


def sessions(sym):
    """Official session open/close times (UTC) per day, from daily_prices.csv."""
    path = os.path.join(HERE, "daily_prices.csv")
    if not os.path.exists(path):
        return pd.DataFrame()
    d = pd.read_csv(path, usecols=["Yahoo Symbol", "Session Open (UTC)", "Session Close (UTC)"])
    d = d[d["Yahoo Symbol"] == sym]
    return pd.DataFrame({"open": pd.to_datetime(d["Session Open (UTC)"], utc=True),
                         "close": pd.to_datetime(d["Session Close (UTC)"], utc=True)})


def view(sym, size, start, end, regular):
    m = pd.read_csv(os.path.join(HERE, "minute", f"{sym}.csv"))
    if m.empty:
        return m
    utc = pd.to_datetime(m["Time (UTC)"], utc=True)
    local = pd.to_datetime(m["Exchange Time"])
    shift = local.iloc[-1] - utc.iloc[-1].tz_convert(None)  # exchange time minus UTC
    m = m.set_index(local)[["Open", "High", "Low", "Close", "Volume"]]
    sess = sessions(sym)
    if regular:
        if sess.empty:
            raise SystemExit(f"{sym}: no session times in daily_prices.csv yet - run daily_prices.py first")
        keep = pd.Series(False, index=range(len(m)))
        for o, c in zip(sess["open"], sess["close"]):
            keep |= ((utc >= o) & (utc < c)).to_numpy()
        m = m[keep.to_numpy()]
    if start:
        m = m[m.index >= start]
    if end:
        m = m[m.index < pd.Timestamp(end) + pd.Timedelta(days=1)]
    if m.empty:
        return m
    offset = None
    if size in ("5min", "15min", "30min", "1h"):
        # Line bars up with the exchange's opening minute (e.g. 09:30 in New York).
        step = int(pd.Timedelta(FREQ[size]).total_seconds() // 60)
        first = sess["open"].iloc[-1].tz_convert(None) + shift if not sess.empty else m.index.min()
        offset = pd.Timedelta(minutes=first.minute % step)
    bars = m.resample(FREQ[size], offset=offset).agg(
        {"Open": "first", "High": "max", "Low": "min", "Close": "last", "Volume": "sum"}).dropna(subset=["Close"])
    bars.index.name = "Bar Start (exchange time)"
    return bars


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("symbol", help="Yahoo symbol, e.g. NVDA or GSK.L, or 'all'")
    ap.add_argument("size", choices=FREQ)
    ap.add_argument("--from", dest="start")
    ap.add_argument("--to", dest="end")
    ap.add_argument("--regular", action="store_true", help="regular session only")
    args = ap.parse_args()

    syms = ([os.path.basename(p)[:-4] for p in sorted(glob.glob(os.path.join(HERE, "minute", "*.csv")))]
            if args.symbol.lower() == "all" else [args.symbol])
    os.makedirs(os.path.join(HERE, "views"), exist_ok=True)
    for sym in syms:
        try:
            bars = view(sym, args.size, args.start, args.end, args.regular)
        except FileNotFoundError:
            print(f"{sym}: no minute file yet")
            continue
        out = os.path.join(HERE, "views", f"{sym}_{args.size}{'_regular' if args.regular else ''}.csv")
        bars.round(6).to_csv(out)
        print(f"{sym}: {len(bars)} bars -> {os.path.relpath(out, HERE)}")


if __name__ == "__main__":
    main()
