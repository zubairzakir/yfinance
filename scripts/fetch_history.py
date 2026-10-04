"""Download the finest-grained price history Yahoo Finance offers for each ticker.

Yahoo limits how far back each bar size goes, so a full view is built from
several tiers, each at the smallest interval available for its time range:

    1d  - full history (back to listing)
    1h  - last ~730 days
    2m  - last ~60 days
    1m  - last ~30 days

Each tier is saved to <out>/<TICKER>/<interval>.<parquet|csv>. Re-running merges
new bars into existing files, so running this regularly accumulates intraday
history beyond Yahoo's rolling windows.

Requests go through a throttled session that backs off on HTTP 429, because
Yahoo blocks clients that send requests too quickly.

Usage:
    python scripts/fetch_history.py NVDA
    python scripts/fetch_history.py NVDA AAPL MSFT --format csv
    python scripts/fetch_history.py --tickers-file tickers.txt
"""

import argparse
import logging
import os
import sys
import threading
import time

import pandas as pd
import requests

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import yfinance as yf  # noqa: E402
from yfinance.exceptions import YFPricesMissingError, YFRateLimitError  # noqa: E402

logger = logging.getLogger("fetch_history")

DAY = 86400

# (interval, lookback in days or None for full history, max days per request).
# Lookbacks sit a day inside Yahoo's limits so the oldest chunk isn't rejected.
TIERS = [
    ("1d", None, None),
    ("1h", 729, 729),
    ("2m", 59, 59),
    ("1m", 29, 7),
]


class ThrottledSession(requests.Session):
    """Session that spaces out requests and retries on 429 / 5xx with backoff."""

    def __init__(self, min_interval=0.5, max_retries=6, backoff=5.0):
        super().__init__()
        self.min_interval = min_interval
        self.max_retries = max_retries
        self.backoff = backoff
        self.request_count = 0
        self._lock = threading.Lock()
        self._last = 0.0

    def _wait_turn(self):
        with self._lock:
            wait = self._last + self.min_interval - time.monotonic()
            if wait > 0:
                time.sleep(wait)
            self._last = time.monotonic()
            self.request_count += 1

    def request(self, method, url, *args, **kwargs):
        for attempt in range(self.max_retries + 1):
            self._wait_turn()
            resp = super().request(method, url, *args, **kwargs)
            if resp.status_code != 429 and resp.status_code < 500:
                return resp
            if attempt == self.max_retries:
                break
            delay = self.backoff * 2 ** attempt
            logger.warning("HTTP %s from Yahoo, retrying in %.0fs", resp.status_code, delay)
            time.sleep(delay)
        return resp


def _history(ticker, retries=3, **kwargs):
    """Ticker.history that raises instead of silently returning empty data."""
    for attempt in range(retries + 1):
        try:
            return ticker.history(auto_adjust=False, actions=True, raise_errors=True, **kwargs)
        except YFPricesMissingError:
            # Range with no trading (e.g. holidays) - genuinely empty.
            return pd.DataFrame()
        except YFRateLimitError:
            if attempt == retries:
                raise
            delay = 60 * (attempt + 1)
            logger.warning("%s rate limited, sleeping %ds", ticker.ticker, delay)
            time.sleep(delay)


def fetch_tier(ticker, interval, lookback_days, chunk_days):
    if lookback_days is None:
        return _history(ticker, period="max", interval=interval)

    end = int(time.time())
    start = end - lookback_days * DAY
    frames = []
    chunk_start = start
    while chunk_start < end:
        chunk_end = min(chunk_start + chunk_days * DAY, end)
        df = _history(ticker, start=chunk_start, end=chunk_end, interval=interval, prepost=True)
        if not df.empty:
            frames.append(df)
        chunk_start = chunk_end
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames)
    return df[~df.index.duplicated(keep="last")].sort_index()


def save(df, path, fmt):
    if os.path.exists(path):
        old = pd.read_parquet(path) if fmt == "parquet" else _read_csv(path, df.index.tz)
        df = pd.concat([old, df])
        df = df[~df.index.duplicated(keep="last")].sort_index()
    if fmt == "parquet":
        df.to_parquet(path)
    else:
        df.to_csv(path)
    return df


def _read_csv(path, tz):
    df = pd.read_csv(path, index_col=0)
    df.index = pd.to_datetime(df.index, utc=True).tz_convert(tz)
    return df


def fetch_ticker(symbol, out_dir, fmt, session):
    ticker = yf.Ticker(symbol, session=session)
    os.makedirs(os.path.join(out_dir, symbol), exist_ok=True)
    summary = []
    for interval, lookback, chunk in TIERS:
        t0 = time.monotonic()
        before = session.request_count
        df = fetch_tier(ticker, interval, lookback, chunk)
        if df.empty:
            logger.warning("%s %s: no data", symbol, interval)
            continue
        path = os.path.join(out_dir, symbol, f"{interval}.{fmt}")
        df = save(df, path, fmt)
        summary.append({
            "ticker": symbol,
            "interval": interval,
            "rows": len(df),
            "first": df.index[0],
            "last": df.index[-1],
            "requests": session.request_count - before,
            "seconds": round(time.monotonic() - t0, 1),
        })
    return summary


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("tickers", nargs="*", help="Ticker symbols")
    parser.add_argument("--tickers-file", help="File with one ticker per line")
    parser.add_argument("--out", default="data", help="Output directory (default: data)")
    parser.add_argument("--format", choices=["parquet", "csv"], default="parquet")
    parser.add_argument("--min-interval", type=float, default=0.5,
                        help="Minimum seconds between requests (default: 0.5)")
    args = parser.parse_args()

    tickers = list(args.tickers)
    if args.tickers_file:
        with open(args.tickers_file) as f:
            tickers += [line.strip().upper() for line in f if line.strip() and not line.startswith("#")]
    if not tickers:
        parser.error("no tickers given")

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    session = ThrottledSession(min_interval=args.min_interval)

    rows, failed = [], []
    t0 = time.monotonic()
    for i, symbol in enumerate(tickers, 1):
        logger.info("[%d/%d] %s", i, len(tickers), symbol)
        try:
            rows += fetch_ticker(symbol, args.out, args.format, session)
        except Exception as e:
            logger.error("%s failed: %s", symbol, e)
            failed.append(symbol)

    if rows:
        print(pd.DataFrame(rows).to_string(index=False))
    print(f"\n{len(tickers) - len(failed)}/{len(tickers)} tickers, "
          f"{session.request_count} requests, {time.monotonic() - t0:.1f}s")
    if failed:
        print("Failed:", " ".join(failed))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
