#!/usr/bin/env python3
"""
JPX Market Analysis - Market Data Fetcher (fetch_market.py)
============================================================
Fetches Nikkei 225 closing price and Nikkei VI from public sources.
Falls back gracefully if sources are unavailable.

Usage:
    python fetch_market.py [--date YYYYMMDD]
    
Output: prints JSON to stdout
    {"nikkei_close": 53739.0, "vi": 49.45, "source": "stooq"}
"""

import argparse
import json
import re
import sys
from datetime import datetime, timedelta
from urllib.request import urlopen, Request
from urllib.error import URLError


def fetch_stooq_nikkei():
    """Fetch Nikkei 225 from stooq.com (no API key needed)."""
    try:
        url = 'https://stooq.com/q/l/?s=^nkx&f=sd2t2ohlcv&h&e=csv'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8').strip()
        lines = text.split('\n')
        if len(lines) >= 2:
            parts = lines[1].split(',')
            # Format: Symbol,Date,Time,Open,High,Low,Close,Volume
            close = float(parts[6])
            if 20000 < close < 80000:
                return close
    except Exception as e:
        print('[fetch] stooq nikkei failed: %s' % e, file=sys.stderr)
    return None


def fetch_yahoo_nikkei():
    """Fetch Nikkei 225 from Yahoo Finance."""
    try:
        url = 'https://finance.yahoo.com/quote/%5EN225/'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8')
        # Look for regularMarketPrice in the page
        m = re.search(r'"regularMarketPrice":\{"raw":([\d.]+)', text)
        if m:
            val = float(m.group(1))
            if 20000 < val < 80000:
                return val
    except Exception as e:
        print('[fetch] yahoo nikkei failed: %s' % e, file=sys.stderr)
    return None


def fetch_stooq_vi():
    """Fetch Nikkei VI from stooq.com."""
    try:
        url = 'https://stooq.com/q/l/?s=^jniv&f=sd2t2ohlcv&h&e=csv'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8').strip()
        lines = text.split('\n')
        if len(lines) >= 2:
            parts = lines[1].split(',')
            close = float(parts[6])
            if 10 < close < 100:
                return close
    except Exception as e:
        print('[fetch] stooq VI failed: %s' % e, file=sys.stderr)
    return None


def run(args):
    nikkei = None
    vi = None
    source = 'none'

    # Try stooq first (most reliable, no auth)
    nikkei = fetch_stooq_nikkei()
    if nikkei:
        source = 'stooq'
    else:
        nikkei = fetch_yahoo_nikkei()
        if nikkei:
            source = 'yahoo'

    vi = fetch_stooq_vi()

    result = {
        'nikkei_close': nikkei,
        'vi': vi,
        'source': source,
        'timestamp': datetime.now().isoformat(),
    }

    print(json.dumps(result))

    # Also write to file for pipeline consumption
    if args.out:
        with open(args.out, 'w') as f:
            json.dump(result, f)
        print('[fetch] Written to %s' % args.out, file=sys.stderr)

    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Fetch Nikkei 225 close and VI')
    parser.add_argument('--out', default=None, help='Output JSON file path')
    args = parser.parse_args()
    run(args)
