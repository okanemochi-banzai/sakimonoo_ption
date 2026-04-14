#!/usr/bin/env python3
"""
JPX Market Analysis - Market Data Fetcher (fetch_market.py)
============================================================
Fetches Nikkei 225 close, VI, and 25-day OHLC history.

Usage:
    python fetch_market.py [--out market_latest.json]
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, timedelta
from urllib.request import urlopen, Request
from urllib.error import URLError


def fetch_stooq_nikkei():
    """Fetch Nikkei 225 cash close from stooq.com."""
    try:
        url = 'https://stooq.com/q/l/?s=^nkx&f=sd2t2ohlcv&h&e=csv'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8').strip()
        lines = text.split('\n')
        if len(lines) >= 2:
            parts = lines[1].split(',')
            close = float(parts[6])
            if 20000 < close < 80000:
                return close
    except Exception as e:
        print('[fetch] stooq nikkei failed: %s' % e, file=sys.stderr)
    return None


def fetch_yahoo_nikkei():
    """Fallback: Fetch from Yahoo Finance Japan."""
    try:
        url = 'https://finance.yahoo.co.jp/quote/998407.O'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8')
        m = re.search(r'([\d,]+\.\d+)', text)
        if m:
            val = float(m.group(1).replace(',', ''))
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


def fetch_investing_vi():
    """Fetch Nikkei VI from Investing.com."""
    try:
        url = 'https://jp.investing.com/indices/nikkei-volatility'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8')
        m = re.search(r'data-test="instrument-price-last"[^>]*>([\d,.]+)', text)
        if m:
            val = float(m.group(1).replace(',', ''))
            if 10 < val < 100:
                return val
    except Exception as e:
        print('[fetch] investing VI failed: %s' % e, file=sys.stderr)
    return None


def fetch_yahoo_vi():
    """Fetch Nikkei VI from Yahoo Finance Japan."""
    try:
        url = 'https://finance.yahoo.co.jp/quote/2035.T'
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=10)
        text = resp.read().decode('utf-8')
        m = re.search(r'<span[^>]*class="[^"]*StyledNumber[^"]*"[^>]*>([\d,.]+)</span>', text)
        if m:
            val = float(m.group(1).replace(',', ''))
            if 10 < val < 100:
                return val
    except Exception as e:
        print('[fetch] yahoo VI failed: %s' % e, file=sys.stderr)
    return None


def fetch_stooq_ohlc_history(days=30):
    """Fetch N225 daily OHLC history from stooq (last N calendar days).

    Returns list of {date, open, high, low, close} sorted oldest-first.
    """
    try:
        end = datetime.now()
        start = end - timedelta(days=days + 10)  # extra buffer for weekends/holidays
        d1 = start.strftime('%Y%m%d')
        d2 = end.strftime('%Y%m%d')
        url = 'https://stooq.com/q/d/l/?s=^nkx&d1=%s&d2=%s&i=d' % (d1, d2)
        req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urlopen(req, timeout=15)
        text = resp.read().decode('utf-8').strip()
        lines = text.split('\n')

        # CSV: Date,Open,High,Low,Close,Volume
        history = []
        for line in lines[1:]:  # skip header
            parts = line.strip().split(',')
            if len(parts) < 5:
                continue
            try:
                o = float(parts[1])
                h = float(parts[2])
                l = float(parts[3])
                c = float(parts[4])
                if 20000 < c < 80000:
                    history.append({
                        'date': parts[0],
                        'open': o,
                        'high': h,
                        'low': l,
                        'close': c,
                    })
            except (ValueError, IndexError):
                continue

        # Sort oldest first
        history.sort(key=lambda x: x['date'])

        # Keep last N trading days
        history = history[-days:]
        print('[fetch] OHLC history: %d trading days' % len(history), file=sys.stderr)
        return history

    except Exception as e:
        print('[fetch] stooq OHLC history failed: %s' % e, file=sys.stderr)
    return []


def run(args):
    nikkei = None
    vi = None
    source = 'none'

    # Nikkei close
    nikkei = fetch_stooq_nikkei()
    if nikkei:
        source = 'stooq'
    else:
        nikkei = fetch_yahoo_nikkei()
        if nikkei:
            source = 'yahoo'

    # VI (try multiple sources)
    vi = fetch_stooq_vi()
    if vi:
        print('[fetch] VI from stooq: %.2f' % vi, file=sys.stderr)
    else:
        vi = fetch_investing_vi()
        if vi:
            print('[fetch] VI from investing.com: %.2f' % vi, file=sys.stderr)
        else:
            vi = fetch_yahoo_vi()
            if vi:
                print('[fetch] VI from yahoo: %.2f' % vi, file=sys.stderr)
            else:
                print('[fetch] VI unavailable from all sources', file=sys.stderr)

    # OHLC history (25 trading days for technical analysis)
    ohlc_history = fetch_stooq_ohlc_history(25)

    result = {
        'nikkei_close': nikkei,
        'vi': vi,
        'source': source,
        'timestamp': datetime.now().isoformat(),
        'ohlc_history': ohlc_history,
    }

    print(json.dumps({k: v for k, v in result.items() if k != 'ohlc_history'}))

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
