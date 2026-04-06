#!/usr/bin/env python3
"""
JPX Market Analysis - Assessment Generator (generate_assessment.py)
====================================================================
Reads data.json and generates ⑧ 総合評価 using Google Gemini API.

Usage:
    python generate_assessment.py --data data.json --key YOUR_API_KEY
    
Environment variable:
    GEMINI_API_KEY=YOUR_API_KEY python generate_assessment.py --data data.json
"""

import argparse
import json
import os
import sys
from urllib.request import urlopen, Request
from urllib.error import URLError


SYSTEM_PROMPT = """あなたはJPXデリバティブ市場の専門アナリストです。
以下のデータを統合し、日本語で簡潔な総合評価を生成してください。

【出力フォーマット】
以下の4項目を、それぞれ2-3文で記述してください。HTMLタグは使わないでください。

■ 需給の構図
先物建玉の増減（ラージ/mini/TOPIX）から、機関投資家と個人の方向性を分析。

■ 手口・建玉からの示唆
J-NET大口取引と建玉変化（±300枚以上）から、市場参加者の意図を読み取る。
建玉壁（P壁/C壁）とMax Painの位置関係にも言及。

■ 参加者ポジション（データがある場合）
海外vs国内の構図。主要参加者の推定戦略。

■ 結論
サポート/レジスタンスの具体的な価格水準、短期予想レンジ、注意すべきリスク要因。
"""


def build_data_summary(data):
    """Extract key data points for the LLM prompt."""
    meta = data.get('metadata', {})
    s01 = data.get('s01', {})
    s02 = data.get('s02', {})
    s03 = data.get('s03', {})
    s04 = data.get('s04', {})
    s05 = data.get('s05', [])
    s06 = data.get('s06', {})
    s07 = data.get('s07', [])
    s09 = data.get('s09', {})
    ind = data.get('indicators', {})
    ohlc = meta.get('ohlc', {})

    lines = []
    lines.append('日付: %s' % meta.get('date_formatted', ''))
    lines.append('ATM（先物清算値）: %s' % meta.get('atm', ''))
    lines.append('日経平均終値: %s' % s01.get('nikkei_close', ''))
    lines.append('VI: %s' % s01.get('vi', ''))
    lines.append('SQ: %s（残り%d営業日）' % (meta.get('sq_label', ''), meta.get('days_to_sq', 0)))

    # OHLC
    if ohlc:
        lines.append('前日4本値: O=%s H=%s L=%s C=%s (値幅%s)' % (
            ohlc.get('open', ''), ohlc.get('high', ''), ohlc.get('low', ''), ohlc.get('close', ''), ohlc.get('range', '')))
        lines.append('ピボット: PP=%s R1=%s R2=%s S1=%s S2=%s' % (
            ohlc.get('pivot', ''), ohlc.get('r1', ''), ohlc.get('r2', ''), ohlc.get('s1', ''), ohlc.get('s2', '')))

    # Futures OI
    lines.append('')
    lines.append('【先物建玉増減】')
    for key, label in [('nk225_large', 'ラージ'), ('nk225_mini', 'mini'), ('topix', 'TOPIX')]:
        sec = s02.get(key, {})
        lines.append('%s: OI %s (前日比 %+d)' % (label, sec.get('total_oi', 0), sec.get('total_change', 0)))

    # Options
    lines.append('')
    lines.append('【OP建玉増減】')
    lg = s04.get('large', {})
    lines.append('ラージ: P %+d / C %+d' % (lg.get('put_total_change', 0), lg.get('call_total_change', 0)))

    # Indicators
    lines.append('')
    lines.append('【指標】')
    lines.append('PCR（取引高）: %s (%s)' % (ind.get('pcr_volume', ''), ind.get('pcr_signal', '')))
    lines.append('Max Pain: %s' % ind.get('max_pain', ''))

    # Distribution - top walls
    dist = s06.get('distribution', [])
    if dist:
        max_p = max(dist, key=lambda d: d['put_oi'])
        max_c = max(dist, key=lambda d: d['call_oi'])
        lines.append('P壁: %s (%d枚)' % (max_p['strike'], max_p['put_oi']))
        lines.append('C壁: %s (%d枚)' % (max_c['strike'], max_c['call_oi']))

    # Wall changes
    reinforced = ind.get('walls_reinforced', [])
    weakened = ind.get('walls_weakened', [])
    if reinforced:
        lines.append('壁の補強: %s' % ', '.join(['%s%d +%d' % (w['type'], w['strike'], w['change']) for w in reinforced[:3]]))
    if weakened:
        lines.append('壁の崩壊: %s' % ', '.join(['%s%d %d' % (w['type'], w['strike'], w['change']) for w in weakened[:3]]))

    # Important OI changes
    if s05:
        lines.append('')
        lines.append('【重要建玉変化（±300枚以上）】')
        for c in s05[:10]:
            lines.append('%s%d: %+d枚 (OI %d)' % (c['type'], c['strike'], c['change'], c.get('oi', 0)))

    # J-NET
    if s07:
        lines.append('')
        lines.append('【J-NET大口手口】')
        for t in s07[:8]:
            pair = '(対当)' if t.get('is_pair') else ''
            cat = {'us': '米系', 'eu': '欧系', 'hf': 'HF代理', 'domestic': '国内'}.get(t.get('category', ''), '')
            lines.append('%s: %s %d枚 [%s]%s' % (t.get('product', '')[:30], t.get('participant', ''), t.get('volume', 0), cat, pair))

    # Participants
    if 'error' not in s09:
        fut = s09.get('futures', {})
        nk = fut.get('nk225_large', {})
        lines.append('')
        lines.append('【参加者ポジション】')
        if nk:
            lines.append('N225: 海外Net %+d / 国内Net %+d' % (nk.get('overseas_net', 0), nk.get('domestic_net', 0)))
        profiles = s09.get('profiles', [])
        for p in profiles[:5]:
            lines.append('%s [%s]: N225=%+d TOPIX=%+d P=%+d C=%+d → %s' % (
                p['name'][:10], p['category_label'],
                p['nk225_large'], p['topix'], p['put_net'], p['call_net'], p['strategy']))
        if s09.get('source') == 'cache':
            lines.append('※ %s時点のキャッシュデータ' % s09.get('data_date', ''))

    return '\n'.join(lines)


def call_gemini(api_key, prompt, data_summary):
    """Call Google Gemini API."""
    url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=%s' % api_key

    payload = {
        'contents': [{
            'parts': [{'text': prompt + '\n\n--- データ ---\n' + data_summary}]
        }],
        'generationConfig': {
            'temperature': 0.3,
            'maxOutputTokens': 1500,
        }
    }

    body = json.dumps(payload).encode('utf-8')
    req = Request(url, data=body, headers={'Content-Type': 'application/json'})

    try:
        resp = urlopen(req, timeout=30)
        result = json.loads(resp.read().decode('utf-8'))
        # Extract text from response
        candidates = result.get('candidates', [])
        if candidates:
            parts = candidates[0].get('content', {}).get('parts', [])
            if parts:
                return parts[0].get('text', '')
        print('[assess] Gemini returned no text', file=sys.stderr)
        return None
    except URLError as e:
        print('[assess] Gemini API error: %s' % e, file=sys.stderr)
        return None
    except Exception as e:
        print('[assess] Unexpected error: %s' % e, file=sys.stderr)
        return None


def run(args):
    # Load data.json
    with open(args.data, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Get API key
    api_key = args.key or os.environ.get('GEMINI_API_KEY', '')
    if not api_key:
        print('[assess] No API key provided. Skipping assessment generation.', file=sys.stderr)
        return

    # Build prompt
    data_summary = build_data_summary(data)
    print('[assess] Data summary: %d chars' % len(data_summary), file=sys.stderr)

    # Call Gemini
    print('[assess] Calling Gemini API...', file=sys.stderr)
    assessment = call_gemini(api_key, SYSTEM_PROMPT, data_summary)

    if assessment:
        print('[assess] Generated %d chars' % len(assessment), file=sys.stderr)
        # Save back into data.json
        data['s08_assessment'] = assessment
        with open(args.data, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print('[assess] Saved to %s' % args.data, file=sys.stderr)
    else:
        print('[assess] Assessment generation failed', file=sys.stderr)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate market assessment via Gemini API')
    parser.add_argument('--data', default='data.json', help='Path to data.json')
    parser.add_argument('--key', default=None, help='Gemini API key (or set GEMINI_API_KEY env)')
    args = parser.parse_args()
    run(args)
