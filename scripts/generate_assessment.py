#!/usr/bin/env python3
"""
JPX Market Analysis - Assessment Generator (generate_assessment.py)
====================================================================
Reads data.json and generates ⑧ 総合評価 using Google Gemini API.

Usage:
    python generate_assessment.py --data data.json [--key YOUR_API_KEY]
"""

import argparse
import json
import os
import sys
import time
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError


# Model: gemini-2.5-flash (stable, free tier)
GEMINI_MODEL = 'gemini-2.5-flash'

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
    lines = []
    meta = data.get('metadata', {})
    s01 = data.get('s01', {})
    lines.append('日付: %s' % meta.get('date_formatted', ''))
    lines.append('ATM: %s' % meta.get('atm', ''))
    lines.append('日経平均終値: %s' % s01.get('nikkei_close', ''))
    lines.append('VI: %s' % s01.get('vi', 'N/A'))
    lines.append('SQ: %s (残り%d営業日)' % (meta.get('sq_label', ''), meta.get('days_to_sq', 0)))

    # Futures OI
    s02 = data.get('s02', {})
    if 'error' not in s02:
        lines.append('')
        lines.append('--- 先物建玉 ---')
        for key, label in [('nk225_large', 'ラージ'), ('nk225_mini', 'mini'), ('topix', 'TOPIX')]:
            sec = s02.get(key, {})
            lines.append('%s: 前日比 %+d' % (label, sec.get('total_change', 0)))

    # Options OI changes
    s04 = data.get('s04', {})
    if 'error' not in s04:
        lg = s04.get('large', {})
        lines.append('')
        lines.append('--- OP建玉増減 ---')
        lines.append('P合計: %+d / C合計: %+d' % (lg.get('put_total_change', 0), lg.get('call_total_change', 0)))

    # Important OI changes
    s05 = data.get('s05', [])
    if s05:
        lines.append('')
        lines.append('--- 重要建玉変化(±300) ---')
        for c in s05[:10]:
            lines.append('%s %s: %+d (OI: %d)' % (c['type'], c['strike'], c['change'], c.get('oi', 0)))

    # OI distribution
    s06 = data.get('s06', {})
    dist = s06.get('distribution', [])
    if dist:
        lines.append('')
        lines.append('--- 建玉分布 ---')
        max_p = max(dist, key=lambda d: d['put_oi'])
        max_c = max(dist, key=lambda d: d['call_oi'])
        lines.append('P壁: %d (%d枚)' % (max_p['strike'], max_p['put_oi']))
        lines.append('C壁: %d (%d枚)' % (max_c['strike'], max_c['call_oi']))

    # J-NET
    s07 = data.get('s07', [])
    if s07:
        lines.append('')
        lines.append('--- J-NET大口 ---')
        for t in s07[:8]:
            pair = ' [対当]' if t.get('is_pair') else ''
            lines.append('%s: %s %d枚%s' % (t['product'][:30], t['participant'], t['volume'], pair))

    # Indicators
    ind = data.get('indicators', {})
    if ind:
        lines.append('')
        lines.append('--- 指標 ---')
        if ind.get('pcr_volume') is not None:
            lines.append('PCR(取引高): %.2f (%s)' % (ind['pcr_volume'], ind.get('pcr_signal', '')))
        if ind.get('max_pain'):
            lines.append('Max Pain: %s (ATM差: %+d)' % (ind['max_pain'], ind.get('max_pain_diff', 0)))
        if ind.get('basis') is not None:
            lines.append('先物ベーシス: %+d (%s)' % (ind['basis'], ind.get('basis_signal', '')))

    # Participant analysis
    s09 = data.get('s09', {})
    if 'error' not in s09:
        profiles = s09.get('profiles', [])
        if profiles:
            lines.append('')
            lines.append('--- 参加者ポジション ---')
            for p in profiles[:8]:
                lines.append('%s [%s]: N225=%+d mini=%+d TOPIX=%+d P=%+d C=%+d → %s' % (
                    p['name'], p['category_label'],
                    p['nk225_large'], p['nk225_mini'], p['topix'],
                    p['put_net'], p['call_net'], p['strategy']))

    return '\n'.join(lines)


def call_gemini(api_key, prompt, data_summary):
    """Call Google Gemini API with retry."""
    import time
    url = 'https://generativelanguage.googleapis.com/v1beta/models/%s:generateContent?key=%s' % (GEMINI_MODEL, api_key)

    payload = {
        'contents': [{
            'parts': [{'text': prompt + '\n\n--- データ ---\n' + data_summary}]
        }],
        'generationConfig': {
            'temperature': 0.3,
            'maxOutputTokens': 2000,
        }
    }

    body = json.dumps(payload).encode('utf-8')

    for attempt in range(3):
        try:
            req = Request(url, data=body, headers={
                'Content-Type': 'application/json',
            })
            resp = urlopen(req, timeout=60)
            result = json.loads(resp.read().decode('utf-8'))
            candidates = result.get('candidates', [])
            if candidates:
                parts = candidates[0].get('content', {}).get('parts', [])
                if parts:
                    return parts[0].get('text', '')
            print('[assess] Gemini returned no text', file=sys.stderr)
            print('[assess] Response: %s' % json.dumps(result, ensure_ascii=False)[:500], file=sys.stderr)
            return None
        except HTTPError as e:
            code = e.code
            body_text = ''
            try:
                body_text = e.read().decode('utf-8')[:300]
            except:
                pass
            if code == 429:
                wait = (attempt + 1) * 15
                print('[assess] Rate limited (429). Retrying in %ds... (%d/3)' % (wait, attempt + 1), file=sys.stderr)
                time.sleep(wait)
            elif code == 400:
                print('[assess] Bad request (400): %s' % body_text, file=sys.stderr)
                return None
            elif code == 403:
                print('[assess] Forbidden (403) - check API key permissions: %s' % body_text, file=sys.stderr)
                return None
            else:
                print('[assess] HTTP %d: %s' % (code, body_text), file=sys.stderr)
                return None
        except URLError as e:
            print('[assess] Network error: %s' % e, file=sys.stderr)
            return None
        except Exception as e:
            print('[assess] Unexpected error: %s' % e, file=sys.stderr)
            return None

    print('[assess] All retries failed', file=sys.stderr)
    return None


def run(args):
    api_key = args.key or os.environ.get('GEMINI_API_KEY', '')
    if not api_key:
        print('[assess] No API key. Skipping.', file=sys.stderr)
        return

    with open(args.data, 'r', encoding='utf-8') as f:
        data = json.load(f)

    data_summary = build_data_summary(data)
    print('[assess] Data summary: %d chars' % len(data_summary), file=sys.stderr)
    print('[assess] Using model: %s' % GEMINI_MODEL, file=sys.stderr)
    print('[assess] Calling Gemini API...', file=sys.stderr)

    assessment = call_gemini(api_key, SYSTEM_PROMPT, data_summary)

    if assessment:
        print('[assess] Generated %d chars' % len(assessment), file=sys.stderr)
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
