#!/usr/bin/env python3
"""
JPX Market Analysis - Renderer (render.py)
============================================
Reads data.json from extract.py and generates:
  1. JPX_market_analysis_YYYYMMDD.md  (Markdown report)
  2. index.html                       (Dashboard)
  3. pnl_simulator.html               (P&L Simulator)
  4. JPX_portal_YYYYMMDD.html         (Archive copy)
  5. Archive snippet (stdout)

Usage:
    python render.py [--data data.json] [--outdir ./output]

Dependencies: None (pure Python, no external libraries)
"""

import argparse
import json
import os
import sys
import copy

# ============================================================
# Formatting Helpers
# ============================================================

def fnum(n, plus=False):
    """Format number with commas. Optionally prefix + for positive."""
    if n is None:
        return '-'
    if isinstance(n, float):
        if n == int(n):
            n = int(n)
        else:
            s = '{:,.1f}'.format(n)
            if plus and n > 0:
                s = '+' + s
            return s
    s = '{:,}'.format(int(n))
    if plus and n > 0:
        s = '+' + s
    return s


def fpct(n):
    """Format percentage."""
    if n is None:
        return '-'
    return '{:.1f}%'.format(n)


def sign_class(n):
    """Return 'positive' or 'negative' for CSS class."""
    if n is None:
        return ''
    return 'positive' if n > 0 else 'negative' if n < 0 else ''


def esc(s):
    """Escape HTML special characters."""
    if s is None:
        return ''
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


# ============================================================
# Markdown Report Builder
# ============================================================

def build_markdown(data):
    """Generate full Markdown report from data.json."""
    meta = data['metadata']
    md = []
    md.append('# JPX Market Analysis %s' % meta['date_formatted'])
    md.append('')
    md.append('> ATM: %s / VI: %s / %s / SQまで%d営業日' % (
        fnum(meta.get('atm')), meta.get('vi', '-'),
        meta.get('sq_label', ''), meta.get('days_to_sq', 0)))
    md.append('')

    # ① Nikkei / VI
    s01 = data.get('s01', {})
    md.append('## ① 日経平均・VI分析')
    md.append('')
    if s01.get('nikkei_close'):
        md.append('- 日経平均終値: **%s**' % fnum(s01['nikkei_close']))
    if s01.get('vi'):
        md.append('- 日経平均VI: **%s**' % s01['vi'])
    r1d = s01.get('range_1d', {})
    r1w = s01.get('range_1w', {})
    if r1d:
        md.append('- 1日予測値幅: %s (%s 〜 %s)' % (fnum(r1d.get('width')), fnum(r1d.get('low')), fnum(r1d.get('high'))))
    if r1w:
        md.append('- 1週予測値幅: %s (%s 〜 %s)' % (fnum(r1w.get('width')), fnum(r1w.get('low')), fnum(r1w.get('high'))))
    md.append('')

    # ② Futures OI
    s02 = data.get('s02', {})
    md.append('## ② 先物建玉残高')
    md.append('')
    if 'error' not in s02:
        md.append('| 銘柄 | 建玉残高 | 前日比 |')
        md.append('|------|---------|--------|')
        for key, label in [('nk225_large', '日経225ラージ'), ('nk225_mini', '日経225mini'), ('topix', 'TOPIX')]:
            sec = s02.get(key, {})
            md.append('| %s | %s | %s |' % (
                label, fnum(sec.get('total_oi')), fnum(sec.get('total_change'), plus=True)))

        # Monthly breakdown
        for key, label in [('nk225_large', 'ラージ'), ('nk225_mini', 'mini'), ('topix', 'TOPIX')]:
            sec = s02.get(key, {})
            months = sec.get('months', [])
            if months:
                md.append('')
                md.append('**%s 限月別**:' % label)
                for m in months:
                    md.append('- %s: OI %s (前日比 %s)' % (m['label'], fnum(m.get('oi')), fnum(m.get('change'), plus=True)))
    else:
        md.append('データなし')
    md.append('')

    # ③ Options Volume
    s03 = data.get('s03', {})
    md.append('## ③ オプション総取引代金')
    md.append('')
    if 'error' not in s03:
        for block_key, block_label in [('large', 'ラージ'), ('mini', 'ミニ')]:
            b = s03.get(block_key, {})
            if b:
                md.append('**%s**:' % block_label)
                md.append('- プット: %s枚 / %s百万円' % (fnum(b.get('put_volume')), fnum(b.get('put_value'))))
                md.append('- コール: %s枚 / %s百万円' % (fnum(b.get('call_volume')), fnum(b.get('call_value'))))
                md.append('- 合計: %s枚 / %s百万円' % (fnum(b.get('total_volume')), fnum(b.get('total_value'))))
                md.append('- J-NET比率: %s' % fpct(b.get('jnet_ratio')))
                md.append('')
    else:
        md.append('データなし')
    md.append('')

    # ④ Options OI Changes
    s04 = data.get('s04', {})
    md.append('## ④ オプション建玉増減')
    md.append('')
    if 'error' not in s04:
        for block_key, block_label in [('large', 'ラージ'), ('mini', 'ミニ')]:
            b = s04.get(block_key, {})
            if b:
                md.append('**%s**:' % block_label)
                md.append('- プット合計: OI %s / 前日比 %s' % (
                    fnum(b.get('put_total_oi')), fnum(b.get('put_total_change'), plus=True)))
                md.append('- コール合計: OI %s / 前日比 %s' % (
                    fnum(b.get('call_total_oi')), fnum(b.get('call_total_change'), plus=True)))
                md.append('')
    md.append('')

    # ⑤ Important OI Changes
    s05 = data.get('s05', [])
    md.append('## ⑤ 重要建玉変化（±300枚以上）')
    md.append('')
    if s05:
        md.append('| 銘柄 | 建玉残高 | 前日比 |')
        md.append('|------|---------|--------|')
        for c in s05:
            md.append('| %s | %s | %s |' % (c['name'], fnum(c.get('oi')), fnum(c['change'], plus=True)))
    else:
        md.append('該当なし')
    md.append('')

    # ⑥ OI Distribution
    s06 = data.get('s06', {})
    md.append('## ⑥ 建玉分布（ATM±5,000円）')
    md.append('')
    if 'distribution' in s06:
        md.append('ATM = %s' % fnum(s06.get('atm')))
        md.append('')
        md.append('| 行使価格 | P建玉 | P前日比 | C建玉 | C前日比 |')
        md.append('|---------|-------|---------|-------|---------|')
        for d in s06['distribution']:
            atm_mark = ' **←ATM**' if d.get('is_atm') else ''
            md.append('| %s%s | %s | %s | %s | %s |' % (
                fnum(d['strike']), atm_mark,
                fnum(d['put_oi']), fnum(d['put_change'], plus=True),
                fnum(d['call_oi']), fnum(d['call_change'], plus=True)))
    md.append('')

    # ⑦ J-NET
    s07 = data.get('s07', [])
    md.append('## ⑦ 大口手口（J-NET）')
    md.append('')
    if s07:
        md.append('| 銘柄 | 参加者 | 取引高 | 分類 |')
        md.append('|------|-------|--------|------|')
        for t in s07:
            pair = ' 🔄' if t.get('is_pair') else ''
            cat_label = {'us': '米系', 'eu': '欧系', 'hf': 'HF代理', 'domestic': '国内'}.get(t['category'], 'その他')
            md.append('| %s | %s%s | %s | %s |' % (
                t['product'], t['participant'], pair, fnum(t['volume']), cat_label))
    else:
        md.append('該当なし')
    md.append('')

    # ⑧ (placeholder for LLM)
    md.append('## ⑧ 総合評価')
    md.append('')
    md.append('> ※ このセクションはLLM（Claude）による定性分析が必要です。')
    md.append('> data.json をClaudeに渡して⑧を生成してください。')
    md.append('')

    # ⑨ Participants
    s09 = data.get('s09', {})
    md.append('## ⑨ 参加者別建玉分析')
    md.append('')
    if 'error' not in s09:
        # Futures summary
        fut = s09.get('futures', {})
        if fut:
            for sec_key, sec_label in [('nk225_large', 'N225ラージ'), ('nk225_mini', 'N225mini'), ('topix', 'TOPIX')]:
                sec = fut.get(sec_key, {})
                if sec:
                    md.append('**%s**: 海外Net %s / 国内Net %s' % (
                        sec_label, fnum(sec.get('overseas_net'), plus=True), fnum(sec.get('domestic_net'), plus=True)))

            md.append('')

            # Top sellers/buyers
            for sec_key, sec_label in [('nk225_large', 'N225ラージ'), ('nk225_mini', 'N225mini'), ('topix', 'TOPIX')]:
                sec = fut.get(sec_key, {})
                if sec:
                    sellers = sec.get('sellers', [])[:5]
                    buyers = sec.get('buyers', [])[:5]
                    if sellers or buyers:
                        md.append('**%s 上位**:' % sec_label)
                        md.append('- 売超: %s' % ', '.join(['%s -%s' % (s['name'], fnum(s['volume'])) for s in sellers]))
                        md.append('- 買超: %s' % ', '.join(['%s +%s' % (b['name'], fnum(b['volume'])) for b in buyers]))
                        md.append('')

        # Profiles
        profiles = s09.get('profiles', [])
        if profiles:
            md.append('### 統合プロファイル')
            md.append('')
            md.append('| 参加者 | 分類 | N225ラージ | mini | TOPIX | P Net | C Net | 推定戦略 |')
            md.append('|-------|------|-----------|------|-------|-------|-------|---------|')
            for p in profiles[:10]:
                md.append('| %s | %s | %s | %s | %s | %s | %s | %s |' % (
                    p['name'], p['category_label'],
                    fnum(p['nk225_large'], plus=True),
                    fnum(p['nk225_mini'], plus=True),
                    fnum(p['topix'], plus=True),
                    fnum(p['put_net'], plus=True),
                    fnum(p['call_net'], plus=True),
                    p['strategy']))
            md.append('')
    else:
        md.append('週次データなし')
    md.append('')

    # ⑩ Stock Flow
    s10 = data.get('s10', {})
    if not s10.get('skipped'):
        md.append('## ⑩ 投資部門別 現物フロー')
        md.append('')
        labels = {'foreigners': '海外投資家', 'individuals': '個人', 'institutions': '法人',
                  'investment_trusts': '投資信託', 'proprietary': '自己'}
        for key, label in labels.items():
            v = s10.get(key, {})
            if v:
                md.append('- %s: %s億円' % (label, fnum(v.get('oku_yen'))))
        md.append('')

    # ⑪ Strategy Map
    s11 = data.get('s11', {})
    md.append('## ⑪ 戦略マップ')
    md.append('')

    otm = s11.get('otm_table', [])
    if otm:
        md.append('### OTM確率テーブル')
        md.append('')
        md.append('| 行使価格 | タイプ | VI-10 | 現在VI | VI+10 | BS価格 |')
        md.append('|---------|--------|-------|--------|-------|--------|')
        for o in otm:
            md.append('| %s | %s | %s | %s | %s | %s |' % (
                fnum(o['strike']), o['label'],
                fpct(o['otm_prob']['vi_minus10']),
                fpct(o['otm_prob']['vi_current']),
                fpct(o['otm_prob']['vi_plus10']),
                fnum(o['bs_price'])))
        md.append('')

    edges = s11.get('edge_scores', [])
    if edges:
        md.append('### ゾーン別売りエッジ')
        md.append('')
        md.append('| ゾーン | タイプ | 壁最大OI | スコア | 評価 |')
        md.append('|-------|--------|---------|--------|------|')
        for e in edges:
            stars = '★' * e['stars'] + '☆' * (5 - e['stars'])
            md.append('| %s | %s | %s @%s | %.2f | %s |' % (
                e['zone'], e['type'], fnum(e['wall_max_oi']), fnum(e['wall_strike']),
                e['total_score'], stars))
        md.append('')

    return '\n'.join(md)


# ============================================================
# HTML Dashboard Builder
# ============================================================

# --- CSS ---
DASHBOARD_CSS = r"""
:root{
  --bg:#06060f;--panel:#0c0c1d;--card:#111128;--border:#1e1e3a;
  --text:#e2e2f0;--sub:#8888aa;--accent:#818cf8;
  --red:#f87171;--green:#4ade80;--blue:#60a5fa;--yellow:#fbbf24;
  --put:#f87171;--call:#60a5fa;
  --us:#93c5fd;--eu:#c4b5fd;--hf:#fbbf24;--dom:#f87171;--overseas:#60a5fa;
}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--text);font-family:'Noto Sans JP','Outfit',sans-serif;font-size:13px;line-height:1.6}
a{color:var(--accent);text-decoration:none}
.topbar{position:sticky;top:0;z-index:100;background:rgba(6,6,15,.92);backdrop-filter:blur(12px);border-bottom:1px solid var(--border);padding:10px 20px;display:flex;align-items:center;justify-content:space-between}
.topbar .logo{font-family:Outfit;font-weight:700;font-size:16px;color:var(--accent)}
.topbar nav a{margin-left:16px;font-size:12px;color:var(--sub)}
.topbar nav a:hover{color:var(--text)}
.hero{text-align:center;padding:28px 16px 10px}
.hero h1{font-family:Outfit;font-size:22px;font-weight:700;color:#fff}
.hero .sub{color:var(--sub);font-size:12px;margin-top:4px}
.kpi-strip{display:flex;justify-content:center;gap:20px;padding:10px 16px 16px;flex-wrap:wrap}
.kpi{text-align:center}
.kpi .label{font-size:10px;color:var(--sub);text-transform:uppercase;letter-spacing:.5px}
.kpi .value{font-family:'DM Mono',monospace;font-size:18px;font-weight:700;color:#fff}
.kpi .value.up{color:var(--green)}
.kpi .value.down{color:var(--red)}
.mobile-nav{display:none;text-align:center;padding:6px;border-bottom:1px solid var(--border)}
.mobile-nav a{margin:0 8px;font-size:11px;color:var(--sub)}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:10px;padding:10px 16px 30px;max-width:1200px;margin:0 auto}
.card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:14px;cursor:pointer;transition:border-color .2s,background .2s}
.card:hover{border-color:var(--accent)}
.card.open{grid-column:1/-1;border-color:var(--accent);background:var(--panel);cursor:default}
.card-hdr{display:flex;align-items:center;gap:8px}
.card-hdr .icon{font-size:18px}
.card-hdr .title{font-family:Outfit;font-weight:600;font-size:14px;color:#fff}
.card-hdr .arrow{margin-left:auto;color:var(--sub);font-size:12px;transition:transform .2s}
.card.open .card-hdr .arrow{transform:rotate(90deg)}
.card-preview{margin-top:10px}
.card-detail{display:none;margin-top:14px;border-top:1px solid var(--border);padding-top:14px}
.card.open .card-detail{display:block}
.mini-metrics{display:flex;gap:12px;flex-wrap:wrap}
.mini-metric{flex:1;min-width:80px}
.mini-metric .mm-label{font-size:10px;color:var(--sub)}
.mini-metric .mm-value{font-family:'DM Mono',monospace;font-size:15px;font-weight:600}
.tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:600;margin:2px}
.tag-put{background:rgba(248,113,113,.15);color:var(--put)}
.tag-call{background:rgba(96,165,250,.15);color:var(--call)}
.tag-up{background:rgba(74,222,128,.15);color:var(--green)}
.tag-down{background:rgba(239,68,68,.15);color:var(--red)}
.tag-us{background:rgba(147,197,253,.15);color:var(--us)}
.tag-eu{background:rgba(196,181,253,.15);color:var(--eu)}
.tag-hf{background:rgba(251,191,36,.15);color:var(--hf)}
.tag-dom{background:rgba(248,113,113,.15);color:var(--dom)}
table{width:100%;border-collapse:collapse;margin:10px 0;font-size:12px}
th{background:var(--card);color:var(--sub);font-weight:600;text-align:left;padding:6px 8px;border-bottom:1px solid var(--border)}
td{padding:5px 8px;border-bottom:1px solid rgba(30,30,58,.5)}
tr.atm-row{background:rgba(251,191,36,.08)}
.bar-row{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:11px}
.bar-label{width:120px;text-align:right;color:var(--sub);flex-shrink:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}
.bar-track{flex:1;height:16px;background:var(--card);border-radius:3px;position:relative;overflow:hidden}
.bar-fill{height:100%;border-radius:3px;min-width:1px}
.bar-fill.put{background:var(--put)}
.bar-fill.call{background:var(--call)}
.bar-fill.up{background:var(--green)}
.bar-fill.down{background:var(--red)}
.bar-val{width:60px;font-family:'DM Mono',monospace;font-size:11px;flex-shrink:0}
.insight{background:rgba(129,140,248,.08);border:1px solid rgba(129,140,248,.2);border-radius:8px;padding:12px;margin-top:12px;font-size:12px;color:var(--sub);line-height:1.7}
.insight strong{color:var(--text)}
.footer{text-align:center;padding:20px;color:var(--sub);font-size:11px;border-top:1px solid var(--border)}
.footer a{margin:0 8px}
.positive{color:var(--green)}
.negative{color:var(--red)}
@media(max-width:768px){
  .topbar nav{display:none}
  .mobile-nav{display:block}
  .grid{grid-template-columns:1fr}
  .kpi .value{font-size:15px}
}
"""

# --- JS ---
DASHBOARD_JS = r"""
(function(){
  var cards=document.querySelectorAll('.card[data-card]');
  var grid=document.querySelector('.grid');
  if(!grid)return;
  grid.addEventListener('click',function(e){
    var el=e.target;
    while(el&&el!==grid){
      if(el.dataset&&el.dataset.card!==undefined){
        toggleCard(el);
        return;
      }
      el=el.parentElement;
    }
  });
  function toggleCard(card){
    var wasOpen=card.classList.contains('open');
    // close all
    for(var i=0;i<cards.length;i++){
      cards[i].classList.remove('open');
    }
    if(!wasOpen){
      // build detail on first open
      if(!card.dataset.built){
        var fn=window['b_'+card.dataset.card];
        if(fn){
          var detail=card.querySelector('.card-detail');
          if(detail) detail.innerHTML=fn();
          card.dataset.built='1';
        }
      }
      card.classList.add('open');
      setTimeout(function(){card.scrollIntoView({behavior:'smooth',block:'start'});},100);
    }
  }
})();
"""


def build_dashboard_html(data):
    """Generate index.html dashboard."""
    meta = data['metadata']
    s01 = data.get('s01', {})
    s02 = data.get('s02', {})
    s03 = data.get('s03', {})
    s04 = data.get('s04', {})
    s05 = data.get('s05', [])
    s06 = data.get('s06', {})
    s07 = data.get('s07', [])
    s09 = data.get('s09', {})
    s11 = data.get('s11', {})

    atm = meta.get('atm', 0)
    nikkei = s01.get('nikkei_close', 0)
    vi = s01.get('vi', 0)

    h = '<!DOCTYPE html>\n<html lang="ja">\n<head>\n'
    h += '<meta charset="UTF-8">\n'
    h += '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
    h += '<title>JPX Market Analysis %s</title>\n' % esc(meta.get('date_formatted', ''))
    h += '<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700&family=Noto+Sans+JP:wght@400;500;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">\n'
    h += '<style>\n%s\n</style>\n' % DASHBOARD_CSS
    h += '</head>\n<body>\n'

    # Topbar
    h += '<div class="topbar">\n'
    h += '  <span class="logo">JPX Dashboard</span>\n'
    h += '  <nav>\n'
    h += '    <a href="index.html">ダッシュボード</a>\n'
    h += '    <a href="pnl_simulator.html">P&Lシミュレーター</a>\n'
    h += '    <a href="archive.html">アーカイブ</a>\n'
    h += '  </nav>\n'
    h += '</div>\n'

    # Hero
    h += '<div class="hero">\n'
    h += '  <h1>%s</h1>\n' % esc(meta.get('date_formatted', ''))
    h += '  <div class="sub">%s / SQまで%d営業日</div>\n' % (esc(meta.get('sq_label', '')), meta.get('days_to_sq', 0))
    h += '</div>\n'

    # KPI Strip
    h += '<div class="kpi-strip">\n'
    nk_cls = ''
    h += '  <div class="kpi"><div class="label">日経平均</div><div class="value %s">%s</div></div>\n' % (nk_cls, fnum(nikkei) if nikkei else '-')
    vi_cls = 'down' if vi and vi > 30 else ''
    h += '  <div class="kpi"><div class="label">VI</div><div class="value %s">%s</div></div>\n' % (vi_cls, vi if vi else '-')
    h += '  <div class="kpi"><div class="label">ATM</div><div class="value">%s</div></div>\n' % (fnum(atm) if atm else '-')
    r1d = s01.get('range_1d', {})
    if r1d:
        h += '  <div class="kpi"><div class="label">1日値幅</div><div class="value">%s</div></div>\n' % fnum(r1d.get('width'))
    h += '</div>\n'

    # Mobile nav
    h += '<div class="mobile-nav">\n'
    h += '  <a href="index.html">ダッシュボード</a>\n'
    h += '  <a href="pnl_simulator.html">P&L</a>\n'
    h += '  <a href="archive.html">アーカイブ</a>\n'
    h += '</div>\n'

    # Card grid
    h += '<div class="grid">\n'

    # Card definitions: (id, icon, title, preview_fn, detail_js_data)
    cards = [
        ('futures', '📈', '先物建玉増減', _preview_futures(s02), _detail_futures_js(s02)),
        ('opval', '💰', 'オプション取引代金', _preview_opval(s03), _detail_opval_js(s03)),
        ('oichg', '📊', 'オプション建玉増減', _preview_oichg(s04), _detail_oichg_js(s04)),
        ('important', '⚡', '重要建玉変化', _preview_important(s05), _detail_important_js(s05)),
        ('dist', '🦋', '建玉分布', _preview_dist(s06), _detail_dist_js(s06)),
        ('jnet', '🏛', '大口手口（J-NET）', _preview_jnet(s07), _detail_jnet_js(s07)),
        ('assess', '🎯', '総合評価', _preview_assess(s01), _detail_assess_js(data)),
        ('participants', '🌏', '参加者別建玉分析', _preview_participants(s09), _detail_participants_js(s09)),
        ('strategy', '🎲', '戦略マップ', _preview_strategy(s11), _detail_strategy_js(s11, atm)),
    ]

    for card_id, icon, title, preview_html, detail_js in cards:
        h += '<div class="card" data-card="%s">\n' % card_id
        h += '  <div class="card-hdr">\n'
        h += '    <span class="icon">%s</span>\n' % icon
        h += '    <span class="title">%s</span>\n' % esc(title)
        h += '    <span class="arrow">▶</span>\n'
        h += '  </div>\n'
        h += '  <div class="card-preview">%s</div>\n' % preview_html
        h += '  <div class="card-detail"></div>\n'
        h += '</div>\n'

    h += '</div>\n'  # grid

    # Footer
    h += '<div class="footer">\n'
    h += '  <a href="pnl_simulator.html">P&Lシミュレーター</a>\n'
    h += '  <a href="archive.html">アーカイブ一覧</a>\n'
    h += '  <span>Generated by JPX Analysis Pipeline</span>\n'
    h += '</div>\n'

    # JavaScript - card detail builder functions
    h += '<script>\n'
    # Embed detail builders
    for card_id, _, _, _, detail_js in cards:
        h += 'function b_%s(){' % card_id
        h += detail_js
        h += '}\n'

    h += DASHBOARD_JS
    h += '</script>\n'
    h += '</body>\n</html>'

    return h


# --- Card Preview Builders ---

def _preview_futures(s02):
    if 'error' in s02:
        return '<span class="mm-label">データなし</span>'
    h = '<div class="mini-metrics">'
    for key, label in [('nk225_large', 'ラージ'), ('nk225_mini', 'mini'), ('topix', 'TOPIX')]:
        sec = s02.get(key, {})
        chg = sec.get('total_change', 0)
        cls = 'positive' if chg > 0 else 'negative' if chg < 0 else ''
        h += '<div class="mini-metric"><div class="mm-label">%s</div>' % label
        h += '<div class="mm-value %s">%s</div></div>' % (cls, fnum(chg, plus=True))
    h += '</div>'
    return h


def _preview_opval(s03):
    if 'error' in s03:
        return '<span class="mm-label">データなし</span>'
    lg = s03.get('large', {})
    h = '<div class="mini-metrics">'
    h += '<div class="mini-metric"><div class="mm-label">プット</div><div class="mm-value">%s枚</div></div>' % fnum(lg.get('put_volume'))
    h += '<div class="mini-metric"><div class="mm-label">コール</div><div class="mm-value">%s枚</div></div>' % fnum(lg.get('call_volume'))
    h += '<div class="mini-metric"><div class="mm-label">J-NET率</div><div class="mm-value">%s</div></div>' % fpct(lg.get('jnet_ratio'))
    h += '</div>'
    return h


def _preview_oichg(s04):
    if 'error' in s04:
        return '<span class="mm-label">データなし</span>'
    lg = s04.get('large', {})
    h = '<div class="mini-metrics">'
    pc = lg.get('put_total_change', 0)
    cc = lg.get('call_total_change', 0)
    h += '<div class="mini-metric"><div class="mm-label">P合計</div><div class="mm-value %s">%s</div></div>' % (sign_class(pc), fnum(pc, plus=True))
    h += '<div class="mini-metric"><div class="mm-label">C合計</div><div class="mm-value %s">%s</div></div>' % (sign_class(cc), fnum(cc, plus=True))
    h += '</div>'
    return h


def _preview_important(s05):
    if not s05:
        return '<span class="mm-label">該当なし</span>'
    h = ''
    for c in s05[:4]:
        cls = 'tag-put' if c['type'] == 'P' else 'tag-call'
        direction = 'tag-up' if c['change'] > 0 else 'tag-down'
        h += '<span class="tag %s">%s%s %s</span>' % (cls, c['type'], fnum(c['strike']), fnum(c['change'], plus=True))
    if len(s05) > 4:
        h += '<span class="tag">+%d件</span>' % (len(s05) - 4)
    return h


def _preview_dist(s06):
    if 'distribution' not in s06:
        return '<span class="mm-label">データなし</span>'
    dist = s06['distribution']
    # Find max P and C walls
    max_p = max(dist, key=lambda d: d['put_oi']) if dist else None
    max_c = max(dist, key=lambda d: d['call_oi']) if dist else None
    h = '<div class="mini-metrics">'
    if max_p:
        h += '<div class="mini-metric"><div class="mm-label">P壁</div><div class="mm-value" style="color:var(--put)">%s (%s)</div></div>' % (fnum(max_p['strike']), fnum(max_p['put_oi']))
    h += '<div class="mini-metric"><div class="mm-label">ATM</div><div class="mm-value" style="color:var(--yellow)">%s</div></div>' % fnum(s06.get('atm'))
    if max_c:
        h += '<div class="mini-metric"><div class="mm-label">C壁</div><div class="mm-value" style="color:var(--call)">%s (%s)</div></div>' % (fnum(max_c['strike']), fnum(max_c['call_oi']))
    h += '</div>'
    return h


def _preview_jnet(s07):
    if not s07:
        return '<span class="mm-label">該当なし</span>'
    h = ''
    seen = set()
    for t in s07[:5]:
        if t['participant'] not in seen:
            cat_cls = {'us': 'tag-us', 'eu': 'tag-eu', 'hf': 'tag-hf', 'domestic': 'tag-dom'}.get(t['category'], '')
            h += '<span class="tag %s">%s %s枚</span>' % (cat_cls, esc(t['participant'][:8]), fnum(t['volume']))
            seen.add(t['participant'])
    return h


def _preview_assess(s01):
    r1w = s01.get('range_1w', {})
    if r1w:
        return '<div class="mini-metrics"><div class="mini-metric"><div class="mm-label">1週予想レンジ</div><div class="mm-value" style="color:var(--yellow)">%s 〜 %s</div></div></div>' % (fnum(r1w.get('low')), fnum(r1w.get('high')))
    return '<span class="mm-label">LLM分析が必要</span>'


def _preview_participants(s09):
    if 'error' in s09:
        return '<span class="mm-label">週次データなし</span>'
    fut = s09.get('futures', {})
    nk = fut.get('nk225_large', {})
    h = '<div class="mini-metrics">'
    on = nk.get('overseas_net', 0)
    dn = nk.get('domestic_net', 0)
    h += '<div class="mini-metric"><div class="mm-label">海外Net</div><div class="mm-value %s">%s</div></div>' % (sign_class(on), fnum(on, plus=True))
    h += '<div class="mini-metric"><div class="mm-label">国内Net</div><div class="mm-value %s">%s</div></div>' % (sign_class(dn), fnum(dn, plus=True))
    h += '</div>'
    return h


def _preview_strategy(s11):
    otm = s11.get('otm_table', [])
    edges = s11.get('edge_scores', [])
    if not otm:
        return '<span class="mm-label">データ不足</span>'
    # Best P sell and C sell zones
    best_p = None
    best_c = None
    for e in edges:
        if e['type'] == 'sell-p' and (best_p is None or e['stars'] > best_p['stars']):
            best_p = e
        if e['type'] == 'sell-c' and (best_c is None or e['stars'] > best_c['stars']):
            best_c = e
    h = '<div class="mini-metrics">'
    if best_p:
        h += '<div class="mini-metric"><div class="mm-label">P売り</div><div class="mm-value" style="color:var(--put)">%s</div></div>' % ('★' * best_p['stars'])
    if best_c:
        h += '<div class="mini-metric"><div class="mm-label">C売り</div><div class="mm-value" style="color:var(--call)">%s</div></div>' % ('★' * best_c['stars'])
    h += '</div>'
    return h


# --- Card Detail JS Builders ---
# Each returns a JS string that builds HTML and returns it.
# Uses h+='...' concatenation (no template literals per v6 rules).

def _js_str(s):
    """Escape a string for JS single-quoted string literal."""
    return str(s).replace('\\', '\\\\').replace("'", "\\'").replace('\n', '').replace('\r', '')


def _detail_futures_js(s02):
    if 'error' in s02:
        return "var h='<div>データなし</div>';return h;"
    js = "var h='';"
    for key, label in [('nk225_large', '日経225ラージ'), ('nk225_mini', '日経225mini'), ('topix', 'TOPIX')]:
        sec = s02.get(key, {})
        chg = sec.get('total_change', 0)
        oi = sec.get('total_oi', 0)
        cls = 'positive' if chg > 0 else 'negative' if chg < 0 else ''
        js += "h+='<div class=\"bar-row\">';"
        js += "h+='<div class=\"bar-label\">%s</div>';" % _js_str(label)
        js += "h+='<div class=\"bar-track\"><div class=\"bar-fill %s\" style=\"width:%dpx\"></div></div>';" % (
            'up' if chg > 0 else 'down', min(abs(chg) // 50 + 5, 200) if chg else 0)
        js += "h+='<div class=\"bar-val %s\">%s</div>';" % (cls, _js_str(fnum(chg, plus=True)))
        js += "h+='</div>';"

        # Monthly details
        for m in sec.get('months', []):
            mc = m.get('change', 0)
            mcls = 'positive' if mc > 0 else 'negative' if mc < 0 else ''
            js += "h+='<div class=\"bar-row\">';"
            js += "h+='<div class=\"bar-label\" style=\"font-size:10px\">%s</div>';" % _js_str(m['label'][:12])
            js += "h+='<div class=\"bar-track\"><div class=\"bar-fill %s\" style=\"width:%dpx\"></div></div>';" % (
                'up' if mc > 0 else 'down', min(abs(mc) // 50 + 3, 150) if mc else 0)
            js += "h+='<div class=\"bar-val %s\" style=\"font-size:10px\">%s</div>';" % (mcls, _js_str(fnum(mc, plus=True)))
            js += "h+='</div>';"

    js += "return h;"
    return js


def _detail_opval_js(s03):
    if 'error' in s03:
        return "var h='<div>データなし</div>';return h;"
    js = "var h='';"
    for bk, bl in [('large', 'ラージ'), ('mini', 'ミニ')]:
        b = s03.get(bk, {})
        if not b:
            continue
        js += "h+='<h3 style=\"color:#fff;font-size:13px;margin:8px 0 4px\">%s</h3>';" % bl
        js += "h+='<table><tr><th></th><th>取引高</th><th>取引代金(百万)</th></tr>';"
        js += "h+='<tr><td style=\"color:var(--put)\">プット</td><td>%s</td><td>%s</td></tr>';" % (_js_str(fnum(b.get('put_volume'))), _js_str(fnum(b.get('put_value'))))
        js += "h+='<tr><td style=\"color:var(--call)\">コール</td><td>%s</td><td>%s</td></tr>';" % (_js_str(fnum(b.get('call_volume'))), _js_str(fnum(b.get('call_value'))))
        js += "h+='<tr><td>合計</td><td>%s</td><td>%s</td></tr>';" % (_js_str(fnum(b.get('total_volume'))), _js_str(fnum(b.get('total_value'))))
        js += "h+='<tr><td>J-NET</td><td>%s</td><td>%s (%s)</td></tr>';" % (
            _js_str(fnum(b.get('jnet_volume'))), _js_str(fnum(b.get('jnet_value'))), _js_str(fpct(b.get('jnet_ratio'))))
        js += "h+='</table>';"
    js += "return h;"
    return js


def _detail_oichg_js(s04):
    if 'error' in s04:
        return "var h='<div>データなし</div>';return h;"
    js = "var h='';"
    for bk, bl in [('large', 'ラージ'), ('mini', 'ミニ')]:
        b = s04.get(bk, {})
        if not b:
            continue
        pc = b.get('put_total_change', 0)
        cc = b.get('call_total_change', 0)
        js += "h+='<h3 style=\"color:#fff;font-size:13px;margin:8px 0 4px\">%s</h3>';" % bl
        js += "h+='<div class=\"bar-row\">';"
        js += "h+='<div class=\"bar-label\" style=\"color:var(--put)\">P合計 %s</div>';" % _js_str(fnum(pc, plus=True))
        js += "h+='<div class=\"bar-track\"><div class=\"bar-fill %s\" style=\"width:%dpx\"></div></div>';" % ('up' if pc > 0 else 'down', min(abs(pc) // 30 + 5, 200) if pc else 0)
        js += "h+='</div>';"
        js += "h+='<div class=\"bar-row\">';"
        js += "h+='<div class=\"bar-label\" style=\"color:var(--call)\">C合計 %s</div>';" % _js_str(fnum(cc, plus=True))
        js += "h+='<div class=\"bar-track\"><div class=\"bar-fill %s\" style=\"width:%dpx\"></div></div>';" % ('up' if cc > 0 else 'down', min(abs(cc) // 30 + 5, 200) if cc else 0)
        js += "h+='</div>';"
    js += "return h;"
    return js


def _detail_important_js(s05):
    if not s05:
        return "var h='<div>該当なし</div>';return h;"
    js = "var h='';"
    js += "h+='<table><tr><th>銘柄</th><th>建玉</th><th>前日比</th></tr>';"
    for c in s05:
        cls = 'positive' if c['change'] > 0 else 'negative'
        js += "h+='<tr><td>%s</td><td>%s</td><td class=\"%s\">%s</td></tr>';" % (
            _js_str(esc(c['name'])), _js_str(fnum(c.get('oi'))), cls, _js_str(fnum(c['change'], plus=True)))
    js += "h+='</table>';"
    js += "return h;"
    return js


def _detail_dist_js(s06):
    if 'distribution' not in s06:
        return "var h='<div>データなし</div>';return h;"
    dist = s06['distribution']
    max_oi = max([max(d['put_oi'], d['call_oi']) for d in dist] or [1])
    js = "var h='';"
    js += "h+='<div style=\"font-size:11px;color:var(--sub);margin-bottom:8px\">ATM = %s</div>';" % _js_str(fnum(s06.get('atm')))

    for d in dist:
        pw = int(d['put_oi'] / max_oi * 150) if max_oi else 0
        cw = int(d['call_oi'] / max_oi * 150) if max_oi else 0
        atm_style = 'background:rgba(251,191,36,.12);' if d.get('is_atm') else ''
        js += "h+='<div style=\"display:flex;align-items:center;gap:4px;padding:2px 0;font-size:11px;%s\">';" % atm_style
        # Put bar (right-aligned)
        pcls = 'positive' if d['put_change'] > 0 else 'negative' if d['put_change'] < 0 else ''
        js += "h+='<div style=\"width:50px;text-align:right;font-family:DM Mono;font-size:10px\" class=\"%s\">%s</div>';" % (pcls, _js_str(fnum(d['put_change'], plus=True)))
        js += "h+='<div style=\"width:50px;text-align:right;font-family:DM Mono;font-size:10px\">%s</div>';" % _js_str(fnum(d['put_oi']))
        js += "h+='<div style=\"width:150px;direction:rtl\"><div style=\"display:inline-block;height:12px;width:%dpx;background:var(--put);border-radius:2px\"></div></div>';" % pw
        # Strike
        strike_color = 'var(--yellow)' if d.get('is_atm') else '#fff'
        js += "h+='<div style=\"width:60px;text-align:center;font-family:DM Mono;font-weight:600;color:%s\">%s</div>';" % (strike_color, _js_str(fnum(d['strike'])))
        # Call bar
        js += "h+='<div style=\"width:150px\"><div style=\"display:inline-block;height:12px;width:%dpx;background:var(--call);border-radius:2px\"></div></div>';" % cw
        js += "h+='<div style=\"width:50px;font-family:DM Mono;font-size:10px\">%s</div>';" % _js_str(fnum(d['call_oi']))
        ccls = 'positive' if d['call_change'] > 0 else 'negative' if d['call_change'] < 0 else ''
        js += "h+='<div style=\"width:50px;font-family:DM Mono;font-size:10px\" class=\"%s\">%s</div>';" % (ccls, _js_str(fnum(d['call_change'], plus=True)))
        js += "h+='</div>';"

    js += "return h;"
    return js


def _detail_jnet_js(s07):
    if not s07:
        return "var h='<div>該当なし</div>';return h;"
    js = "var h='';"
    js += "h+='<table><tr><th>銘柄</th><th>参加者</th><th>取引高</th><th>分類</th></tr>';"
    for t in s07:
        cat_tag = {'us': 'tag-us', 'eu': 'tag-eu', 'hf': 'tag-hf', 'domestic': 'tag-dom'}.get(t['category'], '')
        cat_label = {'us': '米系', 'eu': '欧系', 'hf': 'HF代理', 'domestic': '国内'}.get(t['category'], 'その他')
        pair = ' <span style="color:var(--yellow)">🔄</span>' if t.get('is_pair') else ''
        js += "h+='<tr><td>%s</td><td>%s%s</td><td>%s</td><td><span class=\"tag %s\">%s</span></td></tr>';" % (
            _js_str(esc(t['product'][:30])), _js_str(esc(t['participant'])), pair,
            _js_str(fnum(t['volume'])), cat_tag, cat_label)
    js += "h+='</table>';"
    js += "return h;"
    return js


def _detail_assess_js(data):
    """Assessment card - shows ranges and placeholder for LLM text."""
    s01 = data.get('s01', {})
    r1d = s01.get('range_1d', {})
    r1w = s01.get('range_1w', {})
    js = "var h='';"
    if r1d:
        js += "h+='<div style=\"background:var(--card);border-radius:8px;padding:12px;margin:8px 0\">';"
        js += "h+='<div style=\"color:var(--sub);font-size:10px\">1日予測値幅</div>';"
        js += "h+='<div style=\"font-family:DM Mono;font-size:16px;color:var(--yellow)\">%s 〜 %s</div>';" % (_js_str(fnum(r1d.get('low'))), _js_str(fnum(r1d.get('high'))))
        js += "h+='<div style=\"color:var(--sub);font-size:10px\">幅: %s円</div>';" % _js_str(fnum(r1d.get('width')))
        js += "h+='</div>';"
    if r1w:
        js += "h+='<div style=\"background:var(--card);border-radius:8px;padding:12px;margin:8px 0\">';"
        js += "h+='<div style=\"color:var(--sub);font-size:10px\">1週予測値幅</div>';"
        js += "h+='<div style=\"font-family:DM Mono;font-size:16px;color:var(--yellow)\">%s 〜 %s</div>';" % (_js_str(fnum(r1w.get('low'))), _js_str(fnum(r1w.get('high'))))
        js += "h+='<div style=\"color:var(--sub);font-size:10px\">幅: %s円</div>';" % _js_str(fnum(r1w.get('width')))
        js += "h+='</div>';"
    js += "h+='<div class=\"insight\"><strong>⑧ 総合評価</strong>: LLMによる定性分析が必要です。data.json をClaudeに渡して生成してください。</div>';"
    js += "return h;"
    return js


def _detail_participants_js(s09):
    if 'error' in s09:
        return "var h='<div>週次データなし</div>';return h;"

    js = "var h='';"
    profiles = s09.get('profiles', [])

    if profiles:
        js += "h+='<table><tr><th>参加者</th><th>分類</th><th>N225</th><th>mini</th><th>TOPIX</th><th>P Net</th><th>C Net</th><th>戦略</th></tr>';"
        for p in profiles[:12]:
            cat_tag = {'us': 'tag-us', 'eu': 'tag-eu', 'hf': 'tag-hf', 'domestic': 'tag-dom'}.get(p['category'], '')
            js += "h+='<tr>';"
            js += "h+='<td>%s</td>';" % _js_str(esc(p['name'][:12]))
            js += "h+='<td><span class=\"tag %s\">%s</span></td>';" % (cat_tag, _js_str(p['category_label']))
            for field in ['nk225_large', 'nk225_mini', 'topix', 'put_net', 'call_net']:
                v = p.get(field, 0)
                cls = 'positive' if v > 0 else 'negative' if v < 0 else ''
                js += "h+='<td class=\"%s\" style=\"font-family:DM Mono;font-size:11px\">%s</td>';" % (cls, _js_str(fnum(v, plus=True)))
            js += "h+='<td style=\"font-size:10px\">%s</td>';" % _js_str(esc(p['strategy'][:12]))
            js += "h+='</tr>';"
        js += "h+='</table>';"

        # OP details for top participants
        for p in profiles[:5]:
            if p.get('op_detail'):
                js += "h+='<div style=\"font-size:10px;color:var(--sub);margin:2px 0\">%s: %s</div>';" % (
                    _js_str(esc(p['name'][:10])), _js_str(esc(p['op_detail'][:60])))

    js += "return h;"
    return js


def _detail_strategy_js(s11, atm):
    otm = s11.get('otm_table', [])
    edges = s11.get('edge_scores', [])
    if not otm:
        return "var h='<div>データ不足（ATMまたはVI未設定）</div>';return h;"

    js = "var h='';"

    # OTM probability table
    js += "h+='<h3 style=\"color:#fff;font-size:13px;margin:8px 0 4px\">OTM確率テーブル</h3>';"
    js += "h+='<table><tr><th>行使価格</th><th>タイプ</th><th>VI-10</th><th>現在</th><th>VI+10</th><th>BS価格</th></tr>';"
    for o in otm:
        js += "h+='<tr><td style=\"font-family:DM Mono\">%s</td><td>%s</td><td>%s</td><td style=\"font-weight:600\">%s</td><td>%s</td><td>%s</td></tr>';" % (
            _js_str(fnum(o['strike'])), _js_str(o['label']),
            _js_str(fpct(o['otm_prob']['vi_minus10'])),
            _js_str(fpct(o['otm_prob']['vi_current'])),
            _js_str(fpct(o['otm_prob']['vi_plus10'])),
            _js_str(fnum(o['bs_price'])))
    js += "h+='</table>';"

    # Edge scores
    if edges:
        js += "h+='<h3 style=\"color:#fff;font-size:13px;margin:12px 0 4px\">ゾーン別エッジ評価</h3>';"
        for e in edges:
            stars = '★' * e['stars'] + '☆' * (5 - e['stars'])
            zone_color = 'var(--put)' if 'プット' in e['zone'] else 'var(--call)' if 'コール' in e['zone'] else 'var(--yellow)'
            js += "h+='<div style=\"background:var(--card);border-radius:6px;padding:10px;margin:6px 0\">';"
            js += "h+='<div style=\"display:flex;justify-content:space-between;align-items:center\">';"
            js += "h+='<span style=\"color:%s;font-weight:600\">%s</span>';" % (zone_color, _js_str(e['zone']))
            js += "h+='<span style=\"color:var(--yellow)\">%s</span>';" % stars
            js += "h+='</div>';"
            if e['wall_max_oi']:
                js += "h+='<div style=\"font-size:10px;color:var(--sub)\">壁: %s枚 @%s</div>';" % (
                    _js_str(fnum(e['wall_max_oi'])), _js_str(fnum(e['wall_strike'])))
            js += "h+='</div>';"

    # P&L Simulator link
    js += "h+='<div style=\"margin:14px 0;text-align:center\">';"
    js += "h+='<a href=\"pnl_simulator.html\" style=\"display:inline-block;padding:10px 24px;background:var(--accent);color:#fff;border-radius:8px;text-decoration:none;font-family:Outfit;font-weight:600;font-size:13px\">📊 P&Lシミュレーターを開く →</a>';"
    js += "h+='</div>';"

    js += "return h;"
    return js


# ============================================================
# P&L Simulator Builder (placeholder - full version next)
# ============================================================

def build_simulator_html(data):
    """Generate pnl_simulator.html."""
    meta = data['metadata']
    s01 = data.get('s01', {})
    s11 = data.get('s11', {})
    atm = meta.get('atm', 0)
    vi = s01.get('vi', 0)
    days_to_sq = meta.get('days_to_sq', 0)
    presets = s11.get('presets', [])

    h = '<!DOCTYPE html>\n<html lang="ja">\n<head>\n'
    h += '<meta charset="UTF-8">\n'
    h += '<meta name="viewport" content="width=device-width,initial-scale=1">\n'
    h += '<title>P&L Simulator %s</title>\n' % esc(meta.get('date_formatted', ''))
    h += '<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700&family=Noto+Sans+JP:wght@400;500;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">\n'
    h += '<style>\n%s\n' % DASHBOARD_CSS
    # Additional simulator styles
    h += '.sim-section{max-width:900px;margin:0 auto;padding:16px}\n'
    h += '.preset-btn{display:inline-block;padding:6px 14px;margin:4px;background:var(--card);border:1px solid var(--border);border-radius:6px;color:var(--text);cursor:pointer;font-size:11px;font-family:Outfit}\n'
    h += '.preset-btn:hover{border-color:var(--accent);color:var(--accent)}\n'
    h += '.leg-row{display:flex;gap:8px;align-items:center;margin:4px 0;flex-wrap:wrap}\n'
    h += '.leg-row select,.leg-row input{background:var(--card);border:1px solid var(--border);color:var(--text);padding:4px 8px;border-radius:4px;font-size:12px;font-family:DM Mono}\n'
    h += '.leg-row input{width:80px}\n'
    h += '.btn{padding:6px 16px;background:var(--accent);color:#fff;border:none;border-radius:6px;cursor:pointer;font-family:Outfit;font-size:12px}\n'
    h += '.btn-outline{background:transparent;border:1px solid var(--border);color:var(--text)}\n'
    h += '.btn-outline:hover{border-color:var(--accent)}\n'
    h += '#pnl-canvas{width:100%;height:300px;background:var(--card);border-radius:8px;margin:12px 0}\n'
    h += '.pnl-table{font-size:11px}\n'
    h += '.pnl-table td,.pnl-table th{padding:3px 6px}\n'
    h += '</style>\n'
    h += '</head>\n<body>\n'

    # Topbar
    h += '<div class="topbar">\n'
    h += '  <span class="logo">P&L Simulator</span>\n'
    h += '  <nav>\n'
    h += '    <a href="index.html">← ダッシュボード</a>\n'
    h += '    <a href="archive.html">アーカイブ</a>\n'
    h += '  </nav>\n'
    h += '</div>\n'

    # Hero
    h += '<div class="hero">\n'
    h += '  <h1>P&L Simulator</h1>\n'
    h += '  <div class="sub">%s / ATM %s / VI %s / SQまで%d日</div>\n' % (
        esc(meta.get('date_formatted', '')), fnum(atm), vi, days_to_sq)
    h += '</div>\n'

    # KPI Strip
    h += '<div class="kpi-strip">\n'
    h += '  <div class="kpi"><div class="label">ATM</div><div class="value">%s</div></div>\n' % fnum(atm)
    h += '  <div class="kpi"><div class="label">VI</div><div class="value">%s</div></div>\n' % vi
    h += '  <div class="kpi"><div class="label">SQ</div><div class="value">%s</div></div>\n' % esc(meta.get('sq_date', ''))
    h += '  <div class="kpi"><div class="label">残り営業日</div><div class="value">%d</div></div>\n' % days_to_sq
    h += '</div>\n'

    h += '<div class="sim-section">\n'

    # Presets
    h += '<div style="margin:12px 0">\n'
    h += '<div style="color:var(--sub);font-size:11px;margin-bottom:6px">プリセット戦略</div>\n'
    for i, p in enumerate(presets):
        h += '<span class="preset-btn" data-preset="%d">%s</span>\n' % (i, esc(p['name']))
    h += '</div>\n'

    # Leg builder
    h += '<div style="margin:12px 0">\n'
    h += '<div style="display:flex;gap:8px;margin-bottom:8px">\n'
    h += '  <button class="btn-outline btn" id="add-put">+ プット</button>\n'
    h += '  <button class="btn-outline btn" id="add-call">+ コール</button>\n'
    h += '  <button class="btn-outline btn" id="add-fut">+ 先物mini</button>\n'
    h += '  <button class="btn-outline btn" id="clear-legs">クリア</button>\n'
    h += '  <button class="btn" id="calc-btn">計算</button>\n'
    h += '</div>\n'
    h += '<div id="legs-container"></div>\n'
    h += '</div>\n'

    # Results
    h += '<div id="result-section" style="display:none">\n'
    h += '  <div class="kpi-strip" id="result-kpis"></div>\n'
    h += '  <canvas id="pnl-canvas" width="860" height="300"></canvas>\n'
    h += '  <div id="pnl-table-wrap"></div>\n'
    h += '</div>\n'

    h += '</div>\n'  # sim-section

    # Footer
    h += '<div class="footer"><a href="index.html">ダッシュボード</a> <a href="archive.html">アーカイブ</a></div>\n'

    # JavaScript
    h += '<script>\n'
    h += 'var ATM=%d,VI=%s,T=%s,DAYS=%d;\n' % (atm or 0, vi or 0, round(days_to_sq / 250, 6) if days_to_sq else 0, days_to_sq)

    # Embed presets as JS array
    h += 'var PRESETS=%s;\n' % json.dumps(presets, ensure_ascii=False)

    # P&L simulator JS
    h += _build_simulator_js()
    h += '</script>\n'
    h += '</body>\n</html>'

    return h


def _build_simulator_js():
    """Return the full simulator JavaScript code."""
    return r"""
var legs=[];
var legId=0;

function normCdf(x){
  if(x>=0){var t=1/(1+0.2316419*x)}
  else{var t=1/(1-0.2316419*x)}
  var d=0.3989422804014327;
  var p=((((1.330274429*t-1.821255978)*t+1.781477937)*t-0.356563782)*t+0.319381530)*t;
  if(x>=0)return 1-d*Math.exp(-0.5*x*x)*p;
  return d*Math.exp(-0.5*x*x)*p;
}

function bsPrice(type,K,F,s,T){
  if(T<=0||s<=0){return type==='put'?Math.max(K-F,0):Math.max(F-K,0);}
  var sqT=Math.sqrt(T);
  var d1=(Math.log(F/K)+0.5*s*s*T)/(s*sqT);
  var d2=d1-s*sqT;
  if(type==='put')return K*normCdf(-d2)-F*normCdf(-d1);
  return F*normCdf(d1)-K*normCdf(d2);
}

function addLeg(type,side,strike,premium,qty,mult){
  var id='leg'+legId++;
  legs.push({id:id,type:type,side:side||'short',strike:strike||ATM,premium:premium||0,qty:qty||1,mult:mult||(type==='futures'?100:1000),entry:strike||ATM});
  renderLegs();
}

function renderLegs(){
  var c=document.getElementById('legs-container');
  var h='';
  for(var i=0;i<legs.length;i++){
    var L=legs[i];
    h+='<div class="leg-row" data-lid="'+L.id+'">';
    h+='<select class="leg-side"><option value="short"'+(L.side==='short'?' selected':'')+'>売</option><option value="long"'+(L.side==='long'?' selected':'')+'>買</option></select>';
    h+='<span style="color:var(--sub);font-size:11px;width:50px">'+L.type+'</span>';
    if(L.type==='futures'){
      h+='<label style="font-size:10px;color:var(--sub)">Entry</label><input class="leg-entry" type="number" value="'+L.entry+'" step="500">';
    }else{
      h+='<label style="font-size:10px;color:var(--sub)">K</label><input class="leg-strike" type="number" value="'+L.strike+'" step="500">';
      h+='<label style="font-size:10px;color:var(--sub)">Prem</label><input class="leg-prem" type="number" value="'+L.premium+'" step="10">';
    }
    h+='<label style="font-size:10px;color:var(--sub)">枚</label><input class="leg-qty" type="number" value="'+L.qty+'" step="1" style="width:50px">';
    h+='<select class="leg-mult"><option value="1000"'+(L.mult===1000?' selected':'')+'>x1000</option><option value="100"'+(L.mult===100?' selected':'')+'>x100</option></select>';
    h+='<span style="cursor:pointer;color:var(--red);font-size:14px" class="leg-del">✕</span>';
    h+='</div>';
  }
  c.innerHTML=h;
}

function readLegsFromDOM(){
  var rows=document.querySelectorAll('.leg-row');
  for(var i=0;i<rows.length;i++){
    var lid=rows[i].getAttribute('data-lid');
    for(var j=0;j<legs.length;j++){
      if(legs[j].id===lid){
        var L=legs[j];
        L.side=rows[i].querySelector('.leg-side').value;
        L.qty=parseInt(rows[i].querySelector('.leg-qty').value)||1;
        L.mult=parseInt(rows[i].querySelector('.leg-mult').value)||1000;
        var sk=rows[i].querySelector('.leg-strike');
        if(sk)L.strike=parseInt(sk.value)||ATM;
        var pr=rows[i].querySelector('.leg-prem');
        if(pr)L.premium=parseFloat(pr.value)||0;
        var en=rows[i].querySelector('.leg-entry');
        if(en)L.entry=parseInt(en.value)||ATM;
      }
    }
  }
}

function calculate(){
  readLegsFromDOM();
  if(legs.length===0)return;
  var sqLow=ATM-6000,sqHigh=ATM+6000,step=500;
  var results=[];
  var maxProfit=-Infinity,maxLoss=Infinity;
  for(var sq=sqLow;sq<=sqHigh;sq+=step){
    var total=0;
    var legPnls=[];
    for(var i=0;i<legs.length;i++){
      var L=legs[i];
      var pnl=0;
      if(L.type==='put'){
        var intrinsic=Math.max(L.strike-sq,0);
        pnl=L.side==='short'?(L.premium-intrinsic):(intrinsic-L.premium);
      }else if(L.type==='call'){
        var intrinsic=Math.max(sq-L.strike,0);
        pnl=L.side==='short'?(L.premium-intrinsic):(intrinsic-L.premium);
      }else{
        pnl=L.side==='short'?(L.entry-sq):(sq-L.entry);
      }
      var yen=pnl*L.qty*L.mult;
      legPnls.push(yen);
      total+=yen;
    }
    results.push({sq:sq,total:total,legs:legPnls});
    if(total>maxProfit)maxProfit=total;
    if(total<maxLoss)maxLoss=total;
  }
  drawChart(results,maxProfit,maxLoss);
  drawTable(results);
  // KPIs
  var be=[];
  for(var i=1;i<results.length;i++){
    if((results[i-1].total<=0&&results[i].total>0)||(results[i-1].total>=0&&results[i].total<0)){
      be.push(results[i].sq);
    }
  }
  var kh=document.getElementById('result-kpis');
  var khtml='<div class="kpi"><div class="label">最大利益</div><div class="value up">'+fmtYen(maxProfit)+'</div></div>';
  khtml+='<div class="kpi"><div class="label">最大損失</div><div class="value down">'+fmtYen(maxLoss)+'</div></div>';
  khtml+='<div class="kpi"><div class="label">損益分岐</div><div class="value">'+be.join(' / ')+'</div></div>';
  kh.innerHTML=khtml;
  document.getElementById('result-section').style.display='block';
}

function fmtYen(n){
  if(Math.abs(n)>=10000)return (n/10000).toFixed(1)+'万円';
  return n.toLocaleString()+'円';
}

function drawChart(results,maxP,maxL){
  var canvas=document.getElementById('pnl-canvas');
  var ctx=canvas.getContext('2d');
  var W=canvas.width,H=canvas.height;
  ctx.clearRect(0,0,W,H);
  ctx.fillStyle='#111128';
  ctx.fillRect(0,0,W,H);
  var pad={l:60,r:20,t:20,b:30};
  var cw=W-pad.l-pad.r,ch=H-pad.t-pad.b;
  var range=Math.max(Math.abs(maxP),Math.abs(maxL))*1.1||1;
  // Zero line
  var zy=pad.t+ch/2;
  ctx.strokeStyle='rgba(255,255,255,.15)';
  ctx.setLineDash([4,4]);
  ctx.beginPath();ctx.moveTo(pad.l,zy);ctx.lineTo(W-pad.r,zy);ctx.stroke();
  ctx.setLineDash([]);
  // ATM line
  var atmIdx=-1;
  for(var i=0;i<results.length;i++){if(results[i].sq===ATM){atmIdx=i;break;}}
  if(atmIdx>=0){
    var ax=pad.l+(atmIdx/(results.length-1))*cw;
    ctx.strokeStyle='rgba(251,191,36,.4)';
    ctx.setLineDash([4,4]);
    ctx.beginPath();ctx.moveTo(ax,pad.t);ctx.lineTo(ax,H-pad.b);ctx.stroke();
    ctx.setLineDash([]);
  }
  // P&L curve
  ctx.beginPath();
  for(var i=0;i<results.length;i++){
    var x=pad.l+(i/(results.length-1))*cw;
    var y=pad.t+ch/2-(results[i].total/range)*(ch/2);
    if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y);
  }
  ctx.strokeStyle='#818cf8';ctx.lineWidth=2;ctx.stroke();
  // Fill
  ctx.lineTo(pad.l+cw,zy);ctx.lineTo(pad.l,zy);ctx.closePath();
  ctx.fillStyle='rgba(129,140,248,.1)';ctx.fill();
  // Labels
  ctx.fillStyle='#8888aa';ctx.font='10px DM Mono';ctx.textAlign='center';
  for(var i=0;i<results.length;i+=2){
    var x=pad.l+(i/(results.length-1))*cw;
    ctx.fillText(results[i].sq,x,H-pad.b+14);
  }
  ctx.textAlign='right';
  ctx.fillText(fmtYen(Math.round(range)),pad.l-4,pad.t+10);
  ctx.fillText(fmtYen(Math.round(-range)),pad.l-4,H-pad.b-2);
  ctx.fillText('0',pad.l-4,zy+4);
}

function drawTable(results){
  var w=document.getElementById('pnl-table-wrap');
  var h='<table class="pnl-table"><tr><th>SQ着地</th>';
  for(var i=0;i<legs.length;i++){
    var L=legs[i];
    var lbl=L.type==='futures'?'先物':(L.type==='put'?'P':'C')+L.strike+(L.side==='short'?'売':'買');
    h+='<th>'+lbl+'</th>';
  }
  h+='<th>合計P&L</th></tr>';
  for(var r=0;r<results.length;r++){
    var R=results[r];
    var cls=R.sq===ATM?' class="atm-row"':'';
    h+='<tr'+cls+'><td style="font-family:DM Mono">'+R.sq+'</td>';
    for(var j=0;j<R.legs.length;j++){
      var c=R.legs[j]>=0?'positive':'negative';
      h+='<td class="'+c+'" style="font-family:DM Mono">'+fmtYen(Math.round(R.legs[j]))+'</td>';
    }
    var tc=R.total>=0?'positive':'negative';
    h+='<td class="'+tc+'" style="font-family:DM Mono;font-weight:600">'+fmtYen(Math.round(R.total))+'</td>';
    h+='</tr>';
  }
  h+='</table>';
  w.innerHTML=h;
}

// Event listeners
document.getElementById('add-put').addEventListener('click',function(){
  var prem=Math.round(bsPrice('put',ATM-2000,ATM,VI/100,T));
  addLeg('put','short',ATM-2000,prem,1,1000);
});
document.getElementById('add-call').addEventListener('click',function(){
  var prem=Math.round(bsPrice('call',ATM+2000,ATM,VI/100,T));
  addLeg('call','short',ATM+2000,prem,1,1000);
});
document.getElementById('add-fut').addEventListener('click',function(){
  addLeg('futures','short',ATM,0,1,100);
});
document.getElementById('clear-legs').addEventListener('click',function(){
  legs=[];renderLegs();
  document.getElementById('result-section').style.display='none';
});
document.getElementById('calc-btn').addEventListener('click',calculate);

// Preset loading
document.addEventListener('click',function(e){
  var el=e.target;
  if(el.classList.contains('preset-btn')){
    var idx=parseInt(el.getAttribute('data-preset'));
    if(PRESETS[idx]){
      legs=[];legId=0;
      var P=PRESETS[idx];
      for(var i=0;i<P.legs.length;i++){
        var L=P.legs[i];
        legs.push({id:'leg'+legId++,type:L.type,side:L.side,strike:L.strike||0,premium:L.premium||0,qty:L.qty||1,mult:L.multiplier||1000,entry:L.entry||ATM});
      }
      renderLegs();
      calculate();
    }
  }
});

// Delete leg
document.getElementById('legs-container').addEventListener('click',function(e){
  if(e.target.classList.contains('leg-del')){
    var row=e.target.closest('.leg-row');
    var lid=row.getAttribute('data-lid');
    legs=legs.filter(function(l){return l.id!==lid;});
    renderLegs();
  }
});
"""


# ============================================================
# Archive Helpers
# ============================================================

def build_archive_snippet(data):
    """Generate archive.html insertion snippet."""
    meta = data['metadata']
    s01 = data.get('s01', {})
    nikkei = s01.get('nikkei_close', 0)
    vi = s01.get('vi', 0)

    WEEKDAYS = ['月', '火', '水', '木', '金', '土', '日']

    date_str = meta.get('date', '')
    dt = None
    try:
        from datetime import datetime as _dt
        dt = _dt.strptime(date_str, '%Y%m%d')
    except:
        pass

    weekday = WEEKDAYS[dt.weekday()] if dt else ''
    date_disp = '%s.%s.%s' % (date_str[:4], date_str[4:6], date_str[6:8]) if len(date_str) == 8 else date_str

    # Determine SQ section
    sq_section = meta.get('major_month', '0625')[-4:]  # MMDD of SQ month roughly

    vi_class = 'etag-vi high' if vi and vi > 30 else 'etag-vi'

    snippet = '<!-- archive-list-%s の先頭に追加 -->\n' % sq_section
    snippet += '<a href="JPX_portal_%s.html" class="entry">\n' % date_str
    snippet += '  <span class="entry-date">%s</span>\n' % date_disp
    snippet += '  <span class="entry-weekday">%s</span>\n' % weekday
    snippet += '  <span class="entry-tags">\n'
    snippet += '    <span class="etag etag-nikkei">日経平均 %s</span>\n' % fnum(nikkei)
    snippet += '    <span class="etag %s">VI %s</span>\n' % (vi_class, vi)
    snippet += '  </span>\n'
    snippet += '  <span class="entry-arrow">→</span>\n'
    snippet += '</a>\n'

    return snippet


def update_archive(archive_path, data):
    """Update archive.html by inserting new entry into the correct SQ section.
    
    Finds <div id="archive-list-XXXX"> and inserts the new entry at the top.
    If archive.html doesn't exist, skip silently.
    If the date already exists, skip to avoid duplicates.
    """
    if not os.path.exists(archive_path):
        print('[render.py] archive.html not found at %s — skipping auto-update' % archive_path)
        return False

    meta = data['metadata']
    date_str = meta.get('date', '')
    s01 = data.get('s01', {})
    nikkei = s01.get('nikkei_close', 0)
    vi = s01.get('vi', 0)

    WEEKDAYS = ['月', '火', '水', '木', '金', '土', '日']
    dt = None
    try:
        from datetime import datetime as _dt
        dt = _dt.strptime(date_str, '%Y%m%d')
    except:
        pass

    weekday = WEEKDAYS[dt.weekday()] if dt else ''
    date_disp = '%s.%s.%s' % (date_str[:4], date_str[4:6], date_str[6:8]) if len(date_str) == 8 else date_str

    # Determine SQ section ID: major month as MMDD of SQ
    major_month = meta.get('major_month', '')  # e.g. '202606'
    if len(major_month) >= 6:
        mm = major_month[4:6]  # '06'
        # SQ = 2nd Friday, typically around 12-13th
        section_id = 'archive-list-%s' % mm  # e.g. 'archive-list-06'
    else:
        section_id = 'archive-list-06'

    # Build entry HTML
    vi_class = 'etag-vi high' if vi and vi > 30 else 'etag-vi'
    nk_class = 'etag-nikkei'

    entry = ''
    entry += '    <a href="JPX_portal_%s.html" class="entry">\n' % date_str
    entry += '      <span class="entry-date">%s</span>\n' % date_disp
    entry += '      <span class="entry-weekday">%s</span>\n' % weekday
    entry += '      <span class="entry-tags">\n'
    entry += '        <span class="etag %s">%s</span>\n' % (nk_class, fnum(nikkei) if nikkei else '-')
    entry += '        <span class="etag %s">VI %s</span>\n' % (vi_class, vi if vi else '-')
    entry += '      </span>\n'
    entry += '      <span class="entry-arrow">&rarr;</span>\n'
    entry += '    </a>\n'

    # Read existing archive.html
    with open(archive_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Check for duplicate
    portal_link = 'JPX_portal_%s.html' % date_str
    if portal_link in html:
        print('[render.py] archive.html already contains entry for %s — skipping' % date_str)
        return False

    # Find the target section and insert
    import re
    # Pattern: <div id="archive-list-XX"> (possibly with other attributes)
    pattern = r'(<div[^>]*id=["\']%s["\'][^>]*>)' % re.escape(section_id)
    match = re.search(pattern, html)

    if match:
        # Insert entry right after the opening div tag
        insert_pos = match.end()
        html = html[:insert_pos] + '\n' + entry + html[insert_pos:]
        print('[render.py] Inserted entry into %s' % section_id)
    else:
        # Section not found — try broader patterns
        # Look for any archive-list div
        alt_patterns = [
            r'(<div[^>]*id=["\']archive-list-\d+["\'][^>]*>)',  # archive-list-0625 etc
            r'(<div[^>]*id=["\']archive-list[^"\']*["\'][^>]*>)',
        ]
        inserted = False
        for alt_pat in alt_patterns:
            match = re.search(alt_pat, html)
            if match:
                insert_pos = match.end()
                html = html[:insert_pos] + '\n' + entry + html[insert_pos:]
                print('[render.py] Inserted entry into first available archive-list (section %s not found)' % section_id)
                inserted = True
                break

        if not inserted:
            print('[render.py] WARNING: No archive-list section found in archive.html — entry not inserted')
            return False

    # Write updated archive.html
    with open(archive_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return True


# ============================================================
# Main Pipeline
# ============================================================

def run(args):
    with open(args.data, 'r', encoding='utf-8') as f:
        data = json.load(f)

    outdir = args.outdir
    os.makedirs(outdir, exist_ok=True)

    date_str = data['metadata'].get('date', 'unknown')

    # 1. Markdown report
    md_path = os.path.join(outdir, 'JPX_market_analysis_%s.md' % date_str)
    md = build_markdown(data)
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md)
    print('[render.py] Markdown: %s (%.1f KB)' % (md_path, os.path.getsize(md_path) / 1024))

    # 2. Dashboard HTML
    html_path = os.path.join(outdir, 'index.html')
    html = build_dashboard_html(data)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print('[render.py] Dashboard: %s (%.1f KB)' % (html_path, os.path.getsize(html_path) / 1024))

    # 3. P&L Simulator
    sim_path = os.path.join(outdir, 'pnl_simulator.html')
    sim = build_simulator_html(data)
    with open(sim_path, 'w', encoding='utf-8') as f:
        f.write(sim)
    print('[render.py] Simulator: %s (%.1f KB)' % (sim_path, os.path.getsize(sim_path) / 1024))

    # 4. Archive portal copy
    portal_path = os.path.join(outdir, 'JPX_portal_%s.html' % date_str)
    with open(portal_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print('[render.py] Portal: %s' % portal_path)

    # 5. Archive snippet (for reference)
    snippet = build_archive_snippet(data)
    snippet_path = os.path.join(outdir, 'archive_snippet_%s.txt' % date_str)
    with open(snippet_path, 'w', encoding='utf-8') as f:
        f.write(snippet)
    print('[render.py] Snippet: %s' % snippet_path)

    # 6. Auto-update archive.html
    archive_path = os.path.join(outdir, 'archive.html')
    updated = update_archive(archive_path, data)
    if updated:
        print('[render.py] archive.html updated successfully')

    print('\n[render.py] Done. Files generated.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='JPX Market Analysis - Renderer')
    parser.add_argument('--data', default='data.json', help='Input data.json path')
    parser.add_argument('--outdir', default='.', help='Output directory')
    args = parser.parse_args()
    run(args)
