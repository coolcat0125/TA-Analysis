# -*- coding: utf-8 -*-
"""capacity_retry_nosrc_v4933.py — 容量无源行重试（车型名查询通道，可复跑）

Input：capacity_autohome_evidence_20260927.json 中 ABSTAIN_no_series_brand_guard 行
      （已补 car_name 元数据）。
通道：车型名称 suggest（比通用名更不易撞名）→ brand guard → 全款型矩阵 →
      S4 代际窗口 + 配置签名锚定（续航±8km AND 整备±40kg 严格双信号，缺一弃权）。
产出：audit-output/capacity_retry_nosrc_20260929.json（提名 FIX_ANCHORED 仍须人工复核后 --apply）
"""
import json
import os
import re
import time
from collections import defaultdict

import media_fill as mf
from capacity_evidence_v4925 import query_for, spec_year, batch_ym, PH
from capacity_anchor_v4931 import fetch_matrix, fnum

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'audit-output', 'capacity_autohome_evidence_20260927.json')
OUT = os.path.join(HERE, 'audit-output', 'capacity_retry_nosrc_20260929.json')
CACHE = os.path.join(HERE, 'audit-output', '_cap_retry_cache.json')


def main():
    ev = json.load(open(SRC, encoding='utf-8'))
    jobs = [p for p in ev['proposals']
            if str(p.get('verdict', '')).startswith('ABSTAIN')
            and ('no_series' in p.get('verdict', '') or p.get('verdict') == 'ABSTAIN_NO_SRC')
            and p.get('car_name')]
    print(f'重试目标 {len(jobs)} 行')
    cache = {}
    if os.path.exists(CACHE):
        cache = json.load(open(CACHE, encoding='utf-8'))
    ym = batch_ym()
    BY = {b: (y + (1 if m >= 7 else 0)) for b, (y, m) in ym.items()}

    stats = defaultdict(int)
    for i, p in enumerate(jobs, 1):
        cname = str(p['car_name']).strip()
        q = query_for(re.sub(r'[（(].*?[)）]', '', cname).strip()) or cname
        ckey = q
        if ckey not in cache:
            try:
                sug = mf.suggest(q)
            except Exception as e:
                cache[ckey] = {'err': f'suggest:{type(e).__name__}'}
                continue
            time.sleep(0.55)
            cand = mf.find_series(sug, p.get('power_type', ''), name=cname,
                                  brand=str(p.get('enterprise') or ''))
            if cand is None:
                cache[ckey] = {'err': 'brand_guard_or_no_hit'}
                continue
            try:
                sid = int(str(cand['wordid']))
            except Exception:
                cache[ckey] = {'err': 'wordid'}
                continue
            time.sleep(0.55)
            try:
                rows, err = fetch_matrix(sid)
            except Exception as e:
                rows, err = None, f'{type(e).__name__}'
            cache[ckey] = {'sid': sid, 'series_name': cand['key'], 'rows': rows, 'err': err}
            time.sleep(0.55)
        c = cache[ckey]
        p['retry'] = {'query': q, 'sid': c.get('sid'), 'series_name': c.get('series_name'),
                      'err': c.get('err')}
        rows = c.get('rows')
        if not rows:
            stats['no_src'] += 1
            continue
        b = int(p['batch']) if str(p['batch']).isdigit() else None
        cap_year = BY.get(b)
        rng = fnum(p.get('cur_range'))
        cw = fnum(p.get('cur_cw'))
        elig = [r for r in rows if r['year'] is not None and cap_year is not None and r['year'] <= cap_year]
        if not elig:
            p['retry']['verdict'] = 'FUTURE_ONLY'
            stats['future_only'] += 1
            continue
        cand2 = elig
        sig_ok = True
        if rng:
            c2 = [r for r in cand2 if r['rng'] is not None and abs(r['rng'] - rng) <= 8]
            if c2:
                cand2 = c2
            elif not cw:
                sig_ok = False  # 续航无命中且无整备信号 → 弃权（不兜底）
        if sig_ok and cw:
            c3 = [r for r in cand2 if r['cw'] is not None and abs(r['cw'] - cw) <= 40]
            if c3:
                cand2 = c3
            elif not rng:
                sig_ok = False
        caps_u = sorted({r['cap'] for r in cand2})
        if not sig_ok or len(cand2) != 1 or len(caps_u) != 1:
            p['retry']['verdict'] = f'AMBIG_{min(len(cand2), 9)}'
            stats['ambig'] += 1
            continue
        r0 = cand2[0]
        cur = fnum(p['cur_capacity'])
        if cur is not None and abs(cur - r0['cap']) < 0.5:
            p['retry']['verdict'] = 'KEEP_SAME'
            stats['keep_same'] += 1
        else:
            p['retry']['verdict'] = 'FIX_ANCHORED'
            p['fix_value'] = r0['cap']
            p['anchor_name'] = r0['name']
            p['anchor_rng'] = r0['rng']
            p['anchor_cw'] = r0['cw']
            p['sid'] = c.get('sid')
            p['series_name'] = c.get('series_name')
            p['cap_year'] = cap_year
            stats['fix_anchored'] += 1
            print(f"  FIX提名 {p['batch']}批 {p['model']} {p['generic']}: {p['cur_capacity']}→{r0['cap']}"
                  f"（「{r0['name']}」续航{r0['rng']}/整备{r0['cw']}）")
        if i % 15 == 0:
            json.dump(cache, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    json.dump(cache, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    rep = {'created': '2026-09-29', '纪律': '车型名查询+严格双信号（续航AND整备，缺一弃权）',
           'stats': dict(stats), 'proposals': jobs}
    json.dump(rep, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('重试结果:', dict(stats), '→', OUT)


if __name__ == '__main__':
    main()
