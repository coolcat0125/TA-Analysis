# -*- coding: utf-8 -*-
"""capacity_anchor_v4931.py — 容量矛盾·配置签名锚定轮（可辩护更正通道，可复跑）

背景：媒体矩阵"唯一在售值"不可作历史行更正证据（09-27 教训）。本轮改用
**配置签名锚定**：矩阵内款型若与底表行的 官方续航(±8km) + 官方整备(±40kg)
双信号唯一命中，则该款型容量即行配置容量——配置身份已锚定，可辩护。
用法：
  python capacity_anchor_v4931.py --fetch    # 重抓多值行的全系款型矩阵（含续航/整备）
  python capacity_anchor_v4931.py --apply    # 锚定唯一者落库（浅红+台账）
"""
import argparse
import json
import os
import re
import time
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill

import media_fill as mf
from capacity_evidence_v4925 import QUERY_ALIAS, query_for, spec_year, batch_ym, PH

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SRC_EV = os.path.join(HERE, 'audit-output', 'capacity_autohome_evidence_20260927.json')
OUT = os.path.join(HERE, 'audit-output', 'capacity_anchor_v4931.json')
CACHE = os.path.join(HERE, 'audit-output', '_cap_anchor_cache.json')
RED = PatternFill('solid', fgColor='FFC7CE')


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', ''))
    except Exception:
        return None


def fetch_matrix(seriesid):
    """全款型矩阵：(year, 容量, 续航, 整备, 款型名)[]"""
    ls = mf.list_spec(seriesid)
    specs = ls.get('specs') or []
    if not specs:
        return None, 'no_specs'
    u = (f"https://car-web-api.autohome.com.cn/car/param/getParamConf"
         f"?mode=1&site=2&seriesid={seriesid}&specid={specs[0]['id']}")
    d = mf.get(u)
    r = d.get('result') or {}
    tm = {}
    for grp in r.get('titlelist') or []:
        for it in grp.get('items') or []:
            tm[it.get('titleid')] = str(it.get('itemname') or '')
    tids = {}
    for tid, nm in tm.items():
        if ('电池能量' in nm or '电池容量' in nm) and 'cap' not in tids:
            tids['cap'] = tid
        elif '续航' in nm and 'rng' not in tids:
            tids['rng'] = tid
        elif '整备质量' in nm and 'cw' not in tids:
            tids['cw'] = tid
    if 'cap' not in tids:
        return None, 'no_capacity_param'
    out = []
    for dl in r.get('datalist') or []:
        sn = str(dl.get('specname') or '')
        vals = {}
        for pc in dl.get('paramconflist') or []:
            tid = pc.get('titleid')
            if tid in tids.values():
                v = pc.get('itemname')
                if pc.get('sublist') and pc['sublist']:
                    v = pc['sublist'][0].get('name')
                vals[[k for k, tv in tids.items() if tv == tid][0]] = v
        cap = fnum(vals.get('cap'))
        if cap is None or not (5 <= cap <= 150):
            continue
        out.append({'year': spec_year(sn), 'cap': cap,
                    'rng': fnum(vals.get('rng')), 'cw': fnum(vals.get('cw')),
                    'name': sn[:36]})
    return out, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--fetch', action='store_true')
    ap.add_argument('--apply', action='store_true')
    a = ap.parse_args()
    ev = json.load(open(SRC_EV, encoding='utf-8'))
    multi = [p for p in ev['proposals'] if p.get('verdict') == 'ABSTAIN_MULTI']
    ym = batch_ym()
    BY = {b: (y + (1 if m >= 7 else 0)) for b, (y, m) in ym.items()}

    if a.apply:
        rep = json.load(open(OUT, encoding='utf-8'))
        fixes = [p for p in rep['proposals'] if p.get('verdict') == 'FIX_ANCHORED']
        print(f"锚定提名 {len(fixes)} 格")
        wb = openpyxl.load_workbook(WB)
        ws = wb['NEV公告参数汇总']
        hdr = [str(c.value or '').strip() for c in ws[1]]
        hi = {h: i for i, h in enumerate(hdr)}
        led = wb['变更记录']
        seq = max((r[0] for r in led.iter_rows(min_row=2, values_only=True)
                   if isinstance(r[0], int)), default=0)
        n = 0
        for p in fixes:
            tgt = None
            for row in ws.iter_rows(min_row=2):
                if (str(row[hi['批次']].value).strip() == str(p['batch'])
                        and str(row[hi['产品型号']].value or '').strip() == p['model']):
                    tgt = row
                    break
            if tgt is None:
                continue
            cell = tgt[hi['电池容量(kWh)']]
            old = cell.value
            cell.value = p['fix_value']
            cell.fill = RED
            seq += 1
            led.append((seq, p['batch'], p['model'], p.get('power_type', ''), '电池容量(kWh)',
                        '媒体容量核证更正(配置签名锚定)', str(old), str(p['fix_value']),
                        f"汽车之家款型矩阵（{p['series_name']} sid={p['sid']}；款型「{p['anchor_name']}」"
                        f"续航 {p['anchor_rng']}km/整备 {p['anchor_cw']}kg 与本行官方值双信号唯一命中，"
                        f"检索 2026-09-28）"))
            n += 1
        wb.save(WB)
        print(f"已落库 {n} 格（台账至 {seq}）——三件套+看板+提交")
        return

    # ---- fetch ----
    cache = {}
    if os.path.exists(CACHE):
        cache = json.load(open(CACHE, encoding='utf-8'))
    for i, p in enumerate(multi, 1):
        base = re.split(r'[,，、;；]', str(p['generic'] or ''))[0].strip()
        if PH.search(base):
            base = str(p.get('model') or '')
        if not base or len(base) < 2 or p.get('sid') is None:
            continue
        sid = str(p['sid'])
        if sid in cache:
            continue
        try:
            caps, err = fetch_matrix(int(sid))
        except Exception as e:
            caps, err = None, f'{type(e).__name__}'
        cache[sid] = {'series_name': p.get('series_name'), 'rows': caps, 'err': err}
        print(f'[{i}/{len(multi)}] {base} sid={sid} -> {"ok "+str(len(caps or []))+"款" if caps else err}')
        json.dump(cache, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        time.sleep(0.55)

    # ---- 锚定判定 ----
    wbs = openpyxl.load_workbook(WB, read_only=True)
    wss = wbs['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(wss.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    rowmap = {}
    for r in wss.iter_rows(min_row=2, values_only=True):
        rowmap[(str(r[hi['批次']]).strip(), str(r[hi['产品型号']] or '').strip())] = r
    wbs.close()
    stats = defaultdict(int)
    for p in multi:
        rw = rowmap.get((str(p['batch']), str(p['model']).strip()))
        if rw is not None:
            p['cur_range'] = str(rw[hi['纯电续航里程(km)']] or '')
            p['cur_cw'] = str(rw[hi['整备质量(kg)']] or '')
            p['power_type'] = str(rw[hi['动力类型']] or '')
        sid = str(p.get('sid'))
        rows = (cache.get(sid) or {}).get('rows')
        if not rows:
            p['anchor_verdict'] = 'NO_MATRIX'
            stats['no_matrix'] += 1
            continue
        b = int(p['batch']) if str(p['batch']).isdigit() else None
        cap_year = BY.get(b)
        rng = fnum(p.get('cur_range'))
        cw = fnum(p.get('cur_cw')) if 'cur_cw' in p else None
        # 行整备另行直读
        elig = [r for r in rows
                if r['year'] is not None and cap_year is not None and r['year'] <= cap_year]
        if not elig:
            p['anchor_verdict'] = 'FUTURE_ONLY'
            stats['future_only'] += 1
            continue
        cand = elig
        if rng:
            c2 = [r for r in cand if r['rng'] is not None and abs(r['rng'] - rng) <= 8]
            if c2:
                cand = c2
        if cw:
            c3 = [r for r in cand if r['cw'] is not None and abs(r['cw'] - cw) <= 40]
            if c3:
                cand = c3
        caps_u = sorted({r['cap'] for r in cand})
        if len(cand) == 1 and len(caps_u) == 1:
            cur = fnum(p['cur_capacity'])
            r0 = cand[0]
            p['anchor_verdict'] = 'KEEP_SAME' if (cur is not None and abs(cur - r0['cap']) < 0.5) \
                else 'FIX_ANCHORED'
            if p['anchor_verdict'] == 'FIX_ANCHORED':
                p['fix_value'] = r0['cap']
                p['anchor_name'] = r0['name']
                p['anchor_rng'] = r0['rng']
                p['anchor_cw'] = r0['cw']
            stats[p['anchor_verdict'].lower()] += 1
        else:
            p['anchor_verdict'] = f'AMBIG_{len(cand)}'
            stats[f'amb_{min(len(cand),9)}'] += 1
    rep = {'created': '2026-09-28', '纪律': '配置签名锚定=续航±8km+整备±40kg 双信号唯一命中',
           'stats': dict(stats), 'proposals': multi}
    json.dump(rep, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('锚定结果:', dict(stats), '→', OUT)


if __name__ == '__main__':
    main()
