# -*- coding: utf-8 -*-
"""integrate_dcd_refill.py — 懂车帝补缺批次整合（门禁 Agent · 只补空+双源红修）

输入：raw/dcd_refill/（series_mapping.json 人审 13 车系 + dcd_sXXXX.json 逐款九字段）
规则：
  1. 只补空：9 个映射字段仅写空格；唯一值规则（并列候选值不一致则放弃该格）
  2. 红修例外：行 ∈ T4 补丁锚点 且 dcd 值=T4 新值（双源收敛）→ 更正污染值（浅红+台账双源留痕）
  3. 动力类型守卫：BEV 行只匹配有电池的 dcd 款型；ff=汽油 款型排除
  4. 修剪匹配：续航(名称km或CLTC, ±15) + 整备(±40kg) 打分取最优；平手时字段唯一值才写
  5. 电池类型归一：三元锂电池→三元锂 等；功率 '35kW(48Ps)'→35
"""
import json
import math
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'raw', 'dcd_refill')
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
T4 = os.path.join(HERE, 'candidate', 'v4.9', 'vehicle_length', 'patch_exact_20260918.csv')

GREEN = PatternFill('solid', fgColor='C6EFCE')
RED = PatternFill('solid', fgColor='FFC7CE')
FIELD_MAP = {
    '车长(mm)': 'length', '轴距(mm)': 'wheelbase', '整备质量(kg)': 'curb_weight',
    '纯电续航里程(km)': 'cltc_recharge_mileage', '百公里电耗(kWh/100km)': 'power_consumption',
    '电池容量(kWh)': 'battery_capacity', '电池类型': 'battery_type',
    '电池能量密度(Wh/kg)': 'battery_energy_density',
}
BT_NORM = {'三元锂电池': '三元锂', '磷酸铁锂电池': '磷酸铁锂', '锰酸锂电池': '锰酸锂', '钛酸锂电池': '钛酸锂'}
DCD_URL = 'https://www.dongchedi.com/auto/series/{sid}（懂车帝参数页，检索 2026-09-20）'


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', '').replace('km', '').strip())
    except Exception:
        return None


def parse_power(v):
    """'35kW(48Ps)' → 35；'130kW/270N·m' 首段。"""
    m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*kW', str(v), re.I)
    return float(m.group(1)) if m else None


def main():
    apply = '--apply' in sys.argv
    mapping = json.load(open(os.path.join(SRC, 'series_mapping.json'), encoding='utf-8'))
    accepted = [m for m in mapping['reviewed'] if m.get('ok')]
    try:
        fuel = json.load(open(os.path.join(SRC, 'dcd_fuel_form.json'), encoding='utf-8'))
    except Exception:
        fuel = {}
    # fuel_form: {sid: [{id, name, ff}, ...]} → {(sid, specid): ff}
    fuel_ff = {}
    for sid_f, lst in fuel.items():
        for sp in (lst or []):
            if isinstance(sp, dict) and sp.get('id'):
                fuel_ff[(str(sid_f), str(sp['id']))] = str(sp.get('ff') or '')

    # T4 锚点集（红修白名单）：(批次, 型号, 字段) → 新值
    t4 = {}
    if os.path.exists(T4):
        import csv
        for p in csv.DictReader(open(T4, encoding='utf-8-sig')):
            t4[(p['批次'].strip(), p['产品型号'].strip(), p['字段'].strip())] = p['新值'].strip()

    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in ws[1]]
    hi = {h: i for i, h in enumerate(hdr)}
    ledger = wb['变更记录']
    seq = max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
               if isinstance(r[0], int)), default=0)

    # 估算底账集（v3.7.3 车长估算等）：台账变更类型/来源含"估算"的 (批次,型号,字段)
    est_keys = set()
    for r in ledger.iter_rows(min_row=2, values_only=True):
        if isinstance(r[0], int):
            if '估算' in str(r[5] or '') or '估算' in str(r[8] or ''):
                est_keys.add((str(r[1] or '').strip(), str(r[2] or '').strip(), str(r[4] or '').strip()))

    series_files = {}
    for f in os.listdir(SRC):
        m = re.match(r'dcd_s(\d+)\.json$', f)
        if m:
            series_files[m.group(1)] = json.load(open(os.path.join(SRC, f), encoding='utf-8'))

    rows = list(ws.iter_rows(min_row=2))
    stats = defaultdict(int)
    report = []

    for m in accepted:
        sid = str(m['sid'])
        if sid not in series_files:
            continue
        dcd = series_files[sid]
        kw = m['kw'].lower().replace(' ', '')
        kw2 = kw.replace('plus', '')
        fform = fuel.get(sid, {})
        # 底表行匹配
        for row in rows:
            g = (str(row[hi['通用名称']].value or '') + str(row[hi['车型名称']].value or '')).lower().replace(' ', '')
            pt = str(row[hi['动力类型']].value or '')
            code = str(row[hi['产品型号']].value or '').strip()
            if kw not in g and kw2 not in g:
                continue
            if not ('纯电' in pt or '插电' in pt or '增程' in pt or 'BEV' in pt or 'PHEV' in pt):
                continue
            is_bev = '纯电' in pt or 'BEV' in pt
            # dcd 款型筛选：BEV 行需有电池；排除汽油 ff
            trims = []
            for c in dcd['cars']:
                v = c['vals']
                ff = fuel_ff.get((sid, str(c.get('id'))), '')
                if ff == '汽油':
                    continue
                if is_bev and not v.get('battery_capacity'):
                    continue
                if (not is_bev) and not v.get('battery_capacity') and 'DM' not in c['name'] and 'EV' not in c['name']:
                    continue
                trims.append(c)
            if not trims:
                continue
            # 打分：续航接近（名称 km 或 CLTC）+ 整备接近
            r_rng, r_cw = fnum(row[hi['纯电续航里程(km)']].value), fnum(row[hi['整备质量(kg)']].value)
            scored = []
            for c in trims:
                v = c['vals']
                nm_km = re.search(r'(\d{3,4})\s*km', c['name'])
                cltc = fnum(v.get('cltc_recharge_mileage')) or (float(nm_km.group(1)) if nm_km else None)
                cw = fnum(v.get('curb_weight'))
                score = 0
                if r_rng and cltc:
                    d = abs(r_rng - cltc)
                    score += (0 if d <= 15 else (1 if d <= 30 else 3))
                elif r_rng and not cltc:
                    score += 1
                if r_cw and cw:
                    d = abs(r_cw - cw)
                    score += (0 if d <= 40 else (1 if d <= 90 else 3))
                scored.append((score, c))
            scored.sort(key=lambda t: t[0])
            best_score = scored[0][0]
            tied = [c for s, c in scored if s == best_score]

            # 逐字段：空格补全（唯一值），或 T4 双源红修
            for fld, key in FIELD_MAP.items():
                cell = row[hi[fld]]
                cur = cell.value
                if cur not in (None, ''):
                    key4 = (str(row[hi['批次']].value or '').strip(), code, fld.split('(')[0])
                    key_est = (str(row[hi['批次']].value or '').strip(), code, fld)
                    cand_vals = {fnum(c['vals'].get(key)) for c in tied}
                    cand_vals.discard(None)
                    if key4 in t4 and fnum(cur) is not None and fnum(cur) != fnum(t4[key4]) and cand_vals == {fnum(t4[key4])}:
                        old = cur
                        cell.value = fnum(t4[key4])
                        cell.fill = RED
                        seq += 1
                        ledger.append((seq, row[hi['批次']].value, code, row[hi['动力类型']].value or '',
                                       fld, '双源更正', str(old), t4[key4],
                                       f'懂车帝参数页(sid={sid}) 与 T4补丁 双源一致：{DCD_URL.format(sid=sid)}'))
                        stats['red_fix'] += 1
                    elif key_est in est_keys and len(cand_vals) == 1 and fnum(cur) is not None and fnum(cur) != next(iter(cand_vals)):
                        old = cur
                        newv = next(iter(cand_vals))
                        cell.value = newv
                        cell.fill = RED
                        seq += 1
                        ledger.append((seq, row[hi['批次']].value, code, row[hi['动力类型']].value or '',
                                       fld, '双源更正', str(old), newv,
                                       f'历史估算值被懂车帝参数页(sid={sid})同款型唯一实测值证伪更正：{DCD_URL.format(sid=sid)}'))
                        stats['red_fix_est'] += 1
                else:
                    vals = set()
                    for c in tied:
                        v = c['vals'].get(key)
                        if key == 'battery_type':
                            vals.add(BT_NORM.get(str(v).strip(), str(v).strip()))
                        elif key == 'total_electric_power':
                            pv = parse_power(v)
                            if pv is not None:
                                vals.add(pv)
                        else:
                            fv = fnum(v)
                            vals.add(fv if fv is not None else str(v).strip())
                    vals.discard('')
                    if len(vals) == 1:
                        val = vals.pop()
                        cell.value = val
                        cell.fill = GREEN
                        seq += 1
                        ledger.append((seq, row[hi['批次']].value, code, row[hi['动力类型']].value or '',
                                       fld, '媒体参数补全', '(空)', str(val),
                                       f'懂车帝参数页(sid={sid})：{DCD_URL.format(sid=sid)}'))
                        stats['green_fill'] += 1
                    else:
                        stats['ambiguous_skip'] += 1
            report.append((sid, m['kw'], code, best_score, len(tied)))

    print(f"green_fill={stats['green_fill']} red_fix={stats['red_fix']} ambiguous_skip={stats['ambiguous_skip']} ledger_seq->{seq}")
    if not apply:
        print('DRY-RUN（未写入）')
        json.dump({'stats': dict(stats)}, open(os.path.join(SRC, '_integrate_report.json'), 'w', encoding='utf-8'), ensure_ascii=False)
        return
    wb.save(WB)
    print('saved')
    json.dump({'stats': dict(stats)}, open(os.path.join(SRC, '_integrate_report.json'), 'w', encoding='utf-8'), ensure_ascii=False)


if __name__ == '__main__':
    main()
