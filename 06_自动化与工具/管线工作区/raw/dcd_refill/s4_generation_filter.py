# -*- coding: utf-8 -*-
"""s4_generation_filter.py — Wave 3 / S4：多配置歧义格的代际过滤消解

背景：dcd_batch_run.py apply 以 续航±15/整备±40 打分取最优(tied)，字段值在 tied 内
唯一才补——跨年代多配置款型同分即产生 ambiguous 留档（~1,000 格）。

S4 原理（代际铁律，fix_miniev_gen.py 先例）：DCD 参数页载全系历史款型，历史公告行
禁止套用未来款型。行批次日期 → 圈定 model-year：eligible = year ≤ 批次年+1（公告
先于上市，年内晚批次常挂下一年款），取其中最近年代，缩 tied → tied2 后再验唯一。

铁律继承：只补空 / BEV禁油 / 汽油排除 / W5 车长轴距成对 / 逐格台账 / 底纹绿(C6EFCE)
用法：
  python s4_generation_filter.py --scan            # 量化：ambiguous 池与 S4 可收编数
  python s4_generation_filter.py --scan --verbose  # 附样例明细
  python s4_generation_filter.py --apply           # 落库（先 dry 复核 scan 再 apply）
  python s4_generation_filter.py --apply --dry     # 落库预演不写盘
"""
import argparse
import json
import os
import re
import sys
import time
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
WB = os.path.join(ROOT, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
REFILL = HERE
STATE = os.path.join(REFILL, 'batch_state.json')
FUEL = os.path.join(REFILL, 'dcd_fuel_form.json')

FIELDS = ['车长(mm)', '轴距(mm)', '整备质量(kg)', '纯电续航里程(km)', '百公里电耗(kWh/100km)',
          '电池容量(kWh)', '电池类型', '电池能量密度(Wh/kg)', '电机总功率/扭矩',
          '前电机功率/扭矩', '后电机功率/扭矩']
FIELD_MAP = {
    '车长(mm)': 'length', '轴距(mm)': 'wheelbase', '整备质量(kg)': 'curb_weight',
    '纯电续航里程(km)': 'cltc_recharge_mileage', '百公里电耗(kWh/100km)': 'power_consumption',
    '电池容量(kWh)': 'battery_capacity', '电池类型': 'battery_type',
    '电池能量密度(Wh/kg)': 'battery_energy_density', '电机总功率/扭矩': 'total_electric_power',
    '前电机功率/扭矩': 'front_electric_max_power', '后电机功率/扭矩': 'rear_electric_max_power',
}
BT_NORM = {'三元锂电池': '三元锂', '磷酸铁锂电池': '磷酸铁锂', '锰酸锂电池': '锰酸锂', '钛酸锂电': '钛酸锂'}
GREEN = PatternFill('solid', fgColor='C6EFCE')
TODAY = time.strftime('%Y-%m-%d')


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', '').replace('km', '').replace('kW', '').strip())
    except Exception:
        return None


def load_batch_years():
    """批次时间表 → {批次:int → 公告年份:int, 公告月份:int}"""
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['批次时间表']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    bi = hdr.index('批次'); di = hdr.index('公告日期')
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        b, d = r[bi], r[di]
        if b is None or d is None:
            continue
        m = re.match(r'^(\d{4})-(\d{2})', str(d).strip())
        if m:
            try:
                out[int(float(str(b)))] = (int(m.group(1)), int(m.group(2)))
            except Exception:
                pass
    return out


def year_of(c):
    """款型年代：DCD year 字段为 str('2026')/int，统一转 int；无效返回 None"""
    y = c.get('year')
    try:
        return int(str(y).strip()[:4])
    except Exception:
        return None


def field_values(trims, key):
    """apply 同款取值：battery_type 规范化；电机功率列收裸数字或 kW 后缀；其余 fnum"""
    vals = set()
    for c in trims:
        v = c['vals'].get(key)
        if key == 'battery_type':
            s = str(v).strip()
            if s:
                vals.add(BT_NORM.get(s, s))
        elif key in ('total_electric_power', 'front_electric_max_power', 'rear_electric_max_power'):
            m = re.match(r'^\s*(\d+(?:\.\d+)?)(?:\s*kW)?', str(v), re.I)
            if m:
                vals.add(float(m.group(1)))
        else:
            fv = fnum(v)
            vals.add(fv if fv is not None else str(v).strip())
    vals.discard('')
    vals.discard(None)
    return vals


def main():
    ap = argparse.ArgumentParser(description='S4 代际过滤消解 ambiguous')
    ap.add_argument('--scan', action='store_true')
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--dry', action='store_true', help='apply 预演不写盘')
    ap.add_argument('--verbose', action='store_true')
    args = ap.parse_args()
    if not args.scan and not args.apply:
        ap.print_help(); sys.exit(1)

    batch_years = load_batch_years()
    fuel_ff = {}
    try:
        fuel = json.load(open(FUEL, encoding='utf-8'))
        for sid_f, lst in fuel.items():
            for sp in (lst or []):
                if isinstance(sp, dict) and sp.get('id'):
                    fuel_ff[(str(sid_f), str(sp['id']))] = str(sp.get('ff') or '')
    except Exception:
        pass

    # 目标车系 = batch_state 全部 (kw, sid) 对（与历次 apply 完全同范围）
    st = json.load(open(STATE, encoding='utf-8'))
    targets = {}
    for bid, b in sorted(st['batches'].items()):
        for s in b.get('series', []):
            if isinstance(s, dict) and s.get('sid'):
                kw = str(s['kw']).lower().replace(' ', '')
                if kw and kw not in targets:
                    targets[kw] = str(s['sid'])

    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in ws[1]]
    hi = {h: i for i, h in enumerate(hdr)}
    rows = list(ws.iter_rows(min_row=2))

    # rear_only 清单（gap_reclass）用于专项覆盖率统计
    rear_only_keys = set()
    try:
        gr = json.load(open(os.path.join(REFILL, 'gap_reclass.json'), encoding='utf-8'))
        for g in gr.get('rows', []):
            if g.get('verdict') == 'rear_only':
                rear_only_keys.add((str(g.get('batch')), str(g.get('model')).strip()))
    except Exception:
        pass

    stats = defaultdict(int)
    per_field = defaultdict(lambda: defaultdict(int))
    candidates = []   # (row, field, sid, kw, val, tied2_sizes, batch)
    verbose_samples = []
    rear_stat = defaultdict(int)

    for kw, sid in targets.items():
        f = os.path.join(REFILL, f'dcd_s{sid}.json')
        if not os.path.exists(f):
            stats['no_data_file'] += 1
            continue
        dcd = json.load(open(f, encoding='utf-8'))
        all_cars = dcd.get('cars', [])

        for row in rows:
            g = (str(row[hi['通用名称']].value or '') + str(row[hi['车型名称']].value or '')).lower().replace(' ', '')
            if kw not in g:
                continue
            pt = str(row[hi['动力类型']].value or '')
            if not ('纯电' in pt or '插电' in pt or '增程' in pt or 'BEV' in pt or 'PHEV' in pt):
                continue
            is_bev = '纯电' in pt or 'BEV' in pt
            code = str(row[hi['产品型号']].value or '').strip()
            is_rear_only = (str(row[hi['批次']].value or '').strip(), code) in rear_only_keys
            if is_rear_only:
                rear_stat['rows_seen'] += 1

            try:
                bno = int(float(str(row[hi['批次']].value or '').strip()))
            except Exception:
                bno = None
            binfo = batch_years.get(bno)
            byear = binfo[0] if binfo else None
            # 代际窗口（保守）：公告月份 ≥7 允许圈定到下一年款（Y+1，秋发新车公告常挂
            # 下一年款）；1~6 月批次只允许当年款——343批(2021-04)圈到2022款值即过度。
            cap_year = (byear + (1 if binfo and binfo[1] >= 7 else 0)) if binfo else None
            if byear is None and is_rear_only:
                rear_stat['no_batch_year'] += 1

            # 款型筛选（apply 同款）
            trims = []
            for c in all_cars:
                v = c.get('vals', {})
                ff = fuel_ff.get((sid, str(c.get('id'))), '')
                if ff == '汽油':
                    continue
                if is_bev and not v.get('battery_capacity'):
                    continue
                if (not is_bev) and not v.get('battery_capacity') \
                        and 'DM' not in str(c.get('name')) and 'EV' not in str(c.get('name')):
                    continue
                trims.append(c)
            if not trims:
                continue
            # 打分：续航±15 / 整备±40
            r_rng = fnum(row[hi['纯电续航里程(km)']].value)
            r_cw = fnum(row[hi['整备质量(kg)']].value)
            scored = []
            for c in trims:
                v = c['vals']
                nm_km = re.search(r'(\d{3,4})\s*km', str(c.get('name')))
                cltc = fnum(v.get('cltc_recharge_mileage')) or (float(nm_km.group(1)) if nm_km else None)
                cw = fnum(v.get('curb_weight'))
                sc = 0
                if r_rng and cltc:
                    dd = abs(r_rng - cltc)
                    sc += 0 if dd <= 15 else (1 if dd <= 30 else 3)
                if r_cw and cw:
                    dd = abs(r_cw - cw)
                    sc += 0 if dd <= 40 else (1 if dd <= 90 else 3)
                scored.append((sc, c))
            scored.sort(key=lambda t: t[0])
            best = scored[0][0]
            tied = [c for sc, c in scored if sc == best]

            # ---- S4 代际过滤：year ≤ 批次年+1 取最近年代 ----
            tied2 = tied
            s4_applied = False
            ty = None
            if cap_year is not None and len(tied) > 1:
                elig = [c for c in tied
                        if year_of(c) is not None and year_of(c) <= cap_year]
                if elig:
                    ty = max(year_of(c) for c in elig)
                    tied2 = [c for c in elig if year_of(c) == ty]
                    s4_applied = len(tied2) < len(tied)

            for fld in FIELDS:
                cell = row[hi[fld]]
                cur = cell.value
                if cur is not None and str(cur).strip() != '':
                    continue  # 只补空：已填不问
                vals_before = field_values(tied, FIELD_MAP[fld])
                vals_after = field_values(tied2, FIELD_MAP[fld])
                if len(vals_before) <= 1:
                    stats['no_data' if not vals_before else 'unique_unfilled'] += 1
                    if not vals_before:
                        per_field[fld]['no_data'] += 1
                        if is_rear_only:
                            rear_stat['no_data'] += 1
                    continue
                # 真歧义格（ambiguous_before）
                per_field[fld]['ambiguous_before'] += 1
                stats['ambiguous_before'] += 1
                if is_rear_only:
                    rear_stat['amb_before'] += 1
                pair_ok = True
                if fld in ('车长(mm)', '轴距(mm)'):
                    pair_ok = len(field_values(tied2, 'length') - {None}) == 1 and \
                              len(field_values(tied2, 'wheelbase') - {None}) == 1
                if s4_applied and len(vals_after) == 1 and pair_ok:
                    val = next(iter(vals_after))
                    candidates.append({'row_idx': row[0].row, 'field': fld, 'sid': sid, 'kw': kw,
                                       'val': val, 'batch': bno, 'model': code, 'pt': pt,
                                       'n_tied': len(tied), 'n_tied2': len(tied2)})
                    per_field[fld]['s4_resolved'] += 1
                    stats['s4_resolved'] += 1
                    if is_rear_only:
                        rear_stat['s4_resolved'] += 1
                    if args.verbose and len(verbose_samples) < 25:
                        verbose_samples.append(
                            f"  {bno}批 {code} [{kw}] {fld}: {sorted(vals_before)} --year{ty}--> {val}")
                else:
                    per_field[fld]['still_ambiguous'] += 1
                    stats['still_ambiguous'] += 1
                    if is_rear_only:
                        rear_stat['still_amb'] += 1

    # ---------------- 汇报 ----------------
    print(f"== S4 scan {TODAY} ==")
    print(f"目标 kw-sid 对: {len(targets)}  ambiguous_before: {stats['ambiguous_before']}  "
          f"S4 可收编: {stats['s4_resolved']}  仍歧义: {stats['still_ambiguous']}")
    print(f"{'字段':<22}{'歧义':>6}{'S4收编':>8}{'仍歧义':>8}{'无数据':>8}")
    for fld in FIELDS:
        d = per_field[fld]
        print(f"{fld:<22}{d['ambiguous_before']:>6}{d['s4_resolved']:>8}{d['still_ambiguous']:>8}{d['no_data']:>8}")
    print(f"rear_only 专项: 行命中 {rear_stat['rows_seen']} | 歧义 {rear_stat['amb_before']} → "
          f"S4收编 {rear_stat['s4_resolved']} / 仍歧义 {rear_stat['still_amb']} / 无数据 {rear_stat['no_data']}")
    if args.verbose and verbose_samples:
        print("样例:")
        print('\n'.join(verbose_samples))

    if not args.apply:
        print('DRY-RUN（未写入）')
        return

    # ---------------- 落库 ----------------
    ledger = wb['变更记录']
    seq = max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
               if isinstance(r[0], int)), default=0)
    written = 0
    seen_cells = set()
    for cd in candidates:
        key = (cd['row_idx'], cd['field'])
        if key in seen_cells:
            continue
        seen_cells.add(key)
        row = ws[cd['row_idx']]
        cell = row[hi[cd['field']]]
        if cell.value is not None and str(cell.value).strip() != '':
            continue  # 双保险：只补空
        cell.value = cd['val']
        cell.fill = GREEN
        seq += 1
        ledger.append((seq, cd['batch'], cd['model'], cd['pt'], cd['field'], '媒体参数补全', '(空)',
                       str(cd['val']),
                       f"懂车帝参数页(sid={cd['sid']})：https://www.dongchedi.com/auto/series/{cd['sid']}"
                       f"（懂车帝参数页，S4代际过滤 检索 {TODAY}）"))
        written += 1
    if args.dry:
        print(f"DRY 预演：应写 {written} 格（未写盘）")
        return
    wb.save(WB)
    print(f"已落库 {written} 格（台账 seq {seq - written + 1}~{seq}）——记得三件套+两版看板+提交注明 DCD/S4 通道")


if __name__ == '__main__':
    main()
