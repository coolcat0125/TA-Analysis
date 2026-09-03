# -*- coding: utf-8 -*-
"""
底表信息补全管线 v5
===================
对《NEV公告参数汇总表_合并版（341~410批）.xlsx》做五阶段补全（只补空、不覆盖已有值）：

  A. 官方主链交叉补全 —— 从 data/nev-announcements/db/master_export_fixed.csv（官方公告主链 341~409）
     按 (批次, 产品型号) 匹配，用主链已有非空值补底表空位（绿色 C6EFCE）。
  B. 同产品型号共识补全 —— (产品型号, 动力类型) 维度唯一非空共识值（黄色 FFF2CC，沿用 v3.6 约定）。
  C. 企业+车型共识补全 —— (企业名称, 车型名称, 动力类型) 维度唯一非空共识值（绿色 C6EFCE）。
  D. 百公里电耗公式补全 —— 电耗 = 电池容量×100÷纯电续航，物理守卫 5~40 kWh/100km（蓝色 BDD7EE）。
  E. 车长估算补全 —— 车长 = 轴距 × 车身形式系数（细分市场映射），四舍五入到 10mm（浅紫 D9D2E9）。

原则：同 v3.8（100% 唯一共识才补、不造值、不覆盖、动力类型进维度防混填），
且只处理"空"单元格（None / 空串 / '/' / '—' / '未填报' / '无'）。
逐格记入「变更记录」工作表。

用法：
  python enrich_v5.py <底表.xlsx> <master_export_fixed.csv>
Copyright © 2026 David YE
"""
import sys
import os
import re
import csv
import math
import openpyxl
from collections import defaultdict

GREEN = openpyxl.styles.PatternFill('solid', fgColor='C6EFCE')   # 补充/共识补全
YELLOW = openpyxl.styles.PatternFill('solid', fgColor='FFF2CC')  # 同车型(同型号)共识
BLUE = openpyxl.styles.PatternFill('solid', fgColor='BDD7EE')    # 公式补全
PURPLE = openpyxl.styles.PatternFill('solid', fgColor='D9D2E9')  # 轴距(车长)估算

NUMERIC_FIELDS = {'整备质量(kg)', '轴距(mm)', '纯电续航里程(km)', '电池容量(kWh)',
                  '电池能量密度(Wh/kg)', '百公里电耗(kWh/100km)', '电机峰值功率(kW)',
                  '电机总功率(kW)', '综合油耗(L/100km)', 'B状态油耗(L/100km)',
                  '发动机排量(mL)', '发动机功率(kW)', '车长(mm)'}
TEXT_FIELDS = {'产品商标', '企业名称', '车型名称', '产品名称', '通用名称', '产品类型',
               '细分市场', '是否减免购置税', '电池类型', '电机生产企业', '电机功率/扭矩',
               '电机型号', '发动机生产企业', '发动机型号'}
BAD_TEXT = ('请提供', '无法', '无有效', '未找到', '无符合', 'unknown', 'n/a', '需人工',
            '暂缺', '待确认', '待核实', '略', '无信息', '未填写', '不祥')
MULTI_NUM_RE = re.compile(r'^[\d.]+(\s*/\s*[\d.]+)*$')
DIRTY_WORDS = ('(', ')', '？', '?', '，', ',', '、', '；', ';', '×', '*')

# 车身形式系数（v3.7.3 口径）
BODY_FACTOR = [('Lux Car', 1.66), ('Lux SUV', 1.66), ('Lux MPV', 1.65), ('Car', 1.72),
               ('SUV', 1.69), ('MPV', 1.68), ('Sports', 1.64)]
BODY_FACTOR_OTHER = 1.70


EMPTY_PLACEHOLDER = {'', '/', '-', '—', '0', '未填报', '无', '无信息', '待查', '待核实',
                     '待确认', '未知', '？', '?', 'N/A', 'NA', 'None', '未提供'}
SEMANTIC_PLACEHOLDER = {'不适用', '不适用(BEV)', '不适用(PHEV)', '不适用(HEV)', '不适用(EREV)'}


def norm(v):
    """共识候选值规范化：占位/语义占位一律不作为有效候选值。"""
    if v is None:
        return None
    s = str(v).strip()
    if s.replace('\u3000', '').strip() == '' or s in EMPTY_PLACEHOLDER or s in SEMANTIC_PLACEHOLDER:
        return None
    return s


def fillable(v):
    """是否视为'空'而可以补：真空白或原文空白占位；'不适用(BEV)'等语义占位不可覆盖。"""
    if v is None:
        return True
    s = str(v).strip()
    if s.replace('\u3000', '').strip() == '':
        return True
    if s in EMPTY_PLACEHOLDER:
        return True
    if s in SEMANTIC_PLACEHOLDER:
        return False
    return False


def cast_num(v):
    s = str(v).strip()
    for m in re.finditer(r'[\d.]+', s):
        try:
            f = float(m.group())
            return int(f) if f.is_integer() else f
        except ValueError:
            continue
    return None


def is_clean_numeric(v):
    s = str(v).strip()
    if not MULTI_NUM_RE.match(s):
        return False
    return all(float(x) >= 0 for x in re.split(r'\s*/\s*', s))


def is_clean_text(v):
    s = str(v).strip()
    if not s or len(s) > 60:
        return False
    if any(b in s.lower() for b in BAD_TEXT):
        return False
    if any(d in s for d in DIRTY_WORDS):
        return False
    return True


def body_factor(seg):
    if not seg:
        return BODY_FACTOR_OTHER
    for prefix, f in BODY_FACTOR:
        if str(seg).startswith(prefix):
            return f
    return BODY_FACTOR_OTHER


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return
    src_path, master_path = sys.argv[1], sys.argv[2]
    print(f'[1/6] 加载底表: {os.path.basename(src_path)}')
    wb = openpyxl.load_workbook(src_path)
    ws = wb['NEV公告参数汇总']
    hdr = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(hdr)}
    fi = {f: idx[f] for f in hdr}
    cr = wb['变更记录']
    last_seq = 0
    for r in range(cr.max_row, 1, -1):
        v = cr.cell(r, 1).value
        if isinstance(v, int):
            last_seq = v
            break

    # ---- 读取全部数据行 ----
    rows = []
    for r in range(2, ws.max_row + 1):
        rec = {h: ws.cell(r, fi[h] + 1).value for h in hdr}
        rows.append((r, rec))

    def is_empty(v):
        return fillable(v)

    def get(r, f):
        return norm(r[f])

    # ---- 阶段A: 官方主链交叉补全 ----
    master = defaultdict(list)
    with open(master_path, encoding='utf-8-sig', errors='replace') as f:
        mrows = list(csv.reader(f))
    mhdr = mrows[0]
    midx = {h: i for i, h in enumerate(mhdr)}
    for mr in mrows[1:]:
        if len(mr) != len(mhdr):
            continue
        key = (str(mr[midx['批次']]).strip(), str(mr[midx['产品型号']]).strip())
        if key[0] and key[1]:
            master[key].append(mr)
    print(f'[2/6] 阶段A 官方主链交叉补全（master {len(master)} 键）')
    stats = defaultdict(int)
    a_ledger = []
    for r, rec in rows:
        key = (str(rec['批次']).strip(), str(rec['产品型号']).strip())
        if key[0] and key[1] and key in master:
            mrec = master[key][0]
            for f in hdr:
                if f in ('批次', '产品型号', '数据来源'):
                    continue
                if not is_empty(rec[f]):
                    continue
                mv = mrec[midx[f]] if f in midx and midx[f] < len(mrec) else None
                mv_n = norm(mv)
                if mv_n is None:
                    continue
                if f in NUMERIC_FIELDS and not is_clean_numeric(mv_n):
                    continue
                if f in TEXT_FIELDS and not is_clean_text(mv_n):
                    continue
                cell = ws.cell(r, fi[f] + 1)
                cell.value = cast_num(mv_n) if f in NUMERIC_FIELDS else mv_n
                cell.fill = GREEN
                stats[f] += 1
                a_ledger.append((r, rec, f, mv_n, '官方主链交叉补全'))
    print(f'      阶段A 补 {sum(stats.values())} 格')

    # ---- 阶段B: 同产品型号共识 ----
    print('[3/6] 阶段B 同产品型号共识')
    groups = defaultdict(set)
    ent_count = defaultdict(int)
    for r, rec in rows:
        key = (get(rec, '产品型号'), get(rec, '动力类型'))
        if None in key:
            continue
        ent_count[key] += 1
        for f in hdr:
            v = get(rec, f)
            if v is not None:
                groups[(key, f)].add(v)
    b_ledger = []
    for r, rec in rows:
        key = (get(rec, '产品型号'), get(rec, '动力类型'))
        if None in key or ent_count[key] < 2:
            continue
        for f in hdr:
            if f in ('批次', '产品型号', '数据来源'):
                continue
            if not is_empty(rec[f]):
                continue
            s = groups.get((key, f))
            if s and len(s) == 1:
                val = next(iter(s))
                if f in NUMERIC_FIELDS and not is_clean_numeric(val):
                    continue
                cell = ws.cell(r, fi[f] + 1)
                cell.value = cast_num(val) if f in NUMERIC_FIELDS else val
                cell.fill = YELLOW
                stats[f] += 1
                b_ledger.append((r, rec, f, val, '同产品型号唯一值共识'))
    print(f'      阶段B 补 {len(b_ledger)} 格')

    # ---- 阶段C: 企业+车型共识 ----
    print('[4/6] 阶段C 企业+车型共识')
    cgroups = defaultdict(set)
    cent = defaultdict(int)
    for r, rec in rows:
        key = (get(rec, '企业名称'), get(rec, '车型名称'), get(rec, '动力类型'))
        if None in key:
            continue
        cent[key] += 1
        for f in hdr:
            v = get(rec, f)
            if v is not None:
                cgroups[(key, f)].add(v)
    c_ledger = []
    for r, rec in rows:
        key = (get(rec, '企业名称'), get(rec, '车型名称'), get(rec, '动力类型'))
        if None in key or cent[key] < 2:
            continue
        for f in ('细分市场', '产品类型', '通用名称', '产品名称', '产品商标',
                  '是否减免购置税') + tuple(sorted(NUMERIC_FIELDS)) + \
                ('电池类型', '电机生产企业', '电机功率/扭矩', '电机型号',
                 '发动机生产企业', '发动机型号'):
            if f in ('是否减免购置税', '产品商标'):
                continue  # 减免状态/商标独立于车型共识，保守跳过
            if not is_empty(rec[f]):
                continue
            s = cgroups.get((key, f))
            if s and len(s) == 1:
                val = next(iter(s))
                if f in NUMERIC_FIELDS and not is_clean_numeric(val):
                    continue
                cell = ws.cell(r, fi[f] + 1)
                cell.value = cast_num(val) if f in NUMERIC_FIELDS else val
                cell.fill = GREEN
                stats[f] += 1
                c_ledger.append((r, rec, f, val, '企业+车型唯一共识值'))
    print(f'      阶段C 补 {len(c_ledger)} 格')

    # ---- 阶段D: 电耗公式 ----
    print('[5/6] 阶段D 电耗公式补全')
    d_ledger = []
    for r, rec in rows:
        if not is_empty(rec.get('百公里电耗(kWh/100km)')):
            continue
        cap = cast_num(rec.get('电池容量(kWh)'))
        rng = cast_num(rec.get('纯电续航里程(km)'))
        if cap is None or rng is None or rng <= 0 or cap <= 0:
            continue
        v = round(cap * 100.0 / rng, 1)
        if not (5.0 <= v <= 40.0):
            continue
        cell = ws.cell(r, fi['百公里电耗(kWh/100km)'] + 1)
        cell.value = v
        cell.fill = BLUE
        stats['百公里电耗(kWh/100km)'] += 1
        d_ledger.append((r, rec, '百公里电耗(kWh/100km)', v, '容量×100÷续航公式'))
    print(f'      阶段D 补 {len(d_ledger)} 格')

    # ---- 阶段E: 车长估算 ----
    print('[6/6] 阶段E 车长估算')
    e_ledger = []
    for r, rec in rows:
        if not is_empty(rec.get('车长(mm)')):
            continue
        ab = cast_num(rec.get('轴距(mm)'))
        if ab is None or ab <= 0:
            continue
        seg = get(rec, '细分市场') or get(rec, '产品类型')
        f = body_factor(seg)
        v = round(ab * f / 10.0) * 10
        if not (1800 <= v <= 6500):
            continue
        cell = ws.cell(r, fi['车长(mm)'] + 1)
        cell.value = v
        cell.fill = PURPLE
        stats['车长(mm)'] += 1
        e_ledger.append((r, rec, '车长(mm)', v, '轴距×车身形式系数估算'))
    print(f'      阶段E 补 {len(e_ledger)} 格')

    # ---- 写变更记录 ----
    print(f'      追加变更记录 {len(a_ledger)+len(b_ledger)+len(c_ledger)+len(d_ledger)+len(e_ledger)} 条')
    for ledger, ctype, color_src in (
            (a_ledger, '官方主链交叉补全', None),
            (b_ledger, '同车型共识', None),
            (c_ledger, '共识补全', None),
            (d_ledger, '公式补全', None),
            (e_ledger, '轴距估算', None)):
        for r, rec, f, val, note in ledger:
            last_seq += 1
            cr.append([last_seq, rec['批次'], rec['产品型号'],
                       rec['产品类型'], f, ctype, '(空)', val, note])
    wb.save(src_path)
    print('\n===== 汇总 =====')
    for f, n in sorted(stats.items(), key=lambda x: -x[1]):
        if n:
            print(f'  {f}: +{n}')
    print(f'  合计补全 {sum(stats.values())} 格；变更记录末序号 {last_seq}')


if __name__ == '__main__':
    main()
