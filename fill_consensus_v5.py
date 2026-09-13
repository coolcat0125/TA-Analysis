# -*- coding: utf-8 -*-
"""
共识补全管线 v5.1（v4.1.0）—— 在 enrich_v5.py 四级共识之下新增更宽维度
====================================================================
前置：enrich_v5.py 已执行过 官方主链 > 同型号(产品型号,动力类型) > 企业+车型 > 公式 > 车长 五阶段。
本脚本只处理剩余空格，新增两个置信度更低的通道，全部【唯一非空共识】才补、只补空、不覆盖、逐格台账：

  F. 通用名称+动力类型 唯一共识（浅黄绿 E2EFDA，变更类型=通用名称共识）
     依据：同一通用名称（车系）跨年款/跨配置共享整车平台与三电系统，参数为平台级属性。
     这是「同型号、企业+车型」之后合理的外推：车系身份由通用名称唯一确定时，其配置唯一值可信。
  G. 企业名称+通用名称+动力类型 唯一共识（浅橙 FCE4D6，变更类型=企业车系共识）
     依据：同一企业同一车系的不同配置（续航/功率版本）共用电池/电机/发动机等零部件。
  H. 百公里电耗公式补全（蓝色 BDD7EE，变更类型=公式补全）—— F/G 补出的容量/续航产生新算点。
  I. 车长估算（浅紫 D9D2E9，变更类型=轴距估算）—— F/G 补出的轴距产生新估算点。

排除项（实事求是，不做共识外推）：
  是否减免购置税（随批次目录政策变化）、产品商标、数据来源、批次、产品型号、
  电芯供应商/电池包供应商/官方参数页URL（仅409批官方数据，其余批次无源）。
数值字段额外做物理范围守卫（越界共识值整组不补）。

用法：python fill_consensus_v5.py
Copyright © 2026 David YE
"""
import os
import re
import sys
import time
import openpyxl
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SHEET = 'NEV公告参数汇总'

YELLOWGREEN = openpyxl.styles.PatternFill('solid', fgColor='E2EFDA')  # F 通用名称共识
LIGHTORANGE = openpyxl.styles.PatternFill('solid', fgColor='FCE4D6')  # G 企业车系共识
BLUE = openpyxl.styles.PatternFill('solid', fgColor='BDD7EE')         # H 公式补全
PURPLE = openpyxl.styles.PatternFill('solid', fgColor='D9D2E9')       # I 轴距估算

# 数值字段物理守卫（与看板 generate_dashboard.py NUM_RANGES 同口径，动力类型无关字段用统一区间）
NUM_RANGES = {
    '整备质量(kg)': (300, 4500), '轴距(mm)': (1800, 4200), '纯电续航里程(km)': (20, 1500),
    '电池容量(kWh)': (1, 250), '电池能量密度(Wh/kg)': (30, 400), '百公里电耗(kWh/100km)': (3, 40),
    '综合油耗(L/100km)': (0.1, 20),
    'B状态油耗(L/100km)': (0.1, 20), '发动机排量(mL)': (300, 6000), '发动机功率(kW)': (10, 600),
}
# 参与共识的文本字段（排除批次/型号/来源/减免/商标/供应商/URL）
# v4.4.0：电机功率三列为 "P/T" 文本（前/后/总），归入 TEXT_FILL 唯一共识，不做数值回写
TEXT_FILL = ['企业名称', '车型名称', '产品名称', '通用名称', '产品类型', '细分市场',
             '电池类型', '电机生产企业', '电机型号', '发动机生产企业', '发动机型号',
             '前电机功率/扭矩', '电机总功率/扭矩', '后电机功率/扭矩']
NUM_FILL = ['整备质量(kg)', '轴距(mm)', '纯电续航里程(km)', '电池容量(kWh)', '电池能量密度(Wh/kg)',
            '综合油耗(L/100km)', 'B状态油耗(L/100km)',
            '发动机排量(mL)', '发动机功率(kW)']

EMPTY_PLACEHOLDER = {'', '/', '-', '—', '0', '未填报', '无', '无信息', '待查', '待核实',
                     '待确认', '未知', '？', '?', 'N/A', 'NA', 'None', '未提供'}
SEMANTIC_PLACEHOLDER = {'不适用', '不适用(BEV)', '不适用(PHEV)', '不适用(HEV)', '不适用(EREV)'}
BAD_TEXT = ('请提供', '无法', '无有效', '未找到', '无符合', 'unknown', 'n/a', '需人工',
            '暂缺', '待确认', '待核实', '略', '无信息', '未填写', '不祥')
DIRTY_WORDS = ('(', ')', '？', '?', '，', ',', '、', '；', ';', '×', '*')
MULTI_NUM_RE = re.compile(r'^[\d.]+(\s*/\s*[\d.]+)*$')


def norm(v):
    if v is None:
        return None
    s = str(v).replace('\u3000', '').strip()
    if s == '' or s in EMPTY_PLACEHOLDER or s in SEMANTIC_PLACEHOLDER:
        return None
    return s


def fillable(v):
    if v is None:
        return True
    s = str(v).replace('\u3000', '').strip()
    if s == '' or s in EMPTY_PLACEHOLDER:
        return True
    return False   # 语义占位(不适用(BEV)等)不可补


def cast_num(s):
    m = re.search(r'[\d.]+', str(s))
    if not m:
        return None
    f = float(m.group())
    return int(f) if f.is_integer() else f


def clean_num(s):
    s = str(s).strip()
    if not MULTI_NUM_RE.match(s):
        return None
    return s


def clean_txt(s):
    s = str(s).strip()
    if not s or len(s) > 60:
        return None
    if any(b in s.lower() for b in BAD_TEXT) or any(d in s for d in DIRTY_WORDS):
        return None
    return s


def in_range(field, s):
    """共识数值候选值须落在物理范围内（取首个数值判断），整组一个越界即不补。"""
    f = cast_num(s)
    if f is None:
        return False
    lo, hi = NUM_RANGES[field]
    return lo <= f <= hi


def consensus_fill(ws, rows, fi, key_fields, fields, color, ctype, note_fn, stats, ledger):
    groups = defaultdict(lambda: defaultdict(set))
    gsize = defaultdict(int)
    for _, rec in rows:
        key = tuple(norm(rec.get(k)) for k in key_fields)
        if None in key:
            continue
        gsize[key] += 1
        for f in fields:
            v = norm(rec.get(f))
            if v is None:
                continue
            if f in NUM_FILL and (not clean_num(v) or not in_range(f, v)):
                continue   # 脏值/越界值不作候选
            if f in TEXT_FILL and not clean_txt(v):
                continue
            groups[key][f].add(v)
    for r, rec in rows:
        key = tuple(norm(rec.get(k)) for k in key_fields)
        if None in key or gsize[key] < 2:
            continue
        _ptg = str(rec.get('动力类型') or '')
        _bev_guard = ('BEV' in _ptg or '纯电' in _ptg)
        for f in fields:
            if not fillable(rec.get(f)):
                continue
            if _bev_guard and f in ('综合油耗(L/100km)', 'B状态油耗(L/100km)',
                                    '发动机型号', '发动机排量(mL)', '发动机功率(kW)', '发动机生产企业'):
                continue  # BEV 行禁填油耗/发动机字段（2026-09-13 混系感染教训）
            s = groups[key].get(f)
            if s and len(s) == 1:
                val = next(iter(s))
                cell = ws.cell(r, fi[f] + 1)
                cell.value = cast_num(val) if f in NUM_FILL else val
                cell.fill = color
                rec[f] = cell.value
                stats[f] += 1
                ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), f, ctype,
                               '(空)', val, note_fn(gsize[key])])


def main():
    print(f'[1/5] 加载底表: {os.path.basename(SRC)}')
    wb = openpyxl.load_workbook(SRC)
    ws = wb[SHEET]
    hdr = [c.value for c in ws[1]]
    fi = {h: i for i, h in enumerate(hdr)}
    n_rows = ws.max_row
    rows = []
    for r in range(2, n_rows + 1):
        rec = {h: ws.cell(r, fi[h] + 1).value for h in hdr}
        rows.append((r, rec))
    # 基线填充率
    base_fill = {h: sum(1 for _, rec in rows if norm(rec.get(h)) is not None) for h in hdr}

    cr = wb['变更记录']
    last_seq = 0
    for rr in range(cr.max_row, 1, -1):
        v = cr.cell(rr, 1).value
        if isinstance(v, int):
            last_seq = v
            break

    # 颜色说明追加图例（幂等）
    cs = wb['颜色说明']
    have = {str(row[1] or '') for row in cs.iter_rows(values_only=True) if len(row) > 1}
    for color_name, ctype, desc in (
            ('浅黄绿底纹', '通用名称共识',
             '同一通用名称（车系）+动力类型维度下存在唯一非空共识值时补空。推断依据：同一通用名称车型共享整车平台与三电系统，参数为平台级属性（非官方实测，详见变更记录）。'),
            ('浅橙底纹', '企业车系共识',
             '同一企业+通用名称+动力类型维度下唯一非空共识值补空。推断依据：同一企业同一车系不同配置版本共用电池/电机/发动机等零部件（非官方实测，详见变更记录）。')):
        if ctype not in have:
            cs.append([color_name, ctype, desc])

    stats = defaultdict(int)
    ledgers = []

    print('[2/5] 阶段F 通用名称+动力类型 唯一共识…')
    f_ledger = []
    consensus_fill(ws, rows, fi, ('通用名称', '动力类型'), NUM_FILL + TEXT_FILL,
                   YELLOWGREEN, '通用名称共识',
                   lambda n: f'推断：(通用名称,动力类型)唯一共识(组n={n})，同车系共享平台三电，非实测',
                   stats, f_ledger)
    print(f'      阶段F 补 {len(f_ledger)} 格')
    ledgers.append((f_ledger, '通用名称共识'))

    print('[3/5] 阶段G 企业+通用名称+动力类型 唯一共识…')
    g_ledger = []
    consensus_fill(ws, rows, fi, ('企业名称', '通用名称', '动力类型'), NUM_FILL + TEXT_FILL,
                   LIGHTORANGE, '企业车系共识',
                   lambda n: f'推断：(企业,通用名称,动力类型)唯一共识(组n={n})，同车系不同配置共用零部件，非实测',
                   stats, g_ledger)
    print(f'      阶段G 补 {len(g_ledger)} 格')
    ledgers.append((g_ledger, '企业车系共识'))

    print('[4/5] 阶段H 电耗公式补全（基于F/G新补的容量/续航）…')
    h_ledger = []
    for r, rec in rows:
        if not fillable(rec.get('百公里电耗(kWh/100km)')):
            continue
        cap, rng = cast_num(norm(rec.get('电池容量(kWh)'))), cast_num(norm(rec.get('纯电续航里程(km)')))
        if not cap or not rng or rng <= 0:
            continue
        v = round(cap * 100.0 / rng, 1)
        if not (5.0 <= v <= 40.0):
            continue
        cell = ws.cell(r, fi['百公里电耗(kWh/100km)'] + 1)
        cell.value = v
        cell.fill = BLUE
        rec['百公里电耗(kWh/100km)'] = v
        stats['百公里电耗(kWh/100km)'] += 1
        h_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''),
                         '百公里电耗(kWh/100km)', '公式补全', '(空)', v, '容量×100÷续航公式'])
    print(f'      阶段H 补 {len(h_ledger)} 格')
    ledgers.append((h_ledger, '公式补全'))

    print('[5/5] 阶段I 车长估算（基于F/G新补的轴距）…')
    i_ledger = []
    BODY_FACTOR = [('Lux Car', 1.66), ('Lux SUV', 1.66), ('Lux MPV', 1.65), ('Car', 1.72),
                   ('SUV', 1.69), ('MPV', 1.68), ('Sports', 1.64)]

    def body_factor(seg):
        for p, k in BODY_FACTOR:
            if str(seg or '').startswith(p):
                return k
        return 1.70
    for r, rec in rows:
        if not fillable(rec.get('车长(mm)')):
            continue
        ab = cast_num(norm(rec.get('轴距(mm)')))
        if not ab or not (1800 <= ab <= 4200):
            continue
        k = body_factor(norm(rec.get('细分市场')) or norm(rec.get('产品类型')))
        v = int(round(ab * k / 10.0) * 10)
        if not (1800 <= v <= 6500):
            continue
        cw_check = cast_num(norm(rec.get('整备质量(kg)')))
        if cw_check and not (0.18 <= cw_check / v <= 0.70):
            continue  # 整备/车长比出界→轴距或系数不可信，放弃估算（2026-09-13 教训）
        cell = ws.cell(r, fi['车长(mm)'] + 1)
        cell.value = v
        cell.fill = PURPLE
        rec['车长(mm)'] = v
        stats['车长(mm)'] += 1
        i_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), '车长(mm)',
                         '轴距估算', '(空)', v, f'轴距{int(ab)}×系数{k}估算（非实测，浅紫底纹）'])
    print(f'      阶段I 补 {len(i_ledger)} 格')
    ledgers.append((i_ledger, '轴距估算'))

    total = sum(len(x[0]) for x in ledgers)
    for ledger, ctype in ledgers:
        for row_vals in ledger:
            last_seq += 1
            cr.append([last_seq] + row_vals)
    print(f'      追加变更记录 {total} 条（末序号 {last_seq}）')

    # 保存（防 Excel 占用）
    saved = None
    try:
        wb.save(SRC)
        saved = SRC
        print('      已就地保存')
    except PermissionError:
        tmp = SRC.replace('（341~410批）', '（341~410批）v41tmp')
        wb.save(tmp)
        print(f'      原文件被占用，先写 {os.path.basename(tmp)}，开始重试替换…')
        for i in range(6):
            time.sleep(20)
            try:
                os.replace(tmp, SRC)
                saved = SRC
                print(f'      替换成功（重试 {i+1}）')
                break
            except PermissionError:
                pass
        if not saved:
            kept = SRC.replace('（341~410批）', '（341~410批·v41共识版）')
            os.replace(tmp, kept)
            saved = kept
            print(f'      持续被占用，保留为 {os.path.basename(kept)}')

    # 前后填充率对比
    print('\n===== v4.1.0 补全汇总（前后填充率对比） =====')
    wb2 = openpyxl.load_workbook(saved, read_only=True)
    ws2 = wb2[SHEET]
    it = ws2.iter_rows(values_only=True)
    hdr2 = next(it)
    n2 = 0
    new_fill = [0]*len(hdr2)
    for row in it:
        if not row or row[0] is None:
            continue
        n2 += 1
        for i, v in enumerate(row):
            if i < len(new_fill) and v not in (None, ''):
                new_fill[i] += 1
    wb2.close()
    for i, h in enumerate(hdr2):
        if h in base_fill:
            d = new_fill[i] - base_fill[h]
            if d > 0:
                print(f'  {h}: {base_fill[h]}({base_fill[h]/4494*100:.1f}%) -> {new_fill[i]}({new_fill[i]/n2*100:.1f}%)  +{d}')
    print(f'  合计补全 {sum(stats.values())} 格')


if __name__ == '__main__':
    sys.exit(main())
