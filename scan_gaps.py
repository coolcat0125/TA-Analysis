# -*- coding: utf-8 -*-
"""scan_gaps.py — 缺口盘点器（DCD 自动工作流 · 环节1）

扫描权威底表，按优先级产出缺口清单：
  P1 缺失   ：字段为空（占位 '不适用(BEV)' 等语义占位不计缺失）
  P2 占位   ：'未提供'/'未提供具体数据'/'0' 等无效占位
  P3 存疑   ：逻辑异常（BEV 行带油耗/发动机值、电耗与容量/续航不自洽）
  P4 越界   ：物理范围外异常值
输出：
  raw/dcd_refill/gap_inventory.json   （机器可读，含 DCD 可映射字段标记）
  audit-output/gap_summary.txt        （人读摘要）
用法：python scan_gaps.py [--full]
"""
import json
import os
import re
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OUT_JSON = os.path.join(HERE, 'raw', 'dcd_refill', 'gap_inventory.json')
OUT_TXT = os.path.join(HERE, 'audit-output', 'gap_summary.txt')

# DCD 通道可映射字段（key=底表列, dcd=DCD参数key）
DCD_FIELDS = {
    '车长(mm)': 'length',
    '轴距(mm)': 'wheelbase',
    '整备质量(kg)': 'curb_weight',
    '纯电续航里程(km)': 'cltc_recharge_mileage',
    '百公里电耗(kWh/100km)': 'power_consumption',
    '电池容量(kWh)': 'battery_capacity',
    '电池类型': 'battery_type',
    '电池能量密度(Wh/kg)': 'battery_energy_density',
    '电机总功率/扭矩': 'total_electric_power',
    '前电机功率/扭矩': 'front_electric_max_power',
    '后电机功率/扭矩': 'rear_electric_max_power',
}
# 物理范围（超出即 P4 越界）
NUM_RANGE = {
    '整备质量(kg)': (700, 4000), '轴距(mm)': (1800, 4000), '车长(mm)': (3000, 6000),
    '纯电续航里程(km)': (30, 1500), '电池容量(kWh)': (5, 200), '电池能量密度(Wh/kg)': (50, 260),
    '百公里电耗(kWh/100km)': (5, 60), '发动机排量(mL)': (400, 4000), '发动机功率(kW)': (20, 600),
}
PLACEHOLDERS = {'未提供', '未提供具体数据', '无', '-', '/', '无数据', 'N/A', 'n/a'}
# 逻辑存疑字段（按动力类型）
BEV_FORBIDDEN = ['综合油耗(L/100km)', 'B状态油耗(L/100km)', '发动机排量(mL)',
                 '发动机功率(kW)', '发动机生产企业', '发动机型号']


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', '').replace('km', '').strip())
    except Exception:
        return None


def is_placeholder(v):
    return str(v).strip() in PLACEHOLDERS


def scan():
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    total = len(rows)

    # 逐字段统计
    field_stat = {}
    for f in hdr:
        if f in ('批次', '产品型号'):
            continue
        field_stat[f] = {'filled': 0, 'empty': 0, 'placeholder': 0, 'out_of_range': 0,
                         'dcd_mappable': f in DCD_FIELDS}
    # 缺口明细（行级）——仅记录 DCD 可映射字段 + 存疑/越界（控制体积）
    gaps = []          # {batch, model, field, kind, value}
    logic_issues = []  # 存疑/越界明细

    for r in rows:
        batch = str(r[hi['批次']] or '')
        model = str(r[hi['产品型号']] or '').strip()
        pt = str(r[hi['动力类型']] or '')
        is_bev = '纯电' in pt or 'BEV' in pt

        for f, st in field_stat.items():
            v = r[hi[f]] if f in hi else None
            vs = str(v).strip() if v is not None else ''
            # 语义占位不算缺失
            if f in BEV_FORBIDDEN and is_bev and vs.startswith('不适用'):
                st['filled'] += 1
                continue
            if vs == '':
                st['empty'] += 1
                if f in DCD_FIELDS:
                    gaps.append({'batch': batch, 'model': model, 'field': f, 'kind': 'P1缺失'})
            elif is_placeholder(vs):
                st['placeholder'] += 1
                if f in DCD_FIELDS:
                    gaps.append({'batch': batch, 'model': model, 'field': f, 'kind': 'P2占位'})
            else:
                st['filled'] += 1
                # P4 越界
                if f in NUM_RANGE:
                    n = fnum(v)
                    if n is not None and not (NUM_RANGE[f][0] <= n <= NUM_RANGE[f][1]):
                        st['out_of_range'] += 1
                        logic_issues.append({'batch': batch, 'model': model, 'field': f,
                                             'kind': 'P4越界', 'value': vs})
                # P3 存疑：BEV 行带有效油耗/发动机值
                if is_bev and f in BEV_FORBIDDEN and not vs.startswith('不适用') and vs != '':
                    logic_issues.append({'batch': batch, 'model': model, 'field': f,
                                         'kind': 'P3存疑', 'value': vs})
                # P3 存疑：电耗不自洽（|电耗-容量/续航*100| > 20%）
                if f == '百公里电耗(kWh/100km)' and not is_bev:
                    cap, rng = fnum(r[hi['电池容量(kWh)']]), fnum(r[hi['纯电续航里程(km)']])
                    e = fnum(v)
                    if cap and rng and e and abs(e - cap / rng * 100) / e > 0.2:
                        logic_issues.append({'batch': batch, 'model': model, 'field': f,
                                             'kind': 'P3存疑', 'value': f'{vs}(推算{cap/rng*100:.1f})'})

    # 汇总
    coverage = {}
    for f, st in field_stat.items():
        denom = total - st['empty'] - st['placeholder'] + st['empty'] + st['placeholder']
        eff = st['filled']
        coverage[f] = {
            'filled': st['filled'], 'empty': st['empty'], 'placeholder': st['placeholder'],
            'out_of_range': st['out_of_range'],
            'coverage': round(eff / total * 100, 1),
            'dcd_mappable': st['dcd_mappable'],
        }

    # DCD 可补缺口按字段汇总
    dcd_gap_by_field = defaultdict(int)
    for g in gaps:
        dcd_gap_by_field[g['field']] += 1

    inv = {
        'workbook': os.path.basename(WB),
        'total_rows': total,
        'coverage': coverage,
        'dcd_gaps': gaps,
        'dcd_gap_total': len(gaps),
        'dcd_gap_by_field': dict(sorted(dcd_gap_by_field.items(), key=lambda x: -x[1])),
        'logic_issues': logic_issues,
        'logic_total': len(logic_issues),
        'legend': {
            'P1缺失': '字段为空，DCD可补',
            'P2占位': '无效占位值，DCD可补',
            'P3存疑': '逻辑异常（BEV油耗/电耗不自洽），需人工核证',
            'P4越界': '物理范围外异常值，需甄别',
        },
    }
    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    json.dump(inv, open(OUT_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    os.makedirs(os.path.dirname(OUT_TXT), exist_ok=True)
    lines = [f'缺口盘点 · {inv["workbook"]} · 共 {total} 行',
             f'DCD 可补缺口（P1+P2）：{len(gaps)} 格',
             '按字段：']
    for f, c in inv['dcd_gap_by_field'].items():
        lines.append(f'  {f}: {c}')
    lines.append(f'\n存疑/越界（P3+P4）：{len(logic_issues)} 处')
    lines.append('\n字段覆盖率：')
    for f, st in sorted(coverage.items(), key=lambda x: x[1]['coverage']):
        tag = ' [DCD]' if st['dcd_mappable'] else ''
        lines.append(f"  {f}: {st['coverage']}%  空{st['empty']} 占位{st['placeholder']} 越界{st['out_of_range']}{tag}")
    open(OUT_TXT, 'w', encoding='utf-8').write('\n'.join(lines))
    print('\n'.join(lines[:12]))
    print(f'\n→ {OUT_JSON}')
    print(f'→ {OUT_TXT}')


if __name__ == '__main__':
    scan()
