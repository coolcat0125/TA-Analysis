#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_consistency.py — 底表逻辑自洽校验（夜间迭代每轮必跑）
============================================================
只读校验，不修改底表。四类检查：

  CRITICAL（出现即为不通过，必须修复后重跑）：
    C1 供应商/名称列为纯数字（历史"复制污染"特征，参考 v4.2.5）
    C2 电机/发动机型号列无字母（纯数字/符号）
    C3 (批次, 产品型号) 重复行（排除 410 批公示占位等无型号行）
    C4 车长 < 轴距（物理不可能）
    C5 数值列出现非数值非占位的脏文本（排除公告原生格式与"不适用(x)"语义占位）

  WARN（逐条列出，供甄别；不阻断）：
    W1 八字段物理范围越界（真实值需外部佐证后才可修，参考仰望 U9X 2220kW 案例）
    W2 BEV 电耗守恒：|电耗 − 容量×100÷续航| > 33%（多配置口径混合或需核实）
    W3 BEV 行带数值型综合油耗（动力类型与发动机字段矛盾或脏值）
    W4 PHEV/EREV 行综合油耗与 B 状态油耗均为空
    W5 车长/轴距比值超出 [1.4, 2.4] 平台合理区间

用法：
  python verify_consistency.py [--input 底表.xlsx] [--json 输出.json]
  退出码：0=无 CRITICAL；1=存在 CRITICAL

Copyright © 2026 David YE
"""
import argparse
import json
import os
import re
import sys
from collections import Counter
from datetime import datetime, timezone

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))

MISSING = {'', '/', '-', '--', '—', 'N/A', 'NA', '0', '未填报', '无', '无信息',
           '待查', '待核实', '待确认', '未知', '？', '?', 'None', '未提供', '不适用'}
PLACEHOLDER = re.compile(r'^不适用\(')
NUMERIC_OK = re.compile(r'^\d+(\.\d+)?(/\d+(\.\d+)?)*(,\d+(\.\d+)?)*$')
MODEL_NUM = re.compile(r'^[\d./]+$')  # 无字母且无内部空格的纯数字串（“254 920”类官方代码带空格不算）
NAME_NUM = re.compile(r'^[\d\s./()（）：:kKW]+$')

RANGES = {'整备质量(kg)': (300, 10000), '纯电续航里程(km)': (5, 2000),
          '电池容量(kWh)': (1, 300), '电池能量密度(Wh/kg)': (30, 400),
          '百公里电耗(kWh/100km)': (3, 40), '电机总功率(kW)': (1, 2500),
          '车长(mm)': (1500, 20000), '轴距(mm)': (1000, 5000)}

NAME_COLS = ['电机生产企业', '发动机生产企业', '电芯供应商', '电池包供应商',
             '企业名称', '产品商标']
MODEL_COLS = ['电机型号', '发动机型号']


def fnum(s):
    if not s:
        return None
    m = re.fullmatch(r'(\d+(?:\.\d+)?)', s)
    return float(m.group(1)) if m else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', default=None)
    ap.add_argument('--json', default=os.path.join(HERE, 'audit-output', 'consistency_report.json'))
    args = ap.parse_args()

    path = args.input
    if not path:
        import glob
        cands = sorted(glob.glob(os.path.join(HERE, 'NEV公告参数汇总表_合并版*.xlsx')))
        if not cands:
            print('!! 未找到底表 xlsx')
            sys.exit(1)
        path = cands[-1]

    wb = load_workbook(path, read_only=True)
    ws = wb['NEV公告参数汇总']
    headers = [str(c.value).strip() if c.value is not None else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(headers)}

    def val(row, f):
        i = hi.get(f)
        if i is None or i >= len(row):
            return None
        v = row[i]
        return None if v is None or str(v).strip() in MISSING else str(v).strip()

    total = 0
    crit = Counter()
    warn = Counter()
    det = {'C1': [], 'C2': [], 'C3': [], 'C4': [], 'C5': [],
           'W1': [], 'W2': [], 'W3': [], 'W4': [], 'W5': []}
    seen_keys = Counter()

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, '') for v in row):
            continue
        total += 1
        b = val(row, '批次')
        m = val(row, '产品型号')
        pt = val(row, '动力类型') or ''
        is_bev = ('BEV' in pt) or ('纯电' in pt)
        is_phev = ('PHEV' in pt) or ('插' in pt)

        if b and m:
            seen_keys[(b, m)] += 1

        # C1/C2 类型语义
        for f in NAME_COLS:
            v = val(row, f)
            if v and NAME_NUM.fullmatch(v):
                crit['C1'] += 1
                if len(det['C1']) < 30:
                    det['C1'].append({'批次': b, '型号': m, '字段': f, '值': v[:30]})
        for f in MODEL_COLS:
            v = val(row, f)
            if v and MODEL_NUM.fullmatch(v):
                crit['C2'] += 1
                if len(det['C2']) < 30:
                    det['C2'].append({'批次': b, '型号': m, '字段': f, '值': v[:30]})

        # C5 数值列脏文本：去掉括号注释后不含任何数字的中文/杂文字（如“磷酸铁锂”“否”）。
        # 公告原生格式（512(CLTC)、53(单箱26.5/共2箱)、2040、2060、前200/后285）剥括号后均含数字，视为合法。
        for f in RANGES:
            v = val(row, f)
            if v and not PLACEHOLDER.match(v):
                stripped = re.sub(r'（[^）]*）|\([^)]*\)', '', v).strip()
                if stripped and not re.search(r'\d', stripped):
                    crit['C5'] += 1
                    if len(det['C5']) < 30:
                        det['C5'].append({'批次': b, '型号': m, '字段': f, '值': v[:30]})

        # W1 物理范围
        for f, (lo, rg_hi) in RANGES.items():
            x = fnum(val(row, f))
            if x is not None and not (lo <= x <= rg_hi):
                warn['W1'] += 1
                if len(det['W1']) < 60:
                    det['W1'].append({'批次': b, '型号': m, '字段': f, '值': val(row, f)})

        # W2 BEV 电耗守恒
        if is_bev:
            cap = fnum(val(row, '电池容量(kWh)'))
            rng = fnum(val(row, '纯电续航里程(km)'))
            cons = fnum(val(row, '百公里电耗(kWh/100km)'))
            if cap and rng and cons:
                exp = cap * 100 / rng
                if not (0.75 * exp <= cons <= 1.33 * exp):
                    warn['W2'] += 1
                    if len(det['W2']) < 60:
                        det['W2'].append({'批次': b, '型号': m,
                                          '说明': f'电耗{cons} vs 期望{exp:.1f}'})

        # W3 BEV 带数值油耗
        oil = val(row, '综合油耗(L/100km)')
        if is_bev and oil and NUMERIC_OK.fullmatch(oil) and '/' not in oil:
            warn['W3'] += 1
            if len(det['W3']) < 40:
                det['W3'].append({'批次': b, '型号': m, '油耗': oil})

        # W4 插混缺双油耗
        if is_phev and not oil and not val(row, 'B状态油耗(L/100km)'):
            warn['W4'] += 1
            if len(det['W4']) < 40:
                det['W4'].append({'批次': b, '型号': m})

        # C4/W5 车长轴距
        L = fnum(val(row, '车长(mm)'))
        A = fnum(val(row, '轴距(mm)'))
        if L and A:
            if L < A:
                crit['C4'] += 1
                if len(det['C4']) < 30:
                    det['C4'].append({'批次': b, '型号': m, '车长': L, '轴距': A})
            elif not (1.4 <= L / A <= 2.4):
                warn['W5'] += 1
                if len(det['W5']) < 40:
                    det['W5'].append({'批次': b, '型号': m, '比值': round(L / A, 2)})

    # C3 重复
    for (b, m), n in seen_keys.items():
        if n > 1:
            crit['C3'] += n - 1
            if len(det['C3']) < 30:
                det['C3'].append({'批次': b, '型号': m, '重复次数': n})

    report = {
        'generatedAt': datetime.now(timezone.utc).isoformat(),
        'input': os.path.basename(path),
        'records': total,
        'critical': dict(crit),
        'warn': dict(warn),
        'criticalTotal': sum(crit.values()),
        'warnTotal': sum(warn.values()),
        'detail': det,
        'pass': sum(crit.values()) == 0,
    }

    os.makedirs(os.path.dirname(args.json), exist_ok=True)
    with open(args.json, 'w', encoding='utf-8') as f:
        json.dump(report, f, ensure_ascii=False, indent=1)

    print(f'== verify_consistency ==  记录={total}')
    print(f"CRITICAL: {dict(crit) or '无'}  合计={report['criticalTotal']}")
    print(f"WARN:     {dict(warn) or '无'}  合计={report['warnTotal']}")
    print(f"结论: {'PASS' if report['pass'] else 'FAIL（存在必须修复项）'}")
    sys.exit(0 if report['pass'] else 1)


if __name__ == '__main__':
    main()
