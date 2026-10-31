# -*- coding: utf-8 -*-
"""reclass_rear_motor.py — 后电机缺口重分类（方案 §2.2，2026-09-22）

把 gap_inventory 中「后电机功率/扭矩」缺口按 总功率 vs 前功率 重定性：
  single_legal   总≈前 → 单电机车，空=合法空置（从 DCD 缺口剔除）
  dual_gap       总>前 → 真双电机缺口（DCD 正常回填）
  rear_only      前空+总有 → 疑后驱车，DCD 回填后电机≈总功率（可回填）
  prerequisite   总/前不可解析 → 前置依赖（先补总/前再回头）

只读底表，输出 raw/dcd_refill/gap_reclass.json；不写工作簿。
可复跑：python reclass_rear_motor.py
"""
import io
import json
import os
import re
import sys
from datetime import datetime

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
_cands = [f for f in os.listdir(_ROOT) if f.endswith('.xlsx') and '410' in f and '汇总' in f]
if not _cands:
    raise SystemExit('repo 根未找到底表 xlsx')
WB = os.path.join(_ROOT, _cands[0])
OUT = os.path.join(HERE, 'gap_reclass.json')

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def first_num(s):
    """P/T 文本取首个数值（约定 split('/')[0] 口径的宽松版）。"""
    if s is None:
        return None
    m = re.search(r'([\d.]+)', str(s))
    return float(m.group(1)) if m else None


def main():
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    header = list(next(it))
    hi = {name: i for i, name in enumerate(header) if name}

    need = ['产品型号', '电机总功率/扭矩', '前电机功率/扭矩', '后电机功率/扭矩']
    missing = [c for c in need if c not in hi]
    if missing:
        raise SystemExit(f'缺列: {missing}')
    batch_col = next((c for c in header if c and '批次' in str(c)), None)
    generic_col = next((c for c in header if c == '通用名称'), None)
    ptype_col = next((c for c in header if c == '动力类型'), None)

    rows = []
    counts = {'single_legal': 0, 'dual_gap': 0, 'rear_only': 0, 'prerequisite': 0}
    rear_empty = 0
    by_batch = {}
    for r in it:
        rear = r[hi['后电机功率/扭矩']]
        if rear not in (None, ''):
            continue
        rear_empty += 1
        model = r[hi['产品型号']]
        total_p = first_num(r[hi['电机总功率/扭矩']])
        front_p = first_num(r[hi['前电机功率/扭矩']])
        if total_p is None or front_p is None:
            if front_p is None and total_p is not None:
                verdict = 'rear_only'
            else:
                verdict = 'prerequisite'
        elif abs(total_p - front_p) < 0.5:
            verdict = 'single_legal'
        else:
            verdict = 'dual_gap'
        counts[verdict] += 1
        batch = str(r[hi[batch_col]]) if batch_col else ''
        by_batch.setdefault(verdict, {})
        by_batch[verdict][batch] = by_batch[verdict].get(batch, 0) + 1
        rows.append({
            'batch': batch,
            'model': model,
            'generic': r[hi[generic_col]] if generic_col else '',
            'power_type': r[hi[ptype_col]] if ptype_col else '',
            'verdict': verdict,
            'total_p': total_p,
            'front_p': front_p,
        })
    wb.close()

    out = {
        'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'workbook': os.path.basename(WB),
        'rear_empty_total': rear_empty,
        'counts': counts,
        'by_batch': by_batch,
        'effective_dcd_gap_note': '有效缺口 = 原9766 - single_legal（单电机合法空置不计缺口）',
        'rows': rows,
    }
    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f'rear_empty={rear_empty} | ' + ' | '.join(f'{k}={v}' for k, v in counts.items()))
    print(f'-> {OUT}')


if __name__ == '__main__':
    main()
