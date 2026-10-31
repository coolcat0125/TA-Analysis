# -*- coding: utf-8 -*-
"""export_targets.py — 从缺口队列导出全量抓取目标

产出:
  dcd_auto/state/targets.json    给浏览器抓取工具(extract_browser.js)用
  dcd_auto/state/targets.csv     人审/抽查用(Excel 可开)

排序:缺口格数降序(抓取中断时,优先完成价值最高的车系)。
每条含:qkey / kw / kw_alt / family / part / expect_brand / shared_name / gap_cells / rows
      / gap_fields / priority_rank

用法:
    python export_targets.py                # 导出全部待抓取车系
    python export_targets.py --top 100      # 只导前 100(试跑用)
"""
import argparse
import csv
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dcd_common import QUEUE_FILE, STATE

OUT_JSON = os.path.join(STATE, 'targets.json')
OUT_CSV = os.path.join(STATE, 'targets.csv')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--top', type=int, default=0)
    a = ap.parse_args()
    if not os.path.exists(QUEUE_FILE):
        print('缺队列,先跑 dcd_queue.py'); return 1
    q = json.load(open(QUEUE_FILE, encoding='utf-8'))
    st_path = os.path.join(STATE, 'pipeline_state.json')
    done = set()
    if os.path.exists(st_path):
        st = json.load(open(st_path, encoding='utf-8'))
        done = {k for k, v in st['series'].items()
                if v.get('status') in ('done', 'rejected', 'hold', 'applied', 'rolled_back')}

    targets = []
    for e in q['entries']:
        if e['qkey'] in done:
            continue
        targets.append({
            'qkey': e['qkey'], 'kw': e['kw'], 'kw_alt': e.get('kw_alt'),
            'family': e['family'], 'part': e['part'],
            'expect_brand': e.get('expect_brand', []),
            'shared_name': bool(e.get('shared_name')),
            'already_extracted': bool(e.get('already_extracted')),
            'gap_cells': e['gap_cells'], 'rows': e['rows'],
            'gap_fields': e.get('gap_fields', {}), 'priority_rank': e['priority_rank'],
        })
    if a.top:
        targets = targets[:a.top]

    os.makedirs(STATE, exist_ok=True)
    json.dump({'generated_at': q.get('generated_at'), 'total': len(targets),
               'total_gap_cells': sum(t['gap_cells'] for t in targets), 'targets': targets},
              open(OUT_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    with open(OUT_CSV, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['优先级', '车系(族/名)', '搜索词', '缺口格', '行数', '缺口字段',
                    '预期品牌令牌', '共享名', '已抽取'])
        for t in targets:
            w.writerow([t['priority_rank'], f"{t['family']} / {t['part']}", t['kw'],
                        t['gap_cells'], t['rows'],
                        ';'.join(f'{k}×{v}' for k, v in t['gap_fields'].items()),
                        ' '.join(t['expect_brand']), '是' if t['shared_name'] else '',
                        '是' if t['already_extracted'] else ''])
    print(f'抓取目标已导出: {OUT_JSON}')
    print(f'  待抓取车系 {len(targets)} 个,缺口格 {sum(t["gap_cells"] for t in targets)}')
    print(f'  人审用 CSV: {OUT_CSV}')
    print('  前 8:')
    for t in targets[:8]:
        print(f"    #{t['priority_rank']:<4} {t['family']} / {t['part']}  kw={t['kw']}  gap={t['gap_cells']}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
