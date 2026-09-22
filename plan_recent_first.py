# -*- coding: utf-8 -*-
"""plan_recent_first.py — 由近及远抓取批次计划（用户指令 2026-09-22）

在既有 batch_state.json 基础上追加 recent-first 批次（b09+）：
  排序口径 = 2026年公告批次(403~410)缺口格数降序 → 同格数按批次新→旧
  过滤：泛称占位（纯电动轿车等）、已在采集映射/旧计划中的车系
产物：
  raw/dcd_refill/batch_state.json   追加 r01~r08 批次（status=planned-recent）
  raw/dcd_refill/recent_queue.json  浏览器通道队列（kw/fields/batches/gap_rows）
"""
import json
import os
import re
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
REFILL = os.path.join(HERE, 'raw', 'dcd_refill')
INV = os.path.join(REFILL, 'gap_inventory.json')
STATE = os.path.join(REFILL, 'batch_state.json')
MAPPING = os.path.join(REFILL, 'series_mapping.json')

GENERIC = re.compile(r'纯电动|插电|换电式|增程|燃料电池|轿车|多用途|运动型|越野|乘用车|未知|未提供')


def main():
    inv = json.load(open(INV, encoding='utf-8'))
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    series_of = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        m = str(r[hi['产品型号']] or '').strip()
        series_of[m] = (str(r[hi['通用名称']] or '').strip(),
                        str(r[hi['车型名称']] or '').strip(),
                        str(r[hi['企业名称']] or '').strip())
    bt = {}
    for r in wb['批次时间表'].iter_rows(min_row=2, values_only=True):
        if r[0] and r[2]:
            bt[str(r[0]).strip()] = str(r[2])

    recent_batches = sorted([b for b, d in bt.items() if d.startswith('2026')], key=lambda x: int(x))

    # 车系聚合（2026批缺口）
    agg = defaultdict(lambda: {'rows': 0, 'fields': defaultdict(int), 'batches': defaultdict(int)})
    for g in inv['dcd_gaps']:
        if g['batch'] not in recent_batches:
            continue
        gen, car, ent = series_of.get(g['model'], (g['model'], '', ''))
        s = gen or car or g['model']
        if GENERIC.search(s) or len(s) < 2:
            continue
        agg[s]['rows'] += 1
        agg[s]['fields'][g['field']] += 1
        agg[s]['batches'][g['batch']] += 1

    # 已采集/已计划车系跳过
    mapping = json.load(open(MAPPING, encoding='utf-8'))
    done_kw = {str(m['kw']).lower().replace(' ', '') for m in mapping.get('reviewed', []) if m.get('ok')}
    st = json.load(open(STATE, encoding='utf-8'))
    for b in st['batches'].values():
        for s in b.get('series', []):
            kw = s['kw'] if isinstance(s, dict) else s
            done_kw.add(str(kw).lower().replace(' ', ''))

    def norm(s):
        return s.lower().replace(' ', '')

    ranked = []
    for s, info in agg.items():
        if norm(s) in done_kw:
            continue
        # 搜索关键词：去版本后缀
        kw = re.sub(r'[（(].*?[)）]|\s*\d{4}\s*款.*$', '', s).strip() or s
        if GENERIC.search(kw) or len(kw) < 2:
            continue
        if norm(kw) in done_kw:
            continue
        ranked.append({'kw': kw, 'series_name': s, 'gap_rows': info['rows'],
                       'fields': dict(info['fields']), 'batches': dict(info['batches'])})
    # 由近及远：先按最新批次缺口、再按总量
    def sort_key(x):
        newest = max((int(b) for b in x['batches']), default=0)
        newest_gap = x['batches'].get(str(newest), 0)
        return (-newest, -newest_gap, -x['gap_rows'])
    ranked.sort(key=sort_key)

    # 分批（5车系/批，对齐 Tabbit 180s 限额）
    batches = []
    n_exist = len(st['batches'])
    for i in range(0, min(len(ranked), 40), 5):
        grp = ranked[i:i + 5]
        bid = f'r{len(batches) + 1:02d}'
        batches.append({'id': bid, 'series': grp, 'status': 'planned-recent',
                        'created': __import__('time').strftime('%Y-%m-%d %H:%M:%S'),
                        'priority': 'recent-first(用户指令2026-09-22)'})
    for b in batches:
        st['batches'][b['id']] = b
    json.dump(st, open(STATE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    queue = {'created': __import__('time').strftime('%Y-%m-%d %H:%M:%S'),
             'priority': '由近及远：2026公告批次(403~410)优先',
             'recent_batches': {b: bt[b] for b in recent_batches},
             'targets': ranked,
             'total_series': len(ranked),
             'total_gap_cells': sum(r['gap_rows'] for r in ranked)}
    json.dump(queue, open(os.path.join(REFILL, 'recent_queue.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)

    print(f"2026批缺口车系（去重/去泛称/去已采）: {len(ranked)} 个 / {queue['total_gap_cells']} 格")
    print(f"生成 recent-first 批次: {len(batches)} 批 × 5 车系 → batch_state.json")
    print(f"队列文件: raw/dcd_refill/recent_queue.json（浏览器通道可直接消费）")
    print('\n--- r01~r08 预览 ---')
    for b in batches:
        sers = '、'.join(f"{s['kw']}({s['gap_rows']}格)" for s in b['series'])
        print(f"  {b['id']}: {sers}")
    print('\n--- TOP15 车系明细 ---')
    for r in ranked[:15]:
        bs = ','.join(f"{k}:{v}" for k, v in sorted(r['batches'].items(), key=lambda x: -int(x[0])))
        print(f"  {r['series_name'][:24]:<26} {r['gap_rows']:>3}格 [{bs}]")


if __name__ == '__main__':
    main()
