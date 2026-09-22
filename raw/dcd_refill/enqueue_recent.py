# -*- coding: utf-8 -*-
"""enqueue_recent.py — 把 recent_queue 剩余目标编入新批次 r09+（Wave 2）

按 recent_queue.json 的 targets（261 个），剔除 batch_state 已收录的 kw，
按缺口量降序切 40 车系/批，续号 r09, r10, ...
"""
import io
import json
import os
import re
import sys
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

st = json.load(open('batch_state.json', encoding='utf-8'))
rq = json.load(open('recent_queue.json', encoding='utf-8'))


def norm(s):
    return str(s or '').lower().replace(' ', '').replace('·', '')


have = set()
for b in st['batches'].values():
    for s in b.get('series', []):
        have.add(norm(s['kw']))

targets = [t for t in rq['targets'] if norm(t['kw']) not in have]
targets.sort(key=lambda t: -t.get('gap_rows', 0))
CHUNK = 40
n_new = 0
for i in range(0, len(targets), CHUNK):
    grp = targets[i:i + CHUNK]
    # 续号：现有最大 r 编号 + 1
    maxr = max((int(bid[1:]) for bid in st['batches']
                if re.match(r'^r\d+$', bid)), default=8)
    bid = f'r{maxr + 1:02d}'
    st['batches'][bid] = {
        'id': bid, 'status': 'planned-recent',
        'created': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'series': [{'kw': t['kw'], 'series_name': t.get('series_name', t['kw']),
                    'gap_rows': t.get('gap_rows', 0),
                    'fields': t.get('fields', {}),
                    'batches': t.get('batches', {})} for t in grp],
    }
    n_new += len(grp)
    print(f'{bid}: {len(grp)} 车系 (head: {grp[0]["kw"]} {grp[0].get("gap_rows")}格)')
json.dump(st, open('batch_state.json', 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
print(f'入队 {n_new} 车系 → batch_state.json')
