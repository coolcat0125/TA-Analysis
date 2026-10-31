# -*- coding: utf-8 -*-
"""wave_status.py — Wave 2 实时监控看板（数据源=batch_state/fetch_log/台账）

用法：python wave_status.py [--watch 间隔秒]
"""
import argparse
import glob
import io
import json
import os
import re
import sys
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def snapshot():
    st = json.load(open(os.path.join(HERE, 'batch_state.json'), encoding='utf-8'))
    sids_ok = set()
    rows = []
    for bid in sorted(st['batches']):
        b = st['batches'][bid]
        ss = b.get('series', [])
        mapped = sum(1 for s in ss if s.get('ok') and s.get('sid'))
        for s in ss:
            if s.get('ok') and s.get('sid'):
                sids_ok.add(str(s['sid']))
        rows.append((bid, b.get('status', ''), len(ss), mapped,
                     b.get('applied', {}).get('green')))
    have_data = {re.match(r'dcd_s(\d+)\.json$', f).group(1)
                 for f in os.listdir(HERE) if re.match(r'dcd_s\d+\.json$', f)}
    tot = sum(r[2] for r in rows)
    mapped_all = sum(r[3] for r in rows)
    withdata = len(sids_ok & have_data)
    green = sum(r[4] or 0 for r in rows if r[4] is not None)
    fails = 0
    fl = os.path.join(HERE, 'fetch_log.json')
    if os.path.exists(fl):
        fails = len(json.load(open(fl, encoding='utf-8')))
    return rows, dict(total=tot, mapped=mapped_all, with_data=withdata,
                      green=green, fetch_fail=fails,
                      ts=datetime.now().strftime('%H:%M:%S'))


def render():
    rows, g = snapshot()
    print(f"== DCD Wave 监控 @ {g['ts']} ==")
    print(f"车系: {g['total']} | 已映射: {g['mapped']} | 数据落盘(去重sid): "
          f"{g['with_data']} | 抓取失败: {g['fetch_fail']} | apply绿格累计: {g['green']}")
    for bid, status, n, m, gr in rows:
        flag = {'planned': '…', 'planned-recent': '…', 'ingested': 'D',
                'applied': 'A', 'partial': 'p', 'waiting': 'w'}.get(status, '?')
        print(f"  {bid} [{flag}] {m}/{n} 映射"
              + (f" 绿{gr}" if gr is not None else ''))
    return g


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--watch', type=int, default=0, help='轮询间隔秒，0=单次')
    ap.add_argument('--until-mapped', type=int, default=0,
                    help='映射达到该值即退出(配合watch)')
    a = ap.parse_args()
    if not a.watch:
        render()
    else:
        while True:
            g = render()
            if a.until_mapped and g['mapped'] >= a.until_mapped:
                print('达标退出')
                break
            time.sleep(a.watch)
