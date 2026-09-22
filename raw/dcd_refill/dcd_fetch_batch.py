# -*- coding: utf-8 -*-
"""dcd_fetch_batch.py — 直连批量抓取车系参数（Wave 1-B）

对 batch_state 指定批次内 ok 映射的车系（默认 r*），逐 sid 直连
params-carIds-x-<sid> 抓全量款型参数，落盘 dcd_s<sid>.json：
  {kw: 首kw, kws: [全部别名], channel, fetched_at, series_url, series, cars}
逐文件落盘（断点续跑：已存在即跳过，--force 重抓）；失败记 fetch_log.json。

用法：
  python dcd_fetch_batch.py --batches r01,r02   # 指定批次
  python dcd_fetch_batch.py                     # 全部 r* 批次
"""
import argparse
import io
import json
import os
import sys
import time
from datetime import datetime

import dcd_direct_fetch as d

HERE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(HERE, 'batch_state.json')
FETCH_LOG = os.path.join(HERE, 'fetch_log.json')

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--batches', default='')
    ap.add_argument('--force', action='store_true')
    a = ap.parse_args()

    st = json.load(open(STATE, encoding='utf-8'))
    wanted = set(a.batches.split(',')) if a.batches else None
    sid_kws = {}
    for bid, b in sorted(st['batches'].items()):
        if wanted and bid not in wanted:
            continue
        if not (bid.startswith('r') or bid.startswith('b')):
            continue
        for s in b.get('series', []):
            if s.get('ok') and s.get('sid'):
                sid_kws.setdefault(str(s['sid']), []).append(s['kw'])

    log = {}
    if os.path.exists(FETCH_LOG):
        with open(FETCH_LOG, encoding='utf-8') as f:
            log = json.load(f)
    ch = d.load_cookie_header()
    print(f'待抓取: {len(sid_kws)} 车系')
    nok = nfail = 0
    for i, (sid, kws) in enumerate(sorted(sid_kws.items())):
        out_path = os.path.join(HERE, f'dcd_s{sid}.json')
        if os.path.exists(out_path) and not a.force:
            nok += 1
            continue
        entry = {'kws': kws, 'time': datetime.now().strftime('%m-%d %H:%M')}
        built = None
        for attempt in (1, 2):
            try:
                code, html = d.fetch(
                    f'https://www.dongchedi.com/auto/params-carIds-x-{sid}', ch)
                nd = d.parse_next_data(html) if code == 200 else None
                built = d.extract_series(nd, sid) if nd else None
            except Exception as e:  # noqa: BLE001
                entry['error'] = f'fetch: {e}'
            if built and built['cars']:
                break
            entry['error'] = entry.get('error') or 'rawData 空'
            built = None
            time.sleep(3)
        if built:
            doc = {'kw': kws[0], 'kws': kws, 'channel': 'direct-http',
                   'fetched_at': entry['time'],
                   'series_url': f'https://www.dongchedi.com/auto/series/{sid}',
                   'series': built['series'], 'cars': built['cars']}
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(doc, f, ensure_ascii=False, indent=1)
            entry['ok'] = True
            entry['cars'] = len(built['cars'])
            nok += 1
            print(f'[{i+1}/{len(sid_kws)}] sid={sid} {kws[0]} -> '
                  f"{built['series']['name']} {len(built['cars'])}款")
        else:
            nfail += 1
            log[sid] = entry
            print(f'[{i+1}/{len(sid_kws)}] sid={sid} {kws[0]} -> FAIL '
                  f"({entry.get('error')})")
            with open(FETCH_LOG, 'w', encoding='utf-8') as f:
                json.dump(log, f, ensure_ascii=False, indent=1)
        time.sleep(2)
    print(f'完成: ok={nok} fail={nfail}')


if __name__ == '__main__':
    main()
