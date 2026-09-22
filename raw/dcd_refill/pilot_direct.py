# -*- coding: utf-8 -*-
"""pilot_direct.py — 层A 直连试点：与既有 Tabbit 样本逐字段对拍（Wave 0 验收）

用法：
  python pilot_direct.py            # 对 series_mapping.json 全部 ok 车系对拍
  python pilot_direct.py 2911 3101  # 指定 sid
  python pilot_direct.py --save     # 对拍通过的车系用直连数据覆盖落盘
"""
import io
import json
import os
import sys
import time

import dcd_direct_fetch as d

HERE = os.path.dirname(os.path.abspath(__file__))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def pilot(sid, save=False):
    sp = os.path.join(HERE, f'dcd_s{sid}.json')
    if not os.path.exists(sp):
        return {'sid': sid, 'ok': False, 'reason': 'no sample'}
    with open(sp, encoding='utf-8') as f:
        sample = json.load(f)
    ch = d.load_cookie_header()
    built = None
    used_car = None
    # 修正：params URL 的数字是车系 id（params-carIds-x-<sid>），非款型 id
    try:
        code, html = d.fetch(
            f'https://www.dongchedi.com/auto/params-carIds-x-{sid}', ch)
    except Exception as e:  # noqa: BLE001
        return {'sid': sid, 'ok': False, 'reason': f'fetch: {e}'}
    nd = d.parse_next_data(html) if code == 200 else None
    built = d.extract_series(nd, sid) if nd else None
    used_car = 'sid-direct'
    time.sleep(2)
    if not built:
        return {'sid': sid, 'ok': False, 'reason': 'rawData 空'}
    rep = d.diff_with_sample(built, sample)
    out = {'sid': sid, 'ok': True, 'series': built['series']['name'],
           'entry_car': used_car, 'cars_http': rep['cars_built'],
           'cars_sample': rep['cars_sample'], 'matched': rep['matched'],
           'mismatch_n': len(rep['mismatch']),
           'mismatch_sample': rep['mismatch'][:4],
           'sample_only_n': len(rep['sample_only']),
           'built_only_n': len(rep['built_only'])}
    if save and out['ok'] and out['mismatch_n'] == 0:
        with open(sp, 'w', encoding='utf-8') as f:
            json.dump(built, f, ensure_ascii=False, indent=1)
        out['saved'] = True
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    save = '--save' in sys.argv
    if args:
        sids = args
    else:
        with open(os.path.join(HERE, 'series_mapping.json'), encoding='utf-8') as f:
            sids = [m['sid'] for m in json.load(f)['reviewed'] if m.get('ok')]
    results = []
    for i, sid in enumerate(sids):
        r = pilot(sid, save=save)
        results.append(r)
        flag = 'PASS' if r.get('ok') and r.get('mismatch_n') == 0 else 'DIFF' if r.get('ok') else 'FAIL'
        print(f"[{i+1}/{len(sids)}] sid={sid} {flag} "
              f"{r.get('series','')} cars={r.get('cars_http','-')} "
              f"mismatch={r.get('mismatch_n','-')} {r.get('reason','')}")
        if r.get('ok') and r.get('mismatch_n'):
            print('   ', json.dumps(r['mismatch_sample'], ensure_ascii=False)[:400])
        time.sleep(2)
    npass = sum(1 for r in results if r.get('ok') and r.get('mismatch_n') == 0)
    print(f"\n=== 试点结果: {npass}/{len(results)} 全字段一致 ===")
    with open(os.path.join(HERE, 'pilot_direct_result.json'), 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
