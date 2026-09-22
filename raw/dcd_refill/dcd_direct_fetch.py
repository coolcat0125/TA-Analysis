# -*- coding: utf-8 -*-
"""dcd_direct_fetch.py — 登录态 cookie 直连懂车帝参数页（层A 执行器试点）

用法：
  python dcd_direct_fetch.py --probe --car 256634     # 探测 rawData 结构
  python dcd_direct_fetch.py --sid 4499 --out         # 抓取并写 dcd_s4499.json
  python dcd_direct_fetch.py --sid 4499 --diff        # 与既有样本对拍

cookie 来源：Tabbit 导出（$TEMP/dcd_cookies.json，Playwright 格式），不入 git。
"""
import argparse
import gzip
import io
import json
import os
import re
import sys
import time
import urllib.request
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
TEMP = os.environ.get('TEMP', r'C:\Users\sunji\AppData\Local\Temp')
COOKIE_FILE = os.path.join(TEMP, 'dcd_cookies.json')
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36')


def load_cookie_header():
    with open(COOKIE_FILE, encoding='utf-8') as f:
        cookies = json.load(f)
    pairs = [f"{c['name']}={c['value']}" for c in cookies
             if 'dongchedi.com' in (c.get('domain') or '')]
    return '; '.join(pairs)


def fetch(url, cookie_header, tries=3):
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={
                'User-Agent': UA,
                'Cookie': cookie_header,
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'zh-CN,zh;q=0.9',
                'Referer': 'https://www.dongchedi.com/',
                'Accept-Encoding': 'gzip',
            })
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
                if resp.headers.get('Content-Encoding') == 'gzip':
                    raw = gzip.decompress(raw)
                return resp.getcode(), raw.decode('utf-8', 'replace')
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2 * (i + 1))
    raise RuntimeError(f'fetch failed after {tries} tries: {last}')


def parse_next_data(html):
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.S)
    if not m:
        return None
    return json.loads(m.group(1))


def extract_series(next_data, sid):
    """rawData → 与既有样本同构 {series, cars:[{id,name,year,vals}]}。"""
    pp = next_data.get('props', {}).get('pageProps', {})
    rd = pp.get('rawData')
    if not rd:
        return None
    cars_raw = rd.get('car_info', []) or []
    if not cars_raw:
        return None
    c0 = cars_raw[0]
    out = {'series': {'sid': str(c0.get('series_id') or sid or ''),
                      'name': c0.get('series_name', ''),
                      'brand': c0.get('brand_name', '')},
           'cars': []}
    info_keys = ['length', 'wheelbase', 'curb_weight', 'cltc_recharge_mileage',
                 'recharge_mileage', 'power_consumption', 'battery_capacity',
                 'battery_type', 'battery_energy_density', 'total_electric_power',
                 'total_electric_torque', 'front_electric_max_power',
                 'front_electric_max_torque', 'rear_electric_max_power',
                 'rear_electric_max_torque', 'sub_brand_name']
    for car in cars_raw:
        info = car.get('info') or {}
        vals = {}
        for k in info_keys:
            cell = info.get(k)
            if isinstance(cell, dict):
                vals[k] = cell.get('value', '')
            else:
                vals[k] = cell if cell is not None else ''
        out['cars'].append({'id': str(car.get('car_id', car.get('id', ''))),
                            'name': car.get('car_name') or car.get('name', ''),
                            'year': str(car.get('car_year', car.get('year', ''))),
                            'sale_status': car.get('sale_status', ''),
                            'official_price': car.get('official_price', ''),
                            'vals': vals})
    return out


def diff_with_sample(built, sample):
    rep = {'cars_built': len(built['cars']), 'cars_sample': len(sample['cars']),
           'matched': 0, 'mismatch': [], 'sample_only': [], 'built_only': []}
    smap = {c['id']: c for c in sample['cars']}
    for c in built['cars']:
        s = smap.get(c['id'])
        if not s:
            rep['built_only'].append(c['id'])
            continue
        rep['matched'] += 1
        for k, v in c['vals'].items():
            sv = s['vals'].get(k, '')
            if str(v or '') != str(sv or ''):
                rep['mismatch'].append({'car': c['id'], 'key': k,
                                        'built': v, 'sample': sv})
    bset = {c['id'] for c in built['cars']}
    rep['sample_only'] = [cid for cid in smap if cid not in bset]
    return rep


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sid', help='车系 id')
    ap.add_argument('--car', help='款型 id（params 页入口；不给则用 series 页）')
    ap.add_argument('--probe', action='store_true')
    ap.add_argument('--out', action='store_true')
    ap.add_argument('--diff', action='store_true')
    a = ap.parse_args()

    if not os.path.exists(COOKIE_FILE):
        raise SystemExit(f'缺 cookie 文件 {COOKIE_FILE}（先跑 Tabbit 导出）')
    ch = load_cookie_header()

    url = (f'https://www.dongchedi.com/auto/params-carIds-x-{a.car}' if a.car
           else f'https://www.dongchedi.com/auto/series/{a.sid}')
    code, html = fetch(url, ch)
    print(f'HTTP {code} | {len(html)} bytes | __NEXT_DATA__={"__NEXT_DATA__" in html}')
    if code != 200:
        print('head:', html[:200])
        return
    nd = parse_next_data(html)
    if not nd:
        print('无 __NEXT_DATA__（可能被重定向登录页）')
        return
    rd = nd.get('props', {}).get('pageProps', {}).get('rawData')
    if a.probe:
        print('pageProps keys:', list(nd.get('props', {}).get('pageProps', {}).keys())[:20])
        if rd:
            print('rawData keys:', list(rd.keys())[:30])
            ci = rd.get('car_info') or []
            print('car_info n=', len(ci))
            if ci:
                c0 = ci[0]
                print('car0 keys:', list(c0.keys())[:25])
                info = c0.get('info') or {}
                print('car0 info n=', len(info), '| sample keys:',
                      list(info.keys())[:25])
                one = {k: info[k] for k in list(info.keys())[:3]}
                print('car0 info sample:', json.dumps(one, ensure_ascii=False)[:400])
        else:
            print('rawData 缺失')
        return
    built = extract_series(nd, a.sid)
    if not built:
        print('rawData 缺失，无法构建')
        return
    print(f"series: {built['series']} | cars: {len(built['cars'])}")
    if a.diff:
        sp = os.path.join(HERE, f'dcd_s{a.sid}.json')
        with open(sp, encoding='utf-8') as f:
            sample = json.load(f)
        rep = diff_with_sample(built, sample)
        print(json.dumps(rep, ensure_ascii=False, indent=1)[:2000])
    if a.out:
        op = os.path.join(HERE, f'dcd_s{a.sid}.json')
        with open(op, 'w', encoding='utf-8') as f:
            json.dump(built, f, ensure_ascii=False, indent=1)
        print('->', op, datetime.now().strftime('%H:%M:%S'))


if __name__ == '__main__':
    main()
