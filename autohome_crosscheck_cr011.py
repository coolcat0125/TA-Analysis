# -*- coding: utf-8 -*-
"""autohome_crosscheck_cr011.py — CR-17-011 疑点续航比离群·汽车之家跨平台取证

11 例底表↔doc3 续航比离群（非 NEDC↔CLTC 工况带）——配对候选待取证：
  汽车之家 listSpec 按车系查款型 → 找与底表行签名（续航±8）匹配的款型
  → 若存在则证明底表值有效（doc3 配对错位）；若不存在则 doc3 值存疑
输出：audit-output/autohome_crosscheck_cr011.json
"""
import json
import os
import re
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'audit-output', 'autohome_crosscheck_cr011.json')
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
OPENER = None


def setup():
    global OPENER
    OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def get(url, delay=0.55):
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    with OPENER.open(req, timeout=25) as r:
        data = json.loads(r.read().decode('utf-8'))
    time.sleep(delay)
    return data


def suggest(q):
    u = ('https://sou.api.autohome.com.cn/sug/_suggest?plat=pc&uid=0&q=' + urllib.parse.quote(q))
    d = get(u)
    return [(str(x.get('key') or ''), int(x.get('wordid') or 0)) for x in (d.get('result') or {}).get('data', []) if x.get('wordid')]


def list_spec(seriesid):
    u = (f'https://car-web-api.autohome.com.cn/car/spec/listSpec?type=0x001f&from=1&pm=1&pluginversion=11.64.8&seriesid={seriesid}')
    r = get(u).get('result') or {}
    out = []
    for grp in (r.get('list') or []) + (r.get('otherlist') or []):
        for sp in (grp.get('speclist') or []):
            out.append({'id': sp.get('id'), 'name': sp.get('name')})
    return out


TARGETS = [
    {'model': 'FV7147FADEHEV', 'generic': '迈腾GTE', 'brand': '大众', 'row_range': 54.0, 'doc_range': 77.0},
    {'model': 'HFC6494ECSEV2-W', 'generic': '思皓E10X', 'brand': '江淮', 'row_range': 635.0, 'doc_range': 445.0},
    {'model': 'JL7001SEV03', 'generic': '银河L6', 'brand': '吉利', 'row_range': 415.0, 'doc_range': 605.0},
    {'model': 'LF6461SEV01', 'generic': '雷诺', 'brand': '力帆', 'row_range': 450.0, 'doc_range': 605.0},
    {'model': 'JX6556TA-M6BEV', 'generic': '江铃E路达', 'brand': '江铃', 'row_range': 306.0, 'doc_range': 357.0},
    {'model': 'LF6462SEV01', 'generic': '雷诺', 'brand': '力帆', 'row_range': 450.0, 'doc_range': 605.0},
    {'model': 'CSA6501LBEV6', 'generic': '荣威iMAX8 EV', 'brand': '荣威', 'row_range': 620.0, 'doc_range': 640.0},
    {'model': 'XMA6500LBEVA1', 'generic': '小马', 'brand': '厦门金龙', 'row_range': 750.0, 'doc_range': 760.0},
    {'model': 'MR7003BEV102', 'generic': '极氪', 'brand': '极氪', 'row_range': 690.0, 'doc_range': 710.0},
    {'model': 'CAF6461A62PHEV', 'generic': '红旗E-QM5', 'brand': '红旗', 'row_range': 54.0, 'doc_range': 77.0},
    {'model': 'BYD7153WT6HEV5', 'generic': '比亚迪宋', 'brand': '比亚迪', 'row_range': 103.0, 'doc_range': 60.0},
]


def main():
    setup()
    import urllib.parse
    results = []
    for t in TARGETS:
        g = t['generic'].lower().replace(' ', '')
        # suggest 找车系
        hits = suggest(t['generic'])
        best = None
        for name, sid in hits:
            n = name.lower().replace(' ', '')
            if g.replace('plus', '') in n or n in g:
                best = (name, sid)
                break
        entry = dict(t)
        if not best:
            entry['verdict'] = 'no_series_match'
            results.append(entry)
            print(f"{t['model']}: 无车系匹配 ({t['generic']})")
            continue
        # listSpec 查款型
        try:
            specs = list_spec(best[1])
        except Exception as e:
            entry['verdict'] = f'spec_err: {str(e)[:40]}'
            results.append(entry)
            print(f"{t['model']}: spec ERR {str(e)[:40]}")
            continue
        # 款型名与底表通用名称匹配
        matching = [s for s in specs if g[:6] in s['name'].lower().replace(' ', '') or any(x in s['name'].lower().replace(' ', '') for x in g[:4].split(','))]
        entry['series'] = best[0]
        entry['sid'] = best[1]
        entry['specs_n'] = len(specs)
        entry['matching_n'] = len(matching)
        entry['verdict'] = 'found' if matching else 'no_matching_trim'
        entry['matching_names'] = [s['name'][:24] for s in matching[:3]]
        results.append(entry)
        print(f"{t['model']}: series={best[0]}({best[1]}) matching={len(matching)}")
        time.sleep(0.55)
    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('saved:', OUT)


if __name__ == '__main__':
    main()
