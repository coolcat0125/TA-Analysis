# -*- coding: utf-8 -*-
"""import_bundle.py — 浏览器 bundle 导入器(抓取层 ↔ 回填管线的接缝)

输入:dcd_auto/state/bundles/dcd_bundle_*.json(extract_browser.js 的下载产物)
输出:
  raw/dcd_refill/dcd_s<sid>.json     既有格式 {series, cars:[{id,name,year,vals}]}
  raw/dcd_refill/dcd_fuel_form.json  能源类型合并 {sid:[{id,name,ff}]}
  raw/dcd_refill/series_mapping.json 自动追加人审等效条目(review:"auto")
  dcd_auto/state/import_report.json  导入报告

品牌校验:对每条 bundle 记录跑 decide_series(与 Tabbit 时代同一套判据,
  15 例人审复现 15/15),accept → mapping ok=true;reject/hold → ok=false + 原因,
  数据文件仍落盘但不进入回填(可人工复核后改 ok)。

用法:
    python import_bundle.py [--bundles 目录] [--overwrite]
"""
import argparse
import glob
import json
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dcd_common import QUEUE_FILE, RAW_DCD, STATE, norm_name
from dcd_refill_core import decide_series

MAPPING = os.path.join(RAW_DCD, 'series_mapping.json')
FUEL = os.path.join(RAW_DCD, 'dcd_fuel_form.json')
REPORT = os.path.join(STATE, 'import_report.json')
BUNDLES = os.path.join(STATE, 'bundles')


def load_qkey_index():
    """qkey → 队列表(用于取 expect_brand / shared_name / kw)。

    双索引:qkey 主索引;kw 归一为副索引——bundle 的 qkey 可能因队列重建而失效
    (2026-09-22 实测:人审 kw '北汽EU5' 与队列 part 'EU5' 前缀不匹配时 qkey 落空,
    导致 expect_brand 为空、一律误拒)。副索引按 kw 与前缀锚定双路查找。
    """
    idx, by_kw = {}, {}
    for path in (QUEUE_FILE, os.path.join(STATE, 'targets.json')):
        if not os.path.exists(path):
            continue
        d = json.load(open(path, encoding='utf-8'))
        for e in d.get('entries') or d.get('targets') or []:
            if e.get('qkey'):
                idx.setdefault(e['qkey'], e)
            if e.get('kw'):
                by_kw.setdefault(norm_name(e['kw']), e)
            if e.get('part'):
                by_kw.setdefault(norm_name(e['part']), e)
    return idx, by_kw


def lookup(qidx, by_kw, qkey, kw):
    e = qidx.get(qkey)
    if e:
        return e
    from dcd_common import norm_name as _n
    n = _n(kw)
    for k, v in by_kw.items():
        if k and n and (k == n or k.startswith(n) or n.startswith(k)) and len(min(k, n, key=len)) >= 2:
            return v
    return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--bundles', default=BUNDLES)
    ap.add_argument('--overwrite', action='store_true')
    a = ap.parse_args()

    files = sorted(glob.glob(os.path.join(a.bundles, 'dcd_bundle_*.json')))
    if not files:
        print(f'未找到 bundle: {a.bundles}/dcd_bundle_*.json')
        print('(浏览器抓取完成后,把下载目录里的 dcd_bundle_*.json 拷贝到该目录)')
        return 1

    qidx, by_kw = load_qkey_index()
    mapping = json.load(open(MAPPING, encoding='utf-8')) if os.path.exists(MAPPING) else {'reviewed': []}
    have_kw = {str(m.get('kw') or '') for m in mapping['reviewed']}
    fuel_all = json.load(open(FUEL, encoding='utf-8')) if os.path.exists(FUEL) else {}

    stats = Counter()
    new_entries = []
    details = []
    seen_qkey = set()

    for fp in files:
        try:
            bundle = json.load(open(fp, encoding='utf-8'))
        except Exception as e:
            stats['bundle_parse_error'] += 1
            print(f'  bundle 解析失败 {os.path.basename(fp)}: {e}')
            continue
        for s in bundle.get('series') or []:
            qkey = s.get('qkey')
            if qkey and qkey in seen_qkey:
                stats['dup_qkey_skipped'] += 1
                continue
            if qkey:
                seen_qkey.add(qkey)
            kw = s.get('kw') or ''
            sid = s.get('sid')
            if not sid or s.get('error'):
                stats['no_data'] += 1
                details.append({'kw': kw, 'sid': sid, 'status': 'no_data',
                                'reason': s.get('error') or '搜索未定链'})
                continue
            # 1) 落盘 dcd_s<sid>.json(既有格式:series 为 {sid,name,brand} 字典,cars 为款型列表)
            out_fp = os.path.join(RAW_DCD, f'dcd_s{sid}.json')
            if os.path.exists(out_fp) and not a.overwrite:
                stats['exists_skipped'] += 1
            else:
                json.dump({'series': {'sid': str(sid), 'name': s.get('series_name') or ''},
                           'cars': s.get('cars') or []},
                          open(out_fp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
                stats['written'] += 1
            # 2) 合并能源类型
            if s.get('fuel'):
                fuel_all[str(sid)] = s['fuel']
                stats['fuel_merged'] += 1
            # 3) 品牌校验 → mapping
            e = lookup(qidx, by_kw, qkey, kw)
            cands = s.get('candidates') or []
            r = decide_series(kw, cands, e.get('expect_brand') or [], bool(e.get('shared_name')))
            entry = {'kw': kw, 'sid': str(sid), 'series': s.get('series_name') or '',
                     'ok': r['status'] == 'accept', 'review': 'auto',
                     'reason': r['reason'], 'qkey': qkey,
                     'note': f"浏览器直抓 {s.get('captured_at', '')}"}
            if kw in have_kw:
                stats['mapping_dup_skipped'] += 1
            else:
                mapping['reviewed'].append(entry)
                have_kw.add(kw)
                new_entries.append(entry)
            stats[f"decide_{r['status']}"] += 1
            details.append({'kw': kw, 'sid': sid, 'status': r['status'],
                            'series': s.get('series_name'), 'reason': r['reason'],
                            'cars': len(s.get('cars') or [])})

    if new_entries or stats['written']:
        json.dump(mapping, open(MAPPING, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    if fuel_all:
        json.dump(fuel_all, open(FUEL, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    json.dump({'at': __import__('datetime').datetime.now().isoformat(timespec='seconds'),
               'bundles': [os.path.basename(f) for f in files], 'stats': dict(stats),
               'details': details}, open(REPORT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    print(f"bundle 文件 {len(files)} 个")
    print(f"  数据文件落盘 {stats['written']}(已存在跳过 {stats['exists_skipped']})")
    print(f"  能源类型合并 {stats['fuel_merged']}")
    print(f"  品牌校验:accept={stats['decide_accept']} reject={stats['decide_reject']} "
          f"hold={stats['decide_hold']}")
    print(f"  mapping 新增 {len(new_entries)} 条(review=auto)")
    if stats['no_data']:
        print(f"  ⚠ 无数据 {stats['no_data']} 条(搜索未定链或参数页异常)")
    print(f"报告: {REPORT}")
    rej = [d for d in details if d['status'] in ('reject', 'hold')]
    if rej:
        print('\n需人工复核(品牌校验未过,数据已落盘但不会回填):')
        for d in rej[:12]:
            print(f"  {d['kw']:<16} sid={d['sid']} {d['status']:<7} {d['reason'][:50]}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
