#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_004_005_coverage.py — CR-17-004/005 覆盖率盘点（只读，不写库）

CR-17-004（电池包供应商）：验收 TOP10 品牌覆盖 ≥60%，或给出不可得结论与依据。
CR-17-005（官方参数页URL）：验收覆盖率补至 ≥90%，抽查 20 条 URL 可访问率 ≥85%。

盘点口径：
  - 覆盖率 = 非空值数 / 5,392（与 verify_consistency/audit_data 同族）
  - URL 可用性：本脚本只做**结构体检**（http(s) 前缀 / 去重 / 域名分布 / 与媒体URL 重复度），
    真实外网可访问率需另行发起网络核验，此处不臆测。
输出：audit-output/cr17_004_005_coverage_20261005.json
"""
import json, os, re, collections
from urllib.parse import urlparse
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OUT = os.path.join(HERE, 'audit-output', 'cr17_004_005_coverage_20261005.json')

I_BR, I_ENT, I_CELL, I_PACK, I_URL, I_MEDIA = 2, 3, 31, 32, 33, 34
CR004_TOP10 = ['比亚迪', '特斯拉', '埃安', '吉利', '长安', '奇瑞', '理想', '蔚来',
               '小鹏', '上汽']


def s(v):
    return '' if v is None else str(v).strip()


def core(b):
    t = re.sub(r'[（(].*?[)）]', '', str(b or '').strip())
    return t[:-1] if t.endswith('牌') else t


def urls_of(v):
    return [u for u in re.split(r'[;\s]+', s(v)) if u.startswith('http')]


def main():
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    n = cell = pack = url = 0
    pack_top10 = collections.Counter()
    pack_top10_n = collections.Counter()
    pack_values = collections.Counter()
    url_domains = collections.Counter()
    url_bad = []
    url_dup_with_media = 0
    url_sample = []
    seen_urls = collections.Counter()
    ent_missing_pack = collections.Counter()

    for idx, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if r[0] is None:
            continue
        n += 1
        b, ent = core(s(r[I_BR])), s(r[I_ENT])
        if s(r[I_CELL]):
            cell += 1
        pv = s(r[I_PACK])
        if pv:
            pack += 1
            pack_values[pv] += 1
        for t in CR004_TOP10:
            if t in b or t in ent:
                pack_top10_n[t] += 1
                if pv:
                    pack_top10[t] += 1
        u = s(r[I_URL])
        if u:
            url += 1
            us = urls_of(u)
            seen_urls[u] += 1
            for x in us:
                try:
                    url_domains[urlparse(x).netloc.lower()] += 1
                except Exception:
                    url_bad.append({'row': idx, 'url': x, 'issue': '不可解析'})
            if not u.startswith(('http://', 'https://')):
                url_bad.append({'row': idx, 'url': u[:120], 'issue': '非 http(s) 前缀'})
            if set(us) & set(urls_of(s(r[I_MEDIA]))):
                url_dup_with_media += 1
            if len(url_sample) < 25:
                url_sample.append({'row': idx, 'id': f"{s(r[0])}::{s(r[1])}", 'url': u[:160]})
        if not pv:
            ent_missing_pack[s(r[I_ENT])] += 1

    top10_cov = {t: {'n': pack_top10_n[t], 'covered': pack_top10[t],
                     'rate': round(pack_top10[t] / pack_top10_n[t] * 100, 1) if pack_top10_n[t] else None}
                 for t in CR004_TOP10}
    tot10_n = sum(pack_top10_n.values())
    tot10_c = sum(pack_top10.values())
    out = {
        'generated': '2026-10-05', 'records': n,
        'CR004_packSupplier': {
            'nonempty': pack, 'coverage': round(pack / n * 100, 2),
            'cellSupplier_nonempty': cell, 'cell_coverage': round(cell / n * 100, 2),
            'top10_brands': top10_cov,
            'top10_aggregate': {'n': tot10_n, 'covered': tot10_c,
                                'rate': round(tot10_c / tot10_n * 100, 1) if tot10_n else None,
                                'gate_60pct': (tot10_c / tot10_n >= 0.6) if tot10_n else None},
            'distinct_values': len(pack_values),
            'top_values': dict(pack_values.most_common(15)),
            'top_gap_enterprises': dict(ent_missing_pack.most_common(12)),
        },
        'CR005_officialUrl': {
            'nonempty': url, 'coverage': round(url / n * 100, 2),
            'gate_90pct': url / n >= 0.90,
            'distinct_urls': len(seen_urls),
            'duplicated_urls': {k: v for k, v in seen_urls.most_common(8) if v > 1},
            'malformed': url_bad[:20], 'malformed_count': len(url_bad),
            'domains': dict(url_domains.most_common(15)),
            'overlaps_media_urls': url_dup_with_media,
            'sample': url_sample,
        },
    }
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('记录', n)
    c4 = out['CR004_packSupplier']
    print(f"CR-004 电池包供应商: 非空 {c4['nonempty']} ({c4['coverage']}%)  电芯 {c4['cellSupplier_nonempty']} ({c4['cell_coverage']}%)")
    print('  TOP10 合计', c4['top10_aggregate'])
    for t, v in top10_cov.items():
        print(f"    {t:6s} {v['covered']}/{v['n']} = {v['rate']}%")
    c5 = out['CR005_officialUrl']
    print(f"CR-005 官方URL: 非空 {c5['nonempty']} ({c5['coverage']}%) 门槛90%={c5['gate_90pct']}  去重 {c5['distinct_urls']}  畸形 {c5['malformed_count']}")
    print('  域名Top:', list(c5['domains'].items())[:8])
    print('  与媒体URL重复行', c5['overlaps_media_urls'])
    print('->', OUT)


if __name__ == '__main__':
    main()