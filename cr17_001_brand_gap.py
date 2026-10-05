#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_gap.py — 902 空品牌格的结构性画像：为何四通道都拿不到证据（只读）

回答三个问题：
  Q1 无证据的 318 行（HOLD）里，多少行本身 企业名称 为空？
  Q2 这些行的 产品型号 前缀（MIIT 厂标代号，如 BYD/DFM/SQR/QCJ）能否在表内找到已知品牌的旁证？
  Q3 15号台账对这些行的推导值是什么（是否被我的 id 对齐漏掉）？
输出：audit-output/cr17_001_brand_gap_20261005.json
"""
import json, os, re, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
LEDGER = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                      'db', 'brand_backfill_ledger_2026-10-04.json')
APPLY = os.path.join(HERE, 'audit-output', 'cr17_001_brand_apply_20261005.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_gap_20261005.json')

I_B, I_M, I_BR, I_ENT, I_MN, I_PN, I_GN = 0, 1, 2, 3, 4, 5, 6


def core(b):
    t = re.sub(r'[（(].*?[)）]', '', str(b or '').strip())
    return t[:-1] if t.endswith('牌') else t


def s(v):
    return '' if v is None else str(v).strip()


def main():
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    holds = json.load(open(APPLY, encoding='utf-8'))['hold']
    hold_ids = {h['id'] for h in holds}
    ledger = {r['id']: r for r in json.load(open(LEDGER, encoding='utf-8'))['rows']}

    # 型号前缀 -> 已知牌核心（只用非空品牌行建旁证）
    pref = collections.defaultdict(collections.Counter)
    raw_of = {}
    empties = []
    for idx, r in enumerate(rows, start=2):
        b, m, e = s(r[I_BR]), s(r[I_M]), s(r[I_ENT])
        rid = f"{s(r[I_B])}::{m}"
        if not b:
            empties.append({'row': idx, 'id': rid, 'model': m, 'enterprise': e,
                            'modelName': s(r[I_MN]), 'productName': s(r[I_PN]),
                            'genericName': s(r[I_GN])})
            continue
        c = core(b); raw_of.setdefault(c, b)
        mm = re.match(r'^([A-Za-z]{2,4})', m or '')
        if mm:
            pref[mm.group(1).upper()][c] += 1

    q1 = collections.Counter()
    q2 = {}
    q3 = collections.Counter()
    for rec in empties:
        if rec['id'] not in hold_ids:
            continue
        q1['ent_empty' if not rec['enterprise'] else 'ent_present'] += 1
        lr = ledger.get(rec['id'])
        q3['ledger_hit_brand' if (lr and lr.get('brand')) else
           ('ledger_hit_nobrand' if lr else 'ledger_miss')] += 1
        mm = re.match(r'^([A-Za-z]{2,4})', rec['model'] or '')
        if mm:
            key = mm.group(1).upper()
            cand = pref.get(key) or {}
            if len(cand) == 1:
                c = next(iter(cand))
                q2.setdefault(key, {'brand_core': c, 'brand_raw': raw_of.get(c, c + '牌'),
                                    'support_n': cand[c], 'rows': []})
                q2[key]['rows'].append(rec['id'])
            elif len(cand) > 1:
                q2.setdefault(key, {'brand_core': None, 'multi': dict(cand), 'rows': []})
                q2[key]['rows'].append(rec['id'])
            else:
                q2.setdefault(key, {'brand_core': None, 'multi': {}, 'rows': []})
                q2[key]['rows'].append(rec['id'])

    out = {
        'generated': '2026-10-05',
        'holds': len(holds),
        'q1_enterprise_presence': dict(q1),
        'q3_ledger_alignment': dict(q3),
        'q2_model_prefix_evidence': q2,
        'q2_prefix_summary': {
            'keys': len(q2),
            'single_brand_keys': sum(1 for v in q2.values() if v.get('brand_core')),
            'multi_or_none_keys': sum(1 for v in q2.values() if not v.get('brand_core')),
            'rows_fillable_by_prefix': sum(len(v['rows']) for v in q2.values() if v.get('brand_core')),
        },
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps({k: v for k, v in out.items() if k != 'q2_model_prefix_evidence'}, ensure_ascii=False, indent=1))
    for k, v in sorted(q2.items(), key=lambda kv: -len(kv[1]['rows'])):
        print(f"  {k:6s} rows={len(v['rows']):4d} -> {v.get('brand_raw') or ('多候选:' + str(v.get('multi')))}")
    print('->', OUT)


if __name__ == '__main__':
    main()