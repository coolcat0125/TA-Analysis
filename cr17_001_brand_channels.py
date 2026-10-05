#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_channels.py — 902 空「产品商标」格的**多通道证据可得性**诊断（只读）

按证据强度递增排列四通道，测量每一通道的可回填量（全部只补空、唯一值）：
  C1 同型号跨批：同「产品型号」在其他批次已有非空公告原值（最强，型号=车唯一标识）
  C2 同企业同族：同「企业名称」+ 同「通用名称」下只有一个非空牌核心
  C3 同企业单牌：同「企业名称」全表只有一个非空牌核心
  C4 名称牌词：车型/通用名称含该企业已用过的牌词（推导级，最弱）
按「最高可得通道」归集，统计互斥后的净可回填量，并输出每通道明细。
输出：audit-output/cr17_001_brand_channels_20261005.json
用法：python cr17_001_brand_channels.py
"""
import json, os, re, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
LEDGER = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                      'db', 'brand_backfill_ledger_2026-10-04.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_channels_20261005.json')


def core(b):
    t = re.sub(r'[（(].*?[)）]', '', str(b or '').strip())
    return t[:-1] if t.endswith('牌') else t


def main():
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    def s(v):
        return '' if v is None else str(v).strip()
    I_B, I_M, I_BR, I_ENT, I_MN, I_PN, I_GN = 0, 1, 2, 3, 4, 5, 6

    model_brand = collections.defaultdict(collections.Counter)   # C1
    ent_fam = collections.defaultdict(collections.Counter)      # C2
    ent_brand = collections.defaultdict(collections.Counter)     # C3/C4
    ent_rows = collections.defaultdict(list)
    raw_of = {}
    empties = []
    for idx, r in enumerate(rows, start=2):
        b, m, e = s(r[I_BR]), s(r[I_M]), s(r[I_ENT])
        gn, mn, pn = s(r[I_GN]), s(r[I_MN]), s(r[I_PN])
        rid = f"{s(r[I_B])}::{m}"
        rec = {'row': idx, 'id': rid, 'batch': s(r[I_B]), 'model': m, 'enterprise': e,
               'modelName': mn, 'productName': pn, 'genericName': gn}
        ent_rows[e].append(rec)
        if not b:
            empties.append(rec)
            continue
        c = core(b)
        raw_of.setdefault(c, b)
        if m:
            model_brand[m][c] += 1
        if e and gn:
            ent_fam[(e, gn)][c] += 1
        if e:
            ent_brand[e][c] += 1

    derived = {}
    for r in json.load(open(LEDGER, encoding='utf-8'))['rows']:
        derived[r['id']] = r.get('brand') or ''

    assign = collections.defaultdict(list)
    for rec in empties:
        rid, m, e, gn, mn, pn = rec['id'], rec['model'], rec['enterprise'], rec['genericName'], rec['modelName'], rec['productName']
        names = f'{mn} {pn} {gn}'
        ch = None
        c1 = model_brand.get(m) or {}
        c1 = {k: v for k, v in c1.items() if v}
        if len(c1) == 1:
            ch = 'C1'
        if ch is None and e and gn:
            c2 = ent_fam.get((e, gn)) or {}
            if len(c2) == 1:
                ch = 'C2'
        if ch is None and len(ent_brand.get(e, {})) == 1:
            ch = 'C3'
        if ch is None and e:
            used = set(ent_brand.get(e, {}))
            hits = [k for k in used if k and (k in names)]
            if len(hits) == 1:
                ch = 'C4'
        if ch:
            assign[ch].append(rec)

    def resolve(rec):
        m, e, gn = rec['model'], rec['enterprise'], rec['genericName']
        if len(model_brand.get(m) or {}) == 1:
            return next(iter(model_brand[m]))
        if e and gn and len(ent_fam.get((e, gn)) or {}) == 1:
            return next(iter(ent_fam[(e, gn)]))
        if len(ent_brand.get(e, {})) == 1:
            return next(iter(ent_brand[e]))
        used = set(ent_brand.get(e, {}))
        return [k for k in used if k and k in f"{rec['modelName']} {rec['productName']} {gn}"][0]

    out_rows, disagree = [], []
    for ch, lst in assign.items():
        for rec in lst:
            c = resolve(rec)
            d = core(derived.get(rec['id'], ''))
            row = dict(rec, channel=ch, evidence_core=c, evidence_raw=raw_of.get(c, c + '牌'),
                       derived=derived.get(rec['id'], ''), derived_core=d,
                       agrees_with_derived=(c == d))
            out_rows.append(row)
            if ch in ('C1', 'C2') and d and c != d:
                disagree.append(row)
    unassigned = [r for r in empties if not any(r['id'] in {x['id'] for x in v} for v in assign.values())]

    out = {
        'generated': '2026-10-05',
        'workbook_rows': len(rows),
        'empty_brand_rows': len(empties),
        'channel_counts': {k: len(v) for k, v in sorted(assign.items())},
        'net_fillable': sum(len(v) for v in assign.values()),
        'unassigned': len(unassigned),
        'strong_disagree_with_derivation': len(disagree),
        'rows': sorted(out_rows, key=lambda x: (x['channel'], x['row'])),
        'unassigned_rows': unassigned,
        'strong_disagree': disagree,
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(out['channel_counts'], ensure_ascii=False))
    print('净可回填', out['net_fillable'], '/ 未定', out['unassigned'],
          '｜强证据与映射表推导冲突', out['strong_disagree_with_derivation'])
    for r in disagree[:25]:
        print('DISAGREE', r['channel'], r['id'], r['enterprise'], '| 证据=', r['evidence_raw'],
              '| 推导=', r['derived'], '|', r['modelName'], '/', r['genericName'])
    print('->', OUT)


if __name__ == '__main__':
    main()