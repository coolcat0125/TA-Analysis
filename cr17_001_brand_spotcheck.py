#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_spotcheck.py — 抽查验收（应 15 号 CR-17-TA-analysis.md:86 请求）

抽 N 条 15 号回填台账行，逐条比对「公告原文非空品牌格（同型号跨批 / 同企业同族）」，
检验 15 号映射表推导在公告原文侧是否成立；并对本次拟回填的 607 格做同口径抽验。
输出：audit-output/cr17_001_brand_spotcheck_20261005.json
"""
import json, os, re, random, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
LEDGER = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                      'db', 'brand_backfill_ledger_2026-10-04.json')
APPLY = os.path.join(HERE, 'audit-output', 'cr17_001_brand_apply_20261005.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_spotcheck_20261005.json')

I_B, I_M, I_BR, I_ENT, I_GN = 0, 1, 2, 3, 6
RND = random.Random(17001)


def core(b):
    t = re.sub(r'[（(].*?[)）]', '', str(b or '').strip())
    return t[:-1] if t.endswith('牌') else t


def s(v):
    return '' if v is None else str(v).strip()


def main():
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    model_bc = collections.defaultdict(collections.Counter)
    ent_fam = collections.defaultdict(collections.Counter)
    ent_bc = collections.defaultdict(collections.Counter)
    by_id = {}
    for idx, r in enumerate(rows, start=2):
        rid = f"{s(r[I_B])}::{s(r[I_M])}"
        by_id.setdefault(rid, (idx, r))
        b = s(r[I_BR])
        if not b:
            continue
        c = core(b)
        model_bc[s(r[I_M])][c] += 1
        ent_fam[(s(r[I_ENT]), s(r[I_GN]))][c] += 1
        ent_bc[s(r[I_ENT])][c] += 1

    def baseline(rid):
        idx, r = by_id[rid]
        m, e, gn = s(r[I_M]), s(r[I_ENT]), s(r[I_GN])
        for label, cnt in (('同型号跨批', model_bc.get(m)), ('同企业同通用名称', ent_fam.get((e, gn))),
                           ('同企业全部', ent_bc.get(e))):
            cnt = {k: v for k, v in (cnt or {}).items() if v}
            if len(cnt) == 1:
                c = next(iter(cnt))
                return {'source': label, 'core': c, 'n': cnt[c]}
        return None

    ledger = json.load(open(LEDGER, encoding='utf-8'))['rows']
    sample = RND.sample(ledger, 40)
    checked, verdict = [], collections.Counter()
    for lr in sample:
        rid = lr['id']
        bl = baseline(rid)
        d = core(lr.get('brand'))
        if not bl:
            v = 'NO_BASELINE'
        elif bl['core'] == d:
            v = 'CONFIRM'
        else:
            v = 'DIFFERS'
        verdict[v] += 1
        idx, r = by_id.get(rid, (None, ('',) * 38))
        checked.append({'id': rid, 'enterprise': lr.get('enterprise'), 'modelName': lr.get('modelName'),
                        'ledger_brand': lr.get('brand'), 'ledger_source': lr.get('source'),
                        'baseline': bl, 'verdict': v, 'row': idx,
                        'wb_modelName': s(r[I_ENT]) if idx else None})

    plan = json.load(open(APPLY, encoding='utf-8'))['plan']
    pchk = []
    for p in RND.sample(plan, min(40, len(plan))):
        bl = baseline(p['id'])
        v = ('WRITE_CONFIRM' if bl and bl['core'] == p['evidence_core']
             else ('NO_BASELINE' if not bl else 'WRITE_WITHOUT_BASELINE'))
        pchk.append({'id': p['id'], 'channel': p['channel'], 'write_value': p['evidence_raw'],
                     'derived': p['derived'], 'baseline': bl, 'verdict': v})
        verdict['plan_' + v] += 1

    out = {'generated': '2026-10-05', 'ledger_sample_n': len(sample),
           'plan_sample_n': len(pchk), 'verdict_counts': dict(verdict),
           'ledger_checks': checked, 'plan_checks': pchk}
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('台账抽验', len(sample), '条：', {k: v for k, v in verdict.items() if not k.startswith('plan_')})
    print('拟回填抽验', len(pchk), '条：', {k: v for k, v in verdict.items() if k.startswith('plan_')})
    print('--- 15号台账 CONFIRM 样例 10 条 ---')
    for c in [c for c in checked if c['verdict'] == 'CONFIRM'][:10]:
        print(f"  {c['id']:34s} {c['enterprise'][:18]:18s} 台账={c['ledger_brand']:8s} 原文基线={c['baseline']['source']}/{c['baseline']['n']}")
    print('--- DIFFERS 样例 ---')
    for c in [c for c in checked if c['verdict'] == 'DIFFERS'][:12]:
        print(f"  {c['id']:34s} 台账={c['ledger_brand']:8s} 原文基线={c['baseline']['core']}({c['baseline']['source']}/{c['baseline']['n']})")
    print('->', OUT)


if __name__ == '__main__':
    main()