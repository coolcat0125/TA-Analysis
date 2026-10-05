#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_consensus.py — CR-17-001 产品商标回填的**公告原文共识**通道（只读）

背景：底表「产品商标」902 格为空；派生库侧的映射表回填（企业→牌）属推导证据，
但本项目已有更强的证据层：**底表内同一「通用名称/车型名称」家族的非空公告原值**
（fill_consensus_v5 的通用名称共识通道，与只补空/唯一共识纪律一致）。
本脚本测量三层证据可得性：
  T1 同族唯一共识：同 通用名称（或 车型名称）下只有一个非空 牌核心 → 公告原文级证据；
  T2 同型号多牌：同族有多个牌核心 → 多配置歧义，不补；
  T3 无同族证据 → 留给映射表推导（derived）。
输出：audit-output/cr17_001_brand_consensus_20261005.json
用法：python cr17_001_brand_consensus.py
"""
import json, os, re, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
LEDGER = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                      'db', 'brand_backfill_ledger_2026-10-04.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_consensus_20261005.json')


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

    # 家族 -> 牌核心计数（只用本来非空的行建立共识）
    fam = collections.defaultdict(collections.Counter)
    fam_raw = {}
    for r in rows:
        b = s(r[I_BR])
        if not b:
            continue
        c = core(b)
        for key, tag in ((s(r[I_GN]), 'generic'), (s(r[I_MN]), 'model')):
            if not key:
                continue
            k = f'{tag}::{key}'
            fam[k][c] += 1
            fam_raw.setdefault(c, b)

    ledger = json.load(open(LEDGER, encoding='utf-8'))
    lmap = {r['id']: r for r in ledger['rows']}
    byid = {}
    for idx, r in enumerate(rows, start=2):
        byid.setdefault(f"{s(r[I_B])}::{s(r[I_M])}", (idx, r))

    t1, t2, t3, t1_conflict = [], [], [], []
    for rid, lr in lmap.items():
        if rid not in byid:
            continue
        idx, r = byid[rid]
        ent, mn, gn = s(r[I_ENT]), s(r[I_MN]), s(r[I_GN])
        derived_core = core(lr.get('brand'))
        rec = {'row': idx, 'id': rid, 'enterprise': ent, 'modelName': mn, 'genericName': gn,
               'derived': lr.get('brand'), 'derived_core': derived_core,
               'ledger_source': lr.get('source')}
        # 通用名称优先，其次车型名称
        for key, tag in ((f'generic::{gn}', '通用名称'), (f'model::{mn}', '车型名称')):
            if not key.split('::', 1)[1]:
                continue
            cnt = fam.get(key)
            if not cnt:
                continue
            if len(cnt) == 1:
                c, n = cnt.most_common()[0]
                raw = fam_raw.get(c, c + '牌')
                rec.update(tier='T1', family_key=tag, family_value=key.split('::', 1)[1],
                           consensus_core=c, consensus_raw=raw, consensus_n=n)
                if c == derived_core:
                    t1.append(rec)          # 推导与公告共识一致
                else:
                    t1_conflict.append(rec)  # 公告共识优先，映射表推导错
                break
            rec.update(tier='T2', family_key=tag, family_value=key.split('::', 1)[1],
                       candidates={k: v for k, v in cnt.most_common()})
            t2.append(rec)
            break
        else:
            rec.update(tier='T3')
            t3.append(rec)

    out = {
        'generated': '2026-10-05',
        'workbook_rows': len(rows),
        'ledger_rows': len(lrows := ledger['rows']),
        'tier_counts': {'T1_agrees_with_derived': len(t1),
                        'T1_conflicts_with_derived': len(t1_conflict),
                        'T2_ambiguous_family': len(t2),
                        'T3_no_family_evidence': len(t3)},
        'recommended_writeback': len(t1) + len(t1_conflict),
        't1_conflicts': t1_conflict,
        't2': t2,
        't3': t3,
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(out['tier_counts'], ensure_ascii=False))
    print('可回填(公告共识) =', out['recommended_writeback'])
    for c in t1_conflict[:30]:
        print('CONFLICT', c['row'], c['id'], c['enterprise'], '| 共识=', c['consensus_raw'],
              '| 推导=', c['derived'], '|', c.get('modelName', ''), '/', c.get('genericName', ''))
    print('->', OUT)


if __name__ == '__main__':
    main()