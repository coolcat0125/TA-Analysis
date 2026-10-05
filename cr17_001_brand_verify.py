#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_verify.py — CR-17-001 品牌回填的可辩护性核验（只读，不写库）

背景：15 OEM Portfolio 已对派生库 master_vehicles.json 完成品牌回填（902→5 空），
CR 回执请求 17A 侧「抽查台账任意行比对公告原文」。本脚本以**底表 xlsx 自身的
公告原文值**为基准做交叉核验，不依赖任何外部推导：

1) 企业内公告牌分布：同一「企业名称」下已有 非空 产品商标（公告原文）的行，按牌统计；
   这是库内权威证据（值来自公告附件，非推导）。
2) 牌名规范化：公告原文存在「极狐牌」与「极狐(ARCFOX)牌」、「别克牌」与
   「别克(BUICK)牌」等同牌异写，须先按去括号/去「牌」后缀的核心名归一，否则同牌判成冲突。
3) 分层裁决：
   - 企业内核心牌唯一 → 派生值核心名与之相同 = KEEP；不同 = CONFLICT（多牌集团，需车型名仲裁）；
   - 派生值核心名已出现在该企业的公告牌集合中 → 候选合法，用车型名/通用名称含牌词仲裁
     → KEEP_NAME 或 CONFLICT_NAME；
   - 派生值核心名不在该企业任何公告牌中 → CONFLICT_UNSUPPORTED（推导与企业主体无关）；
   - 该企业无任何公告牌基线 → NO_BASELINE（无库内证据）。
4) 5 条 unresolved：无判定依据，保留为空（不擅改）。

输出：audit-output/cr17_001_brand_verify_20261005.json
用法：python cr17_001_brand_verify.py
"""
import json, os, re, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
LEDGER = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                      'db', 'brand_backfill_ledger_2026-10-04.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_verify_20261005.json')


def core(b):
    """牌名核心化：去括号英文、去尾「牌」，用于同牌异写归一。"""
    t = str(b or '').strip()
    t = re.sub(r'[（(].*?[)）]', '', t)
    return t[:-1] if t.endswith('牌') else t


def main():
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    def s(v):
        return '' if v is None else str(v).strip()
    I_BATCH, I_MODEL, I_BRAND, I_ENT, I_MNAME, I_PNAME, I_GNAME = 0, 1, 2, 3, 4, 5, 6

    # 企业 -> 公告原文牌（按核心名归并，保留原写法样例）
    ent_brand = collections.defaultdict(collections.Counter)
    for r in rows:
        e, b = s(r[I_ENT]), s(r[I_BRAND])
        if e and b:
            ent_brand[e][core(b)] += 1

    ledger = json.load(open(LEDGER, encoding='utf-8'))
    lrows = ledger['rows']
    lmap = {(r['id']): r for r in lrows}

    # 底表行按 id 索引
    byid = {}
    dup = []
    for idx, r in enumerate(rows, start=2):
        rid = f"{s(r[I_BATCH])}::{s(r[I_MODEL])}"
        if rid in byid:
            dup.append(rid)
        byid.setdefault(rid, (idx, r))

    agree, conflict, unbacked, unresolved = [], [], [], []
    for rid, lr in lmap.items():
        if rid not in byid:
            unbacked.append({'id': rid, 'note': '底表无对应行（跨批同码/批次差）'})
            continue
        idx, r = byid[rid]
        ent = s(r[I_ENT])
        derived = (lr.get('brand') or '').strip()
        dcore = core(derived)
        src = lr.get('source') or ''
        base = ent_brand.get(ent, collections.Counter())
        names = ' '.join([s(r[I_MNAME]), s(r[I_PNAME]), s(r[I_GNAME])])
        rec = {'row': idx, 'id': rid, 'enterprise': ent, 'model': s(r[I_MODEL]),
               'modelName': s(r[I_MNAME]), 'derived': derived, 'derived_core': dcore,
               'source': src, 'base_core': list(base), 'base_all': dict(base)}
        if src == 'unresolved' or not derived:
            rec['verdict'] = 'UNRESOLVED'
            unresolved.append(rec)
            continue
        if not base:
            rec['verdict'] = 'NO_BASELINE'
            unbacked.append(rec)
            continue
        if len(base) == 1:
            (b0, n0), = base.most_common()
            if dcore == b0:
                rec.update(verdict='KEEP_CONSISTENT', note=f'企业唯一公告牌={b0}({n0}行)')
                agree.append(rec)
            else:
                rec.update(verdict='CONFLICT_UNSUPPORTED',
                           note=f'企业唯一公告牌={b0}({n0}行)，派生={dcore} 不在其列')
                conflict.append(rec)
            continue
        if dcore in base:
            hit = dcore in names
            rec.update(verdict=('KEEP_NAME' if hit else 'CONFLICT_NAME'),
                       note=f'多牌企业，派生核心={dcore} 在公告牌集合内，车型名含牌词={hit}')
            (agree if hit else conflict).append(rec)
            continue
        rec.update(verdict='CONFLICT_UNSUPPORTED',
                   note=f'多牌企业但派生核心={dcore} 不在公告牌集合 {list(base)}')
        conflict.append(rec)

    byv = collections.Counter()
    for g in (agree, conflict, unbacked, unresolved):
        for x in g:
            byv[x['verdict']] += 1
    out = {
        'generated': '2026-10-05',
        'workbook_rows': len(rows),
        'ledger_rows': len(lrows),
        'verdict_counts': dict(byv),
        'summary': {
            'defensible_keep': len(agree),
            'needs_review': len(conflict),
            'no_baseline_enterprise': sum(1 for x in unbacked if x.get('verdict') == 'NO_BASELINE'),
            'unresolved': len(unresolved),
            'dup_ids': len(dup),
        },
        'conflict': conflict,
        'unresolved': unresolved,
        'no_baseline': unbacked,
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(out['verdict_counts'], ensure_ascii=False))
    print(json.dumps(out['summary'], ensure_ascii=False))
    for c in conflict[:30]:
        print('CONFLICT', c.get('row'), c['id'], c['enterprise'], '|', c.get('note'),
              '|', c.get('modelName', ''))
    for u in unresolved[:10]:
        print('UNRESOLVED', u['id'], u.get('enterprise', ''), '|', u.get('modelName', ''))
    print('->', OUT)


if __name__ == '__main__':
    main()