#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_decide.py — CR-17-001 回填裁决：四通道证据 + 别名表 = 最终可写值（只读）

输入：cr17_001_brand_channels_20261005.json（通道归集）+ media_fill.BRAND_ALIAS（项目自有品牌别名权威）
裁决规则（只补空、唯一值；疑点不擅改）：
  - 证据通道 C1（同型号跨批）/C2（同企业同通用名称）= 公告原文级，可写；
  - C3（同企业单牌）/C4（名称牌词）= 推导级，仅当与 15 号映射表推导**别名等价**时视为双源可辩护，可写；
  - 证据与推导实质冲突（核心名互不含、非别名）→ 不写，列疑点待裁。
输出：audit-output/cr17_001_brand_decide_20261005.json（逐格台账 + 统计）
用法：python cr17_001_brand_decide.py
"""
import json, os, sys, collections

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from media_fill import BRAND_ALIAS  # noqa: E402

CH = os.path.join(HERE, 'audit-output', 'cr17_001_brand_channels_20261005.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_decide_20261005.json')

# 别名关系（含 BRAND_ALIAS 的反向对）
ALIAS = set()
for a, b in BRAND_ALIAS:
    ALIAS.add((a, b)); ALIAS.add((b, a))


def core(b):
    import re
    t = re.sub(r'[（(].*?[)）]', '', str(b or '').strip())
    return t[:-1] if t.endswith('牌') else t


def alias_equiv(a, b):
    """核心名等价：相等 / 互含 / 显式别名关系（含父集团品牌↔子品牌）。"""
    if not a or not b:
        return False
    if a == b or a in b or b in a:
        return True
    return (a, b) in ALIAS or (b, a) in ALIAS


def main():
    ch = json.load(open(CH, encoding='utf-8'))
    rows = ch['rows']
    keep, doubt = [], []
    for r in rows:
        e, d = r['evidence_core'], r['derived_core']
        if r['channel'] in ('C1', 'C2'):
            verdict, note = 'WRITE', '公告原文同族/同型号证据'
        elif r['channel'] in ('C3', 'C4'):
            if alias_equiv(e, d):
                verdict, note = 'WRITE', '推导级证据+映射表推导别名等价（双源）'
            else:
                verdict, note = 'HOLD', '推导级证据与映射表推导冲突，疑点待裁'
        if not r['derived'] and r['channel'] in ('C1', 'C2'):
            verdict = 'WRITE'
        (keep if verdict == 'WRITE' else doubt).append({**r, 'verdict': verdict, 'note': note})

    stats = {
        'candidate_total': len(rows),
        'write_back': len(keep),
        'hold_doubt': len(doubt),
        'write_by_channel': dict(collections.Counter(r['channel'] for r in keep)),
        'hold_by_channel': dict(collections.Counter(r['channel'] for r in doubt)),
    }
    out = {'generated': '2026-10-05', 'rule': 'C1/C2=公告级直写; C3/C4需与推导别名等价; 冲突=疑点待裁',
           'stats': stats, 'write_rows': keep, 'doubt_rows': doubt}
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(stats, ensure_ascii=False))
    print('写出', len(keep), '/ 待裁', len(doubt))
    for r in doubt[:25]:
        print('HOLD', r['channel'], r['id'], r['enterprise'], '| 证据=', r['evidence_raw'],
              '| 推导=', r['derived'], '|', r['modelName'], '/', r['genericName'])
    print('->', OUT)


if __name__ == '__main__':
    main()