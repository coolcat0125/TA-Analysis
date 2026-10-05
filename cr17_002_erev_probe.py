#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_002_erev_probe.py — CR-17-002 增程(EREV)标记：关键词通道可辩护性探查（只读）

CR-17-002 要求：`powertrain` 只有 BEV/PHEV，增程车全部计入 PHEV，需派生 is_erev。
本脚本先回答「关键词通道能不能只用底表原文自证」，不写库。

已知陷阱（实测）：
  - 'AIREV' 是五菱通用名 `Airev`（纯电动轿车）→ 单词 'REV' 子串匹配会误判 9 行 BEV。
    故 REV 必须作独立词/带分隔符匹配，不能用裸 in。
  - '增程' 命中 285 行，全部 PHEV（含 1 行同时含 REEV）。
  - 另有 9 行 BEV 含 'REEV'（待查是否为真增程或命名巧合）。

通道设计（强→弱，全部只补'真值'、不打非增程）：
  E1 名称含 '增程' 关键词                     → 公告原文字面证据，最强
  E2 名称含独立词 REEV/EREV/REV-（非 AIREV）→ 英文型字面证据
  E3 物理参数派生：发动机排量>0 且 纯电续航里程>0 且 综合油耗==0/空
     （增程特征：发动机只发电不驱动，油耗表缺位）→ 推导级，需前两级无命中时作候选
输出：audit-output/cr17_002_erev_probe_20261005.json
"""
import json, os, re, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OUT = os.path.join(HERE, 'audit-output', 'cr17_002_erev_probe_20261005.json')

I_B, I_M, I_MN, I_PN, I_GN, I_PT = 0, 1, 4, 5, 6, 9
I_CURB, I_WB, I_RANGE, I_FUEL, I_BFUEL, I_CC, I_PWR = 11, 12, 13, 23, 24, 25, 26
# 独立词：REV/REEV/EREV 必须词边界，且排除 AIREV
RE_EN = re.compile(r'(?<![A-Za-z])EREV(?![A-Za-z])|(?<![A-Za-z])REEV(?![A-Za-z])|'
                   r'(?<![A-Za-z])REV-(?![A-Za-z])')


def s(v):
    return '' if v is None else str(v).strip()


def num(v):
    try:
        f = float(str(v).strip())
        return f if f == f else None
    except Exception:
        return None


def main():
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0] is not None]

    e1, e2, e3c, noterev, false_pos = [], [], [], [], []
    for idx, r in enumerate(rows, start=2):
        pt = s(r[I_PT])
        blob = f"{s(r[I_MN])} {s(r[I_PN])} {s(r[I_GN])}"
        up = blob.upper()
        rec = {'row': idx, 'id': f"{s(r[I_B])}::{s(r[I_M])}", 'powertrain': pt,
               'modelName': s(r[I_MN]), 'genericName': s(r[I_GN]),
               'enterprise': s(r[3]), 'brand': s(r[2]),
               'engineCc': num(r[I_CC]), 'evRangeKm': num(r[I_RANGE]),
               'fuel': num(r[I_FUEL]), 'bFuel': num(r[I_BFUEL])}
        if 'AIREV' in up:
            false_pos.append(rec)
        en = RE_EN.findall(up)
        if '增程' in blob:
            e1.append({**rec, 'marker': '增程'})
        elif en:
            e2.append({**rec, 'marker': '/'.join(sorted(set(en)))})
        elif pt == 'PHEV' and (rec['engineCc'] or 0) > 0 and (rec['evRangeKm'] or 0) > 0 \
                and not rec['fuel'] and not rec['bFuel']:
            e3c.append({**rec, 'marker': '物理参数派生'})
        elif pt == 'BEV' and en:
            noterev.append(rec)

    out = {'generated': '2026-10-05',
           'totals': {'rows': len(rows),
                      'BEV': sum(1 for r in rows if s(r[I_PT]) == 'BEV'),
                      'PHEV': sum(1 for r in rows if s(r[I_PT]) == 'PHEV'),
                      'other_pt': sorted({s(r[I_PT]) for r in rows} - {'BEV', 'PHEV'})},
           'E1_name_cn': {'n': len(e1), 'pt': dict(collections.Counter(x['powertrain'] for x in e1)),
                          'sample': e1[:8]},
           'E2_name_en': {'n': len(e2), 'pt': dict(collections.Counter(x['powertrain'] for x in e2)),
                          'sample': e2[:8]},
           'E3_param_candidate': {'n': len(e3c),
                                  'pt': dict(collections.Counter(x['powertrain'] for x in e3c)),
                                  'sample': e3c[:8]},
           'airev_false_positive': {'n': len(false_pos), 'sample': false_pos[:3]},
           'rows': {'E1': e1, 'E2': e2, 'E3': e3c}}
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(json.dumps(out['totals'], ensure_ascii=False))
    print('E1 含"增程"', len(e1), out['E1_name_cn']['pt'])
    print('E2 英文独立词', len(e2), out['E2_name_en']['pt'])
    print('E3 物理派生候选', len(e3c), out['E3_param_candidate']['pt'])
    print('AIREV 误判陷阱', len(false_pos), '行（BEV，不作增程）')
    for x in e3c[:6]:
        print('  E3', x['id'], x['enterprise'][:14], x['modelName'], '/', x['genericName'],
              '| 排量', x['engineCc'], '续航', x['evRangeKm'])
    print('->', OUT)


if __name__ == '__main__':
    main()