#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_002_erev_apply.py — CR-17-002 增程(EREV)标记执行器（默认 dry-run）

口径决定（实测证据支撑，见 audit-output/cr17_002_erev_probe_20261005.json）：
  ✅ 采纳 E1 = 车型/产品/通用名称含「增程」字样 → 285 行，全部 PHEV。
     这是公告原文字面证据，属项目「只补可自证值」纪律下的最强级。
  ❌ 不采纳 E2 英文独立词（REEV/EREV）→ 0 行，无可采纳。
  ❌ 不采纳 E3 物理派生（排量>0 且 纯电续航>0 且 油耗缺位）→ 46 行，
     实测混入大量真插混（比亚迪秦PLUS DM-i 1498cc、大通 DMH、极氪），
     油耗缺位是**公告未填报**而非增程特征 → 不可作判据，弃用。
  ❌ 裸子串 'REV' 禁用 → 通用名 `Airev`（五菱纯电轿车）误命中 9 行 BEV。

写库方式：新增独立列 `是否增程(EREV)`（追加式，为第 39 列；既有 38 列语义不变），
值域 1/空（只打真值，不给非增程打 0，避免把「未判定」与「确定非增程」混同）。
副作用（已核）：底表列数 38 → 39，故仓内所有硬编码「38 列」文案须当轮一并修正
（AGENTS 纪律：改 schema 当晚 grep 全仓修工具）：
  build_deliverables_v4923.py:4,197 / build_mobile_v4935.py:175 /
  build_pptx_refresh_v4924.py:562 / export_workbook_csv_v2.py:2,5 /
  generate_dashboard.py:3247（注释）→ 均为 39 列。
用法：python cr17_002_erev_apply.py [--write]
"""
import json, os, re, time, argparse, collections
import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OUT = os.path.join(HERE, 'audit-output', 'cr17_002_erev_apply_20261005.json')

I_B, I_M, I_MN, I_PN, I_GN, I_PT = 0, 1, 4, 5, 6, 9
COL = '是否增程(EREV)'
FILL = PatternFill('solid', fgColor='DDEBF7')


def s(v):
    return '' if v is None else str(v).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--write', action='store_true')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    hdr = [c.value for c in ws[1]]
    if COL in hdr:
        col_no = hdr.index(COL) + 1
        print(f'列已存在（第 {col_no} 列），幂等重跑')
    else:
        col_no = len(hdr) + 1
        ws.cell(1, col_no, COL)
        print(f'新增列「{COL}」第 {col_no} 列（表 {len(hdr)} 列 → {col_no} 列）')

    plan, skipped = [], []
    for idx, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if r[0] is None:
            continue
        blob = f"{s(r[I_MN])} {s(r[I_PN])} {s(r[I_GN])}"
        if '增程' not in blob:
            continue
        pt = s(r[I_PT])
        if pt != 'PHEV':
            skipped.append({'row': idx, 'powertrain': pt, 'blob': blob})
            continue
        plan.append({'row': idx, 'id': f"{s(r[I_B])}::{s(r[I_M])}",
                     'batch': s(r[I_B]), 'model': s(r[I_M]), 'enterprise': s(r[3]),
                     'brand': s(r[2]), 'modelName': s(r[I_MN]), 'genericName': s(r[I_GN]),
                     'marker': '增程'})

    stats = {'erev_candidates': len(plan) + len(skipped), 'write': len(plan),
             'skipped_non_phev': len(skipped),
             'by_enterprise': dict(collections.Counter(p['enterprise'] for p in plan).most_common(12))}
    print(json.dumps(stats, ensure_ascii=False))

    if args.write and plan:
        cs = wb['颜色说明']
        have = {str(r[1] or '') for r in cs.iter_rows(values_only=True) if len(r) > 1}
        if '增程标记' not in have:
            cs.append(['浅蓝底纹', '增程标记',
                       '「是否增程(EREV)」列：车型/产品/通用名称含「增程」字样且动力类型为 PHEV 时置 1。'
                       '系公告原文字面证据（非实测油耗派生，物理派生通道实测混入真插混故弃用）。'
                       '值域 1/空：非增程不打 0，避免「未判定」与「确定非增程」混淆。'])
        log = wb['变更记录']
        last = 0
        for rr in range(log.max_row, 1, -1):
            v = log.cell(rr, 1).value
            if isinstance(v, int):
                last = v
                break
        for p in plan:
            cell = ws.cell(p['row'], col_no)
            cell.value = 1
            cell.fill = FILL
            last += 1
            log.append([last, p['batch'], p['model'], 'PHEV', COL, '派生标记', '(空)', '1',
                        '名称含「增程」字面证据（公告原文）；不影响现有 PHEV 口径；'
                        f"派生自 车型名称/产品名称/通用名称 命中「增程」（企业={p['enterprise']}）"])
        saved = None
        try:
            wb.save(WB)
            saved = WB
            print('      已就地保存')
        except PermissionError:
            tmp = WB.replace('（341~410批）', '（341~410批）v42tmp')
            wb.save(tmp)
            for i in range(6):
                time.sleep(20)
                try:
                    os.replace(tmp, WB)
                    saved = WB
                    print(f'      替换成功（重试 {i+1}）')
                    break
                except PermissionError:
                    pass
            if not saved:
                kept = WB.replace('（341~410批）', '（341~410批·CR17-002版）')
                os.replace(tmp, kept)
                saved = kept
        print(f'已写入：{len(plan)} 行；变更记录追加 {len(plan)} 条（末序号 {last}）')
    else:
        print('dry-run：未写入。加 --write 执行')

    json.dump({'generated': '2026-10-05', 'mode': 'write' if args.write else 'dry-run',
               'column': COL, 'stats': stats, 'plan': plan, 'skipped': skipped},
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('->', OUT)


if __name__ == '__main__':
    main()