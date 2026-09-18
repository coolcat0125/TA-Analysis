# apply_length_patch_v49.py — v4.9 车长/轴距候选补丁执行器（A1 修复包执行预案）
#
# 依据 candidate/v4.9/vehicle_length/patch_exact_20260918.csv（77 格/48 行，全部带证据引用）
# 默认 DRY-RUN：仅校验锚点 + 输出补丁后副本（--output），绝不改动输入底表。
# --apply 才写入：先备份 → 逐格写值（浅红 FFC7CE 底纹）→ 《变更记录》追加台账（变更类型=更正）。
#
# 用法：
#   python3 apply_length_patch_v49.py --input <底表.xlsx> --patch candidate/v4.9/vehicle_length/patch_exact_20260918.csv \
#       --output /tmp/patched_preview.xlsx          # dry-run（默认）
#   python3 apply_length_patch_v49.py ... --apply   # 真写（需 Orchestrator/人工审批后执行）
#
# 铁律：锚点不符即中止该格（当前值≠补丁"当前值(锚点)"列 → 拒绝写入，防错位）；apply 模式零锚点失配才落盘。
# Copyright (c) 2026 David YEAH · MIT License
import argparse, csv, shutil, sys, datetime
import openpyxl
from openpyxl.styles import PatternFill

RED = PatternFill('solid', fgColor='FFC7CE')   # 数据修复更正·浅红（仓库约定）

def g(v): return '' if v is None else str(v).strip()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--patch', required=True)
    ap.add_argument('--output', default=None, help='dry-run 补丁后副本输出路径')
    ap.add_argument('--apply', action='store_true', help='真写（备份+台账+底纹）')
    args = ap.parse_args()

    patch = list(csv.DictReader(open(args.patch, encoding='utf-8-sig')))
    wb = openpyxl.load_workbook(args.input)
    ws = wb['NEV公告参数汇总']
    it = ws.iter_rows(values_only=True)
    hdr = [str(h) if h is not None else '' for h in next(it)]
    def col(*names):
        for i, h in enumerate(hdr):
            if any(n in h for n in names): return i
        raise KeyError(names)
    cB, cM = col('批次'), col('产品型号')
    cL, cW = col('车长'), col('轴距')

    # 行索引（openpyxl 行号从 1 起，表头=1 → 数据自 2；必须复用已消费表头的同一迭代器）
    rowmap = {}
    for ri, r in enumerate(it, start=2):
        m = g(r[cM])
        if m:
            rowmap[(g(r[cB]).split('.')[0], m)] = ri

    fcol = {'车长': cL, '轴距': cW}
    ok = bad = 0
    fails = []
    for p in patch:
        key = (p['批次'].strip(), p['产品型号'].strip())
        ri = rowmap.get(key)
        if ri is None:
            bad += 1; fails.append((p, 'ROW_MISSING')); continue
        ci = fcol[p['字段'].strip()]
        cur = g(ws.cell(ri, ci + 1).value)
        if cur != p['当前值(锚点)'].strip():
            bad += 1; fails.append((p, f'ANCHOR_MISMATCH 实际={cur}')); continue
        ok += 1

    print(f'锚点校验: 一致={ok} 失配/缺行={bad}')
    for p, why in fails[:10]:
        print('  ✗', p['批次'], p['产品型号'], p['字段'], '→', why)
    if bad:
        print('存在失配——按铁律拒绝全部写入（含 dry-run 副本不输出）。')
        sys.exit(2)

    if args.output and not args.apply:
        for p in patch:
            key = (p['批次'].strip(), p['产品型号'].strip())
            ri = rowmap[key]; ci = fcol[p['字段'].strip()] + 1
            ws.cell(ri, ci).value = float(p['新值']) if '.' in p['新值'] else int(p['新值'])
            ws.cell(ri, ci).fill = RED
        wb.save(args.output)
        print(f'dry-run 副本已输出: {args.output}（共 {len(patch)} 格，浅红底纹，未动输入文件）')
    elif args.apply:
        ts = datetime.date.today().strftime('%Y%m%d')
        bak = args.input.replace('.xlsx', f'_pre_v49length_{ts}.xlsx')
        shutil.copy2(args.input, bak)
        wsl = wb['变更记录']
        lit = wsl.iter_rows(values_only=True)
        next(lit)
        seq = max(int(g(r[0])) for r in lit if r[0] is not None and g(r[0]).isdigit())
        for p in patch:
            key = (p['批次'].strip(), p['产品型号'].strip())
            ri = rowmap[key]; ci = fcol[p['字段'].strip()] + 1
            ws.cell(ri, ci).value = float(p['新值']) if '.' in p['新值'] else int(p['新值'])
            ws.cell(ri, ci).fill = RED
            seq += 1
            wsl.append([seq, p['批次'], p['产品型号'], g(ws.cell(ri, col('动力类型') + 1).value),
                        ws.cell(1, ci).value, '更正', p['当前值(锚点)'], p['新值'],
                        f"{p['证据引用']}（{p['证据等级']}级）· v4.9 A1 车长修复包 · {p['定值说明']}"])
        wb.save(args.input)
        print(f'APPLY 完成: 备份={bak}，写入 {len(patch)} 格，台账追加至序号 {seq}')
    else:
        print('仅锚点校验（未输出副本）。')
    wb.close()

if __name__ == '__main__':
    main()
