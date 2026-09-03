# -*- coding: utf-8 -*-
import openpyxl, re, io, sys, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE = r"D:\#AI\00 Project\17 TA Scan"
wb = openpyxl.load_workbook(BASE + r"\TA-Analysis\NEV公告参数汇总表_合并版（341~409批）.xlsx", read_only=True, data_only=True)
ws = wb['NEV公告参数汇总']
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    if row[0] is None: continue
    rows.append(list(row))
wb.close()

# 405-408 批 购置税值分布 + 行数
stat = collections.defaultdict(lambda: collections.Counter())
for r in rows:
    b = int(r[0]) if isinstance(r[0], (int, float)) else 0
    if 400 <= b <= 409:
        v = str(r[10]).strip() if r[10] is not None else '缺'
        if v in ('是','否'): stat[b][v] += 1
        else: stat[b]['缺'] += 1
        stat[b]['总'] += 1
print("批次  总行  是  否  缺")
for b in sorted(stat):
    s = stat[b]
    print(f"  {b}  {s['总']:>4}  {s['是']:>3}  {s['否']:>3}  {s['缺']:>3}")

# 已知"是"行的型号样例 → 用于验证 doc3 解析
known_yes = [(int(r[0]), str(r[1]).strip(), str(r[16]) if r[16] is not None else None) for r in rows
             if isinstance(r[0],(int,float)) and 405 <= int(r[0]) <= 408 and str(r[10]).strip()=='是']
print(f"\n405-408 已知是行: {len(known_yes)} 条, 样例: {known_yes[:5]}")

# 测试: 这些型号是否出现在对应批次 doc3 文本中
for b in [405, 406, 407, 408]:
    txt = open(BASE + rf"\data\nev-announcements\raw\b{b}\doc3.txt", encoding='utf-8', errors='replace').read()
    models_b = [m for bb, m, e in known_yes if bb == b]
    hit = sum(1 for m in models_b if m in txt)
    print(f"b{b}: doc3 {len(txt)}字符, 已知是型号命中 {hit}/{len(models_b)}")
# 缺失行型号在 doc3 中命中情况
miss_rows = [(int(r[0]), str(r[1]).strip()) for r in rows
             if isinstance(r[0],(int,float)) and 405 <= int(r[0]) <= 408 and (r[10] is None or str(r[10]).strip() not in ('是','否'))]
print(f"\n405-408 缺失行: {len(miss_rows)}")
for b in [405, 406, 407, 408]:
    txt = open(BASE + rf"\data\nev-announcements\raw\b{b}\doc3.txt", encoding='utf-8', errors='replace').read()
    ms = [m for bb, m in miss_rows if bb == b]
    hit = sum(1 for m in ms if m and m in txt)
    print(f"b{b}: 缺失型号在doc3命中 {hit}/{len(ms)}")
