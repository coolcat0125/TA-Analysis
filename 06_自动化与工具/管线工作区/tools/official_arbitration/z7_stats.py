# -*- coding: utf-8 -*-
"""_z7_stats.py — 统计 124 行订正的构成：官方车身形式分布 / 官方含「增程」的行数 / 底表是否增程现状。"""
import io
import json
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
A = json.load(io.open("audit-output/_z6_arbitration.json", encoding="utf-8"))
from openpyxl import load_workbook

wb = load_workbook("NEV公告参数汇总表_合并版（341~412批）.xlsx", read_only=True, data_only=True)
ws = wb["NEV公告参数汇总"]
H = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
ix = {n: i for i, n in enumerate(H)}
cur = {}
for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
    cur[i] = r
wb.close()

fix = [a for a in A if a["verdict"] == "订正"]
print("订正行", len(fix))
print("官方车身形式分布：", dict(Counter(a["obody"] for a in fix)))
print("底表 产品类型 原文分布：", dict(Counter(str(a["ptype"]) for a in fix)))
erev_official = [a for a in fix if "增程" in (a["official"] or "")]
print("官方含「增程」的订正行：", len(erev_official))
now = Counter()
for a in fix:
    r = cur[a["row"]]
    v = r[ix["是否增程(EREV)"]]
    now[(("官方增程" if "增程" in (a["official"] or "") else "官方非增程"), str(v))] += 1
print("交叉（官方是否增程 × 底表现值）：", dict(now))

need_erev = [a for a in fix if "增程" in (a["official"] or "") and cur[a["row"]][ix["是否增程(EREV)"]] in (None, "")]
print("\n需同时补 是否增程=1 的行：", len(need_erev))
for a in need_erev[:8]:
    print(f"  {a['batch']}|{a['code']:<18} {a['official']}")
chg = sum(1 for a in fix if str(a["ptype"]) != str(a["official"]))
print("\n产品类型 实际需要改写的行（现值≠官方）：", chg)
for a in fix:
    if str(a["ptype"]) == str(a["official"]):
        print("  已一致但仍被判订正（车身形式同、前缀不同）：", a["batch"], a["code"], a["ptype"], "|", a["official"])

