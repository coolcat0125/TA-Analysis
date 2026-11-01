# -*- coding: utf-8 -*-
"""_z8_apply.py — 用工信部税目录证据把 124 行「产品类型」订正为官方产品名称。
   证据：audit-output/_z6_arbitration.json（来自 audit-output/_z5_official_index_full.json）。
   纪律：只改这 124 格的「产品类型」；逐格写《变更记录》+《错误纠正台账》+《官方目录核实台账》；
        写前整本备份、写后全表逐格对拍（只允许 124 格变化）。
"""
import datetime
import hashlib
import io
import json
import os
import shutil
import sys

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
WB = "NEV公告参数汇总表_合并版（341~412批）.xlsx"
ARB = "audit-output/_z6_arbitration.json"
NOW = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
TAG = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")

fix = [a for a in json.load(io.open(ARB, encoding="utf-8")) if a["verdict"] == "订正"]
miss = [a for a in json.load(io.open(ARB, encoding="utf-8")) if a["verdict"] == "无实物"]
print(f"待写入 {len(fix)} 行；无实物 {len(miss)} 行")

# ── 备份 ─────────────────────────────────────────────────────────────
bak = f"archive/底表_pre-产品类型官方订正_{TAG}.xlsx"
os.makedirs("archive", exist_ok=True)
src_sha = hashlib.sha256(open(WB, "rb").read()).hexdigest()
shutil.copy2(WB, bak)
assert hashlib.sha256(open(bak, "rb").read()).hexdigest() == src_sha, "备份校验失败"
print(f"备份 {bak}  sha={src_sha[:16]}")

wb = load_workbook(WB)
ws = wb["NEV公告参数汇总"]
H = [c.value for c in ws[1]]
ix = {n: i + 1 for i, n in enumerate(H)}
C_PT = ix["产品类型"]

before = {}
for r in range(2, ws.max_row + 1):
    before[r] = tuple(ws.cell(r, c).value for c in range(1, len(H) + 1))

changed = 0
for a in fix:
    r = a["row"]
    cell = ws.cell(r, C_PT)
    assert str(cell.value) == str(a["ptype"]), f"行{r} 产品类型现值漂移：{cell.value!r} != {a['ptype']!r}"
    cell.value = a["official"]
    changed += 1
assert changed == len(fix) == 124, changed

# ── 台账 ─────────────────────────────────────────────────────────────
def _maxnum(ws):
    m = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, (int, float)):
            m = max(m, int(v))
        elif isinstance(v, str) and v.strip().isdigit():
            m = max(m, int(v.strip()))
    return m


lg = wb["变更记录"]
lg_max = _maxnum(lg)
el = wb["错误纠正台账"]
el_max = _maxnum(el)
of = wb["官方目录核实台账"]
of_max = _maxnum(of)

for i, a in enumerate(fix, 1):
    a["dtype"] = ("PHEV" if "插电" in str(a["ptype"]) else "BEV" if "纯电动" in str(a["ptype"]) else "")
    src = f"工信部官方目录核实（{a['odoc']} 第{a['oline']}行）；规则 W3-车身形式不符（型号类别位 6=客车/多用途，底表误标轿车）；{NOW}"
    lg.append([lg_max + i, a["batch"], a["code"], a["dtype"], "产品类型", "错误纠正",
               a["ptype"], a["official"], src])
    el.append([el_max + i, a["batch"], a["code"], a["dtype"], "产品类型", a["ptype"], a["official"],
               "官方目录核实（W3-车身形式不符）",
               f"工信部 {a['odoc']} 第{a['oline']}行 产品名称=「{a['official']}」；型号 {a['code']} 逐字命中", NOW])
    of.append([of_max + i, a["batch"], a["code"], a["ptype"], a["official"],
               "W3-车身形式（官方为多用途/运动型乘用车，底表误标轿车）", "已写入",
               a["odoc"], f"第{a['oline']}行", NOW])

n_extra = 0
for a in miss:
    n_extra += 1
    of.append([of_max + len(fix) + n_extra, a["batch"], a["code"], a["ptype"], "",
               "无实物（税目录索引未命中该完整型号）", "未写入（证据不足，留待公告原件）",
               "-", "-", NOW])

# ── 备注除名 ─────────────────────────────────────────────────────────
wb.save(WB)
new_sha = hashlib.sha256(open(WB, "rb").read()).hexdigest()

# ── 写后逐格对拍 ─────────────────────────────────────────────────────
wb2 = load_workbook(WB, read_only=True, data_only=True)
ws2 = wb2["NEV公告参数汇总"]
after = {}
for r in range(2, ws2.max_row + 1):
    after[r] = tuple(ws2.cell(r, c).value for c in range(1, len(H) + 1)) if False else None
diffs = []
rows2 = list(ws2.iter_rows(min_row=2, values_only=True))
assert len(rows2) == len(before), (len(rows2), len(before))
for i, r in enumerate(range(2, ws.max_row + 1)):
    if rows2[i] != before[r]:
        for c in range(len(H)):
            if rows2[i][c] != before[r][c]:
                diffs.append((r, H[c], before[r][c], rows2[i][c]))
wb2.close()

print(f"\n逐格对拍：变化 {len(diffs)} 格（应 == {len(fix)}）")
for d in diffs[:5]:
    print("  ", d[0], d[1], "|", d[2], "->", d[3])
assert len(diffs) == len(fix), len(diffs)
assert all(d[1] == "产品类型" for d in diffs), "出现非预期字段变化"

print(f"sha {src_sha[:16]} -> {new_sha[:16]}  size={os.path.getsize(WB)}")
print(f"变更记录 {lg_max} -> {lg_max + len(fix)}；错误纠正台账 {el_max} -> {el_max + len(fix)}；"
      f"官方目录核实台账 {of_max} -> {of_max + len(fix) + n_extra}")
io.open("audit-output/_z8_apply_result.json", "w", encoding="utf-8", newline="\n").write(
    json.dumps(dict(sha_before=src_sha, sha_after=new_sha, cells=len(diffs), rows=len(fix),
                    no_evidence=len(miss), backup=bak, at=NOW), ensure_ascii=False, indent=1))
print("已写 audit-output/_z8_apply_result.json")
