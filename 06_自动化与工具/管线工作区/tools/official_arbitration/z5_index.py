# -*- coding: utf-8 -*-
"""_z5_index.py — 从税目录 txt 建「完整车辆型号 → 官方产品名称」索引。
   用法：python _z5_index.py           # 扫描 raw/txt/*.txt，落 audit-output/_z5_official_index_full.json
"""
import glob
import io
import json
import os
import re
import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding="utf-8")
TXT = os.path.join("公告批次扫描_411-412_20261017", "raw", "txt")
OUT = "audit-output/_z5_official_index_full.json"

CODE = re.compile(r"^[A-Z][A-Z0-9]{2,5}\d{2,5}[A-Z0-9\-]*$")
PTYPE = re.compile(r"(纯电动|插电式|混合动力|增程|燃料电池|甲醇|乙醇|天然气).{0,14}(乘用车|轿车|客车|越野|专用车|货车|厢式运输车)$")


def norm_code(s):
    s = (s or "").strip().upper().replace(" ", "")
    return s


idx = defaultdict(list)
files = sorted(glob.glob(os.path.join(TXT, "*.txt")))
print(f"扫描 {len(files)} 个 txt")
for f in files:
    base = os.path.splitext(os.path.basename(f))[0]
    batch = int(base.split("_")[0]) if base.split("_")[0].isdigit() else None
    doc = base.split("_")[-1]
    n = 0
    for ln, line in enumerate(io.open(f, encoding="utf-8", errors="replace"), 1):
        cells = [c.strip() for c in line.split("\t")]
        codes = [norm_code(c) for c in cells if CODE.match(norm_code(c))]
        if not codes:
            continue
        ptypes = [c for c in cells if PTYPE.search(c)]
        if not ptypes:
            # 退一步：取含「乘用车/轿车/客车」且长度合理的单元格
            ptypes = [c for c in cells if re.search(r"(乘用车|轿车)$", c) and 4 <= len(c) <= 30]
        if not ptypes:
            continue
        for c in codes:
            idx[c].append({"batch": batch, "doc": doc, "line": ln,
                           "ptype": ptypes[0], "ptypes_all": ptypes,
                           "cells": cells[:12]})
            n += 1
    print(f"  {base:<14} 命中行 {n}")

print(f"\n去重型号 {len(idx)} 个")
io.open(OUT, "w", encoding="utf-8", newline="\n").write(
    json.dumps({k: v for k, v in idx.items()}, ensure_ascii=False, indent=1))

# 让 402 的 NJL6553HEV1 做自检（已知官方=插电式增程混合动力多用途乘用车）
for probe in ("NJL6553HEV1", "SC6522AAB6HEV"):
    v = idx.get(probe)
    print(f"  自检 {probe}: {v[0]['ptype'] if v else '未命中'}")
