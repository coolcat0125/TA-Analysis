# -*- coding: utf-8 -*-
"""_z6_arbitrate.py — 用官方税目录索引逐行仲裁「6xxx 却标轿车」的目标行。
   产出 audit-output/_z6_arbitration.json + 控制台摘要。
"""
import io
import json
import re
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
T = json.load(io.open("audit-output/_z1_target126.json", encoding="utf-8"))
IDX = json.load(io.open("audit-output/_z5_official_index_full.json", encoding="utf-8"))

BODY = [("多用途乘用车", "多用途乘用车"), ("运动型乘用车", "运动型乘用车"), ("越野乘用车", "越野乘用车"),
        ("轿车", "轿车"), ("客车", "客车"), ("专用车", "专用车")]


def body_of(s):
    for k, v in BODY:
        if k in (s or ""):
            return v
    return ""


out = []
for t in T:
    code = str(t["code"]).strip().upper()
    rec = IDX.get(code)
    if not rec:
        out.append(dict(t, official=None, verdict="无实物", note="税目录索引未命中该完整型号"))
        continue
    # 优先同批次记录
    same = [r for r in rec if r["batch"] == t["batch"]]
    r = (same or rec)[0]
    off = r["ptype"]
    tb, ob = body_of(t["ptype"]), body_of(off)
    if ob == tb:
        v = "一致"
    elif ob and tb and ob != tb:
        v = "订正"
    else:
        v = "待核"
    out.append(dict(t, official=off, obatch=r["batch"], odoc=r["doc"], oline=r["line"],
                    obody=ob, tbody=tb, verdict=v,
                    note="" if v == "一致" else f"官方 {off}"))

c = Counter(o["verdict"] for o in out)
print("仲裁结果：", dict(c), "合计", len(out))
print("\n== 订正候选（官方车身形式与底表不同）==")
for o in [x for x in out if x["verdict"] == "订正"][:25]:
    print(f"  {o['batch']}|{o['code']:<18} 底表={str(o['ptype'])[:22]:<24} 官方={o['official'][:26]:<28} ({o['odoc']}:{o['oline']})")
print(f"  … 共 {sum(1 for x in out if x['verdict']=='订正')} 行")
print("\n== 一致样例（官方也是轿车）==")
for o in [x for x in out if x["verdict"] == "一致"][:8]:
    print(f"  {o['batch']}|{o['code']:<18} {str(o['ptype'])[:30]}  ({o['odoc']}:{o['oline']})")
print("\n== 无实物 ==")
for o in [x for x in out if x["verdict"] == "无实物"][:15]:
    print(f"  {o['batch']}|{o['code']:<18} {str(o['ptype'])[:26]}  {str(o['cname'])[:20]}")

io.open("audit-output/_z6_arbitration.json", "w", encoding="utf-8", newline="\n").write(
    json.dumps(out, ensure_ascii=False, indent=1))
print("\n已写 audit-output/_z6_arbitration.json")
