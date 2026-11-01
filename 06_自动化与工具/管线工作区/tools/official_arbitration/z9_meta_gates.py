# -*- coding: utf-8 -*-
"""_z9_meta_gates.py — 底表 sha 变更后同步 JSON 元信息，并跑两道门禁。"""
import hashlib
import io
import json
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")
WB = "NEV公告参数汇总表_合并版（341~412批）.xlsx"
JS = "audit-output/sales_linkage_analytics_20261019.json"

sha = hashlib.sha256(open(WB, "rb").read()).hexdigest()
J = json.load(io.open(JS, encoding="utf-8"))
old = J["meta"].get("base_table_sha256")
if old != sha:
    J["meta"]["base_table_sha256_prev"] = old
    J["meta"]["base_table_sha256"] = sha
    J["meta"]["base_table_sha256_note"] = (
        "v4.9.40：产品类型官方订正 124 格（工信部税目录逐型号证据）导致字节变化；"
        "统计量未重算（产品类型不参与任何聚合口径）。")
    io.open(JS, "w", encoding="utf-8", newline="\n").write(json.dumps(J, ensure_ascii=False, indent=1))
    print("meta sha 已同步", (old or "")[:16], "->", sha[:16])
else:
    print("meta sha 已一致", sha[:16])
print("base_rows", J["meta"].get("base_rows"), "columns", J["meta"].get("base_table_columns"))
print("by_energy", J["structure"]["by_energy"])

for cmd in (["python", "audit_data.py", "--strict"], ["python", "verify_consistency.py"]):
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (p.stdout or "") + (p.stderr or "")
    print(f"\n== {' '.join(cmd)} exit={p.returncode} ==")
    print("\n".join(out.strip().splitlines()[-12:]))
