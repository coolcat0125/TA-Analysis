# -*- coding: utf-8 -*-
"""_z12_docs.py — v4.9.40 文档收口：CHANGELOG / PROJECT_STATUS / CR-17-041 / X8 L5 结案。
   每条替换均 assert 唯一命中；PROJECT_STATUS 迭代记录只追加。
"""
import io
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.join("..", "..")
CH = os.path.join(ROOT, "05_文档与流程", "仓库文档", "CHANGELOG.md")
PS = os.path.join(ROOT, "05_文档与流程", "仓库文档", "PROJECT_STATUS.md")
CRJ = "Update/cr_register.json"
CRM = "Update/CR登记册.md"
X8 = os.path.join(ROOT, "03_核查信息", "一致性核查_X8_20261101.json")


def rd(p):
    return io.open(p, encoding="utf-8").read()


def wr(p, t):
    io.open(p, "w", encoding="utf-8", newline="\n").write(t)


# ── 1. CHANGELOG ─────────────────────────────────────────────────────
cl = rd(CH)
assert cl.startswith("## [4.9.39]"), cl[:40]
SEC = """## [4.9.40] - 2026-11-01

### 一、目标与方法
- 用户指令：继续做「403~412 批官方附件批量抓取 + 126 行逐行仲裁」。上一轮（v4.9.39）用 402 批官方附件为 50 格找到证据，本轮把同一通道**扩到 390~403 批全量税目录**，对「型号类别位 6（客车/多用途）却标『轿车』」的行逐行仲裁。
- **方法层关键发现**：车船税/购置税目录**以「完整车辆型号」逐行列出，并自带「产品名称（含车身形式）」**——完整型号是比公告附录1「型号族」更强的连接键，可把 402 批的一次性证据扩到全部批次。

### 二、通道与成本
- `_z3_fetch_all.py`：curl（`--http1.1` + 3 次退避重试，站点偶发 502/000）抓 11 个批次页（390/391/392/393/394/395/396/399/401/402/403）及附件，清单 `audit-output/_z3_manifest.json`。
- `_z4_convert.py`：`soffice --convert-to "txt:Text (encoded):UTF8"` 批转 21 份（doc3 优先、doc2 ≤8 MB），53.8 s 全部成功；`399_doc3` 转换 0 B 失败（该批由 399_doc2 覆盖），`403_doc2` 30.8 MB 跳过。
- `_z5_index.py` → `audit-output/_z5_official_index_full.json`：**12,590 个完整车辆型号**（22 份 txt 去重）；自检 `NJL6553HEV1`、`SC6522AAB6HEV` 均得「插电式增程混合动力多用途乘用车」。

### 三、逐行仲裁（`_z6_arbitrate.py` / `_z7_stats.py`）
- 目标 **125 行**（v4.9.39 已修的 `402|NJL6553HEV1` 不再含「轿车」，故 126 → 125），批次分布 390:11 / 391:5 / 392:14 / 393:4 / 394:8 / 395:24 / 396:14 / 399:4 / 401:22 / 402:2 / 403:17。
- 结果：**订正 124 / 无实物 1 / 一致 0**。官方车身形式 **多用途乘用车 112 + 运动型乘用车 12**；底表原文 插电式混合动力轿车 107 + 纯电动轿车 17。
- 官方含「增程」的 28 行，其 `是否增程` 早已为 `1`（0 行需补）→ **本轮不改动力口径，by_energy 不变**。
- 无实物 1 行 `401|SQR6520REEVE024`（风云T11）：税目录索引未命中完整型号，**不改值**，登记待公告原件。

### 四、写入（`_z8_apply.py`）
- 写前整本备份 `archive/底表_pre-产品类型官方订正_20261101T192225.xlsx`（sha 校验一致）。
- 逐格写入 **124 格 `产品类型` = 官方产品名称原文**；写后全表逐格对拍：**变化 124 格且全部落在「产品类型」**（其余 44 列零变化，硬断言）。
- 台账：《变更记录》+124（`变更类型='错误纠正'`，来源含「docX 第N行」）67,245 → **67,369**；《错误纠正台账》1,281 → **1,405**；《官方目录核实台账》410 → **535**（124 已写入 + 1 无实物）。
- 底表 sha `d2d105a0…` → **`e88d749ad6204362…`**，4,832,711 B，5,475 行 × 45 列 / 16 sheets。

### 五、门禁与口径
- `audit_data.py --strict` PASS；`verify_consistency.py` 记录 5,475 / CRITICAL 0 / WARN 212（W2 1 · W4 25 · W5 27 · W6 159）→ PASS。
- 统计口径**零变化**（`产品类型` 不参与任何聚合）：`by_energy=[[BEV,523457],[PHEV,204286],[EREV,69958]]`、`coverage.row 2,202/5,475`、`current_board 68.6131%` / `73.0453%` 全部沿用；仅 `meta.base_table_sha256` 同步（旧值存 `_prev`，附 note 说明）。

### 六、交付物
- 54 表 / 两版看板 / 移动版按新底表重生成并发布（`02_交付物/发布清单.json` 8 件）：数据底表 91,489 B、看板 4,681,632 B、发布版 4,680,571 B、移动版 1,588,369 B、报告 74,224 B、演示稿 622,654 B。

### 七、结案
- **CR-17-041（6xxx 却标「轿车」126 行）→ 已完成**。上一轮「仅 2 行可核、其余无第二见证」的结论被官方税目录推翻：**完整型号连接键使 125/125 行均可核（1 行未命中）**，其中 124 行按官方原文订正、1 行留待公告原件。
- 遗留：①无实物 1 行待公告原件；②403_doc2（30.8 MB）未转换，其覆盖的批次内容由 doc3 承担。

"""
wr(CH, SEC + cl)
print("CHANGELOG [4.9.40] 已插入")

# ── 2. PROJECT_STATUS ────────────────────────────────────────────────
ps = rd(PS)
old = "当前版本 **v4.9.39**（CR-17-031"
new = ("当前版本 **v4.9.40**（v4.9.40 = 产品类型官方订正 124 格 + CR-17-041 结案；"
       "上轮 v4.9.39 = 官方目录实物核实 50 格。历史：CR-17-031")
assert ps.count(old) == 1, ps.count(old)
ps = ps.replace(old, new)

anchor = "## 5. 迭代记录\n"
assert ps.count(anchor) == 1
ENTRY = """
### 2026-11-01 19:2x · DSH 会话 lead · v4.9.40 125 行产品类型官方逐行仲裁（124 格订正 + 1 格无实物）

- **扩通道**：把 v4.9.39 的 402 批单点证据扩到 **390~403 批全量税目录**（车船税/购置税目录 22 份 txt），建成 **12,590 个完整车辆型号**的官方索引 `audit-output/_z5_official_index_full.json`。方法层关键：**税目录以完整车辆型号逐行列出并自带产品名称（含车身形式）**，完整型号比公告附录1「型号族」连接更强。
- **仲裁 125 行**（原 126 行，v4.9.39 已修 1 行不再命中）：**订正 124 / 无实物 1 / 一致 0**；官方车身形式 多用途乘用车 112 + 运动型乘用车 12。
- **写入 124 格 `产品类型` = 官方产品名称原文**；写后全表逐格对拍：变化 124 格、全部落在该列（其余 44 列零变化）。台账：`变更记录` 67,245→**67,369**、《错误纠正台账》→**1,405**、《官方目录核实台账》→**535**。
- **口径零变化**：官方含「增程」的 28 行底表早已标 1，故 `by_energy`/覆盖率/集中度全部沿用；仅同步 `meta.base_table_sha256`（`d2d105a0…`→`e88d749a…`）。
- **门禁**：`audit_data --strict` PASS；`verify_consistency` 5,475 / CRITICAL 0 / WARN 212 / PASS。
- **交付物**：54 表 + 两版看板 + 移动版重生成并发布（8 件，`发布清单.json` 已重算）。
- **结案**：CR-17-041（6xxx 却标轿车）→ 已完成；遗留 1 行无实物（`401|SQR6520REEVE024` 风云T11）待公告原件。
"""
ps = ps.replace(anchor, anchor + ENTRY, 1)
wr(PS, ps)
print("PROJECT_STATUS 迭代记录已追加（只追加，未改历史）")

# ── 3. CR-17-041 结案 ────────────────────────────────────────────────
d = json.load(io.open(CRJ, encoding="utf-8"))
hit = [i for i in d["items"] if i["id"] == "CR-17-041"]
assert len(hit) == 1
it = hit[0]
it["status"] = "已完成"
it["result"] = (
    "【2026-11-01 v4.9.40 结案】上一轮「仅 2 行可核」的结论被官方税目录推翻：把 390~403 批车船税/购置税目录"
    "（22 份 txt、12,590 个完整车辆型号）建成官方索引后，**125/125 行均可核（1 行未命中）**——"
    "**订正 124 行**（官方车身形式：多用途乘用车 112 + 运动型乘用车 12；底表原文为插电式混合动力轿车 107 + 纯电动轿车 17），"
    "**无实物 1 行**（401|SQR6520REEVE024 风云T11，留待公告原件）。写入 124 格 `产品类型`=官方产品名称原文，"
    "写后逐格对拍仅这 124 格变化；台账 `变更记录` 67,245→67,369、《错误纠正台账》→1,405、《官方目录核实台账》→535。"
    "官方含「增程」的 28 行底表现值已为 1，故动力口径与全部统计量零变化。门禁 PASS / CRITICAL 0 / WARN 212。"
)
it["date"] = "2026-11-01"
io.open(CRJ, "w", encoding="utf-8", newline="\n").write(json.dumps(d, ensure_ascii=False, indent=1))
print("CR-17-041 → 已完成")

crm = rd(CRM)
lines = crm.splitlines()
for i, ln in enumerate(lines):
    if "CR-17-041" in ln:
        lines[i] = ln.rstrip() + " · **已完成（v4.9.40）**：390~403 批税目录建成 12,590 型号官方索引，125 行全部可核 → 订正 124 格、无实物 1 行（风云T11）；产物 `audit-output/_z6_arbitration.json`。"
        break
wr(CRM, "\n".join(lines) + "\n")
print("CR登记册.md 已登记结案")

# ── 4. X8 L5 结案 ────────────────────────────────────────────────────
if os.path.exists(X8):
    x = json.load(io.open(X8, encoding="utf-8"))
    k = "L5_产品类型与型号编码交叉核验"
    if k in x:
        x[k]["resolution"] = dict(
            status="已完全仲裁（v4.9.40）",
            method="390~403 批车船税/购置税目录全文 → 完整车辆型号官方索引（12,590 型号）→ 逐行比对官方产品名称",
            corrected_rows=124, no_evidence_rows=1, evidence="audit-output/_z6_arbitration.json",
            no_evidence_item="401|SQR6520REEVE024（风云T11）",
            at="2026-11-01")
        io.open(X8, "w", encoding="utf-8", newline="\n").write(json.dumps(x, ensure_ascii=False, indent=1))
        print("X8 L5 已登记 resolution")
