#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""CR-17-008 · 411 正式批公告 -> 权威底表 映射器（DRY-RUN ONLY）

职责
----
1) 把 411 正式公告扫描包（清单 46 条 / 公示详情 47 条）映射为：
   - new_rows : 底表新增行（25 型号，每型号 1 行，键 = 39 列名）
   - enrich   : 存量 411 行（21 型号）的「只补空、不覆盖」单元格补丁
   - ledger   : 追加到 `变更记录` 的台账行（新增 1 型号 1 条 + 富化 1 格 1 条）
   - qc_issues: 解析/一致性疑点（不改口径，只记录）
   - unmatched: 详情多出的型号（不入行）
2) 产出 audit-output/cr17_008_411_dryrun_20261019.json 与 cr17_008_411_report.md。

铁律
----
* 本脚本对工作簿 **只读**（openpyxl read_only=True），绝不写工作簿。
* 月销量(辆)/销量数据月份/销量更新日期 三列一律不产出、不触碰（并发会话在填）。
* 是否减免购置税 / 细分市场 / 电池能量密度 / 百公里电耗 一律留空（411 公示详情不提供）。
* 富化只填空单元格，任何非空值一律跳过。

用法
----
    python cr17_008_rows_411.py              # 默认 dry-run，写两个交付文件
    python cr17_008_rows_411.py --dry-run    # 同上（显式）
本脚本不存在 --apply 分支：落库由 Lead 的应用脚本执行。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import openpyxl

# ---------------------------------------------------------------- 常量 / 路径
BATCH = 411
# 数据来源标签：对齐既有「工信部公告参数查询系统（第409批官方详情）」风格（Lead p12 先例裁定）
DATA_SOURCE_TAG = "411批正式公告（工信部公告参数查询系统详情页）"
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
WORKBOOK_NAME = "NEV公告参数汇总表_合并版（341~410批）.xlsx"
WORKBOOK = os.path.join(REPO_DIR, WORKBOOK_NAME)
DATA_DIR = r"D:\#AI\00 Project\17 TA analysis\公告批次扫描_411-412_20261017\data"
LIST_PATH = os.path.join(DATA_DIR, "list_411_nev_passenger.json")
DETAIL_PATH = os.path.join(DATA_DIR, "details_411.json")
OUT_JSON = os.path.join(REPO_DIR, "audit-output", "cr17_008_411_dryrun_20261019.json")
OUT_MD = os.path.join(REPO_DIR, "audit-output", "cr17_008_411_report.md")
SHEET = "NEV公告参数汇总"
LEDGER_SHEET = "变更记录"

# 39 列权威表头（必须与工作簿第 1 行完全一致，按列名映射，绝不用列下标）
HEADER = [
    "批次", "产品型号", "产品商标", "企业名称", "车型名称", "产品名称", "通用名称",
    "产品类型", "细分市场", "动力类型", "是否减免购置税", "整备质量(kg)", "轴距(mm)",
    "纯电续航里程(km)", "电池容量(kWh)", "电池类型", "电池能量密度(Wh/kg)",
    "百公里电耗(kWh/100km)", "前电机功率/扭矩", "电机总功率/扭矩", "后电机功率/扭矩",
    "电机生产企业", "电机型号", "综合油耗(L/100km)", "B状态油耗(L/100km)",
    "发动机排量(mL)", "发动机功率(kW)", "发动机生产企业", "发动机型号", "数据来源",
    "车长(mm)", "电芯供应商", "电池包供应商", "官方参数页URL", "媒体校验来源",
    "月销量(辆)", "销量数据月份", "销量更新日期", "是否增程(EREV)",
]

LEDGER_HEADER = ["序号", "批次", "产品型号", "动力类型", "字段", "变更类型",
                 "附件1原值", "合并值", "数据来源"]

# 绝不产出的列（既有口径 / 并发会话领地）
SALES_COLS = ["月销量(辆)", "销量数据月份", "销量更新日期"]
BLANK_BY_POLICY = ["细分市场", "是否减免购置税", "电池能量密度(Wh/kg)",
                   "百公里电耗(kWh/100km)", "媒体校验来源"]
NEVER_EMIT = set(SALES_COLS) | set(BLANK_BY_POLICY)
# 富化候选列的排除项：身份列 / 数据来源（存量已有口径）/ 同上所有绝不产出列
ENRICH_EXCLUDE = NEVER_EMIT | {"产品型号", "数据来源", "批次"}

# 任务书给定的「新增 25 型号」清单（用于交叉复核；实际以差集为准）
EXPECTED_NEW = [
    "BJ7001A51HBEV", "BMW6495CABEV", "BMW7000BABEV", "BYD7000CFBEV3",
    "CC2040BV24APHEV", "CC2040BV24DPHEV", "CC6420BJ00GEBEV", "CC6510DG21HABEV",
    "CC7002DE00HABEV", "CC7002DE00HBBEV", "DFM7002S1F3BEV", "DFM7002S1F4BSEV",
    "LXA6503BEVW04", "LXA6503BEVW05", "LXA6503BEVW06", "LXA6540BEVW12",
    "LZ6460MJD0EV", "LZ6510VLC1EV", "LZ6510VLD1EV", "MR6481BEV106",
    "MR6521BEV101", "SQR6450BEVT1VA2", "SQR6450BEVT1VA3", "SVW646193AEV",
    "SVW646194AEV",
]

TZ_CN = timezone(timedelta(hours=8))
NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")
SUPPLIER_KW_RE = re.compile(r"公司|能源|动力|电池|时代|科技|集团|锂业|新能")
# 工况标注剥离口径：工作簿「字段映射说明」sheet 第 6 行 ——
# 「工况口径原则：BEV=CLTC，PHEV/EREV=WLTC；单元格不再保留 (CLTC)/(WLTC) 标注，工况按动力类型推定」
CONDITION_ANNOTATION_RE = re.compile(
    r"\s*[（(]\s*(?:CLTC|WLTC|NEDC|WLTP)\s*[)）]", re.IGNORECASE)
NORMALIZATIONS = ["strip 工况标注 (CLTC)/(WLTC)/(NEDC)/(WLTP)"]


def strip_condition_annotation(value):
    """剥离工况标注括号（仅 CLTC/WLTC/NEDC/WLTP）。

    返回 (新值, 是否改写)。仅剥离工况词括号，其它语义占位括号（如 `53(单箱26.5/共2箱)`）原样保留；
    剥离后若为单一纯数值则经 num_or_text 回归数值类型，多值（如 `780/770`）保持文本。
    """
    if not isinstance(value, str):
        return value, False
    stripped = CONDITION_ANNOTATION_RE.sub("", value).strip()
    if stripped == value:
        return value, False
    return num_or_text(stripped), True


# ---------------------------------------------------------------- 小工具
def _s(value):
    """宽松字符串化 + strip；None -> ''。"""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def num_or_text(value):
    """单一纯数值 -> int/float；含单位/括号/多值等原生格式 -> 原样文本；空 -> None。"""
    text = _s(value)
    if not text:
        return None
    if NUM_RE.fullmatch(text):
        return float(text) if "." in text else int(text)
    return text


def power_num(value):
    """功率值：整数值用 int，带 '/' 的多值保留文本。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        val = float(value)
        return int(val) if val.is_integer() else val
    text = _s(value)
    if text in ("", "-", "—", "/"):
        return None
    if NUM_RE.fullmatch(text):
        return power_num(float(text))
    return text


def numeric_sum(value):
    """数值或 'a/b/c' 多值求和；不可解析 -> None。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = _s(value)
    if not text:
        return None
    if "/" in text:
        total = 0.0
        for part in text.split("/"):
            part = part.strip()
            if not part:
                continue
            try:
                total += float(part)
            except ValueError:
                return None
        return total
    try:
        return float(text)
    except ValueError:
        return None


def ledger_str(value):
    """台账文本化：整数型 float 去掉小数点尾巴。"""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def join_component_values(values):
    """组件列表逐项拼接：顺序保留、去重、跳过占位符、用 '/' 分隔。"""
    out, seen = [], set()
    for value in values:
        text = _s(value)
        if text in ("", "-", "—", "/"):
            continue
        if text not in seen:
            seen.add(text)
            out.append(text)
    return "/".join(out) if out else None


def power_type_from_fuel(fuel):
    text = _s(fuel)
    if "纯电动" in text and "混合" not in text:
        return "BEV"
    if "混合" in text:
        return "PHEV"
    return None


# ---------------------------------------------------------------- 映射核心
def map_detail(detail, qc):
    """详情一条 -> {底表列名: 值}（只含可映射列；空值为 None）。"""
    model = _s(detail.get("产品型号"))
    url = _s(detail.get("详情页URL")) or _s(detail.get("来源URL"))
    ptype = power_type_from_fuel(detail.get("燃料种类"))
    if ptype is None:
        qc.append({
            "产品型号": model, "字段": "动力类型",
            "原始值": _s(detail.get("燃料种类")) or "(空)",
            "问题": "燃料种类无法判定动力类型，动力类型留空",
        })

    generic_name = _s(detail.get("通用名称")) or None
    # 车型名称 = 通用名称「整串原样」（Lead p12 先例裁定：409 官方详情队列 100/111、397 官方目录 115/126
    # 两列完全相同；多并列项也整串写入，如 车型名称='比亚迪唐,比亚迪唐Pro'，截断会丢别名）
    car_name = generic_name

    mapped = {
        "产品型号": model,
        "产品商标": _s(detail.get("产品商标")) or None,
        "企业名称": _s(detail.get("企业名称")) or None,
        "车型名称": car_name,
        "产品名称": _s(detail.get("产品名称")) or None,
        "通用名称": generic_name,
        "产品类型": _s(detail.get("产品名称")) or None,
        "动力类型": ptype,
        "整备质量(kg)": num_or_text(detail.get("整备质量(kg)")),
        "轴距(mm)": num_or_text(detail.get("轴距(mm)")),
        "纯电续航里程(km)": num_or_text(detail.get("纯电续航里程(km)")),
        "电池容量(kWh)": num_or_text(detail.get("电池容量(kWh)")),
        "电池类型": _s(detail.get("电池类型_公示其它")) or None,
        "综合油耗(L/100km)": num_or_text(detail.get("油耗申报值(L/100km)")),
        "B状态油耗(L/100km)": None,
        "车长(mm)": num_or_text(detail.get("车长(mm)")),
        "官方参数页URL": url or None,
        "数据来源": DATA_SOURCE_TAG,
    }

    # ---- 组件：电机 / 发动机
    motor_rows = [r for r in (detail.get("电机组件_公示详情") or []) if isinstance(r, dict)]
    engine_rows = [r for r in (detail.get("发动机组件_公示详情") or []) if isinstance(r, dict)]

    # BEV 无发动机：若 电机组件 为空而 发动机组件 承载了电机信息（排量为 0/- 的解析归类异常），回退取用
    if not motor_rows and ptype == "BEV":
        fallback = [r for r in engine_rows if _s(r.get("排量")) in ("", "-", "—", "0", "0.0")]
        if fallback:
            motor_rows = fallback
            qc.append({
                "产品型号": model, "字段": "电机组件_公示详情",
                "原始值": json.dumps(detail.get("发动机组件_公示详情"), ensure_ascii=False),
                "问题": "BEV 电机信息被解析进发动机组件，已按 BEV 电机回退规则取用；发动机字段留空",
            })

    mapped["电机型号"] = join_component_values([r.get("型号") for r in motor_rows])
    mapped["电机生产企业"] = join_component_values([r.get("生产企业") for r in motor_rows])

    front = power_num(detail.get("前电机峰值功率_kW"))
    rear = power_num(detail.get("后电机峰值功率_kW"))
    total = power_num(detail.get("电机总功率_kW"))
    total_value = numeric_sum(total)
    parts = [p for p in (front, rear) if p is not None]
    if parts:
        parts_sum = sum(numeric_sum(p) or 0.0 for p in parts)
        if total_value is None:
            total = power_num(round(parts_sum, 3))
            qc.append({
                "产品型号": model, "字段": "电机总功率/扭矩",
                "原始值": "(空)", "问题": "详情缺电机总功率，已用前+后分项功率求和派生",
            })
        elif abs(parts_sum - total_value) > 0.51:
            qc.append({
                "产品型号": model, "字段": "电机总功率/扭矩",
                "原始值": ledger_str(total),
                "问题": "总功率≠前+后分项之和（{} vs {}），按详情总功率为准".format(
                    ledger_str(total), ledger_str(round(parts_sum, 3))),
            })
    mapped["前电机功率/扭矩"] = front
    mapped["后电机功率/扭矩"] = rear
    mapped["电机总功率/扭矩"] = total

    # 发动机仅 PHEV 取（BEV 无发动机列值）
    if ptype == "PHEV" and engine_rows:
        mapped["发动机型号"] = join_component_values([r.get("型号") for r in engine_rows])
        mapped["发动机生产企业"] = join_component_values([r.get("生产企业") for r in engine_rows])
        displacement = join_component_values([r.get("排量") for r in engine_rows])
        power = join_component_values([r.get("功率") for r in engine_rows])
        mapped["发动机排量(mL)"] = num_or_text(displacement) if displacement else None
        mapped["发动机功率(kW)"] = num_or_text(power) if power else None

    # ---- 供应商 QC 守卫
    mapped["电芯供应商"] = supplier_or_none(
        detail.get("电芯供应商_公示其它"), model, "电芯供应商", qc)
    mapped["电池包供应商"] = supplier_or_none(
        detail.get("电池包供应商_公示其它"), model, "电池包供应商", qc)

    # ---- 是否增程(EREV)
    erev = None
    if ptype == "PHEV":
        haystack = [car_name, mapped.get("产品名称"), generic_name]
        if any("增程" in _s(item) for item in haystack):
            erev = 1
    mapped["是否增程(EREV)"] = erev

    # ---- 结构性质检（只记录，不阻断）
    if not motor_rows and any(x is not None for x in (front, rear, total)):
        qc.append({
            "产品型号": model, "字段": "电机组件_公示详情", "原始值": "(空)",
            "问题": "电机组件为空但详情给出电机功率分项：电机型号/生产企业无法映射，留空",
        })

    # 防御：映射键必须 ⊆ 39 列
    extra = set(mapped) - set(HEADER)
    if extra:
        raise SystemExit("FATAL: 映射产出非表头键: {}".format(sorted(extra)))
    return mapped


def supplier_or_none(raw, model, field, qc):
    """供应商守卫：长度 <3 或不含企业关键词 -> 判解析垃圾，留空并记 qc。"""
    text = _s(raw)
    if not text:
        return None
    if len(text) < 3 or not SUPPLIER_KW_RE.search(text):
        qc.append({
            "产品型号": model, "字段": field, "原始值": text,
            "问题": "供应商字段疑似解析垃圾（长度<3 或 不含企业关键词），按 QC 守卫留空",
        })
        return None
    # 已实证的截断形态：品牌首字缺失（「巢能源科技…」应为「蜂巢能源科技…」）
    if "巢能源" in text and not text.startswith("蜂巢"):
        qc.append({
            "产品型号": model, "字段": field, "原始值": text,
            "问题": "疑似品牌首字被截断（「巢能源科技…」疑缺「蜂」）；命中关键词守卫故按规则保留，待人工确认",
        })
    return text


# ---------------------------------------------------------------- 读入
def load_json(path):
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def load_workbook_snapshot(qc_warnings):
    """只读加载工作簿；返回 (header, rows, ledger_header, sha256)；读取期间被改写会追加 warning。"""
    sha_before = sha256_file(WORKBOOK)
    workbook = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
    try:
        if SHEET not in workbook.sheetnames:
            raise SystemExit("FATAL: 缺少 sheet {}".format(SHEET))
        sheet = workbook[SHEET]
        raw_rows = list(sheet.iter_rows(values_only=True))
        ledger_header = []
        if LEDGER_SHEET in workbook.sheetnames:
            ledger_rows = workbook[LEDGER_SHEET].iter_rows(values_only=True)
            first = next(ledger_rows, None)
            ledger_header = list(first) if first else []
    finally:
        workbook.close()
    sha_after = sha256_file(WORKBOOK)
    if sha_before != sha_after:
        qc_warnings.append("工作簿在读取期间被并发修改（sha256 前后不一致），本次快照可能混合两版")
    return list(raw_rows[0]), [list(r) for r in raw_rows[1:]], ledger_header, sha_after


# ---------------------------------------------------------------- 主流程
def build(args):
    qc_warnings = []

    listing = load_json(LIST_PATH)
    details = load_json(DETAIL_PATH)

    header, wb_rows, ledger_header, wb_sha = load_workbook_snapshot(qc_warnings)
    if list(header) != HEADER:
        raise SystemExit(
            "FATAL: 工作簿表头与 39 列口径不符\n实际(共{}列): {}\n期望(共{}列): {}".format(
                len(header), header, len(HEADER), HEADER))
    col_idx = {name: i for i, name in enumerate(header)}

    # ---- 四集合
    list_models = [str(item["产品型号"]).strip() for item in listing]
    detail_models = [str(item["产品型号"]).strip() for item in details]
    if len(set(list_models)) != len(list_models):
        raise SystemExit("FATAL: 清单存在重复产品型号")
    if len(set(detail_models)) != len(detail_models):
        raise SystemExit("FATAL: 详情存在重复产品型号")

    bucket_idx = col_idx["批次"]
    model_idx = col_idx["产品型号"]
    wb_411_models = []
    for row in wb_rows:
        if str(row[bucket_idx]).strip() == str(BATCH):
            wb_411_models.append(str(row[model_idx]).strip())
    if len(set(wb_411_models)) != len(wb_411_models):
        raise SystemExit("FATAL: 表内 411 行存在重复产品型号")

    list_set, wb_set, detail_set = set(list_models), set(wb_411_models), set(detail_models)
    new_set = list_set - wb_set          # 新增行（清单有、表内无）
    enrich_set = list_set & wb_set       # 存量富化
    unmatched_set = detail_set - list_set  # 详情多出、不入行
    missing_detail_set = list_set - detail_set

    if not wb_set <= list_set:
        raise SystemExit("FATAL: 表内 411 型号不在清单内: {}".format(sorted(wb_set - list_set)))
    if missing_detail_set:
        raise SystemExit("FATAL: 清单型号缺详情: {}".format(sorted(missing_detail_set)))

    detail_by_model = {str(item["产品型号"]).strip(): item for item in details}
    wb_411_rows = {}
    for row in wb_rows:
        if str(row[bucket_idx]).strip() == str(BATCH):
            wb_411_rows[str(row[model_idx]).strip()] = row

    qc = []
    norm_log = []

    def apply_norm(model, col, value):
        """统一归一：剥离工况标注（CLTC/WLTC/NEDC/WLTP），并登记被改写格。"""
        new_value, changed = strip_condition_annotation(value)
        if changed:
            norm_log.append({"产品型号": model, "列名": col,
                             "原值": value, "归一值": new_value})
        return new_value

    # ---- 新增行（按清单顺序）
    new_rows = []
    for model in list_models:
        if model not in new_set:
            continue
        mapped = map_detail(detail_by_model[model], qc)
        row = {}
        for col in HEADER:
            if col in SALES_COLS:
                continue  # 绝不产出销量三列
            if col == "批次":
                row[col] = BATCH
            elif col in BLANK_BY_POLICY:
                continue  # 政策留空列：不产出键（键存在即被视为写入）
            else:
                value = mapped.get(col)
                if value is not None:
                    row[col] = apply_norm(model, col, value)
        new_rows.append(row)

    # ---- 富化（只补空，按工作簿行序）
    enrich, ledger = [], []
    enrich_cells = 0
    per_col_stats = {col: {"existing_nonblank": 0, "patched": 0, "still_blank": 0}
                     for col in HEADER}
    for model in wb_411_models:
        row = wb_411_rows[model]
        mapped = map_detail(detail_by_model[model], qc)
        cells = {}
        cell_ledger = []
        for col in HEADER:
            current = row[col_idx[col]]
            if not is_blank(current):
                per_col_stats[col]["existing_nonblank"] += 1
                continue
            if col in ENRICH_EXCLUDE:
                continue  # 不参与富化的列（批次/产品型号/数据来源/销量三列）：只统计存量，不补
            value = apply_norm(model, col, mapped.get(col))
            if is_blank(value):
                per_col_stats[col]["still_blank"] += 1
                continue
            cells[col] = value
            per_col_stats[col]["patched"] += 1
            cell_ledger.append([col, "官方详情页补全", ledger_str(value),
                                mapped.get("官方参数页URL") or ""])
            ledger.append({
                "序号": None,
                "批次": BATCH,
                "产品型号": model,
                "动力类型": mapped.get("动力类型"),
                "字段": col,
                "变更类型": "官方详情页补全",
                "附件1原值": "(空)",
                "合并值": ledger_str(value),
                "数据来源": mapped.get("官方参数页URL") or "",
            })
        enrich_cells += len(cells)
        enrich.append({
            "批次": BATCH,
            "产品型号": model,
            "动力类型": mapped.get("动力类型"),
            "cells": cells,
            "ledger": cell_ledger,
        })

    # ---- 新增行台账（每型号 1 条，序号由应用脚本统一分配 -> None）
    new_ledger = []
    for model in [m for m in list_models if m in new_set]:
        mapped = map_detail(detail_by_model[model], qc)
        new_ledger.append({
            "序号": None,
            "批次": BATCH,
            "产品型号": model,
            "动力类型": mapped.get("动力类型"),
            "字段": "新增行",
            "变更类型": "公告新增行",
            "附件1原值": "(空)",
            "合并值": model,
            "数据来源": "{}；{}".format(DATA_SOURCE_TAG,
                                        mapped.get("官方参数页URL") or ""),
        })
    ledger = new_ledger + ledger

    # ---- unmatched
    unmatched = []
    for model in sorted(unmatched_set):
        item = detail_by_model[model]
        unmatched.append({
            "产品型号": model,
            "原因": "不在46条清单内（详情多出 1 条）",
            "产品名称": _s(item.get("产品名称")) or None,
            "通用名称": _s(item.get("通用名称")) or None,
            "详情页URL": _s(item.get("详情页URL")) or None,
        })
    if unmatched_set != {"BJ6500A602RSHEV"}:
        qc_warnings.append("unmatched 与预期不一致: {}".format(sorted(unmatched_set)))

    # ---- 期望新增清单交叉复核
    expected_vs_actual = {
        "expected_only": sorted(set(EXPECTED_NEW) - new_set),
        "actual_only": sorted(new_set - set(EXPECTED_NEW)),
    }
    if expected_vs_actual["expected_only"] or expected_vs_actual["actual_only"]:
        qc_warnings.append(
            "任务书给定新增 25 型号与实际差集不一致: {}".format(expected_vs_actual))

    # ---- 自检：键 ⊆ 39 列 / 计数 / 禁写列
    for row in new_rows:
        extra = set(row) - set(HEADER)
        if extra:
            raise SystemExit("FATAL: new_rows 出现非表头键: {}".format(sorted(extra)))
        forbidden = set(row) & set(NEVER_EMIT)
        if forbidden:
            raise SystemExit("FATAL: new_rows 出现禁写列: {}".format(sorted(forbidden)))
    for item in enrich:
        extra = set(item["cells"]) - set(HEADER)
        if extra:
            raise SystemExit("FATAL: enrich.cells 出现非表头键: {}".format(sorted(extra)))
        forbidden = set(item["cells"]) & set(NEVER_EMIT)
        if forbidden:
            raise SystemExit("FATAL: enrich.cells 出现禁写列: {}".format(sorted(forbidden)))

    # ---- 自检：工况标注不得残留（口径见 字段映射说明 第 6 行）
    residual = []
    for row in new_rows:
        for col, value in row.items():
            if strip_condition_annotation(value)[1]:
                residual.append((row.get("产品型号"), col, value))
    for item in enrich:
        for col, value in item["cells"].items():
            if strip_condition_annotation(value)[1]:
                residual.append((item["产品型号"], col, value))
    if residual:
        raise SystemExit("FATAL: 输出仍残留工况标注: {}".format(residual[:5]))

    # QC 去重：map_detail 在「新增」「富化」两轮各调用一次，同一疑点会被重复登记
    deduped, seen_qc = [], set()
    for item in qc:
        key = (item["产品型号"], item["字段"], str(item["原始值"]), item["问题"])
        if key in seen_qc:
            continue
        seen_qc.add(key)
        deduped.append(item)
    qc = deduped

    # 系统性特征（属来源特征而非解析缺陷，单独成节，不进 qc_issues 逐行刷屏）
    systematic_notes = []
    dup_brand = sorted({_s(d.get("产品商标")) for d in details if _s(d.get("产品商标")).endswith("牌牌")})
    if dup_brand:
        systematic_notes.append(
            "产品商标重复后缀「牌牌」：{}（官方详情原文如此，按原值落库）".format("、".join(dup_brand)))
    phev_models = [m for m in list_models
                   if power_type_from_fuel(detail_by_model[m].get("燃料种类")) == "PHEV"]
    phev_no_motor = [m for m in phev_models
                     if all(power_num(detail_by_model[m].get(k)) is None
                            for k in ("前电机峰值功率_kW", "后电机峰值功率_kW", "电机总功率_kW"))]
    if phev_models:
        systematic_notes.append(
            "PHEV 详情页不提供电机功率：{}/{} 个 PHEV 型号前/后/总均空（系统性来源特征，非解析缺陷）".format(
                len(phev_no_motor), len(phev_models)))
    no_fuel = [m for m in list_models if is_blank(detail_by_model[m].get("油耗申报值(L/100km)"))]
    if no_fuel:
        systematic_notes.append(
            "油耗申报值(L/100km) 全空：{}/{} 个型号（411 公示详情未给油耗），综合油耗/B 状态油耗留空".format(
                len(no_fuel), len(list_models)))
    no_motor_comp = [m for m in list_models
                     if not (detail_by_model[m].get("电机组件_公示详情") or [])]
    if no_motor_comp:
        with_power = [m for m in no_motor_comp
                      if any(power_num(detail_by_model[m].get(k)) is not None
                             for k in ("前电机峰值功率_kW", "后电机峰值功率_kW", "电机总功率_kW"))]
        systematic_notes.append(
            "电机组件_公示详情 为空：{}/{} 个型号；其中 {} 个仍给出电机功率分项（已逐行记 qc），"
            "其余为 PHEV 系统性缺电机数据".format(
                len(no_motor_comp), len(list_models), len(with_power)))

    counts = {
        "list": len(list_models),
        "details": len(detail_models),
        "workbook_411": len(wb_411_models),
        "new_rows": len(new_rows),
        "enrich_rows": len(enrich),
        "enrich_cells": enrich_cells,
        "ledger": len(ledger),
        "qc_issues": len(qc),
        "unmatched": len(unmatched),
    }
    expected_counts = {"list": 46, "details": 47, "workbook_411": 21,
                       "new_rows": 25, "enrich_rows": 21, "unmatched": 1}
    for key, expected in expected_counts.items():
        if counts[key] != expected:
            raise SystemExit("FATAL: 计数 {} = {}，期望 {}".format(key, counts[key], expected))
    if counts["ledger"] != counts["new_rows"] + counts["enrich_cells"]:
        raise SystemExit("FATAL: 台账条数与 新增行+富化格 不符")

    if ledger_header != LEDGER_HEADER:
        qc_warnings.append("`变更记录` 表头与 9 列口径不符: {}".format(ledger_header))

    # ---- 归一化台账（工况标注剥离）
    norm_by_col = {}
    for item in norm_log:
        norm_by_col[item["列名"]] = norm_by_col.get(item["列名"], 0) + 1
    norm_models = sorted({item["产品型号"] for item in norm_log})
    qc_notes = []
    if norm_log:
        qc_notes.append(
            "工况标注剥离：改写 {} 格（{}）".format(
                len(norm_log),
                "、".join("{}={}".format(c, n) for c, n in sorted(norm_by_col.items()))))
        qc_notes.append("涉及型号 {} 个：{}".format(len(norm_models), "、".join(norm_models)))
        qc_notes.append("逐格明细：" + "；".join(
            "{}.{} {}→{}".format(i["产品型号"], i["列名"], i["原值"], i["归一值"])
            for i in norm_log))
    else:
        qc_notes.append("工况标注剥离：未发现需改写格（0 格）")

    # ---- 详情 `产品名称` 口径审计（47 条：去重取值 / 是否等于通用名称 / 是否非描述串）
    pn_groups = {}
    for item in details:
        key = _s(item.get("产品名称")) or "(空)"
        pn_groups.setdefault(key, []).append(str(item.get("产品型号")).strip())
    pn_equal_generic = []
    pn_non_desc = []
    for item in details:
        model = str(item.get("产品型号")).strip()
        pname, gname = _s(item.get("产品名称")), _s(item.get("通用名称"))
        if pname and pname == gname:
            pn_equal_generic.append(model)
        if pname and not re.search(r"乘用车|轿车|客车|货车|专用车|越野车|挂车", pname):
            pn_non_desc.append({"产品型号": model, "值": pname})
    pn_audit = {
        "details_total": len(details),
        "distinct_count": len(pn_groups),
        "values": [{"值": k, "条数": len(v), "型号": v} for k, v in sorted(pn_groups.items())],
        "equal_to_generic": pn_equal_generic,
        "non_description": pn_non_desc,
        "new_rows_filled": sum(1 for r in new_rows if r.get("产品名称")),
        "new_rows_type_filled": sum(1 for r in new_rows if r.get("产品类型")),
        "new_rows_total": len(new_rows),
    }

    now = datetime.now(TZ_CN)
    meta = {
        "workbook": WORKBOOK,
        "workbook_sha256": wb_sha,
        "header": header,
        "generated_at": now.isoformat(),
        "sources": [
            {"path": LIST_PATH, "sha256": sha256_file(LIST_PATH)},
            {"path": DETAIL_PATH, "sha256": sha256_file(DETAIL_PATH)},
        ],
        "counts": counts,
        "batch": BATCH,
        "data_source": DATA_SOURCE_TAG,
        "normalizations": NORMALIZATIONS,
        "qc_notes": qc_notes,
        "normalization_details": norm_log,
        "detail_product_name_audit": pn_audit,
        "ledger_schema": LEDGER_HEADER,
        "序号": "None 占位：由 Lead 应用脚本统一分配",
        "set_relations": {
            "list_46": len(list_set),
            "workbook_411_21": len(wb_set),
            "new_25": len(new_set),
            "enrich_21": len(enrich_set),
            "details_47": len(detail_set),
            "unmatched_1": len(unmatched_set),
            "new_models": [m for m in list_models if m in new_set],
            "enrich_models": wb_411_models,
            "check_list_eq_new_plus_enrich": (new_set | enrich_set) == list_set,
            "check_new_eq_list_minus_wb": new_set == (list_set - wb_set),
            "check_wb411_subset_of_list": wb_set <= list_set,
            "check_unmatched": unmatched_set == {"BJ6500A602RSHEV"},
            "expected_new_vs_actual": expected_vs_actual,
        },
        "warnings": qc_warnings,
        "systematic_notes": systematic_notes,
        "column_stats": {
            "new": column_fill_stats(new_rows, HEADER),
            "enrich": per_col_stats,
        },
    }
    payload = {
        "meta": meta,
        "new_rows": new_rows,
        "enrich": enrich,
        "ledger": ledger,
        "qc_issues": qc,
        "unmatched": unmatched,
    }
    return payload


def column_fill_stats(rows, header):
    stats = {}
    for col in header:
        stats[col] = sum(1 for row in rows if not is_blank(row.get(col)))
    return stats


# ---------------------------------------------------------------- 报告
def pct(numerator, denominator):
    if not denominator:
        return "-"
    return "{:.0f}%".format(100.0 * numerator / denominator)


def fmt_value(value):
    if value is None:
        return "(空)"
    text = ledger_str(value)
    return text.replace("|", "\\|")


def build_report(payload):
    meta = payload["meta"]
    counts = meta["counts"]
    new_rows = payload["new_rows"]
    enrich = payload["enrich"]
    ledger = payload["ledger"]
    qc = payload["qc_issues"]
    n_new = counts["new_rows"]
    n_enr = counts["enrich_rows"]

    lines = []
    add = lines.append
    add("# CR-17-008 · 411 批正式公告映射 dry-run 报告")
    add("")
    add("- 生成时间：`{}`".format(meta["generated_at"]))
    add("- 工作簿（只读）：`{}`".format(meta["workbook"]))
    add("- 工作簿 sha256：`{}`".format(meta["workbook_sha256"]))
    add("- 权威表头：`{}` 列（按列名映射，未使用列下标）".format(len(meta["header"])))
    add("- 数据来源标签：`{}`".format(meta["data_source"]))
    add("- 落库范围：**dry-run only**（本脚本不写工作簿；新增行/富化格/台账由 Lead 应用脚本落库）")
    add("")

    # 0 结论
    add("## 0. 结论摘要")
    add("")
    add("| 项 | 值 |")
    add("| --- | --- |")
    add("| 清单型号数 | {} |".format(counts["list"]))
    add("| 详情条数 | {} |".format(counts["details"]))
    add("| 表内 411 存量行 | {} |".format(counts["workbook_411"]))
    add("| **新增行** | **{}** |".format(counts["new_rows"]))
    add("| **存量富化行** | **{}** |".format(counts["enrich_rows"]))
    add("| **富化补丁格数** | **{}** |".format(counts["enrich_cells"]))
    add("| 台账条数（新增 1/型号 + 富化 1/格） | {} |".format(counts["ledger"]))
    add("| QC 问题数 | {} |".format(counts["qc_issues"]))
    add("| unmatched（不入行） | {} |".format(counts["unmatched"]))
    add("")

    # 1 集合关系
    rel = meta["set_relations"]
    add("## 1. 四集合关系与差集校验")
    add("")
    add("```")
    add("清单 46（list_411_nev_passenger.json）")
    add(" ├─ 表内已有 21（工作簿 批次=411）  → 富化 21")
    add(" └─ 表内没有 25                    → 新增 25")
    add("详情 47（details_411.json）")
    add(" └─ 额外 1：BJ6500A602RSHEV（不在清单） → unmatched，不入行")
    add("```")
    add("")
    checks = [
        ("清单数 = 46", rel["list_46"] == 46),
        ("详情数 = 47", rel["details_47"] == 47),
        ("表内 411 行 = 21", rel["workbook_411_21"] == 21),
        ("新增 = 46 - 21 = 25", rel["new_25"] == 25 and rel["check_new_eq_list_minus_wb"]),
        ("富化（清单∩表内）= 21", rel["enrich_21"] == 21),
        ("表内 411 型号 ⊆ 清单", rel["check_wb411_subset_of_list"]),
        ("新增 ∪ 富化 = 清单", rel["check_list_eq_new_plus_enrich"]),
        ("unmatched = {BJ6500A602RSHEV}", rel["check_unmatched"]),
        ("任务书给定 25 型号清单与实际差集一致",
         not rel["expected_new_vs_actual"]["expected_only"]
         and not rel["expected_new_vs_actual"]["actual_only"]),
    ]
    add("| 校验 | 结果 |")
    add("| --- | --- |")
    for label, ok in checks:
        add("| {} | {} |".format(label, "PASS" if ok else "**FAIL**"))
    add("")
    add("### 1.1 新增 25 型号（按清单顺序）")
    add("")
    for idx, model in enumerate(rel["new_models"], 1):
        add("{}. `{}`".format(idx, model))
    add("")
    add("### 1.2 富化 21 型号（按工作簿行序）")
    add("")
    for idx, model in enumerate(rel["enrich_models"], 1):
        add("{}. `{}`".format(idx, model))
    add("")
    if rel["expected_new_vs_actual"]["expected_only"] or \
            rel["expected_new_vs_actual"]["actual_only"]:
        add("> 与任务书给定清单的差异：expected_only={} / actual_only={}".format(
            rel["expected_new_vs_actual"]["expected_only"],
            rel["expected_new_vs_actual"]["actual_only"]))
        add("")

    # 2 口径执行
    add("## 2. 落库口径执行情况（逐条对照）")
    add("")
    add("| 口径项 | 执行 |")
    add("| --- | --- |")
    rows = [
        ("批次", "全部写 int `411`（新增行）；富化不改批次"),
        ("数据来源", "新增行 `{}`；存量非空不覆盖".format(meta["data_source"])),
        ("产品名称 / 产品类型", "均取公告描述值（如 `纯电动多用途乘用车`），**禁止填营销名**（Lead p12 先例：两列 111/111 相同）"),
        ("通用名称", "详情原值原样保留，逗号并列**不截断**"),
        ("车型名称", "**= 通用名称整串原样**（Lead p12 先例：409 官方详情队列 100/111、397 官方目录 115/126 两列相同；多并列项整串写入不截断）"),
        ("动力类型", "燃料种类含「纯电动」且不含「混合」→ BEV；含「混合」→ PHEV"),
        ("是否减免购置税", "**留空**（411 存量 21 行亦全空）"),
        ("细分市场", "**留空**（不自动推导）"),
        ("电池能量密度 / 百公里电耗", "**留空**（411 公示详情不提供，不用公式臆造）"),
        ("纯电续航 / 电池容量 / 综合油耗 / B状态油耗 / 发动机排量",
         "取详情原值；多值/带单位保留文本（如 `340/332`、`6.30/6.40`），单一数值写数字；"
         "**工况标注 (CLTC)/(WLTC)/(NEDC)/(WLTP) 一律剥离**（口径见 `字段映射说明` 第 6 行）"),
        ("工况标注归一", "**所有写出值统一剥离** `[（(]CLTC|WLTC|NEDC|WLTP[)）]` → 空串后 strip；"
                     "剥完若为纯数字回归数值类型，多值保持文本；残留即 FATAL 退出（全表 5392 行原生 0 处标注）"),
        ("前/后/总电机功率", "前=前电机峰值功率、后=后电机峰值功率、总=电机总功率；只写功率数字；已校验 总=前+后"),
        ("电机型号 / 电机生产企业", "`电机组件_公示详情` 逐项拼接，`/` 分隔；BEV 解析归类异常时回退 `发动机组件`"),
        ("发动机四列", "仅 PHEV 取 `发动机组件_公示详情`；BEV 留空"),
        ("电芯/电池包供应商", "取 `*_公示其它` + QC 守卫（长度<3 或 无企业关键词 → 留空并记 qc）"),
        ("官方参数页URL", "= 详情页URL"),
        ("媒体校验来源", "**留空**"),
        ("月销量 / 销量数据月份 / 销量更新日期", "**不产出、不触碰**（并发会话在填）"),
        ("是否增程(EREV)", "PHEV 且 车型名称/产品名称/通用名称 含「增程」→ 1，否则留空"),
        ("富化策略", "只补空单元格，**逐格比对非空一律跳过**；每补 1 格 1 条台账；**富化不动 `数据来源` 列**（该列保留行级出处 `第411批公示（车船税目录第九十批）`，逐格证据只进台账「数据来源」字段）"),
    ]
    for key, value in rows:
        add("| {} | {} |".format(key, value))
    add("")

    # 2.1 归一化台账
    add("### 2.1 归一化台账（工况标注剥离，口径 = `字段映射说明` 第 6 行）")
    add("")
    for note in meta.get("qc_notes", []):
        add("- {}".format(note))
    norm_details = meta.get("normalization_details") or []
    if norm_details:
        add("")
        add("| 型号 | 列名 | 原值 | 归一值 |")
        add("| --- | --- | --- | --- |")
        for item in norm_details:
            add("| `{}` | {} | `{}` | `{}` |".format(
                item["产品型号"], item["列名"], item["原值"], item["归一值"]))
    add("")

    # 3 来源映射表
    add("## 3. 来源映射表（详情字段 → 底表列）")
    add("")
    add("| 详情字段 | 底表列 | 处理 |")
    add("| --- | --- | --- |")
    mapping = [
        ("批次", "批次", "常量 411"),
        ("产品型号", "产品型号", "原值"),
        ("产品商标", "产品商标", "原值"),
        ("企业名称", "企业名称", "原值"),
        ("通用名称", "车型名称 / 通用名称", "**整串原样、两列同值**；多并列项不截断（p12 先例）"),
        ("产品名称（公告描述值）", "产品名称 / 产品类型", "同一描述串双列（p12 先例：两列 111/111 相同）；禁止营销名"),
        ("燃料种类", "动力类型", "BEV/PHEV 判定"),
        ("整备质量(kg)", "整备质量(kg)", "数值；多值保留文本"),
        ("轴距(mm)", "轴距(mm)", "数值"),
        ("纯电续航里程(km)", "纯电续航里程(km)", "数值；括号/多值保留文本"),
        ("电池容量(kWh)", "电池容量(kWh)", "数值；多值保留文本"),
        ("电池类型_公示其它", "电池类型", "原值"),
        ("油耗申报值(L/100km)", "综合油耗(L/100km)", "数值；411 详情该字段全空"),
        ("（详情无此字段）", "B状态油耗(L/100km)", "留空"),
        ("前电机峰值功率_kW", "前电机功率/扭矩", "功率数字"),
        ("后电机峰值功率_kW", "后电机功率/扭矩", "功率数字；多值保留文本"),
        ("电机总功率_kW", "电机总功率/扭矩", "功率数字；缺值时用前+后派生并记 qc"),
        ("电机组件_公示详情[型号]", "电机型号", "逐项拼接，`/` 分隔"),
        ("电机组件_公示详情[生产企业]", "电机生产企业", "逐项拼接，`/` 分隔"),
        ("发动机组件_公示详情[排量]", "发动机排量(mL)", "仅 PHEV；数值"),
        ("发动机组件_公示详情[功率]", "发动机功率(kW)", "仅 PHEV；数值"),
        ("发动机组件_公示详情[生产企业]", "发动机生产企业", "仅 PHEV；拼接"),
        ("发动机组件_公示详情[型号]", "发动机型号", "仅 PHEV；拼接"),
        ("车长(mm)", "车长(mm)", "数值"),
        ("电芯供应商_公示其它", "电芯供应商", "原值 + QC 守卫"),
        ("电池包供应商_公示其它", "电池包供应商", "原值 + QC 守卫"),
        ("详情页URL", "官方参数页URL", "原值"),
        ("（口径）", "数据来源", "常量标签"),
        ("（口径）", "是否增程(EREV)", "派生标记"),
        ("（不产出）", "细分市场 / 是否减免购置税 / 电池能量密度 / 百公里电耗 / 媒体校验来源 / 月销量三列", "留空或不产出"),
        ("（未使用）", "车宽/车高/总质量/最高车速/额定载客/动力蓄电池组总质量", "39 列中无对应列，未写入"),
    ]
    for src, dst, note in mapping:
        add("| {} | {} | {} |".format(src, dst, note))
    add("")

    # 3.1 详情 产品名称 口径审计
    pna = meta.get("detail_product_name_audit") or {}
    add("### 3.1 详情 `产品名称` 口径审计（{} 条）".format(pna.get("details_total", 0)))
    add("")
    add("| 去重取值 | 条数 | 型号 |")
    add("| --- | --- | --- |")
    for item in pna.get("values", []):
        add("| `{}` | {} | {} |".format(item["值"], item["条数"], "、".join(item["型号"])))
    add("")
    add("- 去重取值种数：**{}**".format(pna.get("distinct_count", 0)))
    add("- `产品名称` 等于 `通用名称` 的条目：**{} 条**{}".format(
        len(pna.get("equal_to_generic", [])),
        ("（" + "、".join(pna["equal_to_generic"]) + "）") if pna.get("equal_to_generic") else "（无）"))
    add("- 非描述串（不含 乘用车/轿车/客车/货车/专用车/越野车/挂车 任一类别词）的条目：**{} 条**{}".format(
        len(pna.get("non_description", [])),
        ("（" + "；".join("{}=`{}`".format(x["产品型号"], x["值"]) for x in pna["non_description"]) + "）")
        if pna.get("non_description") else "（无）"))
    add("- 新增 25 行 `产品名称` 填充率：**{}/{}**；`产品类型` 填充率：**{}/{}**（两列同值，均取详情 `产品名称` 描述串原样）".format(
        pna.get("new_rows_filled", 0), pna.get("new_rows_total", 0),
        pna.get("new_rows_type_filled", 0), pna.get("new_rows_total", 0)))
    add("- 口径澄清：新增行的 `产品名称` / `产品类型` **自始即为公告描述串**（如 `纯电动多用途乘用车`），"
        "并非通用名称；`车型名称` 才是 `通用名称` 整串。存量 411 行出现 `红旗HS7` 等值是该行**原有**的 产品名称/产品类型 "
        "（本轮遵循「不覆盖非空」未动），其富化补丁只写了 `车型名称`。")
    add("")

    # 4 新增行填充率
    add("## 4. 新增 {} 行逐字段填充率（分母 {}）".format(n_new, n_new))
    add("")
    new_stats = meta["column_stats"]["new"]
    add("| 列名 | 非空数 | 填充率 |")
    add("| --- | --- | --- |")
    for col in meta["header"]:
        add("| {} | {}/{} | {} |".format(col, new_stats.get(col, 0), n_new,
                                         pct(new_stats.get(col, 0), n_new)))
    add("")
    add("> 政策留空列（细分市场 / 是否减免购置税 / 电池能量密度 / 百公里电耗 / 媒体校验来源）"
        "与销量三列（不产出）填充率为 0 属预期。")
    add("")

    # 5 富化统计
    add("## 5. 存量富化 {} 行逐字段统计".format(n_enr))
    add("")
    add("| 列名 | 原有非空 | 本轮补丁 | 富化后非空 | 仍空 |")
    add("| --- | --- | --- | --- | --- |")
    enr_stats = meta["column_stats"]["enrich"]
    for col in meta["header"]:
        stat = enr_stats[col]
        after = stat["existing_nonblank"] + stat["patched"]
        add("| {} | {} | {} | {} | {} |".format(
            col, stat["existing_nonblank"], stat["patched"], after, stat["still_blank"]))
    add("")
    add("> 「原有非空」= 存量 21 行该列已有值（**一律未覆盖**）；「本轮补丁」= 实际写入格数。")
    add("")
    add("### 5.1 富化明细（型号 → 补丁列）")
    add("")
    add("| 型号 | 补丁格数 | 补丁列 |")
    add("| --- | --- | --- |")
    for item in enrich:
        cols = "、".join(item["cells"].keys()) if item["cells"] else "（无）"
        add("| `{}` | {} | {} |".format(item["产品型号"], len(item["cells"]), cols))
    add("")

    # 6 台账
    add("## 6. 台账（`变更记录`，9 列）")
    add("")
    add("| 序号 | 批次 | 产品型号 | 动力类型 | 字段 | 变更类型 | 附件1原值 | 合并值 | 数据来源 |")
    add("| --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for item in ledger:
        add("| {} | {} | `{}` | {} | {} | {} | {} | {} | {} |".format(
            "None" if item["序号"] is None else item["序号"],
            item["批次"], item["产品型号"], item["动力类型"] or "",
            item["字段"], item["变更类型"], item["附件1原值"],
            fmt_value(item["合并值"]), fmt_value(item["数据来源"])))
    add("")
    add("> `序号` 为 `None` 占位，由 Lead 的应用脚本统一分配（延续 `变更记录` 现有最大序号）。"
        "新增行台账 1 型号 1 条（字段=`新增行`/变更类型=`公告新增行`/合并值=产品型号），"
        "富化 1 格 1 条（变更类型=`官方详情页补全`）。")
    add("")

    # 7 QC
    add("## 7. QC 问题清单（共 {} 条）".format(len(qc)))
    add("")
    if qc:
        add("| # | 产品型号 | 字段 | 原始值 | 问题 |")
        add("| --- | --- | --- | --- | --- |")
        for idx, issue in enumerate(qc, 1):
            add("| {} | `{}` | {} | {} | {} |".format(
                idx, issue["产品型号"], issue["字段"],
                fmt_value(issue["原始值"]), issue["问题"]))
    else:
        add("（无）")
    add("")
    if meta["warnings"]:
        add("### 7.1 脚本级警告")
        add("")
        for warning in meta["warnings"]:
            add("- {}".format(warning))
        add("")
    if meta.get("systematic_notes"):
        add("### 7.2 系统性来源特征（非逐行缺陷，故不进 qc_issues）")
        add("")
        for note in meta["systematic_notes"]:
            add("- {}".format(note))
        add("")

    # 8 unmatched
    add("## 8. unmatched（详情多出、不入行）")
    add("")
    add("| 产品型号 | 原因 | 产品名称 | 通用名称 | 详情页 |")
    add("| --- | --- | --- | --- | --- |")
    for item in payload["unmatched"]:
        add("| `{}` | {} | {} | {} | {} |".format(
            item["产品型号"], item["原因"], item.get("产品名称") or "(空)",
            item.get("通用名称") or "(空)", item.get("详情页URL") or ""))
    add("")

    # 9 open questions
    add("## 9. Open questions（只记录，不改规则）")
    add("")
    add("1. **车型名称已按 Lead p12 先例裁定为「通用名称整串原样」**（409 官方详情队列 100/111、397 官方目录 115/126 "
        "两列完全相同；多并列项整串写入如 `比亚迪唐,比亚迪唐Pro`，截断会丢别名）。"
        "**已裁定结论（Lead 2026-10-19，不再变更）：`通用名称` 为空的 4 个新型号"
        "（CC2040BV24DPHEV / BMW6495CABEV / DFM7002S1F4BSEV / CC2040BV24APHEV）车型名称一律留空，"
        "不做「产品商标去牌」等兜底**——商标去牌只会产出品牌级噪声（且表内 车型名称 池已存在 `比亚迪`、`吉利车型` 等噪声值），"
        "本轮不再新增同类噪声。")
    add("2. **动力类型判定依据**：仅按 `燃料种类` 字面判定（含「纯电动」且不含「混合」→ BEV；含「混合」→ PHEV）；"
        "`换电式纯电动轿车`（DFM7002S1F4BSEV）按纯电归 BEV，是否需要单列「换电」子类未定。")
    add("3. **存量 411 行 产品名称 / 产品类型 现有值与 39 列语义不符**（如 `红旗HS7`、`蓝山` 被写入 产品名称/产品类型），"
        "本轮遵循「不覆盖非空」只补空，未纠正；是否在后续专轮统一纠正为类型语义值待定。")
    add("4. **新增行 综合油耗 / B状态油耗 全空**：`details_411.json` 的 `油耗申报值(L/100km)` 47 条全为 null，"
        "PHEV 2 个新型号（CC2040BV24APHEV / CC2040BV24DPHEV）油耗仍缺，需车船税目录/能耗源后续补。")
    add("5. **电芯/电池包供应商覆盖率极低**：47 条详情仅 CC6510DG21HABEV 有值，且 `电池包供应商=\"蜂\"` 被判解析垃圾留空、"
        "`电芯供应商` 首项疑似截断（`巢能源科技…`）。`公示详情原始其它` 自由文本里常含「单体及总成生产企业为…」"
        "（如 SQR6450BEVT1VA3 的合肥国轩高科），本轮按既定口径**未做自由文本再解析**，是否开专轮解析待定。")
    add("6. **`电机型号` / `电机生产企业` 只覆盖 `电机组件_公示详情` 列出的电机**：如 LZ6510VLD1EV / LZ6510VLC1EV 为双电机"
        "（前 55 / 后 125 / 总 180kW），但 `电机组件_公示详情` 仅 1 条 `TZ180XS161`（125kW，对应后电机），"
        "前电机型号无来源故 电机型号 两列只写 1 项；是否从 `公示详情原始其它` 补前电机型号待定。")
    add("7. **BEV 发动机组件误归类**：LZ6460MJD0EV 的电机信息落在 `发动机组件_公示详情`（排量=0），"
        "已按 BEV 回退规则取用并记 qc；解析器口径需上游修正。")
    add("8. **工作簿为并发只读快照**：本次读取 sha256 见报告头，若应用脚本落库前工作簿被其它会话（销量列）更新，"
        "需用新 sha256 复核表头（按列名映射不受追加列影响）。")
    add("")

    # 10 复核
    add("## 10. 复核命令")
    add("")
    add("```powershell")
    add('cd "D:\\#AI\\00 Project\\17 TA analysis\\01-主仓库\\TA-Analysis"')
    add("$env:PYTHONIOENCODING='utf-8'")
    add("python cr17_008_rows_411.py            # 重新生成 dry-run JSON + 本报告（只读工作簿）")
    add("$d = Get-Content -Raw -Encoding utf8 'audit-output/cr17_008_411_dryrun_20261019.json' | ConvertFrom-Json")
    add("$d.meta.counts | Format-List")
    add("$d.meta.set_relations.new_models -join ', '")
    add("$d.meta.normalizations; $d.meta.qc_notes        # 工况标注剥离留痕")
    add("$vals = @($d.new_rows | ForEach-Object { $_.PSObject.Properties.Value }) + "
        "@($d.enrich | ForEach-Object { $_.cells.PSObject.Properties.Value })")
    add("($vals | Where-Object { $_ -match 'CLTC|WLTC|NEDC|WLTP' }).Count   # 期望 0（写出值零残留工况标注）")
    add("```")
    add("")
    add("自检项（脚本内硬校验，任一不满足即退出码非 0）：")
    add("")
    add("- 工作簿表头与 39 列口径**逐名逐序**一致；")
    add("- `new_rows` / `enrich[].cells` 所有键 ⊆ 39 列名；")
    add("- 计数硬校验：清单 46 / 详情 47 / 表内 411 = 21 / 新增 25 / 富化 21 / unmatched 1；")
    add("- 禁写列硬校验：`new_rows` 与 `enrich[].cells` 均不含销量三列与政策留空列；")
    add("- 集合硬校验：新增 = 清单−表内、新增∪富化 = 清单、表内 ⊆ 清单、unmatched = {BJ6500A602RSHEV}；")
    add("- 台账条数 = 新增行数 + 富化格数。")
    add("")
    add("## 11. 边界与未触碰声明")
    add("")
    add("- 本脚本对 `{}` **只读**（`openpyxl.load_workbook(..., read_only=True)`），未写工作簿任何单元格。".format(WORKBOOK_NAME))
    add("- 未创建/修改/删除任何其它仓库文件；仅写本报告与 dry-run JSON 两个交付物"
        "（外加脚本自身 `cr17_008_rows_411.py`）。")
    add("- 未触碰 `月销量(辆)` / `销量数据月份` / `销量更新日期` 三列（dry-run 中亦不产出键）。")
    add("- `new_rows` 每行**只含非空列的键**（留空列不产键，应用脚本按缺失=留空处理）；"
        "`enrich[].cells` 只含实际补丁列（键缺失=不写入）；台账 `ledger` 为完整 9 列。")
    add("- 所有写出值已统一剥离工况标注 (CLTC)/(WLTC)/(NEDC)/(WLTP)（`meta.normalizations`、"
        "`meta.qc_notes`、`meta.normalization_details` 留痕；脚本内含「残留即 FATAL」自检）。")
    add("- 未改 git、未联网、未运行 sync.ps1。")
    add("")
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------- 入口
def main(argv=None):
    parser = argparse.ArgumentParser(description="CR-17-008 · 411 批映射器（dry-run only）")
    parser.add_argument("--dry-run", action="store_true", default=True,
                        help="默认行为：只产出 JSON/MD，不写工作簿（无 --apply 分支）")
    parser.add_argument("--out-json", default=OUT_JSON, help="dry-run JSON 输出路径")
    parser.add_argument("--out-md", default=OUT_MD, help="报告输出路径")
    args = parser.parse_args(argv)

    if not os.path.exists(WORKBOOK):
        raise SystemExit("FATAL: 工作簿不存在: {}".format(WORKBOOK))
    for path in (LIST_PATH, DETAIL_PATH):
        if not os.path.exists(path):
            raise SystemExit("FATAL: 输入不存在: {}".format(path))

    payload = build(args)
    meta = payload["meta"]
    counts = meta["counts"]

    # 写 JSON
    os.makedirs(os.path.dirname(args.out_json), exist_ok=True)
    with open(args.out_json, "w", encoding="utf-8", newline="\n") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")

    # 写报告
    report = build_report(payload)
    with open(args.out_md, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(report)

    # stdout 只输出 ASCII，规避 Windows 控制台 charmap 编码问题
    print("CR-17-008 411 dry-run OK")
    print("list={list} details={details} workbook_411={workbook_411} new_rows={new_rows} "
          "enrich_rows={enrich_rows} enrich_cells={enrich_cells} ledger={ledger} "
          "qc_issues={qc_issues} unmatched={unmatched}".format(**counts))
    print("json -> {}".format(args.out_json))
    print("md   -> {}".format(args.out_md))
    return 0


if __name__ == "__main__":
    sys.exit(main())
