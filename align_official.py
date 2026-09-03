#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
官方渠道横向对齐 v4.0
====================
把《NEV公告参数汇总表_合并版（341~410批）.xlsx》与官方渠道横向拉齐：

1. 第 409 批官方对齐：原底表只有 31 行"无产品型号"的 409 公示名称，
   用 master_export_fixed（官方主链 111 条）+ 工信部公告参数查询系统详情
   （nev-announcement-hub seed：896 条，含电机/发动机/电芯/电池包供应商与详情URL）
   替换为完整的 111 条官方 409 记录（只增不覆盖，逐项记入「参数对齐记录」）。
2. 官方主链补空：按 (批次, 产品型号) 用 master_export_fixed 补其余空位（只补空）。
3. 新增列：电芯供应商 / 电池包供应商 / 官方参数页URL（官方 409 详情，A 级）。
4. 懂车帝已在 2026-09-03 尝试直连，触发滑块验证码；按项目守则不绕过，
   本脚本只使用官方渠道与已存证据，不写入媒体候选值。

用法：
  python align_official.py [--workbook <底表.xlsx>] [--data-root <02-数据底表>/data/nev-announcements]
                           [--hub-seed <miit-2026-21.json>] [--dry-run]
Copyright © 2026 David YE
"""
import sys
import os
import csv
import json
import glob
import shutil
import argparse
import datetime as dt

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
DEFAULT_DATA_ROOT = os.path.join(PROJECT_ROOT, "02-数据底表", "data", "nev-announcements")
DEFAULT_HUB_SEED = os.path.join(PROJECT_ROOT, "03-应用工具", "nev-announcement-hub",
                                "data", "seed", "miit-2026-21.json")

FILL_HUB = PatternFill("solid", fgColor="FCE4D6")   # 官方 409 详情新增/补全（浅橙）
FILL_MASTER = PatternFill("solid", fgColor="C6EFCE")  # 官方主链补空（浅绿，沿用共识色）
RECORD_SHEET = "参数对齐记录"

MISSING = {"", "/", "-", "--", "—", "N/A", "NA", "0", "未填报", "无", "无信息",
           "待查", "待核实", "待确认", "未知", "？", "?", "None", "未提供", "不适用"}


def norm(v):
    if v is None:
        return ""
    return str(v).strip()


def is_missing(v):
    s = norm(v)
    return s == "" or s in MISSING


def load_master(data_root):
    p = os.path.join(data_root, "db", "master_export_fixed.csv")
    with open(p, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def load_hub(seed_path):
    with open(seed_path, encoding="utf-8") as fh:
        d = json.load(fh)
    vehicles = d.get("vehicles", [])
    by_model = {}
    for v in vehicles:
        m = norm(v.get("productModel"))
        if m:
            by_model[m] = v
    return vehicles, by_model


def resolve_segment(product_name, common_name):
    """粗粒度细分市场推断（不造精确级别），仅用于官方 409 新行；无法确定返回 None。"""
    s = norm(product_name) + norm(common_name)
    if "运动型" in s:
        return "Sports"
    if "轿车" in s:
        return "Car"
    if "多用途" in s:
        return "SUV"
    if "MPV" in s.upper() or "多用途" in s:
        return "MPV"
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=None)
    ap.add_argument("--data-root", default=DEFAULT_DATA_ROOT)
    ap.add_argument("--hub-seed", default=DEFAULT_HUB_SEED)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    wb_path = args.workbook or sorted(glob.glob(os.path.join(HERE, "NEV公告参数汇总表_合并版*.xlsx")))[-1]
    if args.dry_run:
        print(f"[dry-run] workbook={wb_path}\ndata-root={args.data_root}\nhub-seed={args.hub_seed}")

    master = load_master(args.data_root)
    hub_by_model = load_hub(args.hub_seed)[1]
    print(f"master rows={len(master)}  hub models={len(hub_by_model)}")

    # backup
    if not args.dry_run:
        backup_dir = os.path.join(PROJECT_ROOT, "04-历史归档", "TA-Analysis仓库剥离件")
        os.makedirs(backup_dir, exist_ok=True)
        bak = os.path.join(backup_dir, "底表_341-410_v3.9_对齐前备份.xlsx")
        shutil.copy2(wb_path, bak)
        print(f"backup -> {bak}")

    wb = load_workbook(wb_path)
    ws = wb["NEV公告参数汇总"]
    headers = [norm(c.value) if c.value is not None else "" for c in ws[1]]
    hi = {h: i for i, h in enumerate(headers)}
    ncol = ws.max_column

    # 新增列（尾部追加，不破坏既有列位）
    new_cols = []
    for name in ("电芯供应商", "电池包供应商", "官方参数页URL"):
        if name not in hi:
            new_cols.append(name)
            hi[name] = ncol  # 0-based row索引
            ws.cell(row=1, column=ncol + 1, value=name)
            ncol += 1
    # 读取现有数据行（保留原行对象为 tuple/list）
    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if any(v not in (None, "") for v in row):
            rows.append(list(row))
    print(f"rows before={len(rows)}  cols={ncol}")

    # 删除 409 无产品型号的公示行（将被官方 111 条替换）
    removed = []
    keep_rows = []
    for r in rows:
        b = norm(r[hi["批次"]]) if hi["批次"] < len(r) else ""
        m = norm(r[hi["产品型号"]]) if hi["产品型号"] < len(r) else ""
        if b == "409" and is_missing(m):
            removed.append(dict(zip(headers, r)))
        else:
            keep_rows.append(r)
    rows = keep_rows
    print(f"removed 409 name-only rows={len(removed)}")

    records = []  # 参数对齐记录行
    for r in removed:
        records.append(["移除", "409公示(无型号)", "", r.get("企业名称", ""), r.get("车型名称", ""), "", "由官方 409（111 条）替代", "https://www.miit.gov.cn/zwgk/zcwj/wjfb/gg/art/2026/art_00d965bfa9ea4bc89bc7613cf911c896.html"])

    # 构造官方 409 新行
    def cell_of(r, name):
        i = hi.get(name)
        if i is None or i >= len(r):
            return None
        return r[i]

    added = 0
    m409 = [m for m in master if norm(m.get("批次")) == "409"]
    for m in m409:
        model = norm(m.get("产品型号"))
        hub = hub_by_model.get(model, {})
        row = [None] * ncol
        for h in headers:
            i = hi[h]
            val = m.get(h)
            if is_missing(val) and hub:
                hmap = {"通用名称": "commonName", "车型名称": "commonName", "产品名称": "productName",
                        "电池类型": "batteryMaterial", "电机峰值功率(kW)": "motorPeakPowerText",
                        "电机总功率(kW)": "motorTotalPowerText", "电机生产企业": "motorSupplier",
                        "电机型号": "motorModel", "发动机排量(mL)": "engineDisplacementText",
                        "发动机功率(kW)": "enginePowerText", "发动机生产企业": "engineSupplier",
                        "纯电续航里程(km)": "electricRangeMinKm", "电池容量(kWh)": "batteryCapacityMinKwh",
                        "产品商标": "brand", "动力类型": "powertrain"}
                if h in hmap:
                    val = hub.get(hmap[h])
            row[i] = val
        # 新列
        row[hi["电芯供应商"]] = hub.get("cellSupplier")
        row[hi["电池包供应商"]] = hub.get("packSupplier")
        row[hi["官方参数页URL"]] = hub.get("detailUrl")
        # 数据来源
        src = "工信部公告参数查询系统（第409批官方详情）" if hub.get("detailUrl") else "工信部主链(master_export)"
        row[hi["数据来源"]] = src
        # 细分市场粗推断（不覆盖，仅空时）
        if is_missing(cell_of(row, "细分市场")):
            row[hi["细分市场"]] = resolve_segment(cell_of(row, "产品名称"), cell_of(row, "通用名称"))
        rows.append(row)
        added += 1
        records.append(["新增", "409官方", model,
                        row[hi["企业名称"]] if hi["企业名称"] < len(row) else "",
                        row[hi["车型名称"]] if hi["车型名称"] < len(row) else "",
                        row[hi["通用名称"]] if hi["通用名称"] < len(row) else "",
                        f"官方详情 {model}", hub.get("detailUrl", "")])
    print(f"added official 409 rows={added}")
    added_start = len(rows) - added

    # 官方主链按 (批次, 产品型号) 补空（只补空）
    mkey = {}
    for m in master:
        k = (norm(m.get("批次")), norm(m.get("产品型号")))
        if k[0].isdigit() and k[1]:
            mkey.setdefault(k, m)
    master_fills = 0
    for r in rows:
        k = (norm(r[hi["批次"]]), norm(r[hi["产品型号"]]))
        m = mkey.get(k)
        if not m:
            continue
        for h in headers:
            i = hi[h]
            if i >= len(r) or is_missing(r[i]):
                val = m.get(h)
                if not is_missing(val):
                    r[i] = val
                    master_fills += 1
    print(f"master chain empty-cells filled={master_fills}")

    # 写回数据行（先删原有数据行，再逐行写入；顶部表头与样式保留）
    # 说明：delete_rows 会移除 2..max 行；为保样式我们先记录数据，随后重建数据区。
    if not args.dry_run:
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        for r in rows:
            ws.append(r)
        # 官方 409 新增行着浅橙，便于复核；主链补空格不着色（已有共识色约定）
        for off in range(added):
            for c in range(1, ncol + 1):
                ws.cell(row=2 + added_start + off, column=c).fill = FILL_HUB
        # 参数对齐记录
        if RECORD_SHEET in wb.sheetnames:
            del wb[RECORD_SHEET]
        rs = wb.create_sheet(RECORD_SHEET)
        rec_head = ["动作", "来源批次", "产品型号", "企业名称", "车型名称", "通用名称", "说明", "证据URL"]
        rs.append(rec_head)
        for line in records:
            rs.append(line)
        wb.save(wb_path)
        print(f"saved rows={len(rows)} total(含表头)={len(rows)+1}")

    # 汇总
    print("=== 对齐汇总 ===")
    print(f"最终数据行={len(rows)}（原 4414 - 移除 {len(removed)} + 新增 {added}）")
    print(f"官方主链补空格数={master_fills}")
    print(f"参数对齐记录条数={len(records)}")


if __name__ == "__main__":
    main()
