# -*- coding: utf-8 -*-
"""CR-17-011: 把 411（正式公告）/412（公示）行写入权威底表。

设计要点
- **按列名映射**：并发会话正在尾插"销量"类列（39→41），本脚本对列数自适应，绝不按下标盲写。
- **只追加、只补空**：新增行追加到末尾；既有 411 行只补空值单元格，非空一律跳过并记录。
- **落盘前强制哈希守卫**：--expect-sha 必须等于当前工作簿 sha256（防并发写入被静默覆盖）。
- **落盘前自动备份**：archive/v4.9.32_pre-411412_backup_<YYYYMMDD>.xlsx。
- 保真预检：本工作簿无公式/无图片/无图表（p13 实证），openpyxl 往返安全。

用法
  python cr17_011_apply_rows.py --check
  python cr17_011_apply_rows.py --apply --expect-sha <sha256>
"""
import argparse, copy, glob, hashlib, io, json, os, re, shutil, sys, time
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 工况标注剥离（依据：字段映射说明 第6行「单元格不再保留 (CLTC)/(WLTC) 标注」；
# 实测全表 纯电续航里程 含标注 = 0 行）——写入前统一剥离并逐格留痕
COND_RE = re.compile(r"[（(]\s*(?:CLTC|WLTC|NEDC|WLTP)\s*[)）]")
# 并发会话（销量工作流）自有列：本脚本绝不写入
PEER_COLS = ("月销量(辆)", "销量数据月份", "销量更新日期")
# 新增行必须留空的列（对齐 410 正式队列先例）
MUST_EMPTY = ("是否减免购置税", "细分市场", "电池能量密度(Wh/kg)", "百公里电耗(kWh/100km)")
NORMALIZATIONS = []

HERE = os.path.dirname(os.path.abspath(__file__))
WB_NAME = "NEV公告参数汇总表_合并版（341~410批）.xlsx"
WB = os.path.join(HERE, WB_NAME)
SHEET_MAIN = "NEV公告参数汇总"
SHEET_LEDGER = "变更记录"
SHEET_TIME = "批次时间表"
SHEET_MAP = "字段映射说明"
ARCHIVE = os.path.join(HERE, "archive")
OUT_DIR = os.path.join(HERE, "audit-output")

FILL_NEW = PatternFill("solid", start_color="00C6EFCE")     # 绿＝补充/官方渠道核实补充（对齐 410 正式队列）
FILL_ER = PatternFill("solid", start_color="00DDEBF7")      # 浅蓝＝增程标记

SRC_411 = "411批正式公告（工信部公告参数查询系统详情页）"
SRC_412 = "412批公示（工信部公示详情页＋车船税目录第九十一批拟定清单）"

LOG = []


def log(*a):
    line = " ".join(str(x) for x in a)
    LOG.append(line)
    print(line)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def is_empty(v):
    return v is None or (isinstance(v, str) and v.strip() in ("", "None"))


def batch_norm(v):
    """批次归一：int / '411' / 411.0 -> int"""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    if s.isdigit():
        return int(s)
    try:
        return int(float(s))
    except Exception:
        return s


def strip_conditions(v, where):
    """剥离 (CLTC)/(WLTC)/(NEDC)/(WLTP) 标注；返回 (新值, 是否改写)。"""
    if not isinstance(v, str):
        return v, False
    new = COND_RE.sub("", v).strip()
    if new == v:
        return v, False
    NORMALIZATIONS.append({"at": where, "from": v, "to": new})
    try:
        if new.isdigit():
            return int(new), True
        f = float(new)
        return (int(f) if f.is_integer() else f), True
    except Exception:
        return new, True


def load_dryruns():
    outs = []
    for pat in ("cr17_008_*_dryrun_*.json", "cr17_009_*_dryrun_*.json"):
        hits = sorted(glob.glob(os.path.join(OUT_DIR, pat)))
        if not hits:
            sys.exit(f"[FATAL] 缺少 dry-run 产物：audit-output/{pat}")
        outs.append(hits[-1])
    res = []
    for p in outs:
        with open(p, encoding="utf-8") as f:
            d = json.load(f)
        d["_path"] = p
        res.append(d)
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--expect-sha", default=None)
    ap.add_argument("--expect-min-cols", type=int, default=39)
    ap.add_argument("--sales-status", default="未匹配-命名待锚定",
                    help="并发销量管线尾插列「销量覆盖状态」的占位值：新增行未经该管线匹配，"
                         "置为该域中的「未匹配」类，待其复跑重判；传空串则留空")
    ap.add_argument("--wb", default=None, help="测试用：覆盖工作簿路径（默认权威底表）")
    ap.add_argument("--no-backup", action="store_true", help="测试用：跳过备份")
    args = ap.parse_args()
    if not (args.apply or args.check):
        ap.error("需要 --check 或 --apply")

    global WB, WB_NAME, ARCHIVE
    if args.wb:
        WB = os.path.abspath(args.wb)
        WB_NAME = os.path.basename(WB)
        ARCHIVE = os.path.join(os.path.dirname(WB), "_test_archive")
        log(f"[TEST] 使用替代工作簿 {WB}")

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    wb_sha = sha256(WB)
    log(f"[i] 工作簿 {WB_NAME}")
    log(f"[i] sha256 = {wb_sha}")
    log(f"[i] mtime  = {datetime.fromtimestamp(os.path.getmtime(WB)):%Y-%m-%d %H:%M:%S}")

    dry = load_dryruns()
    for d in dry:
        m = d.get("meta", {})
        log(f"[i] dry-run {os.path.basename(d['_path'])}: rows={len(d.get('new_rows', []))} "
            f"enrich={len(d.get('enrich', []))} ledger={len(d.get('ledger', []))} "
            f"qc={len(d.get('qc_issues', []))} wb_sha={m.get('workbook_sha256', '?')[:12]}")

    if args.apply:
        if not args.expect_sha:
            sys.exit("[FATAL] --apply 必须带 --expect-sha")
        if wb_sha != args.expect_sha:
            sys.exit(f"[FATAL] 工作簿已被并发修改：现 {wb_sha} ≠ 期望 {args.expect_sha}。"
                     f"重新拉取基线后重跑（不要强写）。")
        for d in dry:
            if d.get("meta", {}).get("workbook_sha256") not in (None, wb_sha):
                log(f"[!] 注意：{os.path.basename(d['_path'])} 的基线 sha 与当前不一致（已按当前列名映射继续）")

    wb = load_workbook(WB)  # data_only=False：本表无公式（实测），保持原样即可
    ws = wb[SHEET_MAIN]
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    idx = {h: i + 1 for i, h in enumerate(header) if h}
    ncol = ws.max_column
    log(f"[i] 主表 {ws.max_row - 1} 数据行 × {ncol} 列")
    for need in ("批次", "产品型号", "数据来源", "动力类型"):
        if need not in idx:
            sys.exit(f"[FATAL] 主表缺少必需列 {need}")
    if ncol < args.expect_min_cols:
        log(f"[!] 列数 {ncol} < 预期 {args.expect_min_cols}")

    # 既有索引
    existing = {}
    for r in range(2, ws.max_row + 1):
        b = batch_norm(ws.cell(row=r, column=idx["批次"]).value)
        m = ws.cell(row=r, column=idx["产品型号"]).value
        existing.setdefault((b, str(m).strip() if m is not None else None), []).append(r)
    log(f"[i] 既有行索引 {len(existing)} 键；批次=411 行数 "
        f"{sum(len(v) for k, v in existing.items() if k[0] == 411)}")

    # ---- 组装写入计划 ----
    plan_new, skipped_dup = [], []
    for d in dry:
        for row in d.get("new_rows", []):
            keys = set(row) - set(idx)
            if keys:
                sys.exit(f"[FATAL] dry-run 出现未知列名：{sorted(keys)}")
            clean = {}
            for k, v in row.items():
                if k in PEER_COLS and not is_empty(v):
                    sys.exit(f"[FATAL] dry-run 试图写并发会话列 {k}：{row.get('产品型号')}")
                if k in MUST_EMPTY and not is_empty(v):
                    sys.exit(f"[FATAL] {k} 按口径必须留空（410 正式队列先例）：{row.get('产品型号')}={v!r}")
                nv, _ = strip_conditions(v, f"new:{row.get('产品型号')}:{k}")
                clean[k] = nv
            if is_empty(clean.get("数据来源")):
                sys.exit(f"[FATAL] 新增行缺 数据来源：{clean.get('产品型号')}")
            cn, gn = clean.get("车型名称"), clean.get("通用名称")
            if not is_empty(cn) and not is_empty(gn) and str(cn) != str(gn):
                sys.exit(f"[FATAL] 车型名称≠通用名称（口径须整串一致）：{clean.get('产品型号')} {cn!r} != {gn!r}")
            b = batch_norm(clean.get("批次"))
            if b not in (411, 412):
                sys.exit(f"[FATAL] 新增行批次异常：{clean.get('产品型号')} -> {b!r}")
            m = str(clean.get("产品型号")).strip()
            if (b, m) in existing:
                skipped_dup.append((b, m))
                continue
            plan_new.append(clean)

    plan_enrich = []   # (row_no, col_name, value, batch, model, power_type)
    for d in dry:
        for e in d.get("enrich", []):
            b = batch_norm(e.get("批次"))
            m = str(e.get("产品型号")).strip()
            rows = existing.get((b, m), [])
            if not rows:
                log(f"[!] 富化目标不存在，跳过：批次{b} 型号{m}")
                continue
            r = rows[0]
            for cname, val in (e.get("cells") or {}).items():
                if cname not in idx:
                    sys.exit(f"[FATAL] 富化列名未知：{cname}")
                if cname in PEER_COLS:
                    sys.exit(f"[FATAL] 富化试图写并发会话列 {cname}：{m}")
                if is_empty(val):
                    continue
                val, _ = strip_conditions(val, f"enrich:{m}:{cname}")
                cur = ws.cell(row=r, column=idx[cname]).value
                if not is_empty(cur):
                    continue  # 只补空
                if cname == "车型名称":
                    gn = ws.cell(row=r, column=idx["通用名称"]).value
                    if not is_empty(gn) and str(gn).strip() != str(val).strip():
                        sys.exit(f"[FATAL] 富化 车型名称≠行内通用名称：{m} {val!r} != {gn!r}")
                pt = ws.cell(row=r, column=idx["动力类型"]).value or e.get("动力类型")
                plan_enrich.append((r, cname, val, b, m, pt))

    # 台账序号起点
    wsl = wb[SHEET_LEDGER]
    mx = 0
    for r in range(2, wsl.max_row + 1):
        v = wsl.cell(row=r, column=1).value
        if isinstance(v, (int, float)):
            mx = max(mx, int(v))
    log(f"[i] 变更记录 {wsl.max_row - 1} 行，最大序号 {mx}；计划 新增行 {len(plan_new)}"
        f" / 富化格 {len(plan_enrich)} → 序号 {mx + 1}..{mx + len(plan_new) + len(plan_enrich)}")
    if skipped_dup:
        log(f"[i] 已存在（幂等跳过）{len(skipped_dup)} 行：{skipped_dup[:5]}")

    n_er = (sum(1 for r, c, v, *_ in plan_enrich if c == "是否增程(EREV)" and str(v) == "1")
            + sum(1 for r in plan_new if str(r.get("是否增程(EREV)")) == "1"))
    log(f"[i] 其中 是否增程(EREV)=1 的新增格 {n_er}")

    if not args.apply:
        log("[i] --check 通过（未写入）。用 --apply --expect-sha %s 落盘。" % wb_sha)
        with open(os.path.join(OUT_DIR, f"cr17_011_check_{ts}.json"), "w", encoding="utf-8") as f:
            json.dump({"mode": "check", "workbook_sha256": wb_sha, "new_rows": len(plan_new),
                       "enrich_cells": len(plan_enrich), "skipped_dups": skipped_dup,
                       "erev_cells": n_er, "normalizations": NORMALIZATIONS,
                       "by_batch": {str(b): sum(1 for r in plan_new if batch_norm(r.get("批次")) == b) for b in (411, 412)},
                       "dryruns": [d["_path"] for d in dry]}, f, ensure_ascii=False, indent=2)
        return

    # ---- 备份 ----
    bak = None
    if args.no_backup:
        log("[TEST] 跳过备份")
    else:
        os.makedirs(ARCHIVE, exist_ok=True)
        bak = os.path.join(ARCHIVE, f"v4.9.32_pre-411412_backup_{datetime.now():%Y%m%d}.xlsx")
        n = 1
        while os.path.exists(bak):
            bak = os.path.join(ARCHIVE, f"v4.9.32_pre-411412_backup_{datetime.now():%Y%m%d}_{n}.xlsx")
            n += 1
        shutil.copy2(WB, bak)
        log(f"[+] 备份 → archive/{os.path.basename(bak)} ({os.path.getsize(bak)} B, sha={sha256(bak)[:12]})")

    # ---- 写新增行 ----
    r0 = ws.max_row
    seq = mx
    ledger_rows = []
    n_status = 0
    for i, row in enumerate(plan_new):
        r = r0 + 1 + i
        for cname, val in row.items():
            c = idx[cname]
            cell = ws.cell(row=r, column=c)
            cell.value = val
            cell.fill = copy.copy(FILL_ER if (cname == "是否增程(EREV)" and str(val) == "1") else FILL_NEW)
        # 并发销量管线尾插列：新增行未经其匹配，置"未匹配"占位（若该列存在）
        if args.sales_status and "销量覆盖状态" in idx:
            ws.cell(row=r, column=idx["销量覆盖状态"]).value = args.sales_status
            n_status += 1
        for c in range(1, ncol + 1):
            if ws.cell(row=r, column=c).fill.fill_type is None:
                ws.cell(row=r, column=c).fill = copy.copy(FILL_NEW)
        seq += 1
        b = batch_norm(row.get("批次"))
        ledger_rows.append([seq, b, row.get("产品型号"), row.get("产品名称"), "整行", "公告新增行",
                            "(新)", None,
                            f"{row.get('数据来源')}；企业={row.get('企业名称') or ''}；通用={row.get('通用名称') or ''}"
                            + (f"；销量三列与车系ID留空（未经销量管线匹配）、销量覆盖状态暂置「{args.sales_status}」待其复跑重判"
                               if args.sales_status and "销量覆盖状态" in idx else "")])

    # ---- 写富化格 ----
    for (r, cname, val, b, m, pt) in plan_enrich:
        cell = ws.cell(row=r, column=idx[cname])
        cell.value = val
        cell.fill = copy.copy(FILL_ER if (cname == "是否增程(EREV)" and str(val) == "1") else FILL_NEW)
        seq += 1
        ledger_rows.append([seq, b, m, pt, cname, "官方详情页补全", "(空)", val,
                            f"411批正式公告查询列详情页 {SRC_411}"])

    # ---- 写台账 ----
    rl = wsl.max_row
    for i, lr in enumerate(ledger_rows):
        for j, v in enumerate(lr):
            wsl.cell(row=rl + 1 + i, column=1 + j).value = v

    # ---- 批次时间表 ----
    wst = wb[SHEET_TIME]
    add_time = [
        [411, "2026-10", "2026-10-16", None, None, "官方精确日期(公告)",
         "工信部第411批公告 发布日期 2026-10-16（成文 2026-10-15）；来源＝公告批次扫描_411-412_20261017《证据与口径》；公示日期待补"],
        [412, None, None, "2026-10", "2026-10-17", "公示期官方(2026-10-17~10-23)·公告日期待定",
         "工信部第412批公示 2026-10-17~2026-10-23（公示中）；来源＝公告批次扫描_411-412_20261017《证据与口径》；正式公告后补公告日期"],
    ]
    for tr in add_time:
        base = wst.max_row + 1          # 必须先固定基准行（max_row 会随写入变化）
        for j, v in enumerate(tr):
            wst.cell(row=base, column=1 + j).value = v

    # ---- 字段映射说明 ----
    wsm = wb[SHEET_MAP]
    add_map = [
        [SRC_411, "批次/产品型号/产品商标/企业名称/通用名称/产品名称/产品类型/动力类型/整备/轴距/续航/容量/电池类型/电机/发动机/车长/电芯·电池包供应商",
         "公告参数查询系统详情页同名字段",
         "CR-17-011：411 批共 46 款＝存量 21 行逐格富化＋新增 25 行。产品名称＝公告描述串；产品类型＝产品名称；车型名称＝通用名称原样（先例：409 官方详情队列 100/111 一致）；是否减免购置税/细分市场/电池能量密度/百公里电耗 留空（对齐 410 正式队列先例 0/35）；只补空不覆盖"],
        [SRC_412, "同上＋纯电续航里程/电池容量/综合油耗/发动机排量（仅税目录拟定清单匹配到的车型）",
         "公示详情页同名字段；车船税目录第九十一批拟定清单复核",
         "CR-17-011：412 批 58 款按公示口径入表，公告后定稿；稀疏字段仅在税目录匹配命中时写入，未命中留空；数据来源列标公示口径"],
    ]
    for mr in add_map:
        base = wsm.max_row + 1          # 同上：基准行必须先固定
        for j, v in enumerate(mr):
            wsm.cell(row=base, column=1 + j).value = v

    wb.save(WB)
    new_sha = sha256(WB)
    log(f"[+] 已写入：新增 {len(plan_new)} 行、富化 {len(plan_enrich)} 格、台账 {len(ledger_rows)} 行、"
        f"时间表 +2、字段映射 +2" + (f"、销量覆盖状态占位 {n_status} 格（{args.sales_status}）" if n_status else ""))
    log(f"[+] 新 sha256 = {new_sha}；主表现 {ws.max_row - 1} 数据行 × {ncol} 列")

    rep = {"mode": "apply", "workbook": WB_NAME, "sha_before": wb_sha, "sha_after": new_sha,
           "backup": (os.path.relpath(bak, HERE) if bak else None), "new_rows": len(plan_new),
           "enrich_cells": len(plan_enrich), "ledger_rows": len(ledger_rows),
           "seq_from": mx + 1, "seq_to": seq, "skipped_dups": skipped_dup,
           "erev_cells": n_er,
           "by_batch": {str(b): sum(1 for r in plan_new if batch_norm(r.get("批次")) == b) for b in (411, 412)},
           "normalizations": NORMALIZATIONS,
           "dryruns": [d["_path"] for d in dry], "at": datetime.now().isoformat(timespec="seconds")}
    with open(os.path.join(OUT_DIR, f"cr17_011_apply_{ts}.json"), "w", encoding="utf-8") as f:
        json.dump(rep, f, ensure_ascii=False, indent=2)
    with open(os.path.join(OUT_DIR, f"cr17_011_apply_{ts}.log"), "w", encoding="utf-8") as f:
        f.write("\n".join(LOG))
    log(f"[i] 报告 → audit-output/cr17_011_apply_{ts}.json")


if __name__ == "__main__":
    main()
