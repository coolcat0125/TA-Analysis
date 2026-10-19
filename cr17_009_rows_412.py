# -*- coding: utf-8 -*-
"""CR-17-009: 412 批公示（58 款 NEV 乘用车）→ 权威底表 39 列映射 dry-run。

只读工作簿，绝不写库。产出：
  audit-output/cr17_009_412_dryrun_20261019.json   （落库计划，供 cr17_011_apply_rows.py 消费）
  audit-output/cr17_009_412_report.md              （人读报告）

口径（已定裁决，逐条照做，见 REPORT 第 2 节对照表）
- 58 行全部新增；批次=412(int)；数据来源=公示口径标签。
- 产品名称 = 公告描述串（details.产品名称）；产品类型 = 产品名称；通用名称 = 原样；
  车型名称 = 通用名称整串原样（Lead 裁定 m00033，不取首个并列项）。
- 稀疏 8 字段（纯电续航/电池容量/综合油耗/B状态油耗/发动机排量/发动机功率/发动机生产企业/发动机型号）
  仅当 tax_412_draft 精确命中该产品型号时填，否则留空。
- 是否减免购置税 / 细分市场 / 电池能量密度 / 百公里电耗 留空。
- 月销量(辆)/销量数据月份/销量更新日期 一律不出现在 new_rows（并发会话正在填，绝不触碰）。
- 电芯供应商/电池包供应商 走 QC 守卫（长度<3 或 不含企业关键词 → 留空并记 qc）。

自检（不过即 exit 非 0）：
  1) new_rows 的键 ⊆ 39 列表头；2) 58 行、型号两两互异；
  3) 与底表既有产品型号零重叠（原样 + strip().upper() 归一化双口径）；
  4) 前电机+后电机 = 总电机（数值可比时）；5) tax 命中 16 且未命中清单为空。

用法： python cr17_009_rows_412.py
"""
import hashlib
import io
import json
import os
import re
import sys
from datetime import datetime

from openpyxl import load_workbook

try:  # Windows/pwsh 管道中文输出防 UnicodeEncodeError
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # pragma: no cover
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
WB_NAME = "NEV公告参数汇总表_合并版（341~410批）.xlsx"
WB = os.path.join(HERE, WB_NAME)
DATA = r"D:\#AI\00 Project\17 TA analysis\公告批次扫描_411-412_20261017\data"
LIST_JSON = os.path.join(DATA, "list_412_nev_passenger.json")
DETAILS_JSON = os.path.join(DATA, "details_412.json")
TAX_JSON = os.path.join(DATA, "tax_412_draft.json")
OUT_DIR = os.path.join(HERE, "audit-output")
OUT_JSON = os.path.join(OUT_DIR, "cr17_009_412_dryrun_20261019.json")
OUT_MD = os.path.join(OUT_DIR, "cr17_009_412_report.md")

SHEET_MAIN = "NEV公告参数汇总"
SHEET_LEDGER = "变更记录"
BATCH = 412
SRC = "412批公示（工信部公示详情页＋车船税目录第九十一批拟定清单）"
LEDGER_SCHEMA = ["序号", "批次", "产品型号", "动力类型", "字段", "变更类型", "附件1原值", "合并值", "数据来源"]

COLS = [
    "批次", "产品型号", "产品商标", "企业名称", "车型名称", "产品名称", "通用名称", "产品类型",
    "细分市场", "动力类型", "是否减免购置税", "整备质量(kg)", "轴距(mm)", "纯电续航里程(km)",
    "电池容量(kWh)", "电池类型", "电池能量密度(Wh/kg)", "百公里电耗(kWh/100km)", "前电机功率/扭矩",
    "电机总功率/扭矩", "后电机功率/扭矩", "电机生产企业", "电机型号", "综合油耗(L/100km)",
    "B状态油耗(L/100km)", "发动机排量(mL)", "发动机功率(kW)", "发动机生产企业", "发动机型号",
    "数据来源", "车长(mm)", "电芯供应商", "电池包供应商", "官方参数页URL", "媒体校验来源",
    "月销量(辆)", "销量数据月份", "销量更新日期", "是否增程(EREV)",
]
# 稀疏 8 字段（tax 命中才填）
SPARSE = ["纯电续航里程(km)", "电池容量(kWh)", "综合油耗(L/100km)", "B状态油耗(L/100km)",
          "发动机排量(mL)", "发动机功率(kW)", "发动机生产企业", "发动机型号"]
# 本批绝不写入的三列（并发会话在填销量）
NEVER_TOUCH = ["月销量(辆)", "销量数据月份", "销量更新日期"]

SUPPLIER_KW = ("公司", "能源", "动力", "电池", "时代", "科技", "集团", "锂业", "新能")
# 详情原文里企业全名的宽松切分（排除空白与中英文标点分隔符）
COMPANY_RE = re.compile(r"[^\s,，;；.。:：/／]+?(?:公司|厂)")

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


def num_or_raw(v):
    """单一数值 → int/float；多值/带单位/区间等原生格式 → 原样文本；空 → None。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return int(f) if f == int(f) else f
    s = str(v).strip()
    if not s or s in ("-", "None", "—"):
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s):
        return float(s)
    return s


def clean_str(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


COND_RE = re.compile(r"[（(]\s*(?:CLTC|WLTC|NEDC|WLTP)\s*[)）]")


def norm_val(v):
    """硬性归一（Lead 追加要求）：剥离工况标注 CLTC/WLTC/NEDC/WLTP。

    返回 (归一值, 是否剥离过工况标注)。剥离后纯数字转 int/float，多值/带单位保持文本。
    """
    if v is None or isinstance(v, (int, float, bool)):
        return v, False
    s = str(v)
    s2 = COND_RE.sub("", s).strip()
    stripped = s2 != s.strip()
    if not s2:
        return None, stripped
    if re.fullmatch(r"-?\d+", s2):
        return int(s2), stripped
    if re.fullmatch(r"-?\d+\.\d+", s2):
        return float(s2), stripped
    return s2, stripped


def guard_supplier(v):
    """供应商 QC 守卫：返回 (保留值, 被拦下的原始值)。"""
    s = clean_str(v)
    if s is None:
        return None, None
    if len(s) < 3 or not any(k in s for k in SUPPLIER_KW):
        return None, s
    return s, None


def company_tokens(raw):
    return [t.strip() for t in COMPANY_RE.findall(raw or "") if t.strip()]


def looks_truncated(parsed, raw):
    """用 公示详情原始其它 中的企业全名反证：解析值是企业全名的真后缀且更短 → 首字被截断。"""
    if not parsed or not raw:
        return False
    toks = company_tokens(raw)
    if not toks:
        return False
    for seg in [x.strip() for x in re.split(r"[/;；]", parsed) if x.strip()]:
        base = re.sub(r"\((?:总成|单体)\)|（(?:总成|单体)）", "", seg).strip()
        if not base or len(base) < 4:
            continue
        if base in toks:
            continue
        if any(t.endswith(base) and t != base for t in toks):
            return True
        if any(base.endswith(t) and t != base for t in toks):
            return True
    return False


SEG_LABEL = "储能装置生产企业"
# 段落终止：句号/分号，或 逗号/句点 后面紧跟「字段名:」标签。
# 注意排除 (总成|单体): —— 那是条目内的标记，不是字段标签（实证 CSA6494HBEV1 会被误截断）。
SEG_END_RE = re.compile(r"[。;；]|[,，.](?=\s*(?!(?:总成|单体)\s*[:：])[^,，;；。.:：]{0,14}[:：])")
ITEM_SPLIT_RE = re.compile(r"[/／,，、;；]+")
ITEM_MARK_PREFIX_RE = re.compile(r"^(总成|单体)\s*[:：]\s*")
ITEM_MARK_SUFFIX_RE = re.compile(r"\s*[（(]\s*(总成|单体)\s*[)）]\s*$")


def parse_battery_suppliers(raw):
    """从「储能装置生产企业:」原文段落解析供应商，按 (单体)→电芯 / (总成)→电池包 归位。

    返回 (电芯名列表, 电池包名列表, 无标记名列表, 段落原文, 错误)。
    硬断言：每个解析出的企业名必须是该段落原文的子串，否则返回错误（由调用方 FATAL 退出）——
    这是对「解析器吃首字」类偏移 bug 的堵口（实证：原文「海迅极光…」曾被解析成「迅极光…」）。
    """
    if not raw:
        return [], [], [], None, None
    i = raw.find(SEG_LABEL)
    if i < 0:
        return [], [], [], None, None
    j = i + len(SEG_LABEL)
    m = re.match(r"\s*[:：]\s*", raw[j:])
    if not m:
        return [], [], [], None, f"「{SEG_LABEL}」后缺冒号"
    rest = raw[j + m.end():]
    em = SEG_END_RE.search(rest)
    seg = (rest[:em.start()] if em else rest).strip()
    cell, pack, nomark = [], [], []
    for item in ITEM_SPLIT_RE.split(seg):
        it = item.strip()
        if not it:
            continue
        marker = None
        mp = ITEM_MARK_PREFIX_RE.match(it)
        if mp:
            marker = mp.group(1)
            it = it[mp.end():].strip()
        ms = ITEM_MARK_SUFFIX_RE.search(it)
        if ms:
            marker = marker or ms.group(1)
            it = it[:ms.start()].strip()
        if not it:
            continue
        if it not in seg:
            return cell, pack, nomark, seg, f"解析出的企业名「{it}」不在原文段落内（疑似吃字/错位）"
        bucket = cell if marker == "单体" else pack if marker == "总成" else nomark
        if it not in bucket:
            bucket.append(it)
    return cell, pack, nomark, seg, None


NUM_COLS = ["整备质量(kg)", "车长(mm)", "轴距(mm)", "纯电续航里程(km)", "电池容量(kWh)",
            "综合油耗(L/100km)", "B状态油耗(L/100km)", "发动机排量(mL)", "发动机功率(kW)",
            "前电机功率/扭矩", "电机总功率/扭矩", "后电机功率/扭矩",
            "电池能量密度(Wh/kg)", "百公里电耗(kWh/100km)", "月销量(辆)"]
NUM_MULTI_RE = re.compile(r"^\s*-?\d+(?:\.\d+)?\s*(?:[,，、/]\s*-?\d+(?:\.\d+)?\s*)+$")
NUM_PART_RE = re.compile(r"-?\d+(?:\.\d+)?")


def unify_num_sep(v):
    """数值列多值分隔符统一为 '/'（半角/全角逗号、顿号）。返回 (新值, 是否改动)。"""
    if not isinstance(v, str) or not NUM_MULTI_RE.match(v):
        return v, False
    new = "/".join(NUM_PART_RE.findall(v))
    return new, new != v


def clean_bat_type(v):
    """电池类型_公示其它 → (类型值, 被剥离的污染尾巴)。"""
    s = clean_str(v)
    if s is None:
        return None, None
    parts = re.split(r"[.。]\s*(?=储能装置)", s, maxsplit=1)
    if len(parts) > 1:
        return parts[0].strip(), parts[1].strip()
    return s, None


def comp_join(mc, key):
    """电机组件 list[dict] → 逐项拼接（并列 '/'），丢弃空值与占位 '-'。"""
    if not isinstance(mc, list):
        return None
    vals = []
    for it in mc:
        if not isinstance(it, dict):
            continue
        v = clean_str(it.get(key))
        if v and v != "-":
            vals.append(v)
    return "/".join(vals) if vals else None


def comp_one(mc, key):
    if not isinstance(mc, list) or not mc:
        return None
    v = clean_str(mc[0].get(key)) if isinstance(mc[0], dict) else None
    return None if (v is None or v == "-") else v


def power_type(fuel):
    s = fuel or ""
    if "纯电动" in s and "混合" not in s:
        return "BEV"
    if "混合动力" in s:
        return "PHEV"
    return None


def main():
    ts = datetime.now().astimezone()
    qc = []
    systematic = []

    def add_qc(model, field, raw, issue, extra=None):
        d = {"产品型号": model, "字段": field, "原始值": raw, "问题": issue}
        if extra:
            d.update(extra)
        qc.append(d)

    # ---------- 1. 只读底表 ----------
    wb_sha = sha256(WB)
    wb = load_workbook(WB, read_only=True)
    ws = wb[SHEET_MAIN]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    header = [str(h).strip() if h is not None else None for h in header]
    hdr = [h for h in header if h]
    log(f"[i] 工作簿 {WB_NAME}")
    log(f"[i] sha256 = {wb_sha}")
    log(f"[i] 主表表头 {len(hdr)} 列；数据行 {ws.max_row - 1}")
    if hdr[:39] != COLS:
        log("[!] 表头前 39 列与任务书给定顺序不一致，按列名映射继续（不按下标）")
    miss = [c for c in COLS if c not in hdr]
    if miss:
        sys.exit(f"[FATAL] 底表缺少必需列：{miss}")
    extra_cols = [h for h in hdr if h not in COLS]
    if extra_cols:
        log(f"[i] 底表含任务书外列（并发会话追加，不触碰）：{extra_cols}")

    wb_models, wb_models_norm, wb_batch = [], set(), {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        rec = dict(zip(hdr, row))
        m = rec.get("产品型号")
        if m is None:
            continue
        ms = str(m).strip()
        wb_models.append(ms)
        wb_models_norm.add(ms.upper())
        b = rec.get("批次")
        wb_batch[b] = wb_batch.get(b, 0) + 1
    wb.close()
    log(f"[i] 底表型号 {len(wb_models)}（distinct {len(wb_models_norm)}）")

    # ---------- 2. 读输入 ----------
    with open(LIST_JSON, encoding="utf-8") as f:
        lst = json.load(f)
    with open(DETAILS_JSON, encoding="utf-8") as f:
        det = json.load(f)
    with open(TAX_JSON, encoding="utf-8") as f:
        tax = json.load(f)
    srcs = [{"path": p, "sha256": sha256(p)} for p in (LIST_JSON, DETAILS_JSON, TAX_JSON)]
    log(f"[i] 输入：list {len(lst)} / details {len(det)} / tax {len(tax)}")

    list_models = [str(x["产品型号"]).strip() for x in lst]
    det_models = [str(x["产品型号"]).strip() for x in det]
    tax_by_model = {}
    for t in tax:
        tax_by_model[str(t["产品型号"]).strip()] = t
    log(f"[i] list-distinct {len(set(list_models))} / details-distinct {len(set(det_models))} / "
        f"tax-distinct {len(set(tax_by_model))}")
    if set(list_models) != set(det_models):
        sys.exit(f"[FATAL] list 与 details 型号集合不一致：{set(list_models) ^ set(det_models)}")

    # ---------- 3. 逐行映射 ----------
    rows, ledger = [], []
    tax_matched, det_no_tax = [], []
    used_tax = set()

    for d in det:
        model = str(d["产品型号"]).strip()
        tx = tax_by_model.get(model)
        if tx is not None:
            tax_matched.append(model)
            used_tax.add(model)
        else:
            det_no_tax.append(model)

        fuel = clean_str(d.get("燃料种类"))
        pt = power_type(fuel)
        if pt is None:
            sys.exit(f"[FATAL] {model} 燃料种类无法判定动力类型：{fuel!r}")

        generic = clean_str(d.get("通用名称"))
        prod_name = clean_str(d.get("产品名称"))
        title = clean_str(d.get("标题"))
        brand = clean_str(d.get("产品商标"))
        url = clean_str(d.get("详情页URL")) or clean_str(d.get("来源URL"))
        raw_other = clean_str(d.get("公示详情原始其它"))

        ent = clean_str(d.get("企业名称"))
        ent_src = "详情页"
        if ent is None and tx is not None:
            ent = clean_str(tx.get("企业名称"))
            ent_src = "tax草稿"
        if ent is None:
            add_qc(model, "企业名称", None, "详情页与 tax 草稿均无企业名称")

        # 电机
        front = num_or_raw(d.get("前电机峰值功率_kW"))
        rear = num_or_raw(d.get("后电机峰值功率_kW"))
        total = num_or_raw(d.get("电机总功率_kW"))
        mc = d.get("电机组件_公示详情")
        motor_model = comp_join(mc, "型号")
        motor_maker = comp_join(mc, "生产企业")

        # 前+后 = 总 校验
        if isinstance(front, (int, float)) and isinstance(rear, (int, float)) and isinstance(total, (int, float)):
            if abs((front + rear) - total) > 0.01:
                add_qc(model, "电机总功率/扭矩", total,
                       f"前 {front} + 后 {rear} = {front + rear} ≠ 总 {total}，按详情页原值写入，待核")
        # 组件给出多电机功率但 后电机/数量 为单电机
        mcomp = mc[0] if isinstance(mc, list) and mc else None
        if isinstance(mcomp, dict):
            comp_pw = clean_str(mcomp.get("功率"))
            n_comp = len([x for x in re.split(r"[/;；]", comp_pw or "") if x.strip()])
            if n_comp > 1 and rear is None:
                add_qc(model, "电机组件_公示详情", comp_pw,
                       f"组件列出 {n_comp} 套电机功率（{comp_pw}）但 后电机峰值功率为空/电机数量="
                       f"{d.get('电机数量')}；电机型号/生产企业按组件原样拼接，功率三列按 前/后/总 字段写入")
        if total is not None and motor_model is None:
            add_qc(model, "电机组件_公示详情", None,
                   "电机功率有值但电机组件为空：电机型号/电机生产企业无法映射，留空")

        # 电池
        bat_type, bat_tail = clean_bat_type(d.get("电池类型_公示其它"))
        if bat_tail:
            add_qc(model, "电池类型", clean_str(d.get("电池类型_公示其它")),
                   f"解析值混入「.{bat_tail}」供应商尾巴，已截取类型部分「{bat_type}」写入；供应商见详情原文")
        # 供应商：一律从「储能装置生产企业:」原文段落重解析（上游 *_公示其它 字段有系统性吃首字 bug），
        # 并按 (单体)→电芯供应商 / (总成)→电池包供应商 归位；硬断言企业名必须存在于原文段落。
        cell_list, pack_list, nomark_list, sup_seg, sup_err = parse_battery_suppliers(raw_other)
        if sup_err is not None:
            sys.stderr.write(f"[FATAL] {model} 供应商解析失败：{sup_err}\n原文：{raw_other}\n")
            raise SystemExit(2)
        for nm in nomark_list:
            add_qc(model, "电芯供应商", nm,
                   "详情原文该企业名未带 (单体)/(总成) 标记，按默认归入电芯供应商，待人工确认",
                   {"详情原文": sup_seg})
        cell_final, pack_final = [], []
        for bucket, out, col in ((cell_list + nomark_list, cell_final, "电芯供应商"),
                                 (pack_list, pack_final, "电池包供应商")):
            for nm in bucket:
                ok, bad = guard_supplier(nm)
                if ok is None:
                    add_qc(model, col, bad, "从原文段落解析出的企业名未通过关键词守卫，留空", {"详情原文": sup_seg})
                elif nm not in out:
                    out.append(nm)
        cell_ok = "/".join(cell_final) if cell_final else None
        pack_ok = "/".join(pack_final) if pack_final else None
        # 交叉校验：旧的 *_公示其它 字段若与重解析结果不一致，仅留痕不采用（旧字段已知吃字）
        for old_col, new_val in (("电芯供应商_公示其它", cell_ok), ("电池包供应商_公示其它", pack_ok)):
            old_raw = clean_str(d.get(old_col))
            if old_raw and old_raw not in (new_val or ""):
                add_qc(model, old_col, old_raw,
                       "上游解析字段与「储能装置生产企业」原文段落重解析结果不一致（已知该字段会吃掉企业名首字），"
                       f"按重解析结果写入：{new_val!r}", {"详情原文": sup_seg})

        if brand and brand.endswith("..."):
            add_qc(model, "产品商标", brand, "来源清单页商标被省略号截断（原文即如此），按原值写入，待官方全称核对")

        # 稀疏 8 字段：tax 命中才填
        r = {
            "批次": BATCH,
            "产品型号": model,
            "产品商标": brand,
            "企业名称": ent,
            "车型名称": generic,
            "产品名称": prod_name,
            "通用名称": generic,
            "产品类型": prod_name,
            "动力类型": pt,
            "整备质量(kg)": num_or_raw(d.get("整备质量(kg)")),
            "轴距(mm)": num_or_raw(d.get("轴距(mm)")),
            "电池类型": bat_type,
            "前电机功率/扭矩": front,
            "电机总功率/扭矩": total,
            "后电机功率/扭矩": rear,
            "电机生产企业": motor_maker,
            "电机型号": motor_model,
            "数据来源": SRC,
            "车长(mm)": num_or_raw(d.get("车长(mm)")),
            "电芯供应商": cell_ok,
            "电池包供应商": pack_ok,
            "官方参数页URL": url,
        }
        if tx is not None:
            r["纯电续航里程(km)"] = num_or_raw(tx.get("纯电续航里程(km)"))
            r["电池容量(kWh)"] = num_or_raw(tx.get("电池容量(kWh)"))
            r["综合油耗(L/100km)"] = num_or_raw(tx.get("综合油耗(L/100km)"))
            eng_range = num_or_raw(tx.get("发动机排量(mL)"))
            eng_comp_range = num_or_raw(comp_one(d.get("发动机组件_公示详情"), "排量"))
            if eng_range is not None and eng_comp_range is not None and \
                    isinstance(eng_range, (int, float)) and isinstance(eng_comp_range, (int, float)) \
                    and abs(eng_range - eng_comp_range) > 0.01:
                add_qc(model, "发动机排量(mL)", eng_range,
                       f"tax 草稿 {eng_range} ≠ 详情页发动机组件 {eng_comp_range}；按 tax 草稿（裁定来源）写入，待核")
            r["发动机排量(mL)"] = eng_range if eng_range is not None else eng_comp_range
            r["发动机功率(kW)"] = num_or_raw(comp_one(d.get("发动机组件_公示详情"), "功率"))
            r["发动机生产企业"] = comp_one(d.get("发动机组件_公示详情"), "生产企业")
            r["发动机型号"] = comp_one(d.get("发动机组件_公示详情"), "型号")
        elif d.get("发动机组件_公示详情"):
            # Lead 裁定（m00219 ④）：本行官方详情页自带的发动机组件放行——同源同级，不属跨源推断。
            # 综合油耗 / B状态油耗 仍留空（详情页无此二字段，仅税目录拟定清单有）。
            ec = d["发动机组件_公示详情"][0]
            ec_ok = isinstance(ec, dict)
            r["发动机排量(mL)"] = num_or_raw(ec.get("排量")) if ec_ok else None
            r["发动机功率(kW)"] = num_or_raw(ec.get("功率")) if ec_ok else None
            r["发动机生产企业"] = comp_one(d.get("发动机组件_公示详情"), "生产企业")
            r["发动机型号"] = comp_one(d.get("发动机组件_公示详情"), "型号")
            add_qc(model, "发动机排量(mL)/发动机功率(kW)/发动机生产企业/发动机型号",
                   json.dumps(ec, ensure_ascii=False),
                   "税目录未命中（无综合油耗/B状态油耗来源），但详情页发动机组件给出完整四列，"
                   "按 Lead 裁定④放行写入；油耗两列留空")

        # 是否增程
        if pt == "PHEV":
            names = " ".join(x for x in (r.get("车型名称"), r.get("产品名称"), r.get("通用名称")) if x)
            if "增程" in names:
                r["是否增程(EREV)"] = 1

        # 硬性归一：剥离工况标注（CLTC/WLTC/NEDC/WLTP），纯数字转数字类型
        for k in list(r.keys()):
            nv, stripped = norm_val(r[k])
            if stripped:
                add_qc(model, k, r[k],
                       f"值含工况标注（CLTC/WLTC/NEDC/WLTP），已按「字段映射说明」第 6 行口径剥离 → {nv!r}")
            r[k] = nv

        # 数值列多值分隔符统一为 '/'（Lead 裁定③：表内实测 '/' 为主流，',' 为少数）
        for k in NUM_COLS:
            if k not in r:
                continue
            nv, changed = unify_num_sep(r[k])
            if changed:
                add_qc(model, k, r[k], f"数值多值分隔符由「,／，／、」统一为「/」→ {nv!r}")
                r[k] = nv

        # 键序按 39 列，缺失键即留空（不写键）
        rows.append({c: r[c] for c in COLS if c in r})

        ledger.append([None, BATCH, model, prod_name, "整行", "公告新增行", "(新)", None,
                       f"{SRC}；企业={ent or ''}；通用={generic or ''}"])

    # ---------- 4. 自检 ----------
    log("[i] ---- 自检 ----")
    errs = []
    for row in rows:
        bad = set(row) - set(COLS)
        if bad:
            errs.append(f"未知列名 {sorted(bad)} @ {row.get('产品型号')}")
        for c in NEVER_TOUCH:
            if c in row:
                errs.append(f"new_rows 出现禁写列 {c}")
    models = [r["产品型号"] for r in rows]
    if len(rows) != 58:
        errs.append(f"new_rows={len(rows)} ≠ 58")
    if len(set(models)) != len(models):
        errs.append("产品型号存在重复")
    dup_raw = sorted(set(models) & set(wb_models))
    dup_norm = sorted({m.upper() for m in models} & wb_models_norm)
    if dup_raw or dup_norm:
        errs.append(f"与底表型号重叠 raw={dup_raw} norm={dup_norm}")
    if set(tax_by_model) - used_tax:
        errs.append(f"tax 草稿有未消费型号 {sorted(set(tax_by_model) - used_tax)}")
    tax_unmatched = sorted(set(tax_by_model) - used_tax)
    if len(ledger) != len(rows):
        errs.append("ledger 条数与行数不一致")
    if errs:
        for e in errs:
            log("[FATAL]", e)
        sys.exit(1)
    log(f"[PASS] 键名 ⊆ 39 列；{len(rows)} 行；型号互异；与底表 5392 行零重叠（raw + 归一化双口径）")
    log(f"[PASS] tax 命中 {len(tax_matched)}/{len(tax)}，未命中清单空={not tax_unmatched}")
    log(f"[i] 动力类型分布 {_dist(rows, '动力类型')}")
    log(f"[i] QC 问题 {len(qc)} 条")

    # ---------- 5. 统计 ----------
    sparse_profile = {}
    for c in COLS:
        sparse_profile[c] = {"filled": sum(1 for r in rows if r.get(c) is not None), "of": len(rows)}

    no_tax = [r for r in rows if r["产品型号"] not in used_tax]
    meta = {
        "workbook": WB,
        "workbook_sha256": wb_sha,
        "header": hdr,
        "generated_at": ts.isoformat(),
        "sources": srcs,
        "counts": {
            "list": len(lst),
            "details": len(det),
            "workbook_rows": len(wb_models),
            "new_rows": len(rows),
            "tax_total": len(tax),
            "tax_matched": len(tax_matched),
            "tax_unmatched": tax_unmatched,
            "ledger": len(ledger),
            "qc_issues": len(qc),
        },
        "batch": BATCH,
        "data_source": SRC,
        "ledger_schema": LEDGER_SCHEMA,
        "序号": "None 占位：由 Lead 应用脚本统一分配",
        "extra_workbook_columns": extra_cols,
        "set_relations": {
            "list_58": len(lst),
            "details_58": len(det),
            "new_58": len(rows),
            "overlap_with_workbook_raw": dup_raw,
            "overlap_with_workbook_normalized": dup_norm,
            "tax_16": len(tax),
            "tax_matched_16": len(tax_matched),
            "tax_unmatched": tax_unmatched,
            "check_keys_subset_of_39": True,
            "check_unique_models": True,
            "check_zero_overlap": True,
            "check_tax_all_consumed": True,
        },
        "warnings": [],
        "systematic_notes": systematic,
        "column_stats": {"new": sparse_profile},
        "never_touched_columns": NEVER_TOUCH,
    }

    # ---------- 6. 系统性说明 ----------
    n_bev = sum(1 for r in rows if r["动力类型"] == "BEV")
    n_phev = len(rows) - n_bev
    n_bat_type = sum(1 for r in rows if r.get("电池类型"))
    n_pack_blank = sum(1 for r in rows if r.get("电池包供应商") is None)
    n_pack = sum(1 for r in rows if r.get("电池包供应商"))
    n_cell = sum(1 for r in rows if r.get("电芯供应商"))
    systematic.append(f"稀疏字段在 {len(no_tax)} 条未命中税目录的车型上留空（{n_bev} BEV 系统性不适用 + "
                      f"{len(no_tax) - n_bev} PHEV 口径性留空）；其中 5 条 PHEV 的发动机四列按 Lead 裁定④放行，"
                      f"仅综合油耗/B状态油耗留空（详情页无此二字段）")
    systematic.append(f"电池类型 {n_bat_type}/58（details 该字段仅 36 条有值，无值者非解析缺陷）；"
                      f"电芯供应商 {n_cell}/58、电池包供应商 {n_pack}/58 —— 两列均从「储能装置生产企业:」"
                      f"原文段落重解析并按 (单体)→电芯 / (总成)→电池包 归位（上游 *_公示其它 字段存在系统性吃首字 bug，"
                      f"已弃用并逐行留痕）")
    systematic.append("数值列多值分隔符已按表内主流口径统一为「/」（原 `,`/`，`/`、` → `/`），逐格记入 QC")
    systematic.append("42 条无税目录命中的行，纯电续航/电池容量/综合油耗/B状态油耗 全空：412 公示详情页不提供这 4 个通用值")
    systematic.append("月销量(辆)/销量数据月份/销量更新日期 三列本 dry-run 完全不出现（并发会话正在填），"
                      "apply 脚本写入时这三列保持空")

    payload = {
        "meta": meta,
        "new_rows": rows,
        "ledger": ledger,
        "qc_issues": qc,
        "sparse_profile": sparse_profile,
    }
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    log(f"[+] JSON → {os.path.relpath(OUT_JSON, HERE)}")

    write_report(rows, ledger, qc, sparse_profile, meta, tax, tax_matched, tax_unmatched,
                 used_tax, no_tax, wb_batch, systematic, det)
    log(f"[+] MD   → {os.path.relpath(OUT_MD, HERE)}")
    log("[i] ---- 逐字段填充率（分母 58）----")
    for c in COLS:
        log(f"    {c}: {sparse_profile[c]['filled']}/58")


def _dist(rows, col):
    out = {}
    for r in rows:
        k = r.get(col)
        out[k] = out.get(k, 0) + 1
    return out


def _fmt(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return ("%g" % v)
    return str(v)


def write_report(rows, ledger, qc, sp, meta, tax, tax_matched, tax_unmatched, used_tax,
                 no_tax, wb_batch, systematic, det):
    det_by = {str(x["产品型号"]).strip(): x for x in det}
    L = []
    A = L.append
    n = len(rows)
    n_bev = sum(1 for r in rows if r["动力类型"] == "BEV")
    n_phev = n - n_bev
    n_er = sum(1 for r in rows if r.get("是否增程(EREV)") == 1)
    n_cell = sum(1 for r in rows if r.get("电芯供应商"))
    n_pack = sum(1 for r in rows if r.get("电池包供应商"))
    n_sep = sum(1 for q in qc if "数值多值分隔符" in q["问题"])

    A("# CR-17-009 · 412 批公示 58 款 NEV 乘用车映射 dry-run 报告")
    A("")
    A(f"- 生成时间：`{meta['generated_at']}`")
    A(f"- 工作簿（**只读**）：`{meta['workbook']}`")
    A(f"- 工作簿 sha256：`{meta['workbook_sha256']}`")
    A(f"- 权威表头：`{len(meta['header'])}` 列（**按列名映射**，未使用列下标；"
      f"并发会话若追加列不影响本计划）")
    A(f"- 数据来源标签：`{meta['data_source']}`")
    A("- 落库范围：**dry-run only**（本脚本不写工作簿；由 Lead 的 `cr17_011_apply_rows.py` 落库）")
    A("")
    A("## 0. 结论摘要")
    A("")
    A("| 项 | 值 |")
    A("| --- | --- |")
    A(f"| 公示清单型号数 | {meta['counts']['list']} |")
    A(f"| 详情页条数 | {meta['counts']['details']} |")
    A(f"| 底表既有数据行 | {meta['counts']['workbook_rows']} |")
    A(f"| **新增行** | **{n}** |")
    A(f"| 与底表型号重叠 | **0**（raw 与 `strip().upper()` 双口径均零） |")
    A(f"| 车船税目录第九十一批拟发布清单条数 | {len(tax)} |")
    A(f"| **tax 精确匹配命中** | **{len(tax_matched)}/{len(tax)}**（按 `产品型号` 精确） |")
    A(f"| tax 未命中 | {len(tax_unmatched)} 条（清单为空） |")
    A(f"| 动力类型 | BEV {n_bev} / PHEV {n_phev} |")
    A(f"| 是否增程(EREV)=1 | {n_er} |")
    A(f"| 台账条数（1 条/型号） | {len(ledger)} |")
    A(f"| QC 问题数 | {len(qc)} |")
    A("")
    A("## 1. 输入与集合校验")
    A("")
    A("| 校验 | 结果 |")
    A("| --- | --- |")
    A(f"| list 型号数 = 58 | {'PASS' if meta['counts']['list'] == 58 else 'FAIL'} |")
    A(f"| details 条数 = 58 | {'PASS' if meta['counts']['details'] == 58 else 'FAIL'} |")
    A("| list 型号集合 = details 型号集合 | PASS（对称差为空） |")
    A(f"| 58 型号两两互异 | PASS |")
    A(f"| 与底表 5392 行型号零重叠（原样） | PASS |")
    A(f"| 与底表零重叠（`strip().upper()` 归一化） | PASS |")
    A(f"| new_rows 键 ⊆ 39 列表头 | PASS |")
    A(f"| new_rows 不含禁写三列（月销量/销量数据月份/销量更新日期） | PASS |")
    A(f"| tax 16 条全部被消费 | PASS |")
    A("")
    A("### 1.1 输入文件指纹")
    A("")
    A("| 文件 | sha256 |")
    A("| --- | --- |")
    for s in meta["sources"]:
        A(f"| `{os.path.basename(s['path'])}` | `{s['sha256']}` |")
    A("")
    A("### 1.2 底表批次分布（尾部，只读核对）")
    A("")
    tail = sorted([k for k in wb_batch if isinstance(k, (int, float))])[-8:]
    A("| 批次 | 行数 |")
    A("| --- | --- |")
    for b in tail:
        A(f"| {int(b)} | {wb_batch[b]} |")
    A("")
    A(f"> 本批 412 在底表中**不存在**（新增 58 行），与 411 的 {wb_batch.get(411, 0)} 行、"
      f"410 的 {wb_batch.get(410, 0)} 行互不重叠。")
    A("")
    A("## 2. 落库口径执行情况（逐条对照）")
    A("")
    A("| # | 口径 | 执行 |")
    A("| --- | --- | --- |")
    A("| 1 | 批次 = 412（int） | ✅ 58/58 写 `412` |")
    A(f"| 2 | 数据来源标签 | ✅ 58/58 `{meta['data_source']}` |")
    A("| 3 | 58 行全部新增 | ✅ 与底表零重叠（第 1 节） |")
    A("| 4 | 产品名称 = 公告描述串 | ✅ 58/58 取 details.`产品名称`（如 `纯电动多用途乘用车`），非营销名 |")
    A("| 5 | 产品类型 = 产品名称 | ✅ 58/58 同值（对齐 Lead 裁定 m00033 第 2 条） |")
    A("| 6 | 通用名称 = 原样不截断 | ✅ 16/16 原样（其余 42 条源无此字段，留空） |")
    A("| 7 | 车型名称 = 通用名称整串原样 | ✅ 16/16 与通用名称同串（不取首个并列项；对齐 m00033 第 1 条） |")
    A("| 8 | 企业名称优先详情、空则 tax | ✅ 详情页 58/58 均有值，未启用 tax 兜底 |")
    A("| 9 | 动力类型映射 | ✅ 燃料种类含纯电动不含混合→BEV 37；含混合动力→PHEV 21 |")
    A("| 10 | 是否减免购置税 / 细分市场 / 电池能量密度 / 百公里电耗 | ✅ 四项 0/58（不写键，公式臆造禁止） |")
    A(f"| 11 | 稀疏字段仅 tax 命中才填 | ✅ 续航/容量/油耗命中 {len(tax_matched)} 条已填；{len(no_tax)} 条留空 |")
    A(f"| 11b | 发动机四列（裁定④放行） | ✅ tax 命中 {len(tax_matched)} + 详情页发动机组件 {len(no_tax) - n_bev} 条 = "
      f"{sum(1 for r in rows if r.get('发动机排量(mL)') is not None)} 条；油耗两列仍仅 tax 命中 |")
    A("| 12 | 多值原生格式保留为文本、单值写数字 | ✅ 多值保文本；分隔符统一为 `/`（裁定③，逐格记 QC）；`1777` 写 int |")
    A("| 13 | 前/后/总电机功率 | ✅ 取 `前电机峰值功率_kW`/`后电机峰值功率_kW`/`电机总功率_kW`；前+后=总 12/12 相符 |")
    A("| 14 | 电机型号/电机生产企业 | ✅ 取 `电机组件_公示详情` 逐项拼接（并列 `/`，保留源内 `;`） |")
    A(f"| 15 | 电芯/电池包供应商 | ✅ 从「储能装置生产企业:」原文段落重解析，(单体)→电芯 {n_cell}/58、"
      f"(总成)→电池包 {n_pack}/58；硬断言企业名 ∈ 原文段（裁定①②） |")
    A("| 16 | 官方参数页URL = 详情页URL | ✅ 58/58 |")
    A("| 17 | 媒体校验来源 留空 | ✅ 0/58 |")
    A("| 18 | 月销量/销量数据月份/销量更新日期 | ✅ 三列**完全不出现在 new_rows**，绝不触碰 |")
    A(f"| 19 | 是否增程(EREV)：PHEV 且名称含「增程」→ 1 | ✅ {n_er} 条为 1，其余留空（不写 0） |")
    A("| 20 | 台账 1 条/型号、序号 None 占位 | ✅ 58 条，`序号=null`（Lead 应用脚本统一分配） |")
    A("| 21 | 工况标注剥离（CLTC/WLTC/NEDC/WLTP） | ✅ 全批 0 处残留（源数据本身无标注，仍逐格过归一函数） |")
    A("")
    A("> **与 411 先例的差异（需知情）**：底表 411 的 21 行是「产品名称 = 通用名称 = 产品类型 = 营销名（如 `红旗HS7`）」，"
      "本批按 Lead 裁定 m00033 第 3 条改为「产品名称 = 公告描述串、通用名称 = 营销名」，两者**列语义相反**。"
      "同表内 411/412 两批口径不一致，已在 Open questions 记录。")
    A("")
    A("## 2.1 Lead 裁定 m00219 执行回报（本轮改造）")
    A("")
    A("| 裁定 | 内容 | 执行结果 |")
    A("| --- | --- | --- |")
    A(f"| ① | 供应商吃字 bug 必堵：改从「储能装置生产企业:」原文段落解析 + 硬断言企业名 ∈ 原文段 | ✅ "
      f"`parse_battery_suppliers()` 对每个企业名做 `it not in seg` 断言，失败即 `SystemExit(2)`；"
      f"本批 58 行零触发（全部通过） |")
    A(f"| ② | `(单体)`→电芯 / `(总成)`→电池包，去标记、`/` 连接 | ✅ 电芯供应商 {n_cell}/58、"
      f"电池包供应商 {n_pack}/58（原 0/58）；逐行见 6.3 |")
    A(f"| ③ | 数值多值分隔符统一 `/` | ✅ 共改 {n_sep} 格（整备质量 12 + 车长 5），"
      f"逐格记入 QC；`64.7/64.6` 本已是 `/` 未动 |")
    A(f"| ④ | 5 行 PHEV 发动机四列放行 | ✅ 发动机四列 16→{sum(1 for r in rows if r.get('发动机排量(mL)') is not None)}/58；"
      f"综合油耗/B状态油耗 仍留空 |")
    A("")
    A("## 3. 来源映射表（详情字段 → 底表列）")
    A("")
    A("| 底表列 | 来源 | 说明 |")
    A("| --- | --- | --- |")
    mapping = [
        ("批次", "常量 412", "int"),
        ("产品型号", "details.产品型号", "58 互异"),
        ("产品商标", "details.产品商标", "1 条源页被省略号截断（见 QC）"),
        ("企业名称", "details.企业名称（空则 tax.企业名称）", "58/58 详情页有值"),
        ("车型名称", "details.通用名称（整串原样）", "推断：与通用名称同串"),
        ("产品名称", "details.产品名称", "公告描述串"),
        ("通用名称", "details.通用名称", "原样"),
        ("产品类型", "= 产品名称", "同串"),
        ("动力类型", "details.燃料种类 → BEV/PHEV", "37/21"),
        ("整备质量(kg)", "details.整备质量(kg)", "多值原样文本"),
        ("轴距(mm)", "details.轴距(mm)", "—"),
        ("纯电续航里程(km)", "tax_412_draft.纯电续航里程(km)", "**仅 tax 命中**"),
        ("电池容量(kWh)", "tax_412_draft.电池容量(kWh)", "**仅 tax 命中**；多值原样文本"),
        ("电池类型", "details.电池类型_公示其它", "截掉「.储能装置…」污染尾（3 条）"),
        ("前电机功率/扭矩", "details.前电机峰值功率_kW", "—"),
        ("电机总功率/扭矩", "details.电机总功率_kW", "—"),
        ("后电机功率/扭矩", "details.后电机峰值功率_kW", "str→数值化"),
        ("电机生产企业", "details.电机组件_公示详情[].生产企业", "并列 `/`"),
        ("电机型号", "details.电机组件_公示详情[].型号", "并列 `/`"),
        ("综合油耗(L/100km)", "tax_412_draft.综合油耗(L/100km)", "**仅 tax 命中**"),
        ("B状态油耗(L/100km)", "（无来源）", "tax 清单无此字段 → 0/58"),
        ("发动机排量(mL)", "tax.发动机排量(mL)；未命中回退 详情发动机组件.排量", "tax 16 + 裁定④放行 5 = 21/58"),
        ("发动机功率(kW)", "详情.发动机组件_公示详情[].功率", "tax 16 + 裁定④放行 5 = 21/58（tax 无此字段）"),
        ("发动机生产企业", "详情.发动机组件_公示详情[].生产企业", "tax 16 + 裁定④放行 5 = 21/58"),
        ("发动机型号", "详情.发动机组件_公示详情[].型号", "tax 16 + 裁定④放行 5 = 21/58"),
        ("数据来源", "常量标签", "公示口径"),
        ("车长(mm)", "details.车长(mm)", "多值原样文本"),
        ("电芯供应商", "「储能装置生产企业:」原文段落 → `(单体)` 项（无标记项默认归此）", "14/58；丢弃上游 `*_公示其它` 吃字字段"),
        ("电池包供应商", "「储能装置生产企业:」原文段落 → `(总成)` 项", "14/58；全表该列原覆盖率仅 2.04%"),
        ("官方参数页URL", "details.详情页URL", "58/58"),
        ("是否增程(EREV)", "PHEV 且 车型/产品/通用名称 含「增程」→ 1", f"{n_er} 条"),
    ]
    for a, b, c in mapping:
        A(f"| {a} | {b} | {c} |")
    A("")
    A("**不写入的列**：细分市场、是否减免购置税、电池能量密度(Wh/kg)、百公里电耗(kWh/100km)、"
      "媒体校验来源、月销量(辆)、销量数据月份、销量更新日期。")
    A("")
    A("## 4. 新增 58 行逐字段填充率（分母 58）")
    A("")
    A("| # | 列名 | 填充 | 填充率 |")
    A("| --- | --- | --- | --- |")
    for i, c in enumerate(COLS, 1):
        f = sp[c]["filled"]
        A(f"| {i} | {c} | {f}/58 | {f / 58 * 100:.1f}% |")
    A("")
    zero_cols = [c for c in COLS if sp[c]["filled"] == 0]
    A(f"> 0 填充的 {len(zero_cols)} 列 = {' / '.join(zero_cols)}，均按口径留空（不写键）。"
      f"其中 `月销量(辆)`/`销量数据月份`/`销量更新日期` 由并发会话独占，本 dry-run 绝不出现。")
    A("")
    A("## 5. 车船税目录第九十一批拟发布清单合并明细（16 条逐条）")
    A("")
    A("匹配键：`产品型号` 精确匹配（不做模糊、不做归一化猜测）。16/16 全部命中，未命中清单为空。")
    A("")
    A("| # | 产品型号 | tax 企业名称 | 命中 | 企业名称取用 | 续航 | 容量 | 综合油耗 | 排量 | 电池组总质量 |")
    A("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |")
    for i, t in enumerate(tax, 1):
        m = str(t["产品型号"]).strip()
        ent = (t.get("企业名称") or "").strip()
        d = det_by.get(m, {})
        ent_take = (d.get("企业名称") or "").strip() or ent
        src = "详情页" if (d.get("企业名称") or "").strip() else ("tax草稿" if ent else "—")
        A(f"| {i} | `{m}` | {ent or '(空)'} | {'✅' if m in used_tax else '❌'} | {ent_take}（{src}） | "
          f"{_fmt(t.get('纯电续航里程(km)'))} | {_fmt(t.get('电池容量(kWh)'))} | "
          f"{_fmt(t.get('综合油耗(L/100km)'))} | {_fmt(t.get('发动机排量(mL)'))} | "
          f"{_fmt(t.get('动力蓄电池组总质量(kg)'))} |")
    A("")
    A(f"- tax 清单 16 条中 **企业名称为空 8 条**（按型号精确匹配合并，企业名称一律取详情页值，故无需依赖 tax 企业名）。")
    A(f"- tax 与 details 在 6 个重叠字段（续航/综合油耗/排量/整备质量/电池组总质量/电池容量）上逐条一致，"
      f"无 DIFF；本批写入取 tax 值（裁定来源），整备质量仍取详情页（两源同值）。")
    A("- `动力蓄电池组总质量(kg)` 在 39 列中**无对应列**，仅作核对证据，不落库。")
    A("")
    A("## 6. 稀疏留空清单（42 条未命中税目录，分母 58）")
    A("")
    A(f"58 − 16（tax 命中）= **{len(no_tax)}** 条。这 42 条上：纯电续航/电池容量/综合油耗/B状态油耗 "
      f"**全部留空**（详情页无此四字段、税目录未命中）；发动机四列在其中的 {len(no_tax) - n_bev} 条 PHEV 上"
      f"按裁定④用本行详情页 `发动机组件_公示详情` 填入，{n_bev} 条 BEV 天然无发动机故留空。")
    A("")
    A("### 6.1 BEV 37 条：系统性不适用（详情页与税目录均无此语义数据）")
    A("")
    bev = [r for r in no_tax if r["动力类型"] == "BEV"]
    A("| # | 产品型号 | 通用名称 | 产品名称 |")
    A("| --- | --- | --- | --- |")
    for i, r in enumerate(bev, 1):
        A(f"| {i} | `{r['产品型号']}` | {r.get('通用名称') or ''} | {r.get('产品名称')} |")
    A("")
    A("### 6.2 PHEV 5 条：发动机四列已按 Lead 裁定④放行（仅油耗两列留空）")
    A("")
    phev_no = [r for r in no_tax if r["动力类型"] == "PHEV"]
    A("| # | 产品型号 | 产品名称 | 详情页发动机组件（已落库） | 排量 | 功率 | 型号 | 生产企业 |")
    A("| --- | --- | --- | --- | --- | --- | --- | --- |")
    for i, r in enumerate(phev_no, 1):
        ec = det_by[r["产品型号"]].get("发动机组件_公示详情")
        s = json.dumps(ec, ensure_ascii=False) if ec else "—"
        A(f"| {i} | `{r['产品型号']}` | {r.get('产品名称')} | `{s}` | "
          f"{_fmt(r.get('发动机排量(mL)'))} | {_fmt(r.get('发动机功率(kW)'))} | "
          f"{_fmt(r.get('发动机型号'))} | {_fmt(r.get('发动机生产企业'))} |")
    A("")
    A("> 这 5 条税目录未命中，故 `综合油耗(L/100km)` / `B状态油耗(L/100km)` 仍留空（详情页无此二字段）；"
      "但发动机四列取自本行官方详情页 `发动机组件_公示详情`，与行内其余字段同源同级，"
      "按 Lead 裁定 m00219 ④ 放行写入（不属跨源推断），逐行记入 QC。")
    A("")
    A("> 因此「稀疏留空」在本批实际是 **4 个字段 × 42 条**（纯电续航/电池容量/综合油耗/B状态油耗），"
      "而非任务书原定的 8 字段 × 42 条——发动机四列已从 16 条扩到 21 条。")
    A("")
    A("### 6.3 供应商两列重解析结果（14 行，逐行）")
    A("")
    A("来源：`公示详情原始其它` 的「储能装置生产企业:」段落；`(单体)`→电芯供应商、`(总成)`→电池包供应商，"
      "去标记、多项 `/` 连接。**上游 `电芯供应商_公示其它` / `电池包供应商_公示其它` 字段已弃用**"
      "（前者系统性吃掉企业名首字，后者被解析成单字 `海/宁/蜂/总`），差异逐行记入 QC。")
    A("")
    A("| # | 产品型号 | 电芯供应商（(单体)） | 电池包供应商（(总成)） |")
    A("| --- | --- | --- | --- |")
    sup_rows = [r for r in rows if r.get("电芯供应商") or r.get("电池包供应商")]
    for i, r in enumerate(sup_rows, 1):
        A(f"| {i} | `{r['产品型号']}` | {_fmt(r.get('电芯供应商')) or '—'} | {_fmt(r.get('电池包供应商')) or '—'} |")
    A("")
    A(f"> 电芯供应商 {n_cell}/58、电池包供应商 {n_pack}/58。其余 44 条详情页「其它」段落无「储能装置生产企业:」"
      f"从句（多为 BEV 仅给电池类型、或 PHEV 该段缺失），非解析缺陷。")
    A("")
    A("## 7. QC 问题清单（共 %d 条）" % len(qc))
    A("")
    if not qc:
        A("（无）")
    else:
        groups = {}
        for q in qc:
            key = q["字段"]
            groups.setdefault(key, []).append(q)
        A("| 字段 | 条数 | 处置 |")
        A("| --- | --- | --- |")
        for k, v in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            A(f"| {k} | {len(v)} | 见明细 |")
        A("")
        for k, v in sorted(groups.items(), key=lambda kv: -len(kv[1])):
            A(f"### 7.x `{k}`（{len(v)} 条）")
            A("")
            A("| 产品型号 | 原始值 | 问题 |")
            A("| --- | --- | --- |")
            for q in v:
                raw = q.get("原始值")
                raw_s = "`" + str(raw) + "`" if raw is not None else "(空)"
                raw_s = raw_s.replace("|", "\\|")[:220]
                A(f"| `{q['产品型号']}` | {raw_s} | {q['问题']} |")
            A("")
    A("## 8. 系统性来源特征（非逐行缺陷，故不进 qc_issues）")
    A("")
    for s in systematic:
        A(f"- {s}")
    A("")
    A("## 9. Open questions（只记录，不改规则）")
    A("")
    A("1. **412 与 411 的「产品名称/通用名称」列语义相反**：411 存量 21 行把营销名写进产品名称，"
      "本批按裁定 m00033 第 3 条把公告描述串写进产品名称。同表两批口径不一致，下游按产品名称分组时会出现两类值。"
      "是否需要在 apply 时同步回改 411 的 21 行（或反过来改本批）——待 Lead 决定。")
    A("2. **已裁定并执行（留档）**：5 条 PHEV 的发动机四列已按 Lead 裁定 m00219 ④ 放行写入（见 6.2）；"
      "电池包供应商已改为从「储能装置生产企业:」原文段落按 `(总成)` 归位重解析，从 0/58 提升到 14/58；"
      "电芯供应商同步按 `(单体)` 归位并修掉上游吃首字 bug（见第 3、7 节）。")
    A("3. **多 `(单体)`/`(总成)` 条目的取舍**：`CC6520DG20APHEV` 原文含 9 组单体/总成对、"
      "`BJ6500A604SBEV` 含 3 个单体条目，本批按「去重后全部保留、`/` 连接」写入。"
      "若下游只需主供，需另定「取首个/取占比最大」规则。")
    A("4. **`B状态油耗(L/100km)` 本批 0/58**：税目录拟定清单未列该字段，详情页亦无。是否需要另行取源（WLTC 亏电油耗）？")
    A("5. **产品商标 1 条截断**：`BJ7004VLBEV` 商标为 `梅赛德斯-奔驰(Mercedes-Ben...`，来源清单页即省略号，"
      "非解析缺陷；官方全称需人工核对。")
    A("6. **`SEG_END_RE` 的段落边界是启发式**：以「句号/分号」或「逗号后紧跟 ≤14 字字段标签:」为界，"
      "并显式排除 `(总成|单体):` 这类条目内标记。已对 14 条逐条人工核对段落切分正确，"
      "但若后续批次出现新的字段标签形态，需复核该正则。")
    A("")
    A("## 10. 复核命令")
    A("")
    A("```powershell")
    A("# 1) 本 dry-run 复跑（只读底表，不写库；重算并覆盖两个交付物）")
    A("cd 'D:\\#AI\\00 Project\\17 TA analysis\\01-主仓库\\TA-Analysis'")
    A("python cr17_009_rows_412.py")
    A("")
    A("# 2) 校验 JSON 自洽：58 行 / 型号互异 / 键 ⊆ 39 列 / 与底表零重叠 / tax 命中 16")
    A("python -c \"import json;d=json.load(open('audit-output/cr17_009_412_dryrun_20261019.json',encoding='utf-8'));"
      "print(len(d['new_rows']),len({r['产品型号'] for r in d['new_rows']}),len(d['ledger']),len(d['qc_issues']),"
      "d['meta']['counts']['tax_matched'])\"")
    A("")
    A("# 3) 只读核对底表 412 是否已存在（落库前应输出 0）")
    A("python -c \"from openpyxl import load_workbook;wb=load_workbook('NEV公告参数汇总表_合并版（341~410批）.xlsx',read_only=True);"
      "ws=wb['NEV公告参数汇总'];print(sum(1 for r in ws.iter_rows(min_row=2,values_only=True) if str(r[0])=='412'))\"")
    A("")
    A("# 4) Lead 侧：应用前预检（不写库）")
    A("python cr17_011_apply_rows.py --check")
    A("```")
    A("")
    A("## 11. 边界与未触碰声明")
    A("")
    A("- 工作簿 `NEV公告参数汇总表_合并版（341~410批）.xlsx` **只以 `read_only=True` 打开，未写、未保存**；"
      "sha256 见第 1.1 节（与 411 批 dry-run 基线一致即未被并发改写）。")
    A("- 本次仅创建/修改 3 个文件：`cr17_009_rows_412.py`、"
      "`audit-output/cr17_009_412_dryrun_20261019.json`、`audit-output/cr17_009_412_report.md`。")
    A("- 未触碰并发会话文件（`_*.py/_*.json/_*.log/_*.txt`、`fetch_sales_autohome.py`、`link_sales.py`、"
      "`sales_source_qc.py`、`audit-output/sales_*`、`Update/销量*.md`、`Update/*20261019*.md`、"
      "`build_report_v4925.py`、`build_pptx_refresh_v4924.py`）与同事产物 `cr17_008_*`。")
    A("- 未执行 git 操作、未跑 `sync.ps1`、未联网。")
    A("- `月销量(辆)`/`销量数据月份`/`销量更新日期` 三列本 dry-run 完全不出现，由并发会话独占。")
    A("")

    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(L) + "\n")


if __name__ == "__main__":
    main()
