#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公开垂直媒体（汽车之家）参数补空 v1
==================================
对底表做公共媒体横向查缺补漏：
- 数据源：汽车之家公开接口（sug/_suggest -> car/spec/listSpec -> car/param/getParamConf）
- 规则（用户确认）：只补空、不覆盖官方；同车系多配置仅写"唯一共识"，多值不写；
  来源写入【媒体校验来源】列并逐条记录到【参数对齐记录】。
- 懂车帝 PC 站被滑块验证码拦截（守则不绕过）；本次主体为汽车之家。

用法：
  python media_fill.py --scope 409-410 [--dry-run] [--delay 0.25]
Copyright © 2026 David YE
"""
import sys
import os
import re
import json
import time
import ssl
import glob
import argparse
import urllib.parse
import urllib.request
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
CTX = ssl.create_default_context()
HERE = os.path.dirname(os.path.abspath(__file__))

FILL_MEDIA = PatternFill("solid", fgColor="F8CBAD")   # 媒体补空（浅橙）
RECORD_SHEET = "参数对齐记录"
MISSING = {"", "/", "-", "--", "—", "N/A", "NA", "0", "未填报", "无", "无信息",
           "待查", "待核实", "待确认", "未知", "？", "?", "None", "未提供", "不适用", "-"}
GENERIC_NAMES = {"纯电动多用途乘用车", "纯电动轿车", "纯电动运动型乘用车", "插电式混合动力多用途乘用车",
                 "插电式增程混合动力多用途乘用车", "插电式混合动力轿车", "插电式增程混合动力轿车",
                 "纯电动客车", "纯电动轻型客车", "纯电动厢式运输车"}
POWER_WORDS = ("纯电", "增程", "DM-i", "DMi", "PHEV", "插混", "混动", "EV", "EVX")

# 汽车之家参数名 -> 底表字段（CLTC 优先，WLTC 兜底；后者以字段已存在为准跳过）
FIELD_MAP = {
    "长度(mm)": "车长(mm)",
    "轴距(mm)": "轴距(mm)",
    "整备质量(kg)": "整备质量(kg)",
    "CLTC纯电续航里程(km)": "纯电续航里程(km)",
    "WLTC纯电续航里程(km)": "纯电续航里程(km)",
    "电池能量(kWh)": "电池容量(kWh)",
    "电动机总功率(kW)": "电机总功率(kW)",
    "电芯品牌": "电芯供应商",
    "电池类型": "电池类型",
    "发动机型号": "发动机型号",
    "WLTC综合油耗(L/100km)": "综合油耗(L/100km)",
    "最低荷电状态油耗(L/100km)WLTC": "B状态油耗(L/100km)",
}


def is_missing(v):
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s in MISSING


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=25, context=CTX) as r:
        return json.loads(r.read().decode("utf-8", "ignore"))


def suggest(q):
    u = ("https://sou.api.autohome.com.cn/sug/_suggest?plat=pc&uid=0&q="
         + urllib.parse.quote(q))
    d = get(u)
    out = []
    for x in (d.get("result") or {}).get("data", []) or []:
        out.append({"key": str(x.get("key") or ""), "wordid": x.get("wordid"),
                    "wordtype": x.get("wordtype"), "score": x.get("score") or 0})
    return out


def list_spec(seriesid):
    u = (f"https://car-web-api.autohome.com.cn/car/spec/listSpec"
         f"?type=0x001f&from=1&pm=1&pluginversion=11.64.8&seriesid={seriesid}")
    r = (get(u).get("result") or {})
    specs = []
    for grp in (r.get("list") or []) + (r.get("otherlist") or []):
        for sp in (grp.get("speclist") or []):
            specs.append({"id": sp.get("id"), "name": sp.get("name"),
                          "type": grp.get("type"), "iselectric": sp.get("iselectric")})
    return {"seriesname": r.get("seriesname"), "seriesid": r.get("seriesid"), "specs": specs}


def param_conf(seriesid, specid):
    u = (f"https://car-web-api.autohome.com.cn/car/param/getParamConf"
         f"?mode=1&site=2&seriesid={seriesid}&specid={specid}")
    d = get(u)
    r = d.get("result") or {}
    tm = {}
    for grp in r.get("titlelist") or []:
        for it in grp.get("items") or []:
            tm[it.get("titleid")] = it.get("itemname")
    vals = {}
    for pc in (r.get("datalist") or [{}])[0].get("paramconflist") or []:
        tid = pc.get("titleid")
        name = tm.get(tid)
        if name is None:
            continue
        if pc.get("sublist"):
            sub = pc["sublist"][0].get("name") if pc["sublist"][0].get("name") else None
            val = sub
        else:
            val = pc.get("itemname")
        vals[name] = val
    return vals


def find_series(sug, row_pt, name="", brand=""):
    """从多个 suggest 结果里挑最优系列；无可靠重叠则返回 None。"""
    cands = [x for x in sug if x.get("wordtype") == 3 and x.get("wordid")]
    if not cands:
        return None
    pt = str(row_pt or "").upper()
    nname = norm_name(name)
    nbrand = norm_name(brand)

    def score(c):
        key = norm_name(c["key"])
        sc = c.get("score") or 0
        # 名称重叠（允许动力词/后缀差异）
        if nname and (nname in key or key in nname or norm_name(re.sub(POWER_RE, "", key)) in nname.replace("ev", "")):
            sc += 1000000
        else:
            return -10**9  # 与名称无关的系列不采用
        if nbrand and nbrand in key:
            sc += 50000
        if pt == "BEV" and ("纯电" in key or "ev" in key):
            sc += 100000
        if pt in ("PHEV", "EREV") and any(k in key for k in ("dm", "增程", "混动", "phev", "插混")):
            sc += 100000
        if pt == "BEV" and any(k in key for k in ("dm", "增程", "混动")):
            sc -= 50000
        return -sc + len(key)
    best = max(cands, key=score)
    return best if score(best) > -10**8 else None


POWER_RE = re.compile("|".join(POWER_WORDS), re.I)


def extract_spec_values(vals):
    out = {}
    # 先处理 CLTC，再处理 WLTC 兜底
    order = [k for k in FIELD_MAP]
    for ah_name in order:
        wb_field = FIELD_MAP[ah_name]
        v = vals.get(ah_name)
        if is_missing(v) or str(v).strip() == "-":
            continue
        if wb_field == "纯电续航里程(km)" and wb_field in out:
            continue  # CLTC 已取到
        out.setdefault(wb_field, str(v).strip())
    eng = vals.get("发动机")
    if eng and not is_missing(eng) and str(eng) != "-":
        m = re.search(r"(\d+\.?\d*)\s*L", str(eng))
        if m:
            out.setdefault("发动机排量(mL)", str(int(float(m.group(1)) * 1000)))
    # 纯电车型：单电机时峰值功率=总功率；功率/扭矩合并列由总功率+最大扭矩推导
    ptxt = str(vals.get("能源类型") or "")
    drive = str(vals.get("驱动电机数") or "")
    total = out.get("电机总功率(kW)")
    if total and "单电机" in drive:
        out.setdefault("电机峰值功率(kW)", total)
    if "纯电" in ptxt and total:
        tq = vals.get("最大扭矩(N·m)")
        if tq and not is_missing(tq) and str(tq) != "-":
            out.setdefault("电机功率/扭矩", f"{total}/{tq}")
    return out


def norm_name(s):
    return re.sub(r"[\s（）()·\-]+", "", str(s or "")).lower()


def clean_name(s):
    """取首个名称片段，去掉通用/车型描述占位。"""
    s = str(s or "").strip()
    for sep in (",", "，", ";", "；", "、"):
        s = s.split(sep)[0]
    s = s.strip()
    if s in GENERIC_NAMES or len(s) < 2:
        return ""
    return s


def strip_power(s):
    return re.sub(POWER_RE, "", str(s or ""))


def pick_name(r, hi):
    """优先通用名称（真实车型名），其次车型名称。"""
    for h in ("通用名称", "车型名称"):
        if h in hi:
            v = clean_name(r[hi[h]] if hi[h] < len(r) else None)
            if v and v not in GENERIC_NAMES:
                return v
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=None)
    ap.add_argument("--scope", default="409-410", help="409-410 | older | all | batch list like '341,342'")
    ap.add_argument("--delay", type=float, default=0.25)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    wb_path = args.workbook or sorted(glob.glob(os.path.join(HERE, "NEV公告参数汇总表_合并版*.xlsx")))[-1]
    wb = load_workbook(wb_path)
    ws = wb["NEV公告参数汇总"]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    hi = {h: i for i, h in enumerate(headers)}
    ncol = ws.max_column
    if "媒体校验来源" not in hi:
        hi["媒体校验来源"] = ncol
        ws.cell(row=1, column=ncol + 1, value="媒体校验来源")
        ncol += 1

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(v not in (None, "") for v in row):
            rows.append(list(row))
    print(f"rows={len(rows)} cols={ncol}")

    if args.scope == "all":
        sel = list(enumerate(rows, start=2))  # (excel row, record)
    elif args.scope == "older":
        sel = [(i, r) for i, r in enumerate(rows, start=2)
               if str(r[hi["批次"]]).strip() not in ("409", "410")]
    elif "-" in args.scope:
        lo, hi_ = args.scope.split("-", 1)
        batches = {str(b) for b in range(int(lo), int(hi_) + 1)}
        sel = [(i, r) for i, r in enumerate(rows, start=2)
               if str(r[hi["批次"]]).strip() in batches]
    else:
        batches = set(x.strip() for x in args.scope.split(","))
        sel = [(i, r) for i, r in enumerate(rows, start=2)
               if str(r[hi["批次"]]).strip() in batches]
    print(f"scope={args.scope} selected={len(sel)}")

    cache = {}
    sugg_cache = {}
    record_rows = []
    stats = defaultdict(int)
    filled_cells = []      # (excel_row, col_idx0, value)
    src_updates = {}       # excel_row -> set(url)
    series_count = 0

    for excel_row, r in sel:
        model = str(r[hi["产品型号"]]).strip() if hi["产品型号"] < len(r) else ""
        name = pick_name(r, hi)
        brand = str(r[hi["产品商标"]] or "").strip().removesuffix("牌")
        pt = str(r[hi["动力类型"]] if hi["动力类型"] < len(r) else "")
        if not name:
            continue
        terms = [name, strip_power(name)]
        if brand and len(name) <= 6:
            terms.append(brand + " " + name)
        # 收集多查询候选
        all_cands = []
        for term in terms:
            tkey = norm_name(term)
            if tkey not in sugg_cache:
                try:
                    sugg_cache[tkey] = suggest(term)
                except Exception as e:
                    print(f"  suggest fail {term}: {e}")
                    sugg_cache[tkey] = []
                time.sleep(args.delay)
            all_cands.extend(sugg_cache[tkey])
        cand = find_series(all_cands, pt, name=name, brand=brand)
        key = norm_name(name) + "::" + pt
        if not cand:
            cache[key] = None
            continue
        if key not in cache or cache[key] is None:
            try:
                ls = list_spec(cand["wordid"])
            except Exception as e:
                print(f"  listSpec fail {name}: {e}")
                cache[key] = None
                continue
            time.sleep(args.delay)
            series_vals = []
            for sp in ls["specs"]:
                try:
                    pv = param_conf(cand["wordid"], sp["id"])
                except Exception:
                    continue
                time.sleep(args.delay)
                sv = extract_spec_values(pv)
                if sv:
                    series_vals.append(sv)
            cache[key] = {"seriesid": cand["wordid"], "seriesname": ls["seriesname"],
                          "values": series_vals,
                          "url": f"https://www.autohome.com.cn/config/series/{cand['wordid']}.html"}
            series_count += 1
        info = cache.get(key)
        if not info or not info["values"]:
            continue

        for wf in set(x for sv in info["values"] for x in sv):
            if wf not in hi:
                continue
            i = hi[wf]
            if not is_missing(r[i]):
                continue
            vals = {sv[wf] for sv in info["values"] if wf in sv and not is_missing(sv[wf])}
            vals = {v for v in vals if v and v != "-"}
            if len(vals) == 1:
                val = vals.pop()
                if not args.dry_run:
                    filled_cells.append((excel_row, i, val))
                src_updates.setdefault(excel_row, set()).add(info["url"])
                stats[wf] += 1
                record_rows.append(["媒体补空", str(r[hi["批次"]]).strip(), model,
                                    str(r[hi["企业名称"]]).strip(), name,
                                    str(r[hi["通用名称"]] or "").strip(),
                                    f"{wf}={val}（汽车之家唯一共识）", info["url"]])

    if not args.dry_run:
        for excel_row, col, val in filled_cells:
            ws.cell(row=excel_row, column=col + 1, value=val)
            ws.cell(row=excel_row, column=col + 1).fill = FILL_MEDIA
        for excel_row, urls in src_updates.items():
            cur = ws.cell(row=excel_row, column=hi["媒体校验来源"] + 1).value
            if cur:
                new = str(cur) + ";" + ";".join(sorted(urls))
            else:
                new = ";".join(sorted(urls))
            ws.cell(row=excel_row, column=hi["媒体校验来源"] + 1, value=new)
        if RECORD_SHEET not in wb.sheetnames:
            rs = wb.create_sheet(RECORD_SHEET)
            rs.append(["动作", "来源批次", "产品型号", "企业名称", "车型名称", "通用名称", "说明", "证据URL"])
        else:
            rs = wb[RECORD_SHEET]
        for line in record_rows:
            rs.append(line)
        wb.save(wb_path)
        print(f"saved. filled_cells={len(filled_cells)} record_rows={len(record_rows)} series_used={series_count}")

    print("=== 媒体补空汇总 ===")
    for k, v in sorted(stats.items(), key=lambda x: -x[1]):
        print(f"{k:24s} +{v}")
    print(f"本次媒体补空合计 {sum(stats.values())} 格；识别车系 {series_count} 个；记录 {len(record_rows)} 条")


if __name__ == "__main__":
    main()
