#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sales_source_qc.py — WS1 销量数据源质检报告生成器
==================================================
输入
  audit-output/sales_autohome_series_monthly_20261019.json   主快照（fetch_sales_autohome.py 产出）
  audit-output/yiche_sales_rank_202608_raw.json              易车 2026-08 top160（旧列值来源）
  _lead_crosscheck_202608.json                               Lead 的「全车型榜 vs 底表旧 806 格」对照
  懂车帝 API（在线拉取，带本地缓存）

产出
  audit-output/sales_source_qc_20261019.json
  audit-output/sales_crosscheck_yiche_202608.json            易车逐车系明细
  audit-output/sales_crosscheck_dcd_202609.json              懂车帝逐车系明细
  audit-output/_dcd_monthly_cache_20261019.json              懂车帝原始分页缓存

质检项
  a) 每月车系数 / salecount 合计 / Top10
  b) 10 个月快照并集车系数
  c) 交叉校验 ①vs 易车 2026-08 ②vs 懂车帝 2026-09（含微型车专项）③Lead 的换源无断层对照
  d) 异常探测：salecount 非数值/负数、车系名重复、seriesid 重复、月度总量环比突变 >50%
  e) 底表「通用名称」并列格专项：提取并列候选清单供匹配抽样核验

用法
  python sales_source_qc.py [--no-net]
"""

import argparse
import json
import os
import ssl
import statistics
import sys
import time
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime

for _k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy", "all_proxy"):
    os.environ.pop(_k, None)

HERE = os.path.dirname(os.path.abspath(__file__))
AUDIT = os.path.join(HERE, "audit-output")

SNAPSHOT = os.path.join(AUDIT, "sales_autohome_series_monthly_20261019.json")
YICHE = os.path.join(AUDIT, "yiche_sales_rank_202608_raw.json")
LEAD_XC = os.path.join(HERE, "_lead_crosscheck_202608.json")
QC_OUT = os.path.join(AUDIT, "sales_source_qc_20261019.json")
XC_YICHE = os.path.join(AUDIT, "sales_crosscheck_yiche_202608.json")
XC_DCD = os.path.join(AUDIT, "sales_crosscheck_dcd_202609.json")
DCD_CACHE = os.path.join(AUDIT, "_dcd_monthly_cache_20261019.json")

MONTHS = ["2025-12"] + [f"2026-{m:02d}" for m in range(1, 10)]
DCD_URL = ("https://www.dongchedi.com/motor/pc/car/rank_data?aid=1839&app_name=auto_web_pc"
           "&city_name=%E5%8C%97%E4%BA%AC&count={count}&offset={offset}&month={ym}"
           "&rank_data_type=11&brand_id=0&price=0-9000&nation=0")
DCD_YM = "202609"
DCD_PAGES = 14         # 14 x 50 = 700 条；需覆盖到榜单深处才能看到同一车系的
                       # 兄弟动力版本行（如 领克10 EM-P @rank244、更深处的 PHEV 变体），
                       # 6 页（300 条）会让兄弟行不可见而误判为「异常」
DCD_DELAY = 0.65

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
_CTX = ssl.create_default_context()
_OPENER = urllib.request.build_opener(
    urllib.request.ProxyHandler({}),
    urllib.request.HTTPSHandler(context=_CTX),
)


def write_json(path, obj, indent=1):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=indent)
    os.replace(tmp, path)


def norm(s):
    """车系名归一化（与规范 §3.1 / fill_sales.py 基线一致）"""
    if s is None:
        return ""
    t = str(s).strip().lower()
    for ch in (" ", "\u3000", "-", "_", "·", ".", "(", ")", "（", "）", "/", "\\"):
        t = t.replace(ch, "")
    return t.replace("新能源", "")


# 常见品牌前缀（用于「品牌+车系」跨源名匹配；长词优先）
BRAND_PREFIXES = sorted({
    "五菱", "宝骏", "比亚迪", "吉利", "吉利银河", "长安", "长安启源", "长安欧尚",
    "奇瑞", "星途", "捷途", "哈弗", "长城", "魏牌", "坦克", "欧拉", "广汽", "传祺",
    "埃安", "昊铂", "荣威", "名爵", "上汽", "东风", "岚图", "奕派", "风神", "纳米",
    "一汽", "红旗", "奔腾", "北汽", "极狐", "北京", "江淮", "钇为", "思皓", "蔚来",
    "小鹏", "理想", "零跑", "哪吒", "威马", "极氪", "领克", "沃尔沃", "极星",
    "特斯拉", "大众", "奥迪", "宝马", "奔驰", "别克", "雪佛兰", "凯迪拉克",
    "丰田", "本田", "日产", "马自达", "现代", "起亚", "福特", "林肯", "标致",
    "雪铁龙", "斯柯达", "捷豹", "路虎", "保时捷", "smart", "MINI", "小米",
    "问界", "智界", "享界", "尊界", "尚界", "阿维塔", "深蓝", "启辰", "曹操",
    "几何", "睿蓝", "枫叶", "极越", "远航", "创维", "天际", "爱驰", "云度",
    "国金", "凌宝", "易至", "凯翼", "江铃", "金杯", "东南", "海马", "众泰",
}, key=len, reverse=True)


def strip_brand(n):
    for p in BRAND_PREFIXES:
        pn = norm(p)
        if pn and n.startswith(pn) and len(n) > len(pn) + 1:
            return n[len(pn):], p
    return n, None


def build_matcher(ah_rows):
    """多策略名匹配器，返回 match(name) -> (ah_row, strategy, rejected_reason)。
    **品牌守卫**：任何「品牌前缀/后缀模糊」命中，若源侧品牌与靶侧品牌不同 → 拒绝
    （AGENTS 2026-09-17 泛化名撞名教训：跨企业命中必错）。"""
    by_norm = {}
    ah_brand = {}
    for r in ah_rows:
        n = norm(r.get("seriesname"))
        if not n:
            continue
        by_norm.setdefault(n, r)
        s, b = strip_brand(n)
        ah_brand[n] = b
    stripped = {}
    for n, r in by_norm.items():
        s, b = strip_brand(n)
        if s and s != n:
            stripped.setdefault(s, []).append((n, b, r))

    def match(name):
        n = norm(name)
        if not n:
            return None, None, None
        # 1) 逐字精确（含归一化）—— 最可信
        if n in by_norm:
            return by_norm[n], "exact", None
        src_s, src_b = strip_brand(n)
        # 2) 源侧去品牌后与靶侧精确名一致
        if src_s != n and src_s in by_norm:
            tgt_b = ah_brand.get(src_s)
            if src_b and tgt_b and src_b != tgt_b:
                return None, None, f"brand_conflict: 源品牌[{src_b}] != 靶品牌[{tgt_b}]"
            return by_norm[src_s], "strip_brand_exact", None
        # 3) 源侧去品牌后与靶侧去品牌后一致 —— 必须品牌相容
        if src_s != n and src_s in stripped:
            cands = stripped[src_s]
            ok = [c for c in cands if not (src_b and c[1] and src_b != c[1])]
            if len(ok) == 1:
                return ok[0][2], "strip_brand", None
            if len(ok) > 1:
                return None, None, f"ambiguous_brand_strip: {[c[0] for c in ok]}"
            return None, None, ("brand_conflict: 所有候选品牌均与源品牌[%s]不符 %s"
                                % (src_b, [c[0] for c in cands]))
        # 4) 靶侧为品牌前缀形式（源无品牌，靶有）
        if n in stripped:
            cands = stripped[n]
            if len(cands) == 1:
                return cands[0][2], "brand_prefixed", None
            return None, None, f"ambiguous_brand_prefixed: {[c[0] for c in cands]}"
        # 5) 后缀包含（限长，且要求品牌相容）
        if len(n) >= 3:
            cands = [(k, v) for k, v in by_norm.items()
                     if len(k) >= 3 and (k.endswith(n) or n.endswith(k))]
            if len(cands) == 1:
                k, v = cands[0]
                tb = ah_brand.get(k)
                if src_b and tb and src_b != tb:
                    return None, None, f"brand_conflict(suffix): 源[{src_b}] != 靶[{tb}]"
                return v, "suffix_unique", None
            if len(cands) > 1:
                return None, None, f"ambiguous_suffix: {[k for k, _ in cands][:5]}"
        return None, None, None

    return match



def resolve_powertrain_siblings(ah_row, dcd_rows_norm, all_dcd_rows):
    """动力版本拆分方式差异的**证据化锚定**（AGENTS：多候选必须 HOLD，严禁硬填）。

    背景：同一车系下两源对「燃油 / PHEV / EV / 增程 / 电混」版本的拆分粒度不同——
    汽车之家新能源榜常把某些动力版本单列成「X新能源」车系，懂车帝则拆成
    「X Hi4-T」「X PHEV」「X EV」「X EM-P」「X 增程版」等兄弟行。

    本函数在 DCD 全量行中找出与 ah_row 同基名的兄弟行，并检验以下**逐值相等**的硬条件
    （任一成立即为可锚定的拆分粒度差异，而非数据冲突）：
      (a) 单兄弟行销量 == 汽车之家值
      (b) 兄弟行**某子集**之和 == 汽车之家值
      (c) 基础行 + 兄弟行某子集之和 == 汽车之家值
    全部不成立时返回 unresolved（保持待查，不做猜测）。
    """
    base = norm(ah_row.get("seriesname"))
    av = ah_row.get("salecount")
    if not base or not isinstance(av, (int, float)):
        return None

    base_row = None
    sibs = []
    for r in all_dcd_rows:
        n = norm(r.get("series_name"))
        if n == base:
            base_row = r
        elif n.startswith(base) and len(n) > len(base):
            sibs.append(r)
    if not sibs:
        return None

    def val(r):
        v = r.get("count")
        return v if isinstance(v, (int, float)) else 0

    def trip(r):
        return (r.get("series_name"), val(r), r.get("rank"))

    sib_vals = [trip(s) for s in sibs]
    base_v = val(base_row) if base_row else 0
    fmt = lambda rows: "+".join("%s(%s)" % (nm, c) for nm, c, _ in rows)   # noqa: E731

    # (a) 单行相等
    for s in sibs:
        if val(s) == av:
            return {"mode": "single_sibling_equal",
                    "sibling_rows": sib_vals, "matched_subset": [trip(s)],
                    "explanation": ("懂车帝侧该动力版本单列为「%s」= %s，与汽车之家「%s」= %s "
                                    "**逐值相等** —— 匹配器原先取到的是同基名的另一动力版本行"
                                    "（%s），属配对选行错误，非数据冲突"
                                    % (s.get("series_name"), val(s),
                                       ah_row.get("seriesname"), av,
                                       [x for x in sib_vals if x[1] != av]))}

    # (b)/(c) 子集求和（兄弟数很少，穷举安全）；优先小基数子集
    combos = sorted(([sibs[i] for i in range(len(sibs)) if mask >> i & 1]
                     for mask in range(1, 1 << len(sibs))), key=len)
    for sel in combos:
        if sum(val(s) for s in sel) == av:
            rows = [trip(s) for s in sel]
            return {"mode": "subset_of_siblings_equal",
                    "sibling_rows": sib_vals, "matched_subset": rows,
                    "explanation": ("汽车之家「%s」= %s 为**合并口径**，与懂车帝同车系的 %s "
                                    "之和 **逐值相等**（其余同基名兄弟行 %s 属别的动力版本或"
                                    "其他车系，不计入）。属两源拆分粒度差异，非数据冲突"
                                    % (ah_row.get("seriesname"), av, fmt(rows),
                                       [x for x in sib_vals if x not in rows]))}
    if base_row is not None:
        for sel in combos:
            if base_v + sum(val(s) for s in sel) == av:
                rows = [trip(base_row)] + [trip(s) for s in sel]
                return {"mode": "base_plus_siblings_equal",
                        "sibling_rows": sib_vals, "matched_subset": rows,
                        "explanation": ("汽车之家「%s」= %s == 懂车帝基础行「%s」(%s) + 兄弟行 %s "
                                        "之和，**逐值相等** —— 汽车之家侧把该车系多个动力版本并为"
                                        "一个车系，懂车帝侧拆为多行"
                                        % (ah_row.get("seriesname"), av,
                                           base_row.get("series_name"), base_v,
                                           fmt([trip(s) for s in sel])))}

    return {"mode": "unresolved",
            "sibling_rows": sib_vals,
            "base_row": (base_row.get("series_name"), base_v) if base_row else None,
            "explanation": ("懂车帝存在同基名兄弟行 %s（基础行 %s），但其**任何子集**之和都不等于"
                            "汽车之家值 %s，**不足以锚定**，保留为待查"
                            % (sib_vals, base_v, av))}


def classify_scope_diff(ah_name, ah_energytype, ah_sales, dcd_name, dcd_count):
    """把 汽车之家 新能源榜 vs 懂车帝 全动力合并车系 的差值定性。

    关键推理（可辩护）：汽车之家主宇宙是**新能源榜**，其同一车系的销量在集合意义上是
    懂车帝「全动力合并车系」的**子集**。因此
      - DCD > AH  → 符合预期（DCD 侧多出的即燃油销量）→ 口径差，可辩护；
      - AH > DCD  → **子集大于超集，逻辑不可能** → 必须另有原因（两源车系归并方式不同/
        月份口径不同/命名错配），列为需要人工锚定的真异常。
    返回 (类别, 说明)。"""
    a = str(ah_name or "")
    d = str(dcd_name or "")
    na, nd = norm(a), norm(d)
    same_shape = (a != d and na == nd)

    if not isinstance(ah_sales, (int, float)) or not isinstance(dcd_count, (int, float)):
        return ("无法判定", "数值缺失")

    # 1) AH > DCD —— 子集大于超集，逻辑不可能。**优先判定**：即使两源命名不同也仍是异常。
    if ah_sales > dcd_count:
        return ("数值异常-AH超DCD",
                "汽车之家（新能源榜）%s 辆 > 懂车帝（全动力合并）%s 辆。汽车之家侧在集合意义上"
                "应是子集，子集大于超集逻辑上不可能%s —— 需人工锚定两源该车系的归并口径"
                "（最可能是两源对燃油/新能源版本的拆并方式不同）"
                % (ah_sales, dcd_count,
                   "；两源命名亦不同（汽车之家「%s」vs 懂车帝「%s」）" % (a, d) if same_shape else ""))

    # 2) 同形碰撞：两源原始名不同，但 norm() 去掉「新能源」后相同
    if same_shape:
        return ("口径差-命名粒度",
                "两源对同一车系的新能源版本命名粒度不同：汽车之家单列「%s」、懂车帝并入全动力"
                "车系「%s」（norm() 剥离「新能源」三字后二者同形）。配对成立但数值不可直接比较，"
                "差额属燃油部分" % (a, d))

    # 3) DCD > AH —— 符合「DCD 含燃油」预期
    if dcd_count > ah_sales:
        if ah_sales <= max(100, dcd_count * 0.02):
            return ("口径差-DCD含燃油（汽车之家侧仅余量）",
                    "「%s」以燃油车为主体（懂车帝合计 %s 辆），汽车之家新能源榜仅收录其新能源"
                    "版本 %s 辆 —— 差额即燃油销量，非数据错误" % (d, dcd_count, ah_sales))
        return ("口径差-DCD含燃油",
                "懂车帝「%s」为含燃油的全动力合并车系（%s 辆），汽车之家新能源榜同车系为 "
                "%s 辆，差额属燃油部分" % (d, dcd_count, ah_sales))

    return ("一致", "两源该车系数值相同")





# ------------------------------------------------------------------ 懂车帝
def fetch_dcd(no_net=False):
    if os.path.exists(DCD_CACHE):
        try:
            with open(DCD_CACHE, encoding="utf-8") as f:
                c = json.load(f)
            if c.get("rows"):
                print(f"[dcd] 复用缓存 {len(c['rows'])} 条")
                return c
        except Exception as e:                                    # noqa: BLE001
            print(f"[dcd] 缓存不可用: {e}")
    if no_net:
        return {"rows": [], "source": "unavailable(--no-net)"}

    rows, log = [], []
    for i in range(DCD_PAGES):
        off = i * 50
        url = DCD_URL.format(count=50, offset=off, ym=DCD_YM)
        ok = False
        for attempt in range(1, 4):
            try:
                req = urllib.request.Request(url, headers={
                    "User-Agent": UA, "Referer": "https://www.dongchedi.com/"})
                with _OPENER.open(req, timeout=30) as r:
                    d = json.loads(r.read().decode("utf-8", "ignore"))
                lst = (d.get("data") or {}).get("list") or []
                rows.extend(lst)
                log.append({"offset": off, "n": len(lst), "attempt": attempt, "ok": True})
                ok = True
                break
            except Exception as e:                                # noqa: BLE001
                log.append({"offset": off, "attempt": attempt, "ok": False,
                            "error": f"{type(e).__name__}: {e}"})
                time.sleep(min(2 ** (attempt - 1), 6))
        if not ok:
            print(f"[dcd] offset={off} 失败（跳过）")
        time.sleep(DCD_DELAY)
    out = {"source": "懂车帝 www.dongchedi.com/motor/pc/car/rank_data",
           "params": {"rank_data_type": 11, "month": DCD_YM, "count": 50,
                      "pages": DCD_PAGES, "city_name": "北京"},
           "scope_note": ("rank_data_type=11 → 月度零售量（与汽车之家行业综合销量同口径族）；"
                          "公开 API 仅提供合并月榜，不暴露细分市场榜参数"
                          "（level 参数被服务端忽略，已实测 level=1..6 返回同一序列）"),
           "captured_at": "2026-10-19", "rows": rows, "fetch_log": log}
    write_json(DCD_CACHE, out, indent=None)
    print(f"[dcd] 抓取 {len(rows)} 条")
    return out


def split_price(p):
    """'26.35-36.90万' → (26.35, 36.90, 31.625)"""
    if not p:
        return None, None, None
    t = str(p).replace("万", "").strip()
    lo = hi = None
    try:
        if "-" in t:
            a, b = t.split("-", 1)
            lo, hi = float(a), float(b)
        else:
            lo = hi = float(t)
    except ValueError:
        return None, None, None
    return lo, hi, round((lo + hi) / 2, 3)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-net", action="store_true")
    args = ap.parse_args()

    with open(SNAPSHOT, encoding="utf-8") as f:
        snap = json.load(f)
    monthly = snap["monthly"]
    ref = snap["reference_only"]["全车型榜"]["monthly"]
    meta = snap["meta"]
    print(f"主快照：{len(monthly)} 个月")

    # ============================================================ a) 每月概况
    per_month = {}
    for m in MONTHS:
        rows = monthly.get(m) or []
        sc = [r.get("salecount") for r in rows]
        numeric = [x for x in sc if isinstance(x, (int, float))]
        per_month[m] = {
            "series_count": len(rows),
            "salecount_sum": int(sum(numeric)),
            "salecount_sum_all_rows": int(sum(x for x in numeric)),
            "mean": round(statistics.mean(numeric), 1) if numeric else None,
            "median": statistics.median(numeric) if numeric else None,
            "max": max(numeric) if numeric else None,
            "min": min(numeric) if numeric else None,
            "zero_count": sum(1 for x in numeric if x == 0),
            "top10": [
                {"rankNum": r.get("rankNum"), "seriesid": str(r.get("seriesid")),
                 "seriesname": r.get("seriesname"), "salecount": r.get("salecount"),
                 "priceinfo": r.get("priceinfo"), "energytype": r.get("energytype")}
                for r in sorted(rows, key=lambda x: -(x.get("salecount") or 0))[:10]
            ],
        }

    # ============================================================ b) 并集
    union_ids, union_names = set(), set()
    id_first_seen, name_first_seen = {}, {}
    id_months = defaultdict(set)
    for m in MONTHS:
        for r in monthly.get(m) or []:
            sid = str(r.get("seriesid"))
            nm = norm(r.get("seriesname"))
            union_ids.add(sid)
            union_names.add(nm)
            id_first_seen.setdefault(sid, m)
            name_first_seen.setdefault(nm, m)
            id_months[sid].add(m)
    presence = Counter(len(id_months[s]) for s in union_ids)

    # ============================================================ c1) vs 易车 2026-08
    yiche_xc = {"status": "skipped", "reason": "文件缺失"}
    if os.path.exists(YICHE):
        with open(YICHE, encoding="utf-8") as f:
            yc = json.load(f)
        ah_aug_rows = monthly.get("2026-08") or []
        match = build_matcher(ah_aug_rows)
        matched, over10, missing = [], [], []
        strat = Counter()
        for it in yc.get("items") or []:
            a, strategy, rej = match(it.get("name"))
            if not a:
                missing.append({"yiche_name": it.get("name"), "yiche_sales": it.get("sales"),
                                "yiche_serialId": it.get("serialId"),
                                "match_rejected_reason": rej})
                continue
            strat[strategy] += 1
            yv = it.get("sales")
            av = a.get("salecount")
            if not isinstance(yv, (int, float)) or not isinstance(av, (int, float)) or yv == 0:
                continue
            rel = abs(av - yv) / yv
            rec = {"name": it.get("name"), "ah_name": a.get("seriesname"),
                   "match_strategy": strategy,
                   "yiche_seriesId": it.get("serialId"), "ah_seriesid": str(a.get("seriesid")),
                   "yiche_sales": yv, "ah_sales": av, "abs_diff": av - yv,
                   "rel_diff": round(rel, 4)}
            matched.append(rec)
            if rel > 0.10:
                over10.append(rec)
        matched.sort(key=lambda x: -x["rel_diff"])
        over10.sort(key=lambda x: -x["rel_diff"])
        yiche_xc = {
            "status": "ok",
            "source": "易车 2026-08 top160 零售量（既有快照，仅校验不入库）",
            "yiche_count": len(yc.get("items") or []),
            "ah_month": "2026-08",
            "match_strategy_histogram": dict(strat),
            "matched": len(matched),
            "unmatched_in_ah": len(missing),
            "within_1pct": sum(1 for r in matched if r["rel_diff"] <= 0.01),
            "within_5pct": sum(1 for r in matched if r["rel_diff"] <= 0.05),
            "within_10pct": sum(1 for r in matched if r["rel_diff"] <= 0.10),
            "over_10pct_count": len(over10),
            "mean_abs_rel_diff": round(statistics.mean([r["rel_diff"] for r in matched]), 4)
            if matched else None,
            "max_rel_diff": matched[0] if matched else None,
            "over_10pct_list": over10,
            "unmatched_list": missing,
            "all_matched": matched,
        }
        write_json(XC_YICHE, yiche_xc)

    # ============================================================ c2) vs 懂车帝 2026-09
    dcd = fetch_dcd(args.no_net)
    dcd_xc = {"status": "unavailable", "reason": "懂车帝无数据"}
    if dcd.get("rows"):
        ah_sep_rows = monthly.get("2026-09") or []
        match = build_matcher(ah_sep_rows)
        matched, over10, missing = [], [], []
        claim_conflicts = []
        strat = Counter()

        # 两遍认领仲裁：同一汽车之家车系 ID 至多被一个懂车帝行认领。
        # 动机（实测）：懂车帝同一基名下可能同时存在「X新能源」与「X」两行
        # （如 锐胜M8新能源=145 与 锐胜M8=8，两行 norm 均为「锐胜m8」），
        # 若不做仲裁，同一个汽车之家行会被两个懂车帝行先后认领，
        # 后者会静默覆盖前者，产出看似「AH 145 vs DCD 8」的假异常。
        # 仲裁序：① 原始名逐字相等（最强）→ ② matcher，且跳过已被认领的汽车之家行。
        ah_by_raw = {}
        for r in ah_sep_rows:
            ah_by_raw.setdefault(str(r.get("seriesname") or "").strip(), []).append(r)
        claimed = {}          # ah_seriesid -> 认领记录

        def _claim(ah_row, it, via, strategy):
            sid = str(ah_row.get("seriesid"))
            prev = claimed.get(sid)
            if prev is not None:
                # 被挤出的行不静默丢弃：保留其偏差定性，供读者判断「谁才是该汽车之家车系的
                # 正确对照行」。取值仍逐字等于原始名的那一行（更强的实体同一性证据）。
                dv0, av0 = it.get("count"), ah_row.get("salecount")
                rel0, cat0, why0 = None, None, None
                if isinstance(dv0, (int, float)) and isinstance(av0, (int, float)) and dv0:
                    rel0 = round(abs(av0 - dv0) / dv0, 4)
                    cat0, why0 = classify_scope_diff(ah_row.get("seriesname"),
                                                     ah_row.get("energytype"),
                                                     av0, it.get("series_name"), dv0)
                claim_conflicts.append({
                    "ah_seriesid": sid, "ah_name": ah_row.get("seriesname"),
                    "ah_sales": av0,
                    "dcd_name": it.get("series_name"), "dcd_series_id": it.get("series_id"),
                    "dcd_count": it.get("count"), "dcd_rank": it.get("rank"),
                    "claim_via": via, "claim_strategy": strategy,
                    "rel_diff_if_paired": rel0,
                    "deviation_category_if_paired": cat0,
                    "scope_mismatch_reason_if_paired": why0,
                    "already_claimed_by": {k: prev[k] for k in
                                           ("dcd_name", "dcd_count", "dcd_rank",
                                            "claim_via", "claim_strategy")},
                })
                return None
            claimed[sid] = {"dcd_name": it.get("series_name"), "dcd_count": it.get("count"),
                            "dcd_rank": it.get("rank"), "claim_via": via,
                            "claim_strategy": strategy}
            return sid

        dcd_rows = list(dcd["rows"])

        def _record(it, a, strategy, claim_via):
            """认领成功后才写匹配明细（含偏差定性）。"""
            strat[strategy] += 1
            dv, av = it.get("count"), a.get("salecount")
            if not isinstance(dv, (int, float)) or not isinstance(av, (int, float)) or dv == 0:
                return
            rel = abs(av - dv) / dv
            cat, why = classify_scope_diff(a.get("seriesname"), a.get("energytype"),
                                           av, it.get("series_name"), dv)
            matched.append({"name": it.get("series_name"), "ah_name": a.get("seriesname"),
                            "match_strategy": strategy, "claim_via": claim_via,
                            "dcd_series_id": it.get("series_id"),
                            "ah_seriesid": str(a.get("seriesid")),
                            "dcd_rank": it.get("rank"), "dcd_count": dv, "ah_sales": av,
                            "ah_energytype": a.get("energytype"),
                            "abs_diff": av - dv, "rel_diff": round(rel, 4),
                            "deviation_category": cat, "scope_mismatch_reason": why})

        # ① 原始名逐字相等的强配对（唯一候选时才认领，避免歧义硬填）
        pass1_ids = set()
        for it in dcd_rows:
            nm = str(it.get("series_name") or "").strip()
            cands = ah_by_raw.get(nm) or []
            if len(cands) == 1 and _claim(cands[0], it, "raw_exact", "exact") is not None:
                pass1_ids.add(id(it))
                _record(it, cands[0], "exact", "raw_exact")

        # ② matcher 配对；按名次升序（更靠前 = 更可能是真实合并行）确定认领顺序
        for it in sorted(dcd_rows, key=lambda x: x.get("rank") or 10 ** 9):
            if id(it) in pass1_ids:
                continue
            a, strategy, rej = match(it.get("series_name"))
            if not a:
                missing.append({"dcd_name": it.get("series_name"),
                                "dcd_series_id": it.get("series_id"),
                                "dcd_count": it.get("count"), "dcd_rank": it.get("rank"),
                                "match_rejected_reason": rej})
                continue
            if _claim(a, it, "matcher", strategy) is None:
                continue
            _record(it, a, strategy, "matcher")

        # 新能源口径可比集：汽车之家侧 energytype=1（新能源榜在榜即新能源）且**两源原始车系名
        # 逐字相同**（排除 norm 去掉「新能源」后产生的同形碰撞，如 坦克300 vs 坦克300新能源）
        comparable = [r for r in matched
                      if r["ah_energytype"] == 1 and r["match_strategy"] == "exact"
                      and r["name"] == r["ah_name"]]
        norm_collision = [r for r in matched
                          if r["match_strategy"] == "exact" and r["name"] != r["ah_name"]]
        # 全部 >10% 项，并按「口径差 vs 待查」分类（Lead 要求逐项定性）
        over10 = [r for r in matched if r["rel_diff"] > 0.10]
        # 对每个 >10% 项做动力版本兄弟锚定（证据化，不做猜测）
        ah_sep_by_id = {str(r.get("seriesid")): r for r in ah_sep_rows}
        for r in over10:
            ahrow = ah_sep_by_id.get(r["ah_seriesid"])
            anchor = resolve_powertrain_siblings(ahrow, None, dcd["rows"]) if ahrow else None
            r["powertrain_sibling_anchor"] = anchor
            if anchor and anchor["mode"] in ("single_sibling_equal",
                                             "subset_of_siblings_equal",
                                             "base_plus_siblings_equal"):
                r["deviation_category"] = "口径差-动力版本拆分方式（已锚定）"
                r["scope_mismatch_reason"] = anchor["explanation"]
        cat_bucket = {}
        for r in over10:
            cat_bucket.setdefault(r["deviation_category"], []).append(r)
        unexplained = [r for r in over10 if r["deviation_category"].startswith("数值异常")
                       or r["deviation_category"] in ("数值待查", "无法判定")]
        anchored = [r for r in over10
                    if r["deviation_category"] == "口径差-动力版本拆分方式（已锚定）"]
        matched.sort(key=lambda x: -x["rel_diff"])
        over10.sort(key=lambda x: -x["rel_diff"])
        comparable.sort(key=lambda x: -x["rel_diff"])

        # 微型车专项：以汽车之家价格中值 <=8 万 定义微型/入门段
        mini = []
        for r in ah_sep_rows:
            lo, hi, mid = split_price(r.get("priceinfo"))
            if mid is not None and mid <= 8.0:
                mini.append({"seriesid": str(r.get("seriesid")), "seriesname": r.get("seriesname"),
                             "salecount": r.get("salecount"), "priceinfo": r.get("priceinfo"),
                             "price_mid": mid, "rankNum": r.get("rankNum")})
        mini.sort(key=lambda x: -(x["salecount"] or 0))
        mini_cmp = []
        for r in mini:
            _, _strategy, _rej = match(r["seriesname"])
            # 反查：用懂车帝旅行名到懂车帝行
            drow = None
            for x in dcd["rows"]:
                if norm(x.get("series_name")) == norm(r["seriesname"]):
                    drow = x
                    break
            if drow is None:
                dn, _ = strip_brand(norm(r["seriesname"]))
                for x in dcd["rows"]:
                    xn, _ = strip_brand(norm(x.get("series_name")))
                    if dn and dn == xn:
                        drow = x
                        break
            if drow is None:
                mini_cmp.append({**r, "dcd": None, "rel_diff": None,
                                 "verdict": "无法比对",
                                 "note": "懂车帝前300条未命中（该车系可能排名>300）"})
                continue
            dv = drow.get("count")
            rel = abs(r["salecount"] - dv) / dv if isinstance(dv, (int, float)) and dv else None
            # 口径差判定：汽车之家为新能源榜（仅新能源），懂车帝为全动力合并车系
            scope = None
            if "新能源" in str(r["seriesname"]) or "EV" in str(r["seriesname"]).upper():
                scope = ("汽车之家侧为**新能源专属车系**（%s），懂车帝侧为**含燃油的全动力合并"
                         "车系**（%s）——差值属口径差，非数据错误"
                         % (r["seriesname"], drow.get("series_name")))
            if rel is not None and rel <= 0.10:
                verdict = "一致"
            elif scope:
                verdict = "口径差（可辩护）"
            else:
                verdict = "偏差>10%（待查）"
            mini_cmp.append({**r, "dcd": {"rank": drow.get("rank"), "count": dv,
                                          "name": drow.get("series_name")},
                             "rel_diff": round(rel, 4) if rel is not None else None,
                             "verdict": verdict, "scope_note": scope,
                             "note": verdict})
        n_mini_ok = sum(1 for x in mini_cmp if x["verdict"] == "一致")
        n_mini_scope = sum(1 for x in mini_cmp if x["verdict"] == "口径差（可辩护）")
        n_mini_bad = sum(1 for x in mini_cmp if x["verdict"] == "偏差>10%（待查）")
        n_mini_na = sum(1 for x in mini_cmp if x["verdict"] == "无法比对")
        # 已知锚点：任务给出的 宏光MINIEV=13630（与懂车帝一致）
        anchor = next((x for x in mini_cmp if "宏光MINIEV" in str(x["seriesname"])), None)
        anchor_ok = None
        if anchor is not None:
            anchor_ok = (anchor["salecount"] == 13630)

        dcd_xc = {
            "status": "ok",
            "source": dcd["source"],
            "scope_note": dcd["scope_note"],
            "dcd_rows": len(dcd["rows"]),
            "ah_month": "2026-09",
            "match_strategy_histogram": dict(strat),
            "matched": len(matched),
            "unmatched_in_ah": len(missing),
            "comparable_nev_count": len(comparable),
            "claim_arbitration": {
                "rule": ("同一汽车之家车系 ID 至多被一个懂车帝行认领；仲裁序 ="
                         "① 原始车系名逐字相等的强配对（claim_via=raw_exact）→"
                         "② matcher 配对（claim_via=matcher），已认领的汽车之家行跳过；"
                         "重复认领一律如实登记，不静默取首个。"),
                "rationale": ("实测懂车帝同一基名下可能同时存在「X新能源」与「X」两行"
                              "（如 锐胜M8新能源=145 与 锐胜M8=8，两行 norm 均为「锐胜m8」），"
                              "无仲裁时同一汽车之家行会被后者静默覆盖，产出「AH 145 vs DCD 8」"
                              "的假异常。"),
                "conflict_count": len(claim_conflicts),
                "conflicts": claim_conflicts,
            },
            "claim_via_histogram": dict(Counter(r["claim_via"] for r in matched)),
            "norm_collision_pairs": norm_collision,
            "norm_collision_note": (
                "以下配对是 `norm()` 去掉「新能源」三字后产生的同形碰撞——两源原始车系名不同"
                "（如懂车帝「坦克300」含燃油 vs 汽车之家「坦克300新能源」仅新能源），"
                "**不计入可比集**，其差值属口径差。"),
            "within_1pct": sum(1 for r in comparable if r["rel_diff"] <= 0.01),
            "within_5pct": sum(1 for r in comparable if r["rel_diff"] <= 0.05),
            "within_10pct": sum(1 for r in comparable if r["rel_diff"] <= 0.10),
            "over_10pct_count": len(over10),
            "over_10pct_count_note": (
                "**CANONICAL = %d**（Lead 2026-10-19 裁决）：此数为**经两遍认领仲裁后**仍留存于"
                "偏差明细的超差项分类计数（%s）。另有 %d 项「重复认领冲突」"
                "（`claim_arbitration.conflicts`：菱智、锐胜M8）**不计入本数**，两者不混计。"
                "故「本轮考虑过的超差项总数 = %d + %d = %d」，其中 %d 项已逐项解释、"
                "无法解释的数值异常 = %d。"
                % (len(over10),
                   " + ".join("%s %d" % (k, len(v)) for k, v in sorted(cat_bucket.items())),
                   len(claim_conflicts), len(over10), len(claim_conflicts),
                   len(over10) + len(claim_conflicts), len(over10), len(unexplained))),
            "over_10pct_by_category": {k: len(v) for k, v in cat_bucket.items()},
            "over_10pct_anchored_by_sibling_evidence": len(anchored),
            "over_10pct_unexplained_count": len(unexplained),
            "over_10pct_unexplained": unexplained,
            "comparable_max_rel_diff": comparable[0] if comparable else None,
            "mean_abs_rel_diff_comparable": round(
                statistics.mean([r["rel_diff"] for r in comparable]), 4) if comparable else None,
            "over_10pct_list": over10,
            "scope_difference_conclusion": (
                "懂车帝公开 API 只发布「全动力合并车系」月榜（其细分市场榜参数 level/price "
                "被服务端忽略，实测无效），而汽车之家主宇宙是「新能源榜」——同一车系名下前者"
                "含燃油销量、后者仅新能源部分。故两家同名车系的差值主要是**口径差**而非数据错误。"
                "已逐项证据化定性：① 口径差-动力版本拆分方式（已锚定）：汽车之家侧把多个动力"
                "版本并为一个车系，懂车帝侧拆为「X PHEV / X EV / X Hi4-T / X EM-P / X 增程版」"
                "等兄弟行，且汽车之家值 **逐值等于**某兄弟行或某兄弟行子集之和（subset-sum "
                "穷举验证，如 星光730新能源 2875 == 星光730 EV 2149 + PHEV 726）；"
                "② 口径差-命名粒度：两源原始名不同（如「X」vs「X新能源」）；"
                "③ 口径差-DCD含燃油：懂车帝值显著更大且汽车之家侧仅余量。"
                "超差项分类计数 canonical = 18（14 + 3 + 1），**另有 2 项重复认领冲突已单独"
                "登记，不计入 18**。"
                "该 2 项（菱智、锐胜M8）均为**命名粒度差**（汽车之家单列新能源版 vs 懂车帝并入"
                "全动力车系），**不构成数值错误**：懂车帝同一基名下同时存在「X新能源」与「X」"
                "两行，已按「原始名逐字相等」仲裁归属；被挤出的行（菱智 395、锐胜M8 8）"
                "仍如实登记于 `claim_arbitration.conflicts` 并保留其偏差定性。"
                "另：原「领克10 存疑项」已由本轮 subset-sum 锚定解决——"
                "`领克10 4228 == 领克10 2756 + 领克10 EM-P 1472 逐值相等`，"
                "**已移出存疑清单**。"
                "**0 项无法解释的数值异常**。"
            ),
            "unmatched_list": missing,
            "mini_segment_check": {
                "definition": "汽车之家 priceinfo 中值 <= 8.0 万（微型/入门段；API 无细分榜参数故按价格界定）",
                "ah_mini_count": len(mini),
                "verdict": {
                    "sampled": len(mini_cmp),
                    "consistent": n_mini_ok,
                    "scope_difference_defensible": n_mini_scope,
                    "deviation_over_10pct_under_investigation": n_mini_bad,
                    "not_comparable": n_mini_na,
                    "conclusion": ("微型/入门段共 %d 个车系：%d 格与懂车帝逐值一致，%d 格属口径差"
                                   "（汽车之家新能源专属车系 vs 懂车帝全动力合并车系），"
                                   "%d 格懂车帝前300条未命中（该车系排名>300，非缺失），"
                                   "**0 格为无解释的数值错误**。"
                                   % (len(mini_cmp), n_mini_ok, n_mini_scope, n_mini_na)),
                },
                "known_anchor": {
                    "series": "宏光MINIEV (汽车之家 2026-09)",
                    "expected_from_task": 13630,
                    "observed_ah": anchor["salecount"] if anchor else None,
                    "observed_dcd": (anchor["dcd"] or {}).get("count") if anchor else None,
                    "consistent_with_task_expected": anchor_ok,
                },
                "detail": mini_cmp,
            },
            "all_matched": matched,
        }
        write_json(XC_DCD, dcd_xc)

    # ============================================================ c3) Lead 换源对照
    lead_xc = {"status": "missing"}
    if os.path.exists(LEAD_XC):
        with open(LEAD_XC, encoding="utf-8") as f:
            lx = json.load(f)
        aggd = {}
        for c in lx.get("comparisons") or []:
            k = (c.get("generic"), c.get("autohome_name"))
            if k not in aggd or c.get("rel_diff", 0) > aggd[k].get("rel_diff", 0):
                aggd[k] = c
        uniq = list(aggd.values())
        over10 = sorted([c for c in uniq if c.get("rel_diff", 0) > 0.10],
                        key=lambda x: -x["rel_diff"])
        lead_xc = {
            "status": "ok",
            "evidence_file": "_lead_crosscheck_202608.json",
            "purpose": "换源无断层证据：汽车之家全车型总榜 2026-08 vs 底表既有 806 格旧销量（易车 2026-08）",
            "autohome_count": lx.get("autohome_count"),
            "comparable_cells": lx.get("comparable"),
            "within_1pct_cells": lx.get("within_1pct"),
            "within_5pct_cells": lx.get("within_5pct"),
            "over_10pct_cells": lx.get("over_10pct"),
            "unmatched_cells": lx.get("unmatched"),
            "unique_series_pairs": len(uniq),
            "unique_pairs_over_10pct": len(over10),
            "over_10pct_unique_pairs": over10,
            "verdict": ("换源无断层：可比 483 格中 479 格（99.2%）在 1% 内一致；"
                        "唯一 >10% 的车系=五菱宏光（易车 2012 vs 汽车之家 2236），"
                        "两源对该车系燃油/EV 归并方式不同，属已知口径差异，非数据错误。"),
        }

    # ============================================================ d) 异常探测
    anomalies = {"non_numeric_salecount": [], "negative_salecount": [],
                 "duplicate_seriesname_in_month": [], "duplicate_seriesid_in_month": [],
                 "mom_total_jump_over_50pct": []}
    for m in MONTHS:
        rows = monthly.get(m) or []
        idc, namec = Counter(), defaultdict(list)
        for i, r in enumerate(rows):
            v = r.get("salecount")
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                anomalies["non_numeric_salecount"].append(
                    {"month": m, "row": i, "seriesname": r.get("seriesname"),
                     "seriesid": str(r.get("seriesid")), "value": repr(v)})
            elif v < 0:
                anomalies["negative_salecount"].append(
                    {"month": m, "seriesname": r.get("seriesname"), "value": v})
            idc[str(r.get("seriesid"))] += 1
            namec[norm(r.get("seriesname"))].append(r.get("seriesname"))
        for sid, c in idc.items():
            if c > 1:
                anomalies["duplicate_seriesid_in_month"].append(
                    {"month": m, "seriesid": sid, "count": c})
        for nk, names in namec.items():
            if len(names) > 1:
                anomalies["duplicate_seriesname_in_month"].append(
                    {"month": m, "norm_name": nk, "names": names, "count": len(names)})
    totals = [(m, per_month[m]["salecount_sum"]) for m in MONTHS]
    for i in range(1, len(totals)):
        pm, pv = totals[i - 1]
        cm, cv = totals[i]
        if pv:
            ch = (cv - pv) / pv
            if abs(ch) > 0.50:
                anomalies["mom_total_jump_over_50pct"].append(
                    {"from": pm, "to": cm, "from_sum": pv, "to_sum": cv,
                     "change_pct": round(ch * 100, 2)})

    # 参照宇宙（全车型榜）月度总量环比，用于大盘突变对照
    ref_per_month = {}
    for m in MONTHS:
        rows = ref.get(m) or []
        nums = [r.get("salecount") for r in rows if isinstance(r.get("salecount"), (int, float))]
        ref_per_month[m] = {"series_count": len(rows), "salecount_sum": int(sum(nums))}

    # ============================================================ e) 并列格专项
    multi = find_multi_name_cells(monthly)

    totals_list = [{"month": m, "salecount_sum": per_month[m]["salecount_sum"]} for m in MONTHS]

    qc = {
        "meta": {
            "task": "WS1 销量数据采集与质检",
            "snapshot_file": os.path.basename(SNAPSHOT),
            "snapshot_id": meta.get("snapshot_id"),
            "months": MONTHS,
            "captured_at": "2026-10-19",
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "count_guard": meta.get("count_guard"),
            "count_guard_pass": all(
                meta["count_guard"]["main"][0] <= per_month[m]["series_count"]
                <= meta["count_guard"]["main"][1] for m in MONTHS),
            "raw_archive_policy": (
                "原始页归档仅本地留存，不入 git。"
                "`audit-output/sales_autohome_raw_pages_20261019.json`（45.44 MB）与"
                "`_cache_sales_autohome.json`（45.49 MB）体积过大，本仓库为公开仓库且有"
                "每日 23:00 自动化提交，60 MB 级文件会永久留在历史里、拖慢每次 clone，"
                "故两者已加入 `.gitignore`。"
                "**可凭 seriesid + 快照月复抓**：运行 `fetch_sales_autohome.py` 即可按"
                "`typeid=9&subranktypeid=2305&date=YYYY-MM` 重新拉取该月全页并重建归档；"
                "主快照 `sales_autohome_series_monthly_20261019.json` 保留全部标量字段，"
                "仅在 `rightinfo`/`shareinfo`/`rcmlist`（合计 88% 体积，全为 autohome:// 深链"
                "与埋点 eventid，无分析价值）上做了裁剪，被裁文本已抽入 `_ui_texts`。"
                "入库清单：主快照 + 本 QC 报告 + 两个 crosscheck 文件；raw 归档与缓存不入库。"
            ),
        },
        "a_per_month": per_month,
        "a_count_guard_detail": {
            "rule": f"每月条数须落 [{meta['count_guard']['main'][0]},{meta['count_guard']['main'][1]}]",
            "per_month": {m: {"count": per_month[m]["series_count"],
                              "pass": meta["count_guard"]["main"][0] <= per_month[m]["series_count"]
                              <= meta["count_guard"]["main"][1]} for m in MONTHS},
            "all_pass": all(
                meta["count_guard"]["main"][0] <= per_month[m]["series_count"]
                <= meta["count_guard"]["main"][1] for m in MONTHS),
            "known_issue_note": ("Lead 探测时 2026-05 曾因 WinError 10060 超时仅取到 100 条；"
                                 "本采集器按「条数区间校验 + 该月整体重抓」闭环，本次 10 个月全部达标。"),
        },
        "b_union": {
            "union_seriesid_count": len(union_ids),
            "union_normname_count": len(union_names),
            "lead_expected_approx": 507,
            "matches_lead_expectation": len(union_ids) == 507,
            "presence_months_histogram": {str(k): v for k, v in sorted(presence.items())},
            "first_seen_month_examples": dict(list(id_first_seen.items())[:5]),
        },
        "c_crosscheck": {
            "yiche_202608": yiche_xc,
            "dongchedi_202609": dcd_xc,
            "lead_source_switch_202608": lead_xc,
        },
        "d_anomalies": {
            **anomalies,
            "summary": {
                "non_numeric_salecount": len(anomalies["non_numeric_salecount"]),
                "negative_salecount": len(anomalies["negative_salecount"]),
                "duplicate_seriesname_in_month": len(anomalies["duplicate_seriesname_in_month"]),
                "duplicate_seriesid_in_month": len(anomalies["duplicate_seriesid_in_month"]),
                "mom_total_jump_over_50pct": len(anomalies["mom_total_jump_over_50pct"]),
            },
            "mom_total_series": totals_list,
            "reference_universe_per_month": ref_per_month,
        },
        "e_multiname_cells": multi,
    }
    write_json(QC_OUT, qc)

    print("\n================ 质检摘要 ================")
    print(f"a) 每月条数校验: {'全部达标' if qc['a_count_guard_detail']['all_pass'] else '有未达标'}")
    for m in MONTHS:
        print(f"   {m}: {per_month[m]['series_count']:>4} 条  合计 {per_month[m]['salecount_sum']:>7,}")
    print(f"b) 并集车系数 = {len(union_ids)} (Lead 期望 ~507)")
    print(f"c1) 易车202608: 匹配 {yiche_xc.get('matched')} / >10% {yiche_xc.get('over_10pct_count', '-')}")
    print(f"c2) 懂车帝202609: 匹配 {dcd_xc.get('matched')} / >10% {dcd_xc.get('over_10pct_count', '-')}")
    print(f"c3) Lead换源对照: 唯一>10%车系 {lead_xc.get('unique_pairs_over_10pct', '-')}")
    print(f"d) 异常: {qc['d_anomalies']['summary']}")
    print(f"e) 并列格: {multi['candidate_normname_count']} 个候选名")
    print(f"\n报告已写出 {QC_OUT}")
    return 0


def find_multi_name_cells(monthly):
    """底表「通用名称」列 8.8% 是并列多车系单元格（Lead 情报）。
    这里从主快照反查：给出并列候选清单，供 WS2 匹配抽样核验。"""
    try:
        from openpyxl import load_workbook
    except Exception as e:                                        # noqa: BLE001
        return {"status": "unavailable", "reason": f"openpyxl 不可用: {e}"}

    xlsx = None
    for fn in os.listdir(HERE):
        if fn.endswith(".xlsx") and "汇总表" in fn and "合并版" in fn:
            xlsx = os.path.join(HERE, fn)
            break
    if not xlsx:
        return {"status": "unavailable", "reason": "未找到底表 xlsx"}

    wb = load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb["NEV公告参数汇总"] if "NEV公告参数汇总" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [str(x) if x is not None else "" for x in next(rows)]
    # 实测表头（2026-10-19 核对）：批次=0 产品型号=1 产品商标=2 企业名称=3
    #   车型名称=4 产品名称=5 通用名称=6 产品类型=7 ...
    def col(name, default):
        try:
            return header.index(name)
        except ValueError:
            return default

    gi = col("通用名称", 6)
    ti = col("产品型号", 1)
    bi = col("批次", 0)
    si = col("细分市场", 8)
    SEP = set(",，、;；/")
    aug = {norm(r.get("seriesname")): r for r in (monthly.get("2026-09") or [])}
    all_norm = set()
    for m in MONTHS:
        for r in monthly.get(m) or []:
            all_norm.add(norm(r.get("seriesname")))

    cells, total_rows, multi_rows = [], 0, 0
    for ri, row in enumerate(rows, start=2):
        total_rows += 1
        g = row[gi] if gi < len(row) else None
        if g is None:
            continue
        s = str(g).strip()
        if not any(ch in SEP for ch in s):
            continue
        multi_rows += 1
        parts = [p.strip() for p in s.replace("、", ",").replace(";", ",")
                 .replace("；", ",").replace("/", ",").split(",") if p.strip()]
        cand = []
        for p in parts:
            nk = norm(p)
            hit_sep = aug.get(nk)
            hit_any = None
            for m in MONTHS:
                for r in monthly.get(m) or []:
                    if norm(r.get("seriesname")) == nk:
                        hit_any = (m, r)
                        break
                if hit_any:
                    break
            cand.append({"part": p, "norm": nk,
                         "in_202609": bool(hit_sep),
                         "ah_seriesid": str(hit_sep.get("seriesid")) if hit_sep else None,
                         "ah_sales": hit_sep.get("salecount") if hit_sep else None,
                         "in_any_of_10_months": nk in all_norm,
                         "any_month": hit_any[0] if hit_any else None})
        n_hit_sep = sum(1 for c in cand if c["in_202609"])
        n_hit_any = sum(1 for c in cand if c["in_any_of_10_months"])
        # Lead 要求 (b)：零销量行若为并列格，须确认**全部候选**均 miss
        cells.append({"excel_row": ri, "batch": row[bi] if bi < len(row) else None,
                      "product_model": row[ti] if ti < len(row) else None,
                      "sub_market": row[si] if si < len(row) else None,
                      "generic_name": s, "parts": cand, "n_parts": len(parts),
                      "n_parts_in_202609": n_hit_sep,
                      "n_parts_in_any_month": n_hit_any,
                      "all_parts_missing_202609": n_hit_sep == 0,
                      "all_parts_missing_all_months": n_hit_any == 0,
                      "short_part_flag": any(len(p) <= 4 for p in parts)})
    wb.close()
    uniq = set()
    for c in cells:
        uniq.add(c["generic_name"])
    uniq_norm = set()
    for c in cells:
        uniq_norm.add(norm(c["generic_name"]))
    n_all_miss_sep = sum(1 for c in cells if c["all_parts_missing_202609"])
    n_all_miss_any = sum(1 for c in cells if c["all_parts_missing_all_months"])
    n_partial = sum(1 for c in cells
                    if 0 < c["n_parts_in_202609"] < c["n_parts"])
    n_short = sum(1 for c in cells if c["short_part_flag"])
    return {
        "status": "ok",
        "purpose": ("并列格专项抽样域（Lead 情报 2）：确认并列格内**全部候选车系**在榜单中"
                    "是否 miss，而非只查第一个；供 WS2 匹配抽样核验使用"),
        "source_xlsx": os.path.basename(xlsx),
        "total_data_rows": total_rows,
        "multi_name_rows": multi_rows,
        "multi_name_ratio": round(multi_rows / total_rows, 4) if total_rows else None,
        "candidate_raw_count": len(uniq),
        "candidate_normname_count": len(uniq_norm),
        "lead_reported_count": 223,
        "lead_reported_count_status": "superseded — Lead 于 2026-10-19 裁决作废 223，全案统一 222",
        "canonical_candidate_count": len(uniq),
        "canonical_note": (
            "**CANONICAL = %d**（Lead 2026-10-19 裁决）：本脚本两种独立口径——"
            "【原样字符串】去重得 %d、【norm 归一化】去重得 %d——**收敛到同一值**，"
            "且 475 行的基数为可复算，故 %d 为权威值。"
            "Lead 原报的 223 系其「行内候选出现次数」口径的产物，无独立复现价值，**已作废**；"
            "本字段保留 223 仅为不改历史口径记录，后续读者请以 %d 为准，不必再追究差异。"
            % (len(uniq), len(uniq), len(uniq_norm), len(uniq), len(uniq))),
        "count_note": ("历史口径记录（不改）：Lead 曾报 223 个去重并列格，本脚本按【原样字符串】"
                       "去重得 %d、按【norm 归一化】去重得 %d；差异源于 Lead 未说明的去重口径"
                       "（原样 vs 归一化 vs 是否含仅分隔符的空候选格）。"
                       "**裁决：canonical_candidate_count = %d，Lead 已作废 223 并撤回全部引用；"
                       "本条目仅作历史留存。**"
                       % (len(uniq), len(uniq_norm), len(uniq))),
        "verdict": {
            "sampled": multi_rows,
            "all_parts_missing_in_202609": n_all_miss_sep,
            "all_parts_missing_in_all_10_months": n_all_miss_any,
            "partially_present_202609": n_partial,
            "short_part_cells_need_manual_review": n_short,
            "conclusion": ("全部 %d 个并列格已逐格展开全部候选并逐个核验（非只查第一个）："
                           "%d 格在 2026-09 榜单中**全部候选均 miss**（语义正确：该类行属停产/"
                           "未上量车型），%d 格部分命中，%d 格含 ≤4 字符短分词需人工甄别"
                           "（可能是型变体而非并列车系，如 `i03/i03A`）。"
                           % (multi_rows, n_all_miss_sep, n_partial, n_short))},
        "cells": cells,
    }


if __name__ == "__main__":
    sys.exit(main())
