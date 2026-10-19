# -*- coding: utf-8 -*-
"""link_sales.py — WS2：车系锚定 + 底表销量写入 + 联动指标计算

口径锚点：Update/销量口径规范_20261019.md（v1.0，Lead 定稿）
  主源 = 汽车之家 · 行业综合销量（车系级月度零售量，新能源榜 typeid=9&subranktypeid=2305）
  快照 = 2025-12 + 2026-01 … 2026-09（10 个月）
  底表写入 = 2026-09 车系销量 → 列35 月销量(辆) / 列36 销量数据月份 / 列37 销量更新日期
              + 尾插 列39 车系ID / 列40 销量覆盖状态（§2.2，不得插中间）

铁律（§4）：月销量是**车系级**指标；同一车系跨批次/动力类型多行共享同值。
  任何汇总**必须以「车系」为统计单元**——本脚本先按 seriesid 建「车系记录」
  （数值型=median，分类型=mode），再对车系记录聚合销量；严禁对底表行直接 SUM。

数据源优先级（--source auto）：
  1) audit-output/sales_autohome_series_monthly_20261019.json   （WS1 正式产物）
  2) _cache_sales_autohome.json                                 （WS1 分页原始缓存，结构等价）
  3) _recon_ah_cache.json                                       （Lead 开发快照，字段不同）

用法：
  python link_sales.py --dry-run                 # 只出统计与 JSON，不写底表（**默认即只读**）
  python link_sales.py --dry-run --no-suggest    # 跳过汽车之家 suggest 兜底（快速）
  python link_sales.py --apply --no-suggest --expect-sha <sha256>
        # ← **验收基线口径（唯一认可写法）**：Lead m00878③ 裁决本轮不启用通道 D。
        #   漏写 --no-suggest 会静默启用通道 D，导致 matched_rows 2185→2190 等基线漂移
        #   （2026-10-19 实际发生过一次，靠 archive 备份还原 + 本命令重跑修正）。
        #   自检：若 apply 时 suggest_calls > 0，会打印 [WARN] 通道D已启用，与验收基线不符。
  python link_sales.py --show-sha                # 只打印工作簿 sha256（小写 hex）
  python link_sales.py --apply --expect-sha <sha256>
        # 写底表。四项护具（Lead m00729，并发写入方 cr17_011 同簿）：
        #   ① sha 守卫：--apply 必须带 --expect-sha，落盘前比对（小写 hex），不等即中止
        #   ② 落盘前整本备份 archive/v4.9.32_pre-sales_backup_<YYYYMMDD>.xlsx
        #   ③ 行数不硬编码：base_rows 实读；六类/四类分桶和 == 实读行数
        #   ④ 行身份复核：按 (批次, 产品型号) 逐行 zip 配对，任一不符即中止
        # 另：两个 JSON 均写 meta.dry_run 与 meta.base_table_sha256（本 JSON 依据的工作簿版本）
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import ssl
import sys
import time
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from statistics import median

import openpyxl
from openpyxl.styles import PatternFill

# --------------------------------------------------------------------------- 环境
for _k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy", "all_proxy"):
    os.environ.pop(_k, None)

HERE = os.path.dirname(os.path.abspath(__file__))
AUDIT = os.path.join(HERE, "audit-output")
WB_PATH = os.path.join(HERE, "NEV公告参数汇总表_合并版（341~410批）.xlsx")
SHEET = "NEV公告参数汇总"
LEDGER_SHEET = "变更记录"

OFFICIAL_JSON = os.path.join(AUDIT, "sales_autohome_series_monthly_20261019.json")
PAGE_CACHE = os.path.join(HERE, "_cache_sales_autohome.json")
DEV_CACHE = os.path.join(HERE, "_recon_ah_cache.json")

LEDGER_OUT = os.path.join(AUDIT, "sales_link_20261019.json")
ANALYTICS_OUT = os.path.join(AUDIT, "sales_linkage_analytics_20261019.json")
BACKUP_OUT = os.path.join(AUDIT, "sales_prior_yiche_202608_backup_20261019.json")
SUGGEST_CACHE = os.path.join(AUDIT, "sales_link_suggest_cache_20261019.json")
ANCHOR_OUT = os.path.join(AUDIT, "sales_series_anchor_candidates_20261019.json")

MONTHS = ["2025-12"] + ["2026-%02d" % m for m in range(1, 10)]
CUR_MONTH = "2026-09"
CAPTURED_AT = "2026-10-19"
SNAPSHOT = "20261019"
MONTH_MIN, MONTH_MAX = 330, 460          # 新能源榜月度条数区间硬校验（同 WS1 采集器）

# 列号（0-based）
C_BATCH, C_CODE, C_TRADEMARK, C_COMPANY, C_MODEL, C_PRODUCT = 0, 1, 2, 3, 4, 5
C_GENERIC, C_PTYPE, C_SEGMENT, C_POWER, C_TAX = 6, 7, 8, 9, 10
C_CURB, C_WHEEL, C_RANGE, C_CAP, C_BATT, C_DENS, C_ECONS = 11, 12, 13, 14, 15, 16, 17
C_FMOTOR, C_MOTOR_T, C_RMOTOR, C_MOTOR_SUP, C_MOTOR_MDL = 18, 19, 20, 21, 22
C_FUEL, C_BFUEL, C_DISP, C_EPOW, C_ESUP, C_EMDL = 23, 24, 25, 26, 27, 28
C_SRC, C_LENGTH, C_CELL_SUP, C_PACK_SUP, C_URL, C_MEDIA = 29, 30, 31, 32, 33, 34
C_SALES, C_SMONTH, C_SDATE, C_EREV = 35, 36, 37, 38
C_SERIESID, C_STATUS = 39, 40            # 新增尾插列

NCOLS_OLD = 39
NCOLS_NEW = 41

FILL_MEDIA = PatternFill("solid", fgColor="F8CBAD")
NO_FILL = PatternFill(fill_type=None)

SOURCE_TAG = "汽车之家·行业综合销量（车系级月度零售量，新能源榜 typeid=9&subranktypeid=2305）"
TRANSFER_TAG = "口径统一：易车202608 → 汽车之家202609"

STATUS_ONSALE = "在售-已匹配"
STATUS_ZERO = "在售-零销量"
STATUS_HOLD = "未匹配-命名待锚定"
STATUS_GONE = "未匹配-已停产或未上量"
STATUS_NOID = "无车系身份"
STATUS_NOMATCH = "未匹配-无车系ID"
STATUS_ENUM = [STATUS_ONSALE, STATUS_ZERO, STATUS_HOLD, STATUS_GONE, STATUS_NOID, STATUS_NOMATCH]

DEDUP_RULE = ("先按 seriesid 建车系记录（数值型取中位数 median、分类型取众数 mode），"
              "再对车系记录聚合销量；同车系多行不重复计数")
UNIT = "车系"

# --------------------------------------------------------------------------- 归一化


def norm(s):
    """§3.1 归一化：小写→去空格/连字符/下划线/·/./括号(半角全角)/；去「新能源」。"""
    s = str(s or "").strip().lower()
    s = re.sub(r"[\s\-_·．.（）()/]+", "", s)
    return s.replace("新能源", "")


_POW = r"(?:纯电动|插电式混合动力|插电式增程混合动力|增程式混合动力|增程式|混合动力|燃料电池)"
_BODY = r"(?:运动型乘用车|多用途乘用车|越野乘用车|轿车|乘用车|客车|货车)"
# §3.3 占位符/动力类型泄漏值：本身不是车系名，不参与匹配 → `无车系身份`
PLACEHOLDER_GENERIC_RE = re.compile(r"^(?:换电式)?" + _POW + r"$|^" + _POW + r"(?:[^,，、;；/]{0,10})" + _BODY + r"$")
PLACEHOLDER_CODE_RE = re.compile(r"^[A-Z0-9]{4,}[^,，、;；/]{0,10}?" + _POW + r"[^,，、;；/]{0,10}?" + _BODY + r"$")

SPLIT_RE = re.compile(r"[,，、;；/]+")


def is_placeholder(g):
    g = str(g or "").strip()
    if not g:
        return False
    return bool(PLACEHOLDER_GENERIC_RE.match(g) or PLACEHOLDER_CODE_RE.match(g))


def clean_num(v):
    """数值清洗（沿用 generate_dashboard.clean_num 语义 + 公告原生格式）：
    '不适用(BEV)'→None；'285/302'→285；'512(CLTC)'→512；'53(单箱26.5/共2箱)'→53。
    """
    if v is None:
        return None
    s = str(v).strip()
    if not s or "不适用" in s or s in ("-", "—", "/", "None", "nan", "无"):
        return None
    s = s.split("/")[0].strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def _as_int(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().replace(",", "").replace("辆", "")
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def brand_of(trademark):
    """产品商标 → 品牌词（去「牌/汽车/股份/…」）"""
    t = str(trademark or "").strip()
    if not t:
        return ""
    t = re.sub(r"(牌|汽车|股份|有限公司|集团|制造|工业)+$", "", t)
    t = re.sub(r"(牌|汽车|股份|有限公司|集团|制造|工业)", "", t)
    return t.strip()


def norm_power(v):
    """动力类型双形态归一（BEV/PHEV/EREV 与中文全称并存）。"""
    s = str(v or "").strip()
    if not s:
        return ""
    u = s.upper()
    if u in ("BEV", "PHEV", "EREV", "HEV", "FCV"):
        return u
    if "增程" in s:
        return "EREV"
    if "插电" in s or "混合动力" in s or "混动" in s:
        return "PHEV"
    if "纯电" in s:
        return "BEV"
    if "燃料电池" in s:
        return "FCV"
    return s


def row_energy(r):
    """行级动力口径（契约 meta、§2.3 第③级消歧用）：增程 → EREV；否则 动力类型 列。"""
    ev = r[C_EREV]
    if _as_int(ev) == 1 or str(ev or "").strip() in ("是", "Y", "y", "TRUE", "true"):
        return "EREV"
    e = norm_power(r[C_POWER])
    return e if e in ("BEV", "PHEV", "EREV", "FCV") else (e or "未知")


def series_energy_label(r):
    """行级动力标签（Lead m00806 R1，**车系级=行级标签众数**）：
    EREV 列取真值 → EREV；否则 动力类型 含 BEV → BEV；否则 → PHEV。
    """
    ev = r[C_EREV]
    if _as_int(ev) == 1 or str(ev or "").strip() in ("是", "Y", "y", "TRUE", "true"):
        return "EREV"
    return "BEV" if "BEV" in norm_power(r[C_POWER]) else "PHEV"


def parse_price_mid(p):
    """'26.35-36.90万' → 31.625（中值）；'6.48万' → 6.48；解析失败 → None。"""
    s = str(p or "").strip()
    if not s:
        return None
    s = s.replace("万", "").replace("元", "").replace(",", "")
    m = re.match(r"^\s*([\d.]+)\s*-\s*([\d.]+)\s*$", s)
    if m:
        try:
            return (float(m.group(1)) + float(m.group(2))) / 2.0
        except ValueError:
            return None
    m = re.match(r"^\s*([\d.]+)\s*$", s)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    return None


PRICE_BANDS = [(0, 10, "≤10万"), (10, 15, "10-15万"), (15, 20, "15-20万"),
               (20, 30, "20-30万"), (30, 40, "30-40万"), (40, 10 ** 9, "40万+")]
PRICE_BINS = [lab for _, _, lab in PRICE_BANDS]
# ★ 契约 v1.1 tech_sales.range_bins_def —— 消费方严禁自行猜边界，必须读 JSON 该字段
RANGE_BINS_DEF = ["<300km", "300-400km", "400-500km", "500-600km", "600-700km", "≥700km"]
RANGE_BINS = [(0, 300, RANGE_BINS_DEF[0]), (300, 400, RANGE_BINS_DEF[1]),
              (400, 500, RANGE_BINS_DEF[2]), (500, 600, RANGE_BINS_DEF[3]),
              (600, 700, RANGE_BINS_DEF[4]), (700, 10 ** 9, RANGE_BINS_DEF[5])]


def price_band(mid):
    if mid is None:
        return None
    for lo, hi, lab in PRICE_BANDS:
        if lo <= mid < hi:
            return lab
    return None


def range_bin(rng):
    if rng is None:
        return None
    for lo, hi, lab in RANGE_BINS:
        if lo <= rng < hi:
            return lab
    return PRICE_BANDS[0][2] if False else None


# --------------------------------------------------------------------------- 销量源


def _std_item(x):
    sid = x.get("seriesid", x.get("sid"))
    name = x.get("seriesname", x.get("name"))
    return {
        "sid": "" if sid is None else str(sid).strip(),
        "name": str(name or "").strip(),
        "brand": "" if x.get("brandid", x.get("brand")) is None else str(x.get("brandid", x.get("brand"))),
        "n": _as_int(x.get("salecount", x.get("n"))),
        "price": str(x.get("priceinfo") or x.get("price") or "").strip(),
        "et": _as_int(x.get("energytype", x.get("et"))),
        "rank": _as_int(x.get("rankNum")),
        "rankchange": _as_int(x.get("rankchange")),
        "specname": str(x.get("specname") or "").strip(),
    }


def load_sales(source="auto", log=print):
    """返回 (monthly, meta)。monthly = {month: [std_item, ...]}"""
    tried = []
    cands = []
    if source in ("auto", "official"):
        cands.append(("official", OFFICIAL_JSON, "ws1_official"))
    if source in ("auto", "pagecache"):
        cands.append(("pagecache", PAGE_CACHE, "ws1_pagecache"))
    if source in ("auto", "devcache"):
        cands.append(("devcache", DEV_CACHE, "dev_snapshot"))

    for kind, path, tag in cands:
        if not os.path.exists(path):
            tried.append(f"{kind}: missing")
            continue
        try:
            with open(path, encoding="utf-8") as f:
                d = json.load(f)
        except Exception as e:                                    # noqa: BLE001
            tried.append(f"{kind}: read-error {e}")
            continue
        monthly, meta = {}, {}
        if kind == "official" or (kind == "pagecache" and "main" in d):
            def _page_items(v):
                """分页缓存每月字典里除页码外还有一个 int 型键 "pc"（=pagecount），必须过滤。"""
                out = []
                for p in sorted((k for k in v if str(k).isdigit()), key=int):
                    out.extend((v[p] or {}).get("items") or [])
                return out

            raw = d.get("monthly") if kind == "official" else {
                k: _page_items(v) for k, v in (d.get("main") or {}).items()}
            for m in MONTHS:
                if m in raw:
                    monthly[m] = [_std_item(x) for x in raw[m]]
            meta = {"source_kind": kind, "source_path": path, "source_tag": tag,
                    "source_meta": d.get("meta") if kind == "official" else None,
                    "reference_monthly": (
                        ((d.get("reference_only") or {}).get("全车型榜") or {}).get("monthly")
                        if kind == "official" else None),
                    "has_rank_fields": any(it.get("rank") is not None for v in monthly.values() for it in v),
                    "tried": tried}
        elif kind == "devcache":
            for m, v in d.items():
                if m in MONTHS:
                    monthly[m] = [_std_item(x) for x in v]
            meta = {"source_kind": kind, "source_path": path, "source_tag": tag,
                    "source_meta": None, "has_rank_fields": False, "tried": tried}
        else:
            tried.append(f"{kind}: unrecognized structure")
            continue
        if not monthly:
            tried.append(f"{kind}: no months parsed")
            continue
        counts = {m: len(v) for m, v in monthly.items()}
        bad = [m for m, n in counts.items() if not (MONTH_MIN <= n <= MONTH_MAX)]
        meta["month_counts"] = counts
        meta["month_count_guard_violations"] = bad
        log(f"[sales-src] {kind} <- {os.path.basename(path)}  months={len(monthly)} "
            f"counts={counts}  guard_violations={bad}")
        if bad:
            log(f"[warn] 月度条数越界（应∈[{MONTH_MIN},{MONTH_MAX}]）：{bad} —— "
                f"开发快照可能不完整，正式数据应以 WS1 产物为准")
        return monthly, meta

    raise RuntimeError("无可用销量数据源；tried=" + "; ".join(tried))


# --------------------------------------------------------------------------- 匹配引擎


class Matcher:
    """§3 通道匹配：A 精确 → B 品牌+名称 → C 分词 → D suggest 兜底。多候选一律 HOLD。"""

    def __init__(self, monthly, use_suggest=True, max_suggest=0, log=print):
        self.monthly = monthly
        self.union = {}                        # sid -> best record (name/price)
        self.byname = defaultdict(set)         # norm(name) -> {sid}
        self.name_of = {}
        for m in MONTHS:
            for x in monthly.get(m, []):
                if not x["sid"]:
                    continue
                self.union.setdefault(x["sid"], x)
                self.byname[norm(x["name"])].add(x["sid"])
                self.name_of.setdefault(x["sid"], x["name"])
                if m == CUR_MONTH:
                    self.union[x["sid"]] = x
        self.keys = sorted(self.byname.keys(), key=len, reverse=True)
        self.cur = {x["sid"]: x for x in monthly.get(CUR_MONTH, [])}
        self.multi_sid_names = {k: sorted(v) for k, v in self.byname.items() if len(v) > 1}
        self.use_suggest = use_suggest
        self.max_suggest = max_suggest
        self.log = log
        self._sug_cache = {}
        self._neg_cache = {}
        self._sug_path = SUGGEST_CACHE
        if use_suggest and os.path.exists(self._sug_path):
            try:
                with open(self._sug_path, encoding="utf-8") as f:
                    self._sug_cache = json.load(f)
            except Exception:                                     # noqa: BLE001
                self._sug_cache = {}
        self._sug_dirty = False
        self._sug_calls = 0
        self._sug_fail = []
        self._ctx = ssl.create_default_context()
        self._ctx.check_hostname = False
        self._ctx.verify_mode = ssl.CERT_NONE
        self._last_req = 0.0

    # -------- 内部工具（Lead 验收的保守版：A 精确 → B 品牌+名称 → C 分词双向包含(min len>=4)）
    def _add(self, bucket, key):
        if key in self.byname:
            for sid in self.byname[key]:
                bucket.add(sid)

    def _brand_ok(self, bp, key):
        return bool(bp) and norm(bp) in key

    def channel_A(self, g):
        """A 精确：norm(通用名称) == norm(seriesname)"""
        c = set()
        self._add(c, norm(g))
        return c

    def channel_B(self, tokens, bp):
        """B 品牌+名称：品牌词须出现在车系名中（品牌守卫）+ 双向包含。"""
        c = set()
        nb = norm(bp)
        for t in tokens:
            # B0：通用名自带品牌前缀 → 去品牌后精确
            for pre in (nb,):
                if pre and t.startswith(pre) and len(t) > len(pre):
                    stripped = t[len(pre):]
                    self._add(c, stripped)
            # B1：norm(品牌+名称) 精确
            if nb:
                self._add(c, nb + t)
            # B2：norm(品牌+名称) 双向包含 + 品牌词在车系名中
            if nb:
                cand = nb + t
                for y in self.keys:
                    if len(y) < 3:
                        continue
                    if (y in cand or cand in y) and self._brand_ok(bp, y):
                        self._add(c, y)
        return c

    def channel_C(self, tokens):
        """C 分词：任一分词与车系名双向包含且 min(len)>=4；分词精确亦计入。"""
        c = set()
        for t in tokens:
            self._add(c, t)
            for y in self.keys:
                if min(len(y), len(t)) >= 4 and (y in t or t in y):
                    self._add(c, y)
        return c

    # -------- suggest
    def _suggest(self, q):
        if q in self._sug_cache:
            return self._sug_cache[q]
        if self.max_suggest and self._sug_calls >= self.max_suggest:
            return None
        url = ("https://sou.api.autohome.com.cn/sug/_suggest?plat=pc&uid=0&q="
               + urllib.parse.quote(q))
        out = []
        for attempt in range(3):
            dt = time.time() - self._last_req
            if dt < 0.4:
                time.sleep(0.4 - dt)
            self._last_req = time.time()
            self._sug_calls += 1
            try:
                req = urllib.request.Request(url, headers={
                    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                   "(KHTML, like Gecko) Chrome/126 Safari/537.36"),
                    "Accept": "application/json"})
                with urllib.request.urlopen(req, timeout=20, context=self._ctx) as r:
                    d = json.loads(r.read().decode("utf-8", "ignore"))
                for x in (d.get("result") or {}).get("data") or []:
                    out.append({"key": str(x.get("key") or ""), "wordid": x.get("wordid"),
                                "wordtype": x.get("wordtype"), "score": x.get("score") or 0})
                self._sug_cache[q] = out
                self._sug_dirty = True
                return out
            except Exception as e:                                # noqa: BLE001
                if attempt == 2:
                    self._sug_fail.append({"q": q, "error": f"{type(e).__name__}: {e}"})
                    self._sug_cache[q] = []
                    self._sug_dirty = True
                    return []
                time.sleep(1.5 * (attempt + 1))
        return out

    def save_suggest_cache(self):
        if self._sug_dirty:
            try:
                tmp = self._sug_path + ".tmp"
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(self._sug_cache, f, ensure_ascii=False)
                os.replace(tmp, self._sug_path)
            except Exception as e:                                # noqa: BLE001
                self.log(f"[warn] suggest 缓存写出失败: {e}")

    def channel_D(self, g, tokens, trademark, neg_tokens):
        """D suggest 兜底：seriesid 必须在当期(2026-09)榜内；严格名称关系 + 品牌守卫。"""
        if not self.use_suggest:
            return set(), []
        queries = []
        gt = norm(g)
        if gt:
            queries.append(gt)
        for t in tokens:
            if len(t) >= 2 and t not in queries:
                queries.append(t)
        c = set()
        detail = []
        for q in queries:
            res = self._suggest(q)
            if res is None:
                break
            for x in res:
                if x.get("wordtype") != 3:
                    continue
                wid = x.get("wordid")
                try:
                    sid = str(int(wid))
                except (TypeError, ValueError):
                    continue                                  # AGENTS：suggest 泛词 wordid 非数字必须拒
                if sid not in self.cur:
                    continue                                  # 必须在当期榜内
                key = norm(x["key"])
                if not key:
                    continue
                if sid in neg_tokens:
                    continue                                  # 品牌守卫：跨品牌 → 拒
                rel = min(len(key), len(gt)) >= 4 and (key in gt or gt in key)
                if not rel:
                    for t in tokens:
                        if min(len(key), len(t)) >= 4 and (key in t or t in key):
                            rel = True
                            break
                if not rel:
                    continue
                # 品牌守卫（行品牌词必须能与候选车系名互容；候选无品牌词时要求名称强重叠）
                if not self._brand_guard(trademark, x["key"]):
                    continue
                c.add(sid)
                detail.append({"q": q, "key": x["key"], "sid": sid, "score": x.get("score")})
        return c, detail

    def _brand_guard(self, trademark, series_key):
        """沿用 media_fill 品牌守卫：双方都有品牌词且互不包含、无别名 → 冲突（拒）。"""
        btoks = brand_tokens_of(trademark)
        stoks = brand_tokens_of(series_key)
        if not btoks or not stoks:
            return True
        for bt in btoks:
            for st in stoks:
                if bt == st or bt in st or st in bt:
                    return True
                if (bt, st) in BRAND_ALIAS or (st, bt) in BRAND_ALIAS:
                    return True
        return False

    def _neg_sids(self, trademark):
        """行品牌已知 → 该品牌不可能的车系 sid 集合（跨企业=必错）。"""
        ck = str(trademark or "")
        if ck in self._neg_cache:
            return self._neg_cache[ck]
        btoks = set(brand_tokens_of(trademark))
        if not btoks:
            self._neg_cache[ck] = set()
            return set()
        neg = set()
        sids_by_tok = defaultdict(set)
        for key, sids in self.byname.items():
            for st in brand_tokens_of(key):
                sids_by_tok[st] |= sids
        for tok, sids in sids_by_tok.items():
            if tok in btoks:
                continue
            aliased = any((tok, b) in BRAND_ALIAS or (b, tok) in BRAND_ALIAS for b in btoks)
            if not aliased:
                neg |= sids
        self._neg_cache[ck] = neg
        return neg

    # -------- 主入口
    def split_candidates(self, g):
        """并列格 → 候选名列表（Lead m00093：必须按分隔符切候选，不得拿整串比）。"""
        s = str(g or "").strip()
        if not s:
            return []
        return [p for p in (x.strip() for x in SPLIT_RE.split(s)) if p]

    def match_name(self, g, trademark, dtypes=None):
        """单候选名：A 精确 → B 品牌+名称 → C 分词 → D suggest 兜底。

        保守性依据（Lead m00239 撤回 ④、AGENTS 2026-09-17 泛化名撞名事故）：
        短名（S3/E5/07/熊猫Mini 等）不得跨品牌/跨代硬配 —— C 通道 min(len)>=4 是防线。
        dtypes 仅为签名兼容保留，本版不做动力类型消歧（命中 >=2 一律 HOLD）。
        """
        gt = norm(g)
        if not gt:
            return {"channel": "empty", "sids": set(), "cands": [], "note": "候选名为空"}
        if is_placeholder(g):
            return {"channel": "placeholder", "sids": set(), "cands": [],
                    "note": "占位符/动力类型泄漏值"}
        tokens = [gt]
        bp = brand_of(trademark)
        neg = self._neg_sids(trademark)

        a = self.channel_A(g)
        if a:
            return {"channel": "A_exact", "sids": a, "cands": sorted(self.name_of.get(s, s) for s in a)}
        b = self.channel_B(tokens, bp)
        if b:
            return {"channel": "B_brand", "sids": b, "cands": sorted(self.name_of.get(s, s) for s in b)}
        c = self.channel_C(tokens)
        if c:
            return {"channel": "C_substr", "sids": c, "cands": sorted(self.name_of.get(s, s) for s in c)}
        d, ddetail = self.channel_D(g, tokens, trademark, neg)
        if d:
            return {"channel": "D_suggest", "sids": d,
                    "cands": sorted(self.name_of.get(s, s) for s in d), "suggest": ddetail}
        if ddetail:
            return {"channel": "D_rejected", "sids": set(), "cands": [],
                    "note": "suggest 命中但被守卫拒绝", "suggest": ddetail}
        return {"channel": "unmatched", "sids": set(), "cands": [], "note": "10 个月并集内无匹配"}

    def match_cell(self, g, trademark, dtypes=None):
        """整格匹配：并列格先切候选再逐候选匹配。

        通道计数（Lead 契约 v1.1）：
          A_exact / B_brand / C_substr / D_suggest —— 单候选格直接命中
          A2_split            —— 并列格拆分后唯一命中
          multi_series_cell   —— 并列格拆分后命中 ≥2 个 seriesid（→ HOLD）
        """
        cands = self.split_candidates(g)
        if not cands:
            return {"channel": "empty", "sids": set(), "cands": [], "per_cand": [],
                    "n_cands": 0, "note": "通用名称为空"}
        per = []
        all_sids = set()
        for c in cands:
            r = self.match_name(c, trademark, dtypes)
            per.append({"name": c, "cand": c, "channel": r["channel"], "sids": sorted(r["sids"]),
                        "cands": r.get("cands", [])[:8], "note": r.get("note", ""),
                        "suggest": r.get("suggest")})
            all_sids |= set(r["sids"])
        ph = [p for p in per if p["channel"] in ("placeholder", "empty")]
        if len(ph) == len(per):
            return {"channel": "placeholder", "sids": set(), "cands": [], "per_cand": per,
                    "n_cands": len(cands), "note": "整格为占位符/动力类型泄漏值"}
        hit = [p for p in per if p["sids"]]
        base = hit[0]["channel"] if hit else "unmatched"
        hit_detail = [{"cand": p["name"], "sid": s, "channel": p["channel"]}
                      for p in hit for s in p["sids"]]
        if len(cands) > 1 and len(all_sids) > 1:
            ch = "multi_series_cell"
        elif len(cands) > 1 and len(all_sids) == 1:
            ch = "A2_split"
        elif len(cands) == 1 and len(all_sids) > 1:
            ch = "hold"
        else:
            ch = base
        return {"channel": ch, "base_channel": base, "sids": all_sids,
                "cands": sorted(self.name_of.get(s, s) for s in all_sids),
                "per_cand": per, "n_cands": len(cands), "hit_detail": hit_detail,
                "note": "" if all_sids else (hit[0]["note"] if hit else "10 个月并集内无匹配")}


# 品牌词表 + 同司别名（沿用 media_fill.py 2026-09-17 品牌审计教训）
BRAND_TOKENS = ['比亚迪', '特斯拉', '北汽', '北京汽车制造厂', '广汽', '埃安', 'aion', '长安', '深蓝',
                '阿维塔', '启源', '奇瑞', '星途', 'iCAR', '吉利', '银河', '极氪', 'zeekr', '几何',
                '领克', '长城', '欧拉', '魏牌', '坦克', '红旗', '一汽', '奔腾', '上汽', '荣威',
                '名爵', '智己', '东风', '岚图', '猛士', '纳米', '奕派', '风神', '风行', '极狐',
                '蔚来', '乐道', '萤火虫', '小鹏', '理想', '零跑', '哪吒', '腾势', '方程豹', '仰望',
                '问界', '智界', '享界', '尊界', '小米', '五菱', '宝骏', '大众', '捷达', '奥迪',
                '丰田', '本田', '日产', '别克', '凯迪拉克', '雪佛兰', '奔驰', '宝马', '沃尔沃',
                '起亚', '现代', '福特', '大通', '依维柯', '赛力斯', '蓝电', '开瑞', '凯翼',
                '江淮', '江铃', '海马', '神龙', '示界', '摩登', '金杯', '华晨', '威马',
                '合众', '大运', '远航', '高合', 'hiphi', '极石', '雷丁', '恒驰', '爱驰', '天际',
                '云度', '创维', 'skyworth', '岚图', 'smart', '极越', '启辰', '思皓', '钇为',
                '睿蓝', '金琥', '远航', '知豆', '凌宝', '朋克', '吉麦', '野马', '汉腾', '斯威',
                '东南', '陆风', '猎豹', '中兴', '福田', '御捷', '丽驰', '雷驰', '橙仕']
BRAND_ALIAS = {('北京汽车制造厂', '北汽'), ('北汽', '北京汽车制造厂'), ('广汽', '埃安'), ('埃安', 'aion'),
               ('广汽', 'aion'), ('吉利', '极氪'), ('极氪', 'zeekr'), ('吉利', '银河'), ('吉利', '几何'),
               ('吉利', '领克'), ('吉利', '睿蓝'), ('长城', '欧拉'), ('长城', '魏牌'), ('长城', '坦克'),
               ('一汽', '奔腾'), ('一汽', '红旗'), ('上汽', '荣威'), ('上汽', '名爵'), ('上汽', '智己'),
               ('上汽', '大通'), ('上汽', '五菱'), ('五菱', '宝骏'), ('东风', '岚图'), ('东风', '猛士'),
               ('东风', '纳米'), ('东风', '奕派'), ('东风', '风神'), ('东风', '风行'), ('东风', '蓝电'),
               ('东风', '启辰'), ('东风', '本田'), ('东风', '日产'), ('东风', '示界'), ('奇瑞', '星途'),
               ('奇瑞', '智界'), ('奇瑞', '享界'), ('奇瑞', '凯翼'), ('奇瑞', '开瑞'), ('比亚迪', '腾势'),
               ('比亚迪', '方程豹'), ('比亚迪', '仰望'), ('赛力斯', '问界'), ('赛力斯', '蓝电'),
               ('江淮', '尊界'), ('江淮', '思皓'), ('江淮', '钇为'), ('北汽', '极狐'), ('长安', '深蓝'),
               ('长安', '阿维塔'), ('合众', '哪吒'), ('大运', '远航'), ('神龙', '示界'), ('吉利', 'smart'),
               ('吉利', '极越'), ('上汽', '金琥'), ('东风', '御捷')}


def brand_tokens_of(s):
    n = str(s or "").lower().replace(" ", "")
    return [b for b in BRAND_TOKENS if b in n]


# --------------------------------------------------------------------------- 底表读取


def read_base_table(path, read_only=True):
    wb = openpyxl.load_workbook(path, read_only=read_only, data_only=True)
    ws = wb[SHEET]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    header = [c.value for c in wb[SHEET][1]] if not read_only else None
    if read_only:
        it = ws.iter_rows(min_row=1, max_row=1, values_only=True)
        header = list(next(iter(it)))
    return wb, rows, header


def read_ledger_format(wb, read_only=True):
    ws = wb[LEDGER_SHEET]
    if read_only:
        it = ws.iter_rows(min_row=1, max_row=1, values_only=True)
        header = list(next(iter(it)))
        nrow = ws.max_row
    else:
        header = [c.value for c in ws[1]]
        nrow = ws.max_row
    return header, nrow


# --------------------------------------------------------------------------- 车系记录


NUM_FIELDS = [("range_km", C_RANGE), ("capacity_kwh", C_CAP), ("density_wh_kg", C_DENS),
              ("econs_kwh_100km", C_ECONS), ("power_kw", C_MOTOR_T), ("curb_kg", C_CURB),
              ("wheelbase_mm", C_WHEEL), ("length_mm", C_LENGTH), ("fuel_l_100km", C_FUEL),
              ("bfuel_l_100km", C_BFUEL), ("disp_ml", C_DISP), ("engine_kw", C_EPOW)]
CAT_FIELDS = [("segment", C_SEGMENT), ("power_type", C_POWER), ("erev", C_EREV),
              ("company", C_COMPANY), ("trademark", C_TRADEMARK), ("ptype", C_PTYPE),
              ("batt_type", C_BATT), ("tax_free", C_TAX)]


def _mode(vals):
    vals = [v for v in vals if v not in (None, "")]
    if not vals:
        return None
    c = Counter(str(v).strip() for v in vals)
    top = max(c.values())
    return sorted(k for k, v in c.items() if v == top)[0]     # 确定性 tie-break


def build_series_records(rows, gen_anchor):
    """按 seriesid 建车系记录（数值 median / 分类 mode），再参与聚合。"""
    per = defaultdict(lambda: {"rows": [], "nums": defaultdict(list), "cats": defaultdict(list)})
    for idx, r in enumerate(rows):
        a = gen_anchor.get(str(r[C_GENERIC] or "").strip())
        if not a or not a.get("sid"):
            continue
        sid = a["sid"]
        per[sid]["rows"].append(idx)
        for f, ci in NUM_FIELDS:
            v = clean_num(r[ci] if ci < len(r) else None)
            if v is not None:
                per[sid]["nums"][f].append(v)
        for f, ci in CAT_FIELDS:
            v = r[ci] if ci < len(r) else None
            if v not in (None, ""):
                per[sid]["cats"][f].append(v)
        # R1 行级动力标签（Lead m00806）：车系级 = 行级标签众数（并列取字典序最小）
        per[sid]["cats"]["energy_label"].append(series_energy_label(r))
    recs = {}
    for sid, d in per.items():
        rec = {"seriesid": sid, "n_rows": len(d["rows"]),
               "batch_min": _as_int(min((rows[i][C_BATCH] for i in d["rows"]), key=lambda x: _as_int(x) or 0)),
               "batch_max": _as_int(max((rows[i][C_BATCH] for i in d["rows"]), key=lambda x: _as_int(x) or 0))}
        for f, _ in NUM_FIELDS:
            vs = d["nums"].get(f) or []
            rec[f] = round(median(vs), 3) if vs else None
            rec[f + "_n"] = len(vs)
        for f, _ in CAT_FIELDS:
            rec[f] = _mode(d["cats"].get(f) or [])
        el = d["cats"].get("energy_label") or []
        rec["energy_label"] = _mode(el)
        rec["energy_label_counts"] = dict(Counter(el))
        rec["erev_flag_rows"] = sum(1 for i in d["rows"]
                                    if _as_int(rows[i][C_EREV]) == 1)
        rec["power_type"] = norm_power(rec.get("power_type"))
        rec["generic_names"] = sorted({str(rows[i][C_GENERIC] or "").strip() for i in d["rows"]})
        rec["channels"] = sorted({gen_anchor[str(rows[i][C_GENERIC] or "").strip()]["channel"]
                                  for i in d["rows"]})
        recs[sid] = rec
    return recs


# --------------------------------------------------------------------------- 联动指标


def _share(x, tot):
    return round(100.0 * x / tot, 4) if tot else 0.0


def _group_sum(records, keyfn, sales_of):
    g = defaultdict(lambda: {"sales": 0, "series": []})
    for rec in records:
        k = keyfn(rec)
        g[k]["sales"] += sales_of(rec)
        g[k]["series"].append(rec["seriesid"])
    return g


def _cr_hhi(items):
    """items = [(key, sales)] 降序；返回 CR3/5/10 + HHI（0-10000）。"""
    vals = sorted((s for _, s in items if s), reverse=True)
    tot = sum(vals)
    out = {"total": tot, "n": len(vals),
           "CR3": _share(sum(vals[:3]), tot), "CR5": _share(sum(vals[:5]), tot),
           "CR10": _share(sum(vals[:10]), tot),
           "HHI": round(sum((100.0 * v / tot) ** 2 for v in vals), 2) if tot else 0.0}
    return out


def _wavg(records, field, sales_of):
    num = den = 0.0
    n_used = 0
    for r in records:
        v = r.get(field)
        if v is None:
            continue
        s = sales_of(r)
        if not s:
            continue
        num += v * s
        den += s
        n_used += 1
    return (round(num / den, 3) if den else None), n_used


def _mean(recs, field):
    vs = [r[field] for r in recs if r.get(field) is not None]
    return (round(sum(vs) / len(vs), 3) if vs else None), len(vs)


METRIC_KEYS = [("续航", "range_km"), ("容量", "capacity_kwh"), ("电耗", "econs_kwh_100km"),
               ("整备", "curb_kg"), ("功率", "power_kw"), ("能量密度", "density_wh_kg")]

RANKCHANGE_BINS = [("≤-50", -10 ** 9, -50), ("-49~-10", -49, -10), ("-9~-1", -9, -1), ("0", 0, 0),
                   ("+1~+9", 1, 9), ("+10~+49", 10, 49), ("≥+50", 50, 10 ** 9)]

ENERGY_ORDER = ["BEV", "PHEV", "EREV"]


def _ref_universe(src_meta):
    """参照宇宙（全车型总榜）逐月 {month: {series, sales}}（Lead m01157 #2 槽位）。"""
    ref = (src_meta or {}).get("reference_monthly") or {}
    out = {}
    for m in MONTHS:
        lst = ref.get(m) or []
        if not lst:
            continue
        out[m] = {"series": len(lst),
                  "sales": sum(int(x.get("salecount") or 0) for x in lst)}
    return out


def build_analytics(rows, gen_anchor, series_recs, monthly, src_meta, match_stats,
                    pending, status_counts, batch_year, cross_check, sid_by_boardname, log=print):
    """★ 严格按《销量联动JSON契约_20261019.md》v1.1 产出（生产方可加字段，不得改名/改结构/改单位）。

    统计单元一律 = 车系（seriesid）：先按 seriesid 建车系记录（数值 median / 分类 mode），
    再对车系记录聚合销量；**严禁**对底表行直接 SUM。
    """
    cur = {x["sid"]: x for x in monthly.get(CUR_MONTH, [])}
    prev_month = MONTHS[-2]

    def sales_of(rec):
        c = cur.get(rec["seriesid"])
        return int(c["n"] or 0) if c else 0

    def sales_in(m, sid):
        for x in monthly.get(m, []):
            if x["sid"] == sid:
                return int(x["n"] or 0)
        return 0

    linked = [r for r in series_recs.values() if r["seriesid"] in cur]
    linked_onsale = [r for r in linked if sales_of(r) > 0]
    board_total = sum(int(x["n"] or 0) for x in cur.values())
    linked_total = sum(sales_of(r) for r in linked)
    naive_total = 0
    for r in rows:
        a = gen_anchor.get(str(r[C_GENERIC] or "").strip()) or {}
        sid = a.get("sid")
        if sid in cur:
            naive_total += int(cur[sid]["n"] or 0)

    def energy_of(rec):
        """车系级动力口径（Lead m00806 R1）：行级标签众数（并列取字典序最小）。"""
        return rec.get("energy_label") or rec.get("power_type") or "(未定)"

    idcheck = []                       # 去重结构断言：各分桶 seriesid 集合两两不交且并集==linked

    def grp(recs, keyfn, keyname):
        acc = defaultdict(lambda: {"sales": 0, "ids": set()})
        for r in recs:
            k = keyfn(r)
            acc[k]["sales"] += sales_of(r)
            acc[k]["ids"].add(r["seriesid"])
        idcheck.append({"by": keyname, "buckets": {k: sorted(v["ids"]) for k, v in acc.items()}})
        return acc

    def desc_pairs(acc):
        return [[k, v["sales"]] for k, v in sorted(acc.items(), key=lambda kv: (-kv[1]["sales"], kv[0]))]

    def desc_pct(acc):
        tot = sum(v["sales"] for v in acc.values()) or 1
        return [[k, _share(v["sales"], tot)] for k, v in
                sorted(acc.items(), key=lambda kv: (-kv[1]["sales"], kv[0]))]

    # ================= 1 市场结构 =================
    seg = grp(linked, lambda r: (r.get("segment") or "").strip() or "(未定)", "细分市场")
    eng = grp(linked, energy_of, "动力类型")
    oem = grp(linked, lambda r: (r.get("company") or "").strip() or "(未定)", "企业")
    brd = grp(linked, lambda r: brand_of(r.get("trademark")) or "(未定)", "品牌")
    pri = grp(linked, lambda r: price_band(parse_price_mid(cur.get(r["seriesid"], {}).get("price")))
              or "(价格缺失)", "价格带")
    board_price = defaultdict(int)
    for x in cur.values():
        board_price[price_band(parse_price_mid(x.get("price"))) or "(价格缺失)"] += int(x["n"] or 0)

    eng_sales = [[k, eng[k]["sales"]] for k in ENERGY_ORDER if k in eng]
    eng_sales += [[k, v["sales"]] for k, v in eng.items() if k not in ENERGY_ORDER]

    # ---- 未与底表建立关联的当期在榜车系（Lead m00237③ / m00239）
    def _forms_of(g):
        out = []
        for x in [g] + [p.strip() for p in SPLIT_RE.split(str(g or "")) if p.strip()]:
            f = norm(x)
            if f and f not in out:
                out.append(f)
        return out

    linked_sids = {r["seriesid"] for r in linked}
    unlinked = sorted([x for x in cur.values() if x["sid"] not in linked_sids],
                      key=lambda x: (-int(x["n"] or 0), str(x["name"])))
    unlinked_sales = sum(int(x["n"] or 0) for x in unlinked)
    # 底表侧名称痕迹池 = **全部**通用名称形态（含已锚定到别的车系者；占位符/空除外）。
    # 依据 Lead m00239 canonical 分桶（88 车系 / 223,077 辆 / 75.78%）：痕迹探测须用全底表名称，
    # 例如 `海豹05 DM-i` 的痕迹名 `海豹` 自身已锚定到另一在榜车系；只用未锚定池会少算 8 个车系。
    trace_forms = []
    for g, v in gen_anchor.items():
        if v.get("channel") in ("placeholder", "empty"):
            continue
        anchored = v.get("sid") is not None
        for f in _forms_of(g):
            if len(f) >= 2:
                trace_forms.append((f, g, anchored))
    base_brand_tokens = set()
    for r in rows:
        base_brand_tokens |= set(brand_tokens_of(r[C_TRADEMARK])) | set(brand_tokens_of(r[C_COMPANY]))
    unlinked_traced, unlinked_clean = [], []
    for x in unlinked:
        bn = norm(x["name"])
        hs = [(g, an) for f, g, an in trace_forms if f and (f in bn or bn in f)]
        hits = sorted({g for g, _ in hs})
        hits_unanchored = sorted({g for g, an in hs if not an})
        (unlinked_traced if hits else unlinked_clean).append((x, hits, hits_unanchored))
    _tr_sales = sum(int(x["n"] or 0) for x, _, _ in unlinked_traced)
    _cl_sales = sum(int(x["n"] or 0) for x, _, _ in unlinked_clean)
    _clean_with_brand = [(x, h, hu) for x, h, hu in unlinked_clean
                         if set(brand_tokens_of(x["name"])) & base_brand_tokens]
    cause_bucket = {
        # Lead m00237③/m00239 canonical 两桶
        "命名未锚定": {"series": len(unlinked_traced), "sales": _tr_sales,
                       "pct": _share(_tr_sales, unlinked_sales),
                       "note": ("底表存在短名/异形名痕迹（如 Q05←长安启源Q05、S3←极狐贝塔S3、"
                                "07←极氪007GT），放宽守卫会跨品牌/跨代误锚 → 须逐条人工裁决")},
        "公告库无痕迹": {"series": len(unlinked_clean), "sales": _cl_sales,
                         "pct": _share(_cl_sales, unlinked_sales),
                         "note": ("底表无任何名称痕迹；Lead 已核验批次时间表"
                                  "（409/410 批=2026-08）排除公告时间边界成因")},
        # 生产方新增诊断细分（不改变上面两桶）
        "无痕迹-品牌在底表": {"series": len(_clean_with_brand),
                              "sales": sum(int(x["n"] or 0) for x, _, _ in _clean_with_brand)},
        "无痕迹-品牌不在底表": {"series": len(unlinked_clean) - len(_clean_with_brand),
                                "sales": _cl_sales - sum(int(x["n"] or 0) for x, _, _ in _clean_with_brand)},
        "trace_pool_rule": ("痕迹池 = 全部底表通用名称形态（norm 后 len>=2，整串+分隔符切分；"
                            "占位符/空除外），含已锚定到别的车系者"),
        "unit": "车系",
        "unlinked_series": len(unlinked),
        "unlinked_sales": unlinked_sales,
        "pending_lead_confirm": True,
        "note": ("link-engine 独立计算，与 Lead m00239 的 88 车系/223,077 辆/75.78% 与 "
                 "41 车系/71,286 辆/24.22% 互核"),
    }

    # ---- 混合动力家族三口径（Lead m00878 裁决 2：三者不得合并，各自定义）
    def _mix_payload(recs):
        _lst = [[self_name(cur, _r["seriesid"]), sales_of(_r), _share(sales_of(_r), linked_total),
                 dict(sorted((_r.get("energy_label_counts") or {}).items()))]
                for _r in sorted(recs, key=lambda z: (-sales_of(z), z["seriesid"]))]
        _s = sum(x[1] for x in _lst)
        _items = [{"车系": x[0], "销量": x[1], "占比": x[2], "行级标签计数": x[3]} for x in _lst]
        return {"series": len(_lst), "sales": _s, "pct_of_linked": _share(_s, linked_total),
                "items": _items, "list": _lst}

    mixed_energy_families = _mix_payload(
        [r for r in linked if len(r.get("energy_label_counts") or {}) >= 2])
    mixed_with_erev = _mix_payload(
        [r for r in linked if len(r.get("energy_label_counts") or {}) >= 2
         and (r.get("energy_label_counts") or {}).get("EREV")])
    r2_flip_impact = _mix_payload(
        [r for r in linked if (r.get("energy_label_counts") or {}).get("EREV")
         and r.get("energy_label") != "EREV"])

    structure = {
        "by_segment": desc_pairs(seg),
        "by_segment_pct": desc_pct(seg),
        "by_energy": eng_sales,
        "by_price": [[b, pri[b]["sales"]] for b in PRICE_BINS],
        "by_price_pct": [[b, _share(pri[b]["sales"], linked_total)] for b in PRICE_BINS],
        "by_price_coverage": {
            "binned": sum(pri[b]["sales"] for b in PRICE_BINS),
            "total": linked_total,
            "pct": _share(sum(pri[b]["sales"] for b in PRICE_BINS), linked_total),
            "uncovered_sales": linked_total - sum(pri[b]["sales"] for b in PRICE_BINS),
            "uncovered_series": len(pri["(价格缺失)"]["ids"]) if "(价格缺失)" in pri else 0,
            "note": ("价格带份额分母统一 = 当月已锚定销量总量；未覆盖 806 辆（0.101%）"
                     "为无有效指导价车系，未计入任何价格带（Lead m00990 F3）"),
        },
        "by_oem": desc_pairs(oem),
        "by_brand": desc_pairs(brd),
        "sales_total_dedup": linked_total,
        "sales_total_rowlevel_naive": naive_total,
        # ---- 生产方新增（不改变以上契约字段） ----
        "unit": UNIT,
        "dedup_rule": DEDUP_RULE,
        "universe": "已锚定车系（seriesid ∈ 2026-09 新能源榜，且通用名称可锚定到底表）",
        "series_count": len(linked),
        "series_count_onsale": len(linked_onsale),
        "board_sales_total": board_total,
        "board_series_count": len(cur),
        "linked_share_of_board_pct": _share(linked_total, board_total),
        "duplication_ratio": (round(naive_total / linked_total, 4) if linked_total else 0.0),
        "by_price_board": [[b, board_price.get(b, 0)] for b in PRICE_BINS],
        "by_price_board_pct": [[b, _share(board_price.get(b, 0), board_total)] for b in PRICE_BINS],
        "by_energy_board": None,
        # ---- 聚合口径（Lead m00806 裁决，报告/演示稿照抄） ----
        "agg_rule": ("分类型字段：众数（mode），并列取字典序最小（lexicographic-min）；"
                     "数值型字段：中位数（median，剔除空值后计算）；"
                     "空白值归 `(未定)` 而非 null；统计单元 = 车系（seriesid 去重）"),
        "oem_tiebreak": "lexicographic-min",
        "energy_rule": ("全库动力三分一律 **R1**：行级标签 = EREV（`是否增程(EREV)` 列取真值）"
                        "否则 BEV（`动力类型` 含 BEV）否则 PHEV；车系级 = 行级标签众数"
                        "（并列取字典序最小）。`by_energy` 与 `trend.monthly` **同源同规则**；"
                        "旧 R2（任一行为增程→整系 EREV）已废弃，仅留 `trend.monthly_prev_r2` 审计对照"
                        "（Lead m00806/m00990 F1）"),
        "share_denominator_rule": ("全篇份额（动力/品牌/企业/价格带/细分市场）分母统一 = **当月已锚定销量**"
                                   "（structure 用 2026-09 的 linked_total=797,701）；"
                                   "价格带未覆盖部分为无有效指导价车系，不计入任何价格带（Lead m00990 F3）"),
        "segment_rule": ("车系级 `细分市场` = 该车系行值的众数（值先 strip，去空白变体）；"
                         "行值全为空 → 该车系 segment=None → 归入 `(未定)` 桶；"
                         "`(未定)` 是显式桶，计入 by_segment 的 n"),
        "segment_empty_rule": "空值归入 `(未定)` 桶（不计入任何具体细分市场）；by_segment 的 n 含 `(未定)`",
        "tech_agg_rule": ("数值型字段先按行取中位数（剔除空值）得到车系值；加权均值分母 = "
                          "该字段非空的已锚定车系数（sample_n.weighted）；等权均值分母同"),
        "mixed_energy_families": mixed_energy_families,
        "mixed_with_erev": mixed_with_erev,
        "r2_flip_impact": r2_flip_impact,
        "mixed_family_defs": {
            "mixed_energy_families": "同一车系含 ≥2 种行级动力标签（含 BEV+PHEV 混搭）",
            "mixed_with_erev": "其中含 EREV 行者",
            "r2_flip_impact": "R2（任一行为增程→整系 EREV）相对 R1（行级众数）会实际翻转的车系 —— "
                              "被否决方案的成本证据（Lead m00806/m00878）",
        },
        "mixed_family_fields": ["车系名", "销量", "占已锚定比例", "行级标签计数"],
        "unlinked_board": {
            "series_count": len(unlinked),
            "sales": unlinked_sales,
            "sales_pct": _share(unlinked_sales, board_total),
            "board_series": len(cur),
            "board_sales": board_total,
            "top": [[str(x["name"]), int(x["n"] or 0), x.get("price") or ""] for x in unlinked[:20]],
            "top_fields": ["seriesname", "salecount", "priceinfo"],
            "cause_bucket": cause_bucket,
            "traced_series": [[str(x["name"]), int(x["n"] or 0), x.get("price") or "", hits[:6],
                               hu[:6]]
                              for x, hits, hu in unlinked_traced[:60]],
            "traced_series_fields": ["seriesname", "salecount", "priceinfo", "底表痕迹名称",
                                     "其中未锚定的痕迹名"],
            "traced_series_total": len(unlinked_traced),
            "clean_series_total": len(unlinked_clean),
            "note": ("2026-09 新能源榜 411 车系中未与底表建立关联者；"
                     "cause_bucket 为初判，待 Lead 核验（pending_lead_confirm）。"),
        },
        "missing": {
            "price": {"sales": pri["(价格缺失)"]["sales"] if "(价格缺失)" in pri else 0,
                      "series_count": len(pri["(价格缺失)"]["ids"]) if "(价格缺失)" in pri else 0},
            "segment": {"sales": seg["(未定)"]["sales"] if "(未定)" in seg else 0,
                        "series_count": len(seg["(未定)"]["ids"]) if "(未定)" in seg else 0},
        },
        "note": ("by_segment/by_energy/by_oem/by_brand/by_price 均为已锚定车系口径，合计 == "
                 "sales_total_dedup；by_price_board 为 2026-09 全榜（含未锚定车系）价格结构。"
                 "by_energy 由底表「动力类型」+「是否增程(EREV)」推导（接口 energytype 全为 1，无区分力）。"
                 "企业/品牌均取底表原值，不做集团合并。"),
    }

    # ================= 2 集中度 =================
    series_items = [[self_name(cur, s), sales_of(r)] for s, r in
                    ((r["seriesid"], r) for r in linked)]
    series_items = [[k, v] for k, v in series_items]
    brand_acc = defaultdict(int)
    for r in linked:
        brand_acc[brand_of(r.get("trademark")) or "(未定)"] += sales_of(r)
    oem_acc = defaultdict(int)
    for r in linked:
        oem_acc[(r.get("company") or "").strip() or "(未定)"] += sales_of(r)

    def top_n(pairs, n=20):
        return sorted(pairs, key=lambda kv: (-kv[1], kv[0]))[:n]

    def with_pct(pairs, tot):
        return [[k, v, _share(v, tot)] for k, v in pairs]

    ser_sorted = top_n(series_items, 10 ** 9)
    ser_tot = sum(v for _, v in ser_sorted) or 1
    top_series = [[k, v, _share(v, ser_tot),
                   "%s/%s" % (oem_of(series_recs, cur, k), brd_of(series_recs, cur, k))]
                  for k, v in top_n(series_items, 20)]
    concentration = {
        "series": _cr_hhi(series_items),
        "brand": _cr_hhi(list(brand_acc.items())),
        "oem": _cr_hhi(list(oem_acc.items())),
        "top_series": top_series,
        "top_brand": with_pct(top_n(list(brand_acc.items()), 20), linked_total),
        "top_oem": with_pct(top_n(list(oem_acc.items()), 20), linked_total),
        "unit": UNIT,
        "universe": "已锚定车系（seriesid ∈ 2026-09 新能源榜且可锚定到底表）",
        "series_board": _cr_hhi([[x["name"], int(x["n"] or 0)] for x in cur.values()]),
        "top_series_board": [[x["name"], int(x["n"] or 0)] for x in
                             sorted(cur.values(), key=lambda y: -int(y["n"] or 0))[:20]],
        "note": ("CR3/CR5/CR10 单位=百分数数值（如 42.1 表示 42.1%）；HHI 为 0-10000 点值。"
                 "series_board 为 2026-09 全榜 411 车系口径（未锚定车系无企业/品牌归属，故契约 "
                 "top_series 用已锚定口径以保证第 4 位企业/品牌可溯）。"),
    }

    # ================= 3 月度趋势 =================
    # F1（Lead m00990）：trend.monthly 动力拆分统一 R1 三层 → [月, 合计, BEV, PHEV, EREV]；
    # 旧 R2 二元拆分保留为 monthly_prev_r2（标注已废弃·仅审计留档）。
    def r2_energy(rec):
        """R2（已废弃）：任一行为增程 → 整系归 EREV。仅用于 monthly_prev_r2 审计对照。"""
        if rec.get("erev_flag_rows"):
            return "EREV"
        return rec.get("power_type") or "(未定)"

    monthly_rows, energy_rows, price_rows, seg_rows, prev_r2_rows = [], [], [], [], []
    price_cov = {}
    for m in MONTHS:
        lst = monthly.get(m, [])
        mrecs = [series_recs[x["sid"]] for x in lst if x["sid"] in series_recs]
        m_tot = m_bev = m_phev = m_erev = 0
        r2_bev = r2_pe = 0
        pb = defaultdict(int)
        sg = defaultdict(int)
        unpriced_sales = 0
        unpriced_series = 0
        for r in mrecs:
            s = sales_in(m, r["seriesid"])
            if not s:
                continue
            e = energy_of(r)
            m_tot += s
            if e == "BEV":
                m_bev += s
            elif e == "EREV":
                m_erev += s
            elif e == "PHEV":
                m_phev += s
            e2 = r2_energy(r)
            if e2 == "BEV":
                r2_bev += s
            else:
                r2_pe += s
            _band = price_band(parse_price_mid(next((x for x in lst if x["sid"] == r["seriesid"]),
                                                    {}).get("price")))
            if _band:
                pb[_band] += s
            else:
                unpriced_sales += s
                unpriced_series += 1
            sg[(r.get("segment") or "").strip() or "(未定)"] += s
        binned = sum(pb.get(b, 0) for b in PRICE_BINS)
        monthly_rows.append([m, m_tot, m_bev, m_phev, m_erev])
        prev_r2_rows.append([m, m_tot, r2_bev, r2_pe])
        energy_rows.append([m, m_bev, m_phev, m_erev, _share(m_bev, m_tot)])
        price_rows.append([m] + [pb.get(b, 0) for b in PRICE_BINS])
        seg_rows.append([m, {k: v for k, v in sorted(sg.items(), key=lambda kv: (-kv[1], kv[0]))}])
        price_cov[m] = {"binned": binned, "total": m_tot, "pct": _share(binned, m_tot),
                        "uncovered_sales": m_tot - binned, "uncovered_series": unpriced_series}

    def month_top20(m):
        lst = sorted(monthly.get(m, []), key=lambda x: -int(x["n"] or 0))
        return [[x["name"], int(x["n"] or 0)] for x in lst[:20]]

    first_ids = {x["sid"] for x in monthly.get(MONTHS[0], [])}
    last_ids = set(cur)
    prev_ids = {x["sid"] for x in monthly.get(prev_month, [])}
    prev_name = {x["sid"]: x["name"] for x in monthly.get(prev_month, [])}
    new_entries = sorted(([x["name"], int(x["n"] or 0)] for x in cur.values() if x["sid"] not in prev_ids),
                         key=lambda kv: (-kv[1], kv[0]))
    drops = sorted(([prev_name[s], sales_in(prev_month, s)] for s in prev_ids - last_ids),
                   key=lambda kv: (-kv[1], kv[0]))
    rch = Counter()
    for x in cur.values():
        v = x.get("rankchange")
        if v is None:
            rch["缺失"] += 1
            continue
        for lab, lo, hi in RANKCHANGE_BINS:
            if lo <= v <= hi:
                rch[lab] += 1
                break
    trend = {
        "monthly": monthly_rows,
        "monthly_prev_r2": prev_r2_rows,
        "monthly_fields": ["月份", "已锚定合计", "BEV", "PHEV", "EREV"],
        "monthly_prev_r2_fields": ["月份", "已锚定合计", "BEV(R2)", "PHEV+EREV(R2)"],
        "monthly_prev_r2_note": ("**已废弃·仅审计留档**：R2 口径（任一行为增程→整系 EREV）的二元拆分。"
                                 "对外一律使用 `monthly`（R1 三层），两者不得混引（Lead m00990 F1）"),
        "energy_monthly": energy_rows,
        "price_monthly": price_rows,
        "price_monthly_fields": ["月份"] + PRICE_BINS,
        "price_coverage": price_cov,
        "price_coverage_note": ("价格带份额分母统一 = 当月已锚定销量总量；2026-09 未覆盖 806 辆（0.101%）"
                                "为无有效指导价车系，未计入任何价格带（Lead m00990 F3）"),
        "segment_monthly": seg_rows,
        "top20_first": month_top20(MONTHS[0]),
        "top20_last": month_top20(CUR_MONTH),
        "rankchange_hist": [[lab, rch.get(lab, 0)] for lab, _, _ in RANKCHANGE_BINS]
                           + [["缺失", rch.get("缺失", 0)]],
        "new_entries": new_entries[:60],
        "drops": drops[:60],
        # ---- 新增 ----
        "unit": UNIT,
        "universe": ("monthly/energy_monthly/price_monthly/segment_monthly 为已锚定车系口径"
                     "（total == bev+phev+erev 自洽）；top20_*/new_entries/drops/rankchange_hist "
                     "为 2026-09 新能源榜全榜口径"),
        "monthly_board": [[m, sum(int(x["n"] or 0) for x in monthly.get(m, []))] for m in MONTHS],
        "price_monthly_board": [[m] + [sum(int(x["n"] or 0) for x in monthly.get(m, [])
                                          if (price_band(parse_price_mid(x.get("price"))) or "") == b)
                                       for b in PRICE_BINS] for m in MONTHS],
        "new_entries_total": len(new_entries),
        "drops_total": len(drops),
        "roster_rule": ("new_entries/drops = 月度榜单集合差（2026-09 vs %s 新能源榜名册），"
                        "非 Top20 位次差" % prev_month),
        "rankchange_bins": [lab for lab, _, _ in RANKCHANGE_BINS] + ["缺失"],
        "note": ("价格带取自当月榜单 priceinfo 中值；动力类型分解来自底表标注，故仅覆盖已锚定车系。"
                 "动力三分一律 R1（与 structure.by_energy 同源同规则）；"
                 "价格带/动力/品牌/企业份额分母统一 = 已锚定销量。"),
    }

    # ================= 4 销量-技术关联 =================
    weighted, unweighted = {}, {}
    n_w = n_u = 0
    for cn, f in METRIC_KEYS:
        w, nw = _wavg(linked, f, sales_of)
        m, nu = _mean(linked, f)
        weighted[cn] = w
        unweighted[cn] = m
        n_w = max(n_w, nw)
        n_u = max(n_u, nu)

    rng_acc = defaultdict(lambda: {"sales": 0, "ids": set()})
    for r in linked:
        b = range_bin(r.get("range_km")) or "(续航缺失)"
        rng_acc[b]["sales"] += sales_of(r)
        rng_acc[b]["ids"].add(r["seriesid"])
    idcheck.append({"by": "续航分箱", "buckets": {k: sorted(v["ids"]) for k, v in rng_acc.items()}})
    range_bins = [[b, _share(rng_acc[b]["sales"] if b in rng_acc else 0, linked_total),
                   len(rng_acc[b]["ids"]) if b in rng_acc else 0] for b in RANGE_BINS_DEF]

    pb_tech, pr_matrix = [], []
    for b in PRICE_BINS:
        recs = [r for r in linked
                if (price_band(parse_price_mid(cur.get(r["seriesid"], {}).get("price"))) or "") == b]
        pb_tech.append([b, _wavg(recs, "range_km", sales_of)[0], _wavg(recs, "capacity_kwh", sales_of)[0],
                        _wavg(recs, "power_kw", sales_of)[0], _wavg(recs, "curb_kg", sales_of)[0],
                        sum(sales_of(r) for r in recs), len(recs)])
        m_sales, m_cnt = [], []
        for rb in RANGE_BINS_DEF:
            sub = [r for r in recs if (range_bin(r.get("range_km")) or "(续航缺失)") == rb]
            m_sales.append(sum(sales_of(r) for r in sub))
            m_cnt.append(len(sub))
        pr_matrix.append([b, m_sales, m_cnt])
    by_energy_tech = []
    for e in ENERGY_ORDER + sorted({energy_of(r) for r in linked} - set(ENERGY_ORDER)):
        sub = [r for r in linked if energy_of(r) == e]
        if not sub:
            continue
        by_energy_tech.append([e, {cn: _wavg(sub, f, sales_of)[0] for cn, f in METRIC_KEYS}])
    tech = {
        "weighted": weighted,
        "unweighted": unweighted,
        "sample_n": {"weighted": n_w, "unweighted": n_u},
        "range_bins": range_bins,
        "range_bins_def": RANGE_BINS_DEF,
        "price_tech": pb_tech,
        "price_range_matrix": pr_matrix,
        "by_energy_tech": by_energy_tech,
        # ---- 新增 ----
        "unit": UNIT,
        "metric_rule": "车系指标取该车系底表多行的中位数（数值）/众数（分类）",
        "weight_rule": "2026-09 车系销量（salecount）加权",
        "metric_fields": {cn: f for cn, f in METRIC_KEYS},
        "universe": "已锚定车系；weighted 为销量加权均值，unweighted 为等权（公告）均值",
        "missing": {"range": {"sales": rng_acc["(续航缺失)"]["sales"] if "(续航缺失)" in rng_acc else 0,
                              "series_count": len(rng_acc["(续航缺失)"]["ids"])
                              if "(续航缺失)" in rng_acc else 0}},
        "note": ("range_bins 顺序 == range_bins_def（消费方勿自行猜边界）；"
                 "price_range_matrix 外层 6 行 == meta.price_bins 顺序，内层 6 元 == range_bins_def 顺序。"),
    }

    # ================= 5 公告-市场转化 =================
    batch_rows = defaultdict(list)
    for idx, r in enumerate(rows):
        batch_rows[_as_int(r[C_BATCH])].append(idx)

    def anchor_sids(idxs):
        out = set()
        for i in idxs:
            a = gen_anchor.get(str(rows[i][C_GENERIC] or "").strip()) or {}
            if a.get("sid"):
                out.add(a["sid"])
        return out

    def snap_sales(sids):
        return sum(int(cur[s]["n"] or 0) for s in sids if s in cur)

    by_batch, year_rows = [], defaultdict(list)
    for b in sorted(batch_rows, key=lambda x: (x is None, x)):
        idxs = batch_rows[b]
        sids = anchor_sids(idxs)
        y = batch_year.get(b)
        by_batch.append([b, y, len(idxs), snap_sales(sids)])
        if y is not None:
            year_rows[y].append((idxs, sids))
    by_year = []
    for y in sorted(year_rows):
        idxs = [i for ii, _ in year_rows[y] for i in ii]
        sids = set()
        for _, ss in year_rows[y]:
            sids |= ss
        by_year.append([str(y), len(idxs), snap_sales(sids)])
    # 412 批（公示口径）在「批次时间表」无公告年月 → year=None；显式列出，**不得静默吞掉**（Lead m01157）
    _undated = [(b, idxs) for b, idxs in batch_rows.items() if batch_year.get(b) is None]
    by_year_undated = []
    for b, idxs in sorted(_undated):
        by_year_undated.append({"batch": b, "year": None,
                                "note": "412 批=公示口径（无正式发文字号），批次时间表无公告年月",
                                "models": len(idxs), "sales": snap_sales(anchor_sids(idxs))})

    brd_rows = defaultdict(list)
    for idx, r in enumerate(rows):
        brd_rows[brand_of(r[C_TRADEMARK]) or "(未定)"].append(idx)
    brand_scatter, rate_table = [], []
    for b, idxs in brd_rows.items():
        sids = anchor_sids(idxs)
        sc = snap_sales(sids)
        insale = {s for s in sids if s in cur}
        brand_scatter.append([b, len(idxs), sc])
        rate_table.append([b, len(idxs), len(insale), sc,
                           round(sc / len(insale), 1) if insale else 0.0])
    brand_scatter.sort(key=lambda x: (-x[2], x[0]))
    rate_table.sort(key=lambda x: (-x[3], x[0]))
    conversion = {
        "by_batch": by_batch,
        "by_year": by_year,
        "by_year_undated": by_year_undated,
        "by_year_note": ("by_year 仅含「批次时间表」有公告年月的批次；"
                         "`by_year_undated` 显式列出年份缺失的批次（412 公示批），"
                         "下游按年汇总时**必须显式排除或单列『待定』桶**，不得静默丢弃"),
        "brand_scatter": brand_scatter,
        "rate_table": rate_table,
        "by_batch_note": ("by_batch 第 3 位 = 该批次公告车型数（底表行数，=公告车型/型号条数）；"
                          "第 4 位 = 该批次公告车型在 2026-09 的**车系级快照销量合计**"
                          "（统一快照口径，非批次发布当期切片）；同批次内同车系只计一次。"),
        # ---- 新增 ----
        "unit": UNIT,
        "year_source": "底表「批次时间表」sheet 的公告年月（例 341→2021-03）",
        "by_batch_detail": [[b, len(batch_rows[b]), len({str(rows[i][C_CODE] or "").strip()
                                                         for i in batch_rows[b]}),
                             len(anchor_sids(batch_rows[b]))] for b in
                            sorted(batch_rows, key=lambda x: (x is None, x))],
        "by_batch_detail_fields": ["批次", "公告行数", "去重产品型号数", "已锚定车系数"],
        "note": ("跨批次/跨年同车系会分别计入各自批次/年份（车系级去重仅在组内进行），"
                 "故 by_batch 求和 > sales_total_dedup 属正确现象。"),
    }

    # ================= 6 覆盖状态（契约 v1.1） =================
    total_rows = len(rows)
    buckets = {k: int(status_counts.get(k, 0)) for k in
               (STATUS_NOID, STATUS_GONE, STATUS_HOLD, STATUS_NOMATCH)}
    coverage = {
        "row": {"covered": int(match_stats["matched_rows"]), "total": total_rows,
                "pct": _share(match_stats["matched_rows"], total_rows)},
        "series": {"matched": int(match_stats["distinct_seriesid_anchored"]),
                   "total": int(match_stats["generic_denom"]),
                   "pct": _share(match_stats["distinct_seriesid_anchored"],
                                 match_stats["generic_denom"]),
                   "generic_names_matched": int(match_stats["generic_anchored"]),
                   "generic_names_total": int(match_stats["generic_denom"]),
                   "generic_names_pct": _share(match_stats["generic_anchored"],
                                               match_stats["generic_denom"])},
        # Lead m00237①：对外唯一可主张的「在售车系覆盖率」（当期在榜口径，非 10 月并集）
        "current_board": {
            "board_series": len(cur),
            "linked_series": len(linked),
            "pct": _share(len(linked), len(cur)),
            "linked_generic_cells": sum(1 for v in gen_anchor.values() if v.get("sid") in cur),
            "generic_cells_anchored": int(match_stats["generic_anchored"]),
            "linked_sales": linked_total,
            "board_sales": board_total,
            "sales_pct": _share(linked_total, board_total),
            "definition": ("已锚定车系 ∩ 2026-09 新能源榜 / 2026-09 全榜车系数；"
                           "**对外任何「在售车系覆盖率」主张只准引用本字段**"),
        },
        # Lead m00237②：自指指标（分母由分子派生），非覆盖率，不得对外主张
        "self_consistent_ratio": {
            "matched": int(match_stats["distinct_seriesid_anchored"]),
            "total": int(match_stats["distinct_seriesid_anchored"] + match_stats["generic_hold"]),
            "pct": _share(match_stats["distinct_seriesid_anchored"],
                          match_stats["distinct_seriesid_anchored"] + match_stats["generic_hold"]),
            "generic_names_matched": int(match_stats["generic_anchored"]),
            "generic_names_hold": int(match_stats["generic_hold"]),
            "generic_names_pct": _share(match_stats["generic_anchored"],
                                        match_stats["generic_anchored"] + match_stats["generic_hold"]),
            "definition": "分母（已匹配+命名待锚定）由分子派生，属同义反复，**不是覆盖率**",
        },
        # Lead m01097：col39 非空 2603 ≠ col35 非空 2185，差 418（设计使然，主动披露）
        "anchored_not_on_board": {
            "rows": sum(1 for r in rows
                        if (gen_anchor.get(str(r[C_GENERIC] or "").strip()) or {}).get("sid")
                        and gen_anchor[str(r[C_GENERIC] or "").strip()]["sid"] not in cur),
            "series": len({v["sid"] for v in gen_anchor.values() if v.get("sid")} - set(cur)),
            "note": ("已锚定车系ID、但该车系不在 2026-09 榜（属 10 个月并集），"
                     "按口径当期销量留空、不写旧月值不写 0；与『未匹配-已停产或未上量』"
                     f"{int(status_counts.get(STATUS_GONE, 0)):,} 行为包含关系"
                     f"（该状态含本类 {sum(1 for r in rows if (gen_anchor.get(str(r[C_GENERIC] or '').strip()) or {}).get('sid') and gen_anchor[str(r[C_GENERIC] or '').strip()]['sid'] not in cur):,} 行）"),
        },
        "status_counts": {k: int(status_counts.get(k, 0)) for k in STATUS_ENUM},
        "buckets": buckets,
        "pending_naming": pending[:50],
        "definitions": {
            "row": f"已填销量行 / 底表数据行（{len(rows):,}，实读，非硬编码）",
            "series": ("已匹配独一无二车系数 / (去重通用名称数 − 无车系身份)；"
                       f"当前分子 {int(match_stats['distinct_seriesid_anchored'])}，"
                       f"分母 {int(match_stats['generic_denom'])}"),
            "current_board": "对外唯一在售覆盖率口径：已锚定车系 ∩ 2026-09 在榜 / 2026-09 全榜车系数（411）",
            "self_consistent_ratio": "自指指标，分母由分子派生 → 非覆盖率，不得对外主张",
            "buckets": "四类未覆盖行数（= status_counts 后 4 类之和）",
        },
        "detail": {
            "generic_names_all": int(match_stats["generic_all"]),
            "generic_names_no_identity": int(match_stats["generic_noident"]),
            "generic_names_denominator": int(match_stats["generic_denom"]),
            "generic_names_anchored": int(match_stats["generic_anchored"]),
            "generic_names_hold": int(match_stats["generic_hold"]),
            "generic_names_unmatched": int(match_stats["generic_unmatched"]),
            "distinct_seriesid_anchored": int(match_stats["distinct_seriesid_anchored"]),
            "pending_naming_total": int(match_stats["generic_hold"]),
            "pending_naming_shown": min(50, len(pending)),
        },
    }

    # ================= 7 去重自检 =================
    selfcheck = dedup_selfcheck(series_recs, gen_anchor, cur, rows, sid_by_boardname)
    assertions = _dedup_assertions(idcheck, linked, selfcheck)

    # ================= meta（契约 v1.1） =================
    mc = {m: len(monthly.get(m, [])) for m in MONTHS}          # 键集合必须 == months
    ch = match_stats["channel_counts"]
    meta = {
        "snapshot_month": CUR_MONTH,
        "months": list(MONTHS),
        "fetch_date": CAPTURED_AT,
        "source": "汽车之家 · 行业综合销量（车系级月零售）",
        "source_tip": "数据源于行业综合销量，每月10日左右更新",
        "source_url_template": ("https://car-web-api.autohome.com.cn/car/rank/getList?..."
                               "&typeid=9&subranktypeid=2305&date={month}"),
        "reference_source": "全车型总榜 typeid=1（仅参照） + 易车202608快照 + 懂车帝微型车榜",
        "scope": "新能源榜（与本项目仅 NEV 乘用车口径同宇宙）",
        "dedup_rule": "车系级（seriesid 唯一）；底表行级值仅供单行参考，禁止加总",
        "price_bins": list(PRICE_BINS),
        "price_bin_rule": "priceinfo 指导价区间取中值后分箱",
        "backup_old": "audit-output/sales_prior_yiche_202608_backup_20261019.json",
        "base_table_path": os.path.basename(WB_PATH),
        "base_rows": total_rows,
        "matched_rows": int(match_stats["matched_rows"]),
        "match_channel_counts": {
            "A_exact": int(ch.get("A_exact", 0)),
            "B_brand": int(ch.get("B_brand", 0)),
            "C_substr": int(ch.get("C_substr", 0)),
            "D_suggest": int(ch.get("D_suggest", 0)),
            "hold": int(ch.get("hold", 0)),
            "A2_split": int(ch.get("A2_split", 0)),
            "multi_series_cell": int(ch.get("multi_series_cell", 0)),
            "unmatched": int(ch.get("unmatched", 0)),
            "placeholder": int(ch.get("placeholder", 0)),
            "empty": int(ch.get("empty", 0)),
        },
        "qc": {
            "monthly_counts": mc,
            "union_series": int(match_stats["union_series"]),
            "current_month_series": len(cur),
            "reference_universe_per_month": _ref_universe(src_meta),
            "reference_universe_note": (
                "参照宇宙 = 汽车之家**全车型总榜**（typeid=1&subranktypeid=1）逐月车系数与销量合计；"
                "仅作分母/渗透率参照，**不入底表**。来源：WS1 官方产物 "
                "`audit-output/sales_autohome_series_monthly_20261019.json` → "
                "`reference_only.全车型榜.monthly`（与 WS1 QC 报告 "
                "`sales_source_qc_20261019.json` 逐月一致）"),
            "cross_source": cross_check,
        },
        "dedup_selfcheck": selfcheck["star"],
    }
    meta["match_channel_counts_rows"] = match_stats["channel_counts_rows"]
    meta["match_channel_detail"] = {
        "unit": f"通用名称单元格（去重后 {int(match_stats['generic_all']):,} 格，实读）",
        "A2_split": "并列格按分隔符切候选后唯一命中",
        "multi_series_cell": "一枚并列格命中 ≥2 个 seriesid → 该行 HOLD",
        "A2_split_base_channels": match_stats["a2_base_channels"],
    }
    meta["source_meta"] = {k: src_meta.get(k) for k in
                           ("source_kind", "source_path", "source_tag", "has_rank_fields")}
    meta["ledger_file"] = "audit-output/sales_link_20261019.json"
    meta["coverage_status_enum"] = STATUS_ENUM

    out = {
        "meta": meta,
        "coverage": coverage,
        "structure": structure,
        "concentration": concentration,
        "trend": trend,
        "tech_sales": tech,
        "conversion": conversion,
        # ---- 新增顶层 ----
        "dedup_selfcheck": selfcheck,
        "dedup_assertions": assertions,
        "unit": UNIT,
        "dedup_rule": DEDUP_RULE,
        "generated_at": CAPTURED_AT,
    }

    # ---- 契约前置断言（Lead R3 验收）
    asserts = []
    asserts.append({"name": "coverage.status_counts 六类之和 == base_rows",
                    "value": sum(coverage["status_counts"].values()), "expect": total_rows,
                    "pass": sum(coverage["status_counts"].values()) == total_rows})
    asserts.append({"name": "coverage.buckets 四类之和 == status_counts 后四类之和",
                    "value": sum(coverage["buckets"].values()),
                    "expect": sum(coverage["status_counts"][k] for k in
                                  (STATUS_NOID, STATUS_GONE, STATUS_HOLD, STATUS_NOMATCH)),
                    "pass": sum(coverage["buckets"].values()) ==
                            sum(coverage["status_counts"][k] for k in
                                (STATUS_NOID, STATUS_GONE, STATUS_HOLD, STATUS_NOMATCH))})
    asserts.append({"name": "trend.price_monthly 每行长度 == len(meta.price_bins)",
                    "value": sorted({len(r) for r in price_rows}), "expect": 1 + len(PRICE_BINS),
                    "pass": all(len(r) == 1 + len(PRICE_BINS) for r in price_rows)})
    asserts.append({"name": "meta.qc.monthly_counts 键集合 == meta.months",
                    "value": sorted(mc), "expect": sorted(MONTHS),
                    "pass": sorted(mc) == sorted(MONTHS)})
    asserts.append({"name": "trend.monthly 10 行、5 元（R1 三层）且 total == bev+phev+erev",
                    "value": [[r[0], len(r), r[1] - r[2] - r[3] - r[4]] for r in monthly_rows],
                    "pass": (len(monthly_rows) == 10
                             and all(len(r) == 5 for r in monthly_rows)
                             and all(r[1] == r[2] + r[3] + r[4] for r in monthly_rows))})
    asserts.append({"name": "F1：trend.monthly 与 structure.by_energy 动力口径一致（R1 同源）",
                    "value": {"by_energy": dict((k, v) for k, v in eng_sales),
                              "monthly_last": monthly_rows[-1][2:]},
                    "expect": [dict((k, v) for k, v in eng_sales).get(k, 0)
                               for k in ("BEV", "PHEV", "EREV")],
                    "pass": (dict((k, v) for k, v in eng_sales).get("BEV", 0) == monthly_rows[-1][2]
                             and dict((k, v) for k, v in eng_sales).get("PHEV", 0) == monthly_rows[-1][3]
                             and dict((k, v) for k, v in eng_sales).get("EREV", 0) == monthly_rows[-1][4])})
    asserts.append({"name": "F3：trend.price_monthly 各月分箱合计 == 当月已锚定合计 − 未覆盖",
                    "value": [[m, price_cov[m]["binned"], price_cov[m]["total"],
                               price_cov[m]["uncovered_sales"]] for m in MONTHS],
                    "pass": all(price_cov[m]["binned"] + price_cov[m]["uncovered_sales"]
                                == price_cov[m]["total"] for m in MONTHS)})
    asserts.append({"name": "price_range_matrix 6×6（销量/车系数各 6 元）",
                    "value": len(pr_matrix),
                    "pass": len(pr_matrix) == 6 and all(len(r[1]) == 6 and len(r[2]) == 6 for r in pr_matrix)})
    asserts.append({"name": "range_bins 顺序 == range_bins_def",
                    "value": [r[0] for r in range_bins],
                    "pass": [r[0] for r in range_bins] == RANGE_BINS_DEF})
    asserts.append({"name": "去重结构断言（各分桶 seriesid 两两不交）",
                    "value": assertions["all_pass"], "pass": bool(assertions["all_pass"])})
    out["contract_assertions"] = {"assertions": asserts, "all_pass": all(a["pass"] for a in asserts)}
    return out


def self_name(cur, sid):
    return (cur.get(sid) or {}).get("name") or sid


def oem_of(series_recs, cur, name):
    for r in series_recs.values():
        if r["seriesid"] in cur and (cur[r["seriesid"]].get("name") or "") == name:
            return (r.get("company") or "").strip() or "(未定)"
    return "(未定)"


def brd_of(series_recs, cur, name):
    for r in series_recs.values():
        if r["seriesid"] in cur and (cur[r["seriesid"]].get("name") or "") == name:
            return brand_of(r.get("trademark")) or "(未定)"
    return "(未定)"


DEDUP_CASES = [("星愿", "星愿"), ("Model Y", "Model Y"), ("Model 3", "Model 3"),
               ("比亚迪宋Ultra", "宋Ultra"), ("元UP", "元UP"), ("海豚", "海豚")]


def _resolve_case_sid(key, sid_by_boardname):
    k = norm(key)
    if k in sid_by_boardname and len(sid_by_boardname[k]) == 1:
        return sorted(sid_by_boardname[k])[0]
    for b in BRAND_TOKENS:
        nb = norm(b)
        if k.startswith(nb) and len(k) > len(nb):
            rest = k[len(nb):]
            if rest in sid_by_boardname and len(sid_by_boardname[rest]) == 1:
                return sorted(sid_by_boardname[rest])[0]
    return None


def dedup_selfcheck(series_recs, gen_anchor, cur, rows, sid_by_boardname):
    """Lead m00037 指定用例：同车系多枚通用名称必须塌缩到唯一 seriesid 且排名表只占一席。"""
    top20_ids = [x["sid"] for x in sorted(cur.values(), key=lambda y: -int(y["n"] or 0))[:20]]
    cases = []
    for label, key in DEDUP_CASES:
        sid = _resolve_case_sid(key, sid_by_boardname)
        cells = sorted({g for g, v in gen_anchor.items() if v.get("sid") and v["sid"] == sid}) if sid else []
        nrows = sum(1 for r in rows if (gen_anchor.get(str(r[C_GENERIC] or "").strip()) or {}).get("sid") == sid)
        in_top = sum(1 for s in top20_ids if s == sid)
        own = [g for g in cells if gen_anchor[g]["sid"] == sid]
        cases.append({
            "case": label,
            "resolved_seriesid": sid,
            "seriesname_in_board": (cur.get(sid) or {}).get("name"),
            "generic_names": cells,
            "generic_names_count": len(cells),
            "collapsed_rows": nrows,
            "distinct_series_in_top20": in_top,
            "in_top20": in_top > 0,
            "pass": bool(sid) and len(cells) >= 1 and len(own) == len(cells),
        })
    star = next((c for c in cases if c["case"] == "星愿"), {})
    multi = [{"generic": g, "rows": sum(1 for r in rows
                                        if str(r[C_GENERIC] or "").strip() == g),
              "sids": sorted(gen_anchor[g]["sids"]) if gen_anchor[g].get("sids") else [],
              "candidate_names": gen_anchor[g].get("cands", [])[:8]}
             for g, v in gen_anchor.items() if v.get("channel") == "multi_series_cell"]
    multi.sort(key=lambda x: (-x["rows"], x["generic"]))
    return {
        "star": star,
        "cases": cases,
        "summary": {"cases_checked": len(cases), "cases_pass": sum(1 for c in cases if c["pass"]),
                    "rule": DEDUP_RULE},
        "multi_series_cells": multi[:20],
        "multi_series_cell_count": len(multi),
        "note": ("star = meta.dedup_selfcheck（契约字段）；cases 为全部指定用例。"
                 "distinct_series_in_top20 = 该用例在 2026-09 Top20 中占据的席位数，去重正确时应 == 1"
                 "（未进 Top20 时为 0，见 in_top20）。"),
    }


def _dedup_assertions(idcheck, linked, selfcheck):
    """结构性断言：任一分桶型聚合表中同一 seriesid 只出现一次。"""
    linked_ids = {r["seriesid"] for r in linked}
    res = []
    for blk in idcheck:
        seen = Counter()
        for ids in blk["buckets"].values():
            for s in ids:
                seen[s] += 1
        dup = {s: c for s, c in seen.items() if c > 1}
        union = set(seen)
        res.append({"by": blk["by"], "buckets": len(blk["buckets"]),
                    "series_count_sum": sum(len(v) for v in blk["buckets"].values()),
                    "distinct_series": len(union), "linked_series": len(linked_ids),
                    "duplicated": list(dup.items())[:10],
                    "pass": not dup and union == linked_ids})
    return {"assertions": res, "all_pass": all(x["pass"] for x in res),
            "selfcheck_pass": selfcheck["summary"]["cases_pass"] == selfcheck["summary"]["cases_checked"]}

# --------------------------------------------------------------------------- 主流程


CROSSCHECK_FILE = os.path.join(HERE, "_lead_crosscheck_202608.json")
LEAD_FP_FILE = os.path.join(AUDIT, "_lead_prewrite_fingerprint_20261019.json")
QC_FILE = os.path.join(AUDIT, "sales_source_qc_20261019.json")
BATCH_SHEET = "批次时间表"


def read_batch_years(wb):
    """批次 → 公告年份（底表「批次时间表」sheet 的公告年月列，如 341→2021-03）。"""
    out = {}
    try:
        ws = wb[BATCH_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            b = _as_int(row[0] if row else None)
            ym = str((row[1] if row and len(row) > 1 else "") or "").strip()
            m = re.match(r"^(\d{4})", ym)
            if b is not None and m:
                out[b] = int(m.group(1))
    except Exception as e:                                        # noqa: BLE001
        print(f"[warn] 批次时间表 读取失败：{type(e).__name__}: {e}")
    return out


def read_batch_months(wb):
    """批次 → 公告年月原文（如 341→2021-03），用于候选表代际判据。"""
    out = {}
    try:
        ws = wb[BATCH_SHEET]
        for row in ws.iter_rows(min_row=2, values_only=True):
            b = _as_int(row[0] if row else None)
            ym = str((row[1] if row and len(row) > 1 else "") or "").strip()
            if b is not None and ym:
                out[b] = ym
    except Exception as e:                                        # noqa: BLE001
        print(f"[warn] 批次时间表 读取失败：{type(e).__name__}: {e}")
    return out


def load_cross_check(log=print):
    """QC 跨源校验。易车202608 引用 Lead 独立交叉验证产物；懂车帝202609 只读镜像 WS1 QC 报告。"""
    cc = {"yiche_202608": {"compared": 0, "within_1pct": 0, "over_10pct": [], "mean_rel_diff": 0.0,
                           "source": "_lead_crosscheck_202608.json（Lead 独立交叉验证，只读引用）",
                           "note": "汽车之家全车型总榜(typeid=1) 2026-08 vs 底表既有 806 格易车 202608"}}
    try:
        with open(CROSSCHECK_FILE, encoding="utf-8") as f:
            d = json.load(f)
        comps = d.get("comparisons") or []
        rels = [abs(float(c.get("rel_diff") or 0)) for c in comps]
        over = [[c.get("generic"), c.get("base_sales"), c.get("autohome_sales"),
                 round(float(c.get("rel_diff") or 0), 4)]
                for c in comps if abs(float(c.get("rel_diff") or 0)) > 0.10]
        cc["yiche_202608"] = {
            "compared": int(d.get("comparable") or 0),
            "within_1pct": int(d.get("within_1pct") or 0),
            "within_5pct": int(d.get("within_5pct") or 0),
            "unmatched": int(d.get("unmatched") or 0),
            "autohome_count": int(d.get("autohome_count") or 0),
            "over_10pct": over,
            "over_10pct_fields": ["generic", "base_sales_yiche", "autohome_sales", "rel_diff"],
            "mean_rel_diff": round(sum(rels) / len(rels), 6) if rels else 0.0,
            "source": "_lead_crosscheck_202608.json（Lead 独立交叉验证，只读引用）",
            "note": ("易车202608 与 汽车之家202608 同属行业综合销量家族；换源唯一实质差异是覆盖率，"
                     "相对差 >10% 仅五菱宏光一例口径尾数"),
        }
    except Exception as e:                                        # noqa: BLE001
        log(f"[warn] 跨源校验文件读取失败（{type(e).__name__}: {e}）：compared 置 0")

    # ---- 懂车帝 202609：只读镜像 WS1 QC 报告（Lead m00341②；替换旧的 compared=0 空槽位）
    cc["qc_report_path"] = "audit-output/sales_source_qc_20261019.json"
    cc["qc_report_note"] = "以 WS1 QC 报告为权威，本处为只读摘要（Lead m00341②）"
    cc["deprecated"] = {"dongchedi_micro_202609":
                        "已废弃：WS1 完成懂车帝 2026-09 全量抓取（660 条）前的旧状态（compared=0），"
                        "禁止引用；请改用 dongchedi_202609"}
    try:
        with open(QC_FILE, encoding="utf-8") as f:
            q = json.load(f)
        x = ((q.get("c_crosscheck") or {}).get("dongchedi_202609")) or {}
        cc["dongchedi_202609"] = {
            "status": x.get("status"),
            "source": x.get("source"),
            "scope_note": x.get("scope_note"),
            "dcd_rows": x.get("dcd_rows"),
            "dcd_month": x.get("ah_month"),
            "matched": x.get("matched"),
            "comparable_nev_count": x.get("comparable_nev_count"),
            "within_1pct": x.get("within_1pct"),
            "mean_abs_rel_diff": x.get("mean_abs_rel_diff_comparable"),
            "over_10pct_count": x.get("over_10pct_count"),
            "over_10pct_by_category": x.get("over_10pct_by_category"),
            "over_10pct_anchored_by_sibling_evidence": x.get("over_10pct_anchored_by_sibling_evidence"),
            "over_10pct_unexplained_count": x.get("over_10pct_unexplained_count"),
            "claim_arbitration_conflict_count": (x.get("claim_arbitration") or {}).get("conflict_count"),
            "scope_difference_conclusion": (x.get("scope_difference_conclusion") or "")[:800],
            "mirror_of": "audit-output/sales_source_qc_20261019.json#c_crosscheck.dongchedi_202609",
            "note": "只读镜像，未在本地重算",
        }
        log(f"[qc  ] 懂车帝202609 镜像：matched={cc['dongchedi_202609']['matched']} "
            f"comparable={cc['dongchedi_202609']['comparable_nev_count']} "
            f"within1%={cc['dongchedi_202609']['within_1pct']} "
            f"over10%={cc['dongchedi_202609']['over_10pct_count']}"
            f"(无法解释={cc['dongchedi_202609']['over_10pct_unexplained_count']})")
    except Exception as e:                                        # noqa: BLE001
        cc["dongchedi_202609"] = {"status": "qc_report_unreadable",
                                  "error": f"{type(e).__name__}: {e}",
                                  "mirror_of": "audit-output/sales_source_qc_20261019.json"}
        log(f"[warn] WS1 QC 报告读取失败（{type(e).__name__}: {e}）：懂车帝镜像槽位不可用")
    return cc


def _anchor_forms(g):
    """通用名称的候选形态（整串 + 分隔符切分），norm 后去重去空。"""
    out = []
    for s in [g] + [p for p in SPLIT_RE.split(str(g or ""))]:
        f = norm(s)
        if f and f not in out:
            out.append(f)
    return out


def _digits_of(s):
    return re.sub(r"\D", "", str(s or ""))


def build_anchor_candidates(rows, gen_anchor, monthly, batch_year, batch_month, log=print):
    """Lead m00239：未关联在榜车系的锚定候选证据（**只出证据，不写底表**）。

    verdict ∈ {可锚定, 须人工裁决, 跨代疑云, 无候选}；多候选一律 须人工裁决，严禁自动定值。
    """
    cur = {x["sid"]: x for x in monthly.get(CUR_MONTH, [])}
    linked_sids = {v["sid"] for v in gen_anchor.values() if v.get("sid") is not None}
    # 痕迹池 = 全部底表通用名称形态（含已锚定到别的车系者），与 structure.unlinked_board.cause_bucket 同池，
    # 保证「88 个有痕迹车系」两侧一致（Lead m00239）。
    form2gen = defaultdict(set)
    for g, v in gen_anchor.items():
        if v.get("channel") in ("placeholder", "empty"):
            continue
        for f in _anchor_forms(g):
            if len(f) >= 2:
                form2gen[f].add(g)
    gen_sid = {g: (v.get("sid")) for g, v in gen_anchor.items()}
    gen2rows = defaultdict(list)
    for i, r in enumerate(rows):
        gen2rows[str(r[C_GENERIC] or "").strip()].append(i)

    out = []
    for x in sorted(cur.values(), key=lambda y: (-int(y["n"] or 0), str(y["name"]))):
        if x["sid"] in linked_sids:
            continue
        bn = norm(x["name"])
        btok = set(brand_tokens_of(x["name"])) | set(brand_tokens_of(x.get("brand") or ""))
        gens = sorted({g for f, gs in form2gen.items() if f and (f in bn or bn in f) for g in gs})
        cands = []
        for g in gens:
            idxs = gen2rows.get(g) or []
            if not idxs:
                continue
            forms_g = _anchor_forms(g)
            fa = max(forms_g, key=len)
            # 按 (产品商标, 企业名称) 分组：裸通用名常跨企业复用（AGENTS 2026-09-17），
            # 品牌签名必须**分组**判定，否则并集会把「某一行恰好同品牌」掩盖成整格品牌匹配。
            groups = defaultdict(list)
            for i in idxs:
                groups[(str(rows[i][C_TRADEMARK] or "").strip(),
                        str(rows[i][C_COMPANY] or "").strip())].append(i)
            bgroups = []
            for (tm, cp), gi in groups.items():
                ctok = set(brand_tokens_of(tm)) | set(brand_tokens_of(cp))
                alias_hit = sorted("%s|%s" % (a, c) for a in btok for c in ctok
                                   if (a, c) in BRAND_ALIAS or (c, a) in BRAND_ALIAS)
                bgroups.append({"产品商标": tm, "企业名称": cp, "rows": len(gi),
                                "brand_ok": bool(btok & ctok) or bool(alias_hit),
                                "brand_alias_hit": alias_hit, "brand_tokens_base": sorted(ctok),
                                "excel_rows": [i + 2 for i in gi[:40]]})
            bgroups.sort(key=lambda d: (not d["brand_ok"], -d["rows"], d["企业名称"]))
            best = bgroups[0]
            bi = groups[(best["产品商标"], best["企业名称"])]
            r0 = rows[bi[0]]
            b = _as_int(r0[C_BATCH])
            brand_ok = any(x["brand_ok"] for x in bgroups)
            # 品牌签名三态：confirmed / conflict（明确冲突）/ unknown（两侧缺品牌词，证据不足）
            if brand_ok:
                bstate = "confirmed"
            elif (not btok) or any(not g["brand_tokens_base"] for g in bgroups):
                bstate = "unknown"
            else:
                bstate = "conflict"
            suffix = [f for f in forms_g if f and bn.endswith(f) and bn != f]
            weak = [f for f in forms_g if f and f not in suffix and len(f) >= 2 and (f in bn or bn in f)]
            fam_ok, fam_note = False, ""
            if suffix:
                f = max(suffix, key=len)
                prefix = bn[:len(bn) - len(f)]
                fam_ok = bool(brand_tokens_of(prefix)) or prefix == ""
                fam_note = ("底表名 norm=「%s」为榜名后缀，前缀=「%s」→ %s"
                            % (f, prefix or "(空)",
                               "前缀含品牌词：真锚定" if fam_ok else "前缀不含品牌词：存疑"))
            elif weak:
                fam_note = "仅弱包含（%s），非后缀关系 → 假锚定风险（如 07⊂极氪007GT）" % "/".join(weak[:3])
            else:
                fam_note = "与榜名无包含关系"
            gen_hint = ""
            if re.search(r"\d", bn) and re.search(r"\d", fa) and _digits_of(bn) != _digits_of(fa):
                gen_hint = ("数字片段不一致：底表「%s」vs 榜名「%s」→ 跨代/假锚定疑云（如 EZ-6 vs EZ-60）"
                            % (_digits_of(fa), _digits_of(bn)))
            elif re.fullmatch(r"[0-9a-z]+", fa) and len(fa) <= 3:
                gen_hint = "底表名为纯数字/字母碎片（%s），非车系名 → 假锚定风险" % fa
            cands.append({
                "generic_name": g, "rows": len(idxs), "excel_rows": [i + 2 for i in idxs[:60]],
                "企业名称": best["企业名称"], "产品商标": best["产品商标"],
                "细分市场": str(r0[C_SEGMENT] or "").strip(),
                "动力类型": str(r0[C_POWER] or "").strip(),
                "批次": b, "公告年月": batch_month.get(b) or "", "年份": batch_year.get(b),
                "整备质量": r0[C_CURB], "轴距": r0[C_WHEEL], "车长": r0[C_LENGTH],
                "纯电续航": r0[C_RANGE], "电池容量": r0[C_CAP],
                "product_models": sorted({str(rows[i][C_CODE] or "").strip() for i in bi})[:8],
                "brand_ok": brand_ok,
                "brand_state": bstate,
                "brand_groups": bgroups,
                "suffix_of_board": suffix, "weak_forms": weak[:4],
                "base_anchored_sid": gen_sid.get(g),
                "base_anchored_elsewhere": bool(gen_sid.get(g)),
                "signature_match": {"brand_ok": brand_ok, "brand_state": bstate,
                                    "model_family_ok": fam_ok,
                                    "generation_hint": gen_hint},
                "family_note": fam_note,
            })
        strong = [c for c in cands if c["signature_match"]["brand_ok"]]
        gen_first = next((c for c in strong if c["signature_match"]["generation_hint"]), None)
        if not cands:
            verdict = "无候选"
            basis = "底表无任何名称痕迹"
        elif not strong:
            if any(c["brand_state"] == "conflict" for c in cands):
                verdict = "无候选"
                basis = "有名称痕迹但品牌明确冲突（跨品牌误配风险，如 S3 同形不同厂）"
            else:
                verdict = "须人工裁决"
                basis = "有名称痕迹但一侧无品牌词，品牌签名无法证实（证据不足，须人工裁决）"
        elif len(strong) > 1:
            verdict = "须人工裁决"
            basis = "多个候选品牌签名通过（%s）" % "、".join(c["generic_name"] for c in strong[:4])
        elif gen_first:
            verdict = "跨代疑云"
            basis = gen_first["signature_match"]["generation_hint"]
        elif strong[0]["signature_match"]["model_family_ok"]:
            verdict = "可锚定"
            basis = strong[0]["family_note"]
        else:
            verdict = "须人工裁决"
            basis = strong[0]["family_note"]
        out.append({"seriesname": x["name"], "salecount": int(x["n"] or 0), "seriesid": x["sid"],
                    "priceinfo": x.get("price") or "", "rank": x.get("rank"),
                    "traced": bool(cands),
                    "base_candidates": cands, "verdict": verdict, "verdict_basis": basis})

    vc = Counter(a["verdict"] for a in out)
    traced = [a for a in out if a["traced"]]
    payload = {
        "meta": {
            "task": "WS2 未关联在榜车系锚定候选证据（Lead m00239）",
            "captured_at": CAPTURED_AT, "snapshot_month": CUR_MONTH,
            "source": SOURCE_TAG, "base_table_path": os.path.basename(WB_PATH),
            "scope": "2026-09 新能源榜车系中未与底表建立 seriesid 关联者",
            "board_series": len(cur),
            "board_linked_series": sum(1 for s in linked_sids if s in cur),
            "linked_seriesid_union": len(linked_sids),
            "unlinked_series": len(out), "unlinked_sales": sum(a["salecount"] for a in out),
            "traced_series": len(traced), "traced_sales": sum(a["salecount"] for a in traced),
            "note": ("只出证据、不写底表；底表本轮仍按 `未匹配-命名待锚定` 披露。"
                     "多候选一律 须人工裁决，严禁自动定值（AGENTS 2026-09-17 泛化名撞名跨品牌事故）。"),
            "verdict_defs": {
                "可锚定": "唯一候选且品牌三重签名通过、底表名为榜名后缀且前缀含品牌词",
                "须人工裁决": "多候选，或型号族关系不成立",
                "跨代疑云": "候选数字片段与榜名不一致（如 EZ-6 vs EZ-60）",
                "无候选": "底表无痕迹，或痕迹候选品牌签名不通过（跨品牌误配风险）",
            },
        },
        "verdict_counts": dict(vc),
        "verdict_counts_traced": dict(Counter(a["verdict"] for a in traced)),
        "candidates": out,
    }
    log(f"[anch] 未关联在榜={len(out)}（销量 {sum(a['salecount'] for a in out)}） 有痕迹={len(traced)}"
        f"（销量 {sum(a['salecount'] for a in traced)}） verdict={dict(vc)}")
    return payload


def load_lead_fingerprint_check(rows, wb_sha, log=print):
    """与 Lead 写入前全表指纹（`_lead_prewrite_fingerprint_20261019.json`，只读）对拍：
    ① 旧值 806 格逐行核对 ② `(批次,产品型号)` 重复键陷阱一致性（Lead m00925②）。
    """
    out = {"path": "audit-output/_lead_prewrite_fingerprint_20261019.json", "available": False}
    try:
        with open(LEAD_FP_FILE, encoding="utf-8") as f:
            d = json.load(f)
        out["available"] = True
        out["lead_sha256"] = d.get("sha256")
        out["sha_match"] = (str(d.get("sha256") or "") == wb_sha)
        out["lead_n_rows"] = d.get("n_rows")
        old = d.get("old_sales_cells") or []
        mine = {i + 2: r for i, r in enumerate(rows)}
        mism = []
        for c in old:
            r = mine.get(c.get("row"))
            if r is None:
                mism.append({"row": c.get("row"), "reason": "行不存在"})
                continue
            # 比对口径：两侧一律 strip（Lead 侧保留原始空白，本脚本读入后 strip）
            got = {"batch": _as_int(r[C_BATCH]), "code": _cell_key(r[C_CODE]),
                   "generic": _cell_key(r[C_GENERIC]), "v35": _as_int(r[C_SALES]),
                   "v36": _cell_key(r[C_SMONTH]), "v37": _cell_key(r[C_SDATE])}
            want = {"batch": _as_int(c.get("batch")), "code": _cell_key(c.get("code")),
                    "generic": _cell_key(c.get("generic")), "v35": _as_int(c.get("v35")),
                    "v36": _cell_key(c.get("v36")), "v37": _cell_key(c.get("v37"))}
            if got != want:
                mism.append({"row": c.get("row"), "lead": want, "mine": got})
        out["old_sales_checked"] = len(old)
        out["old_sales_mismatch"] = mism[:20]
        out["old_sales_mismatch_count"] = len(mism)
        out["old_sales_mine_count"] = sum(1 for r in rows if r[C_SALES] not in (None, ""))
        out["old_sales_match"] = (not mism) and len(old) == out["old_sales_mine_count"]
        _dup = Counter("%s|%s" % (_as_int(r[C_BATCH]),
                                  str(r[C_CODE]).strip() if r[C_CODE] is not None else "None")
                       for r in rows)
        my_dups = sorted(k for k, v in _dup.items() if v > 1)
        out["lead_dup_row_keys"] = list(d.get("dup_row_keys") or [])
        out["my_dup_row_keys"] = my_dups
        out["dup_row_keys_match"] = (out["lead_dup_row_keys"] == my_dups)
        out["dup_row_keys_note"] = ("本表 `(批次,产品型号)` 非唯一键 → 行身份复核必须"
                                    "桶内出现序配对 + 整行内容指纹（Lead m00925②）")
        out["lead_dirty_generic"] = d.get("dirty_generic")
        out["lead_dirty_segment_whitespace"] = d.get("dirty_segment_whitespace")
    except Exception as e:                                        # noqa: BLE001
        out["error"] = f"{type(e).__name__}: {e}"
        log(f"[warn] Lead 写入前指纹读取失败：{out['error']}")
    log(f"[lead-fp] 旧值 {out.get('old_sales_checked')} 格对拍="
        f"{'PASS' if out.get('old_sales_match') else 'FAIL'}；重复键(我)={out.get('my_dup_row_keys')}"
        f" 与 Lead 一致={out.get('dup_row_keys_match')}")
    return out


def run(dry_run=True, source="auto", use_suggest=True, max_suggest=0, expect_sha=None, log=print):
    t0 = time.time()
    wb_sha = _wb_sha256()
    monthly, src_meta = load_sales(source, log=log)
    matcher = Matcher(monthly, use_suggest=use_suggest, max_suggest=max_suggest, log=log)
    log(f"[match] union series={len(matcher.union)}  name_keys={len(matcher.byname)}  "
        f"2026-09 series={len(matcher.cur)}  multi-sid names={len(matcher.multi_sid_names)}")

    wb, rows, header = read_base_table(WB_PATH, read_only=True)
    led_header, led_nrow = read_ledger_format(wb, read_only=True)
    try:
        wb.close()
    except Exception:                                             # noqa: BLE001
        pass
    if len(rows) != 5392:
        log(f"[note] 底表数据行 = {len(rows)}（≠ 基线 5,392：并发写入方已追加公告行，属预期，"
            f"本脚本按实读行数计算、不硬编码）；sha256={wb_sha}")

    # ---- 通用名称**单元格**级匹配（并列格先按分隔符切候选，再逐候选分层匹配）
    bmodel = defaultdict(Counter)          # 单元格 -> 产品商标计数（取众数做品牌守卫，保证可复现）
    dmodel = defaultdict(set)              # 单元格 -> 行级动力类型（§2.3 第③级消歧）
    for r in rows:
        g = str(r[C_GENERIC] or "").strip()
        bmodel[g][str(r[C_TRADEMARK] or "").strip()] += 1
        dmodel[g].add(row_energy(r))
    gen_info = {}
    for g, tc in sorted(bmodel.items()):
        trademark = sorted(tc.items(), key=lambda kv: (-kv[1], kv[0]))[0][0]
        dtypes = sorted(d for d in dmodel.get(g, ()) if d in ("BEV", "PHEV", "EREV"))
        m = matcher.match_cell(g, trademark, dtypes)
        ch = m["channel"]
        sids = sorted(m["sids"])
        base = {"channel": ch, "cands": m["cands"], "per_cand": m["per_cand"],
                "n_cands": m["n_cands"], "trademark": trademark, "dtypes": dtypes,
                "hit_detail": m.get("hit_detail"), "in_union": m.get("in_union")}
        if ch in ("placeholder", "empty"):
            gen_info[g] = {**base, "sid": None, "sids": [], "note": m["note"]}
        elif ch in ("hold", "multi_series_cell"):
            gen_info[g] = {**base, "sid": None, "sids": sids, "chan_before": m.get("base_channel"),
                           "note": m["note"]}
        elif len(sids) == 1:
            sid = sids[0]
            gen_info[g] = {**base, "sid": sid, "sids": sids, "seriesname": matcher.name_of.get(sid),
                           "note": ""}
        else:
            gen_info[g] = {**base, "channel": "unmatched", "sid": None, "sids": [],
                           "note": m.get("note") or "当期新能源榜内无匹配"}
    matcher.save_suggest_cache()

    # ---- 行级状态
    buckets = Counter()
    status_rows = Counter()
    row_plan = []
    matched_rows = 0
    for idx, r in enumerate(rows):
        g = str(r[C_GENERIC] or "").strip()
        info = gen_info.get(g, {"channel": "empty", "sid": None})
        sid = info.get("sid")
        ch = info.get("channel")
        sub = None
        if ch in ("placeholder", "empty"):
            status = STATUS_NOID
        elif ch in ("hold", "multi_series_cell"):
            status = STATUS_HOLD
        elif sid and sid in matcher.cur:
            status = STATUS_ONSALE if int(matcher.cur[sid]["n"] or 0) > 0 else STATUS_ZERO
        elif sid:
            status = STATUS_GONE
            sub = "anchored_not_current_month"
        else:
            status = STATUS_GONE
            sub = "name_absent_from_union"
        n = int(matcher.cur[sid]["n"] or 0) if (sid and sid in matcher.cur) else None
        if sid and sid in matcher.cur and n and n > 0:
            matched_rows += 1
        buckets[status] += 1
        status_rows[status] += 1
        row_plan.append({"sheet_row": idx + 2, "batch": r[C_BATCH], "code": r[C_CODE],
                         "generic": g, "sid": sid, "status": status, "sub": sub,
                         "sales_new": n if (n and n > 0) else None,
                         "seriesname": matcher.name_of.get(sid) if sid else None,
                         "channel": ch})

    # ---- 覆盖率（§5 三口径）与匹配统计
    nonph = {g: v for g, v in gen_info.items() if v["channel"] not in ("placeholder", "empty")}
    anchored = {g: v for g, v in nonph.items() if v.get("sid")}
    hold = {g: v for g, v in nonph.items() if v["channel"] in ("hold", "multi_series_cell")}
    unmatched = {g: v for g, v in nonph.items() if v["channel"] == "unmatched"}
    n_generic_all = len(gen_info)
    n_generic_noident = n_generic_all - len(nonph)
    n_generic_denom = n_generic_all - n_generic_noident               # §5.2 分母
    row_counts = Counter(str(r[C_GENERIC] or "").strip() for r in rows)
    channel_counts = Counter(v["channel"] for v in gen_info.values())          # 单元格（通用名称）级
    channel_rows = Counter()
    for r in rows:
        channel_rows[gen_info[str(r[C_GENERIC] or "").strip()]["channel"]] += 1
    a2_base = Counter(v.get("chan_before") for v in gen_info.values()
                      if v["channel"] == "A2_split")

    # ---- 车系记录（先建车系、再聚合销量）+ 联动指标
    series_recs = build_series_records(rows, gen_info)
    match_stats = {
        "matched_rows": matched_rows,
        "channel_counts": dict(channel_counts.most_common()),
        "channel_counts_rows": dict(channel_rows.most_common()),
        "a2_base_channels": dict(a2_base.most_common()),
        "generic_all": n_generic_all,
        "generic_noident": n_generic_noident,
        "generic_denom": n_generic_denom,
        "generic_anchored": len(anchored),
        "generic_hold": len(hold),
        "generic_unmatched": len(unmatched),
        "distinct_seriesid_anchored": len({v["sid"] for v in anchored.values()}),
        "union_series": len(matcher.union),
    }
    _hold_sorted = sorted(hold.items(), key=lambda kv: (-row_counts.get(kv[0], 0), kv[0]))
    pending = [[g, row_counts.get(g, 0), (v.get("cands") or [])[:8]] for g, v in _hold_sorted]
    pending_detail = [{"generic": g, "rows": row_counts.get(g, 0), "channel": v["channel"],
                       "candidates": (v.get("cands") or [])[:8], "candidate_sids": v.get("sids", []),
                       "chan_before": v.get("chan_before"), "trademark": v.get("trademark"),
                       "per_cand": v.get("per_cand")}
                      for g, v in _hold_sorted]
    _wb_t = openpyxl.load_workbook(WB_PATH, read_only=True)
    batch_year = read_batch_years(_wb_t)
    batch_month = read_batch_months(_wb_t)
    try:
        _wb_t.close()
    except Exception:                                             # noqa: BLE001
        pass
    log(f"[batch] 批次时间表 {len(batch_year)} 个批次映射年份（如 341→{batch_year.get(341)}）")
    cross_check = load_cross_check(log=log)
    sid_by_boardname = defaultdict(set)
    for x in matcher.cur.values():
        sid_by_boardname[norm(x["name"])].add(x["sid"])
    analytics = build_analytics(rows, gen_info, series_recs, monthly, src_meta, match_stats,
                                pending, dict(status_rows), batch_year, cross_check,
                                dict(sid_by_boardname), log=log)
    analytics["meta"]["pending_naming"] = pending
    analytics["coverage"]["pending_naming"] = pending
    analytics["meta"]["pending_naming_detail"] = pending_detail[:80]
    analytics["meta"]["suggest_decision"] = {
        "enabled": bool(use_suggest),
        "decision": "本轮**未启用**通道 D（汽车之家 suggest），保持 --no-suggest（Lead m00878 裁决 3）",
        "reasons": [
            "AGENTS 已载两次事故模式：2026-09-17 泛化名撞名跨品牌（E5/ET7/S3/M6/G6/A05）、"
            "2026-09-27 代际过滤；suggest 是本项目唯一会引入外部未验证车系名的通道",
            "已有两条可人工裁决通道：pending_naming（33 条）与 "
            "sales_series_anchor_candidates_20261019.json（129 条，逐条证据）",
            "跑 suggest 会改变已验收的全部数字（match_channel_counts、覆盖率三口径、六类状态），"
            "需整体重新验收；收益（最多几十格）不抵风险",
        ],
        "consequence": "未锚定的在榜/待锚定车系一律 HOLD，不做推测性补名（保守设计）",
    }
    analytics["coverage"]["definitions"]["suggest"] = (
        "通道 D（suggest）本轮**未启用**，属保守设计；未锚定的 129 个在榜车系一律 HOLD，"
        "不做推测性补名（Lead m00878 裁决 3）")
    _batches = sorted({_as_int(r[C_BATCH]) for r in rows if _as_int(r[C_BATCH]) is not None})
    _agg_rule_text = ("分类型字段取众数，并列取字典序最小（lexicographic-min）；"
                      "数值型字段取中位数（剔除空值）；空白值归 `(未定)` 而非 null；"
                      "统计单元 = 车系（seriesid 去重）")
    analytics["meta"].update({
        "base_rows": len(rows), "base_table_sha256": wb_sha, "dry_run": bool(dry_run),
        # 写入后表宽 41（39 原列 + 车系ID + 销量覆盖状态）；dry-run 时为原始 39
        "base_table_columns": (NCOLS_NEW if not dry_run else (len(header) if header else None)),
        "base_table_columns_pre": (len(header) if header else None),
        "base_table_sha256_note": "本 JSON 所依据的工作簿 sha256（小写 hex）；apply 需带 --expect-sha 同值",
        "suggest_enabled": bool(use_suggest),
        "suggest_calls": int(getattr(matcher, "_sug_calls", 0) or 0),
        "suggest_note": ("验收基线要求 `--no-suggest`：suggest_enabled=False 且 suggest_calls=0；"
                         "若为 True/非 0 则与 Lead 验收基线（matched_rows=2185）不符"),
        "agg_rule": _agg_rule_text,
        "batch_range": ([_batches[0], _batches[-1]] if _batches else []),
        "batch_count": len(_batches),
        "batch_range_note": "末批 412 为**公示口径**（未正式公告）；交付物封面/说明页引用批次范围时须标注",
    })
    _ca = analytics.get("contract_assertions") or {"all_pass": None, "assertions": []}
    log(f"[assert] contract all_pass={_ca['all_pass']}  ({len(_ca['assertions'])} 条)")
    for a in _ca["assertions"]:
        if not a.get("pass"):
            log(f"  [FAIL] {a.get('name')}  value={str(a.get('value'))[:240]}")

    # ---- 数据质量登记（Lead m00806④）：ERA 行级脏标记——含 EREV 行但车系级标签非 EREV
    # （R2「任一行为增程→整系 EREV」会把这些整系误判；R1 行级众数不受影响）
    erev_anom = []
    for i, r in enumerate(rows):
        if _as_int(r[C_EREV]) != 1:
            continue
        _a = gen_info.get(str(r[C_GENERIC] or "").strip()) or {}
        _rec = series_recs.get(_a.get("sid")) if _a.get("sid") else None
        if not _rec or _rec.get("energy_label") == "EREV":
            continue
        erev_anom.append({"sheet_row": i + 2, "批次": _as_int(r[C_BATCH]),
                          "产品型号": str(r[C_CODE] or "").strip(),
                          "通用名称": str(r[C_GENERIC] or "").strip(),
                          "车型名称": str(r[C_MODEL] or "").strip(),
                          "动力类型": str(r[C_POWER] or "").strip(),
                          "是否增程(EREV)": r[C_EREV],
                          "seriesid": _a.get("sid"), "系列名": _a.get("seriesname"),
                          "车系级标签": _rec.get("energy_label"),
                          "行级标签计数": dict(sorted((_rec.get("energy_label_counts") or {}).items()))})
    erev_anom_summary = {
        "rows": len(erev_anom),
        "distinct_series": len({x["seriesid"] for x in erev_anom if x["seriesid"]}),
        "distinct_series_names": sorted({x["系列名"] for x in erev_anom if x["系列名"]}),
        "rule": ("`是否增程(EREV)=1` 的行，其所属车系按行级标签众数归类**不是** EREV "
                 "→ 单行标记与车系主身份矛盾（列级脏值嫌疑；也是 R2 会误判的集合）"),
        "action": "本轮只登记，不清洗底表（Lead m00806④）",
    }

    # ---- 数据质量登记（Lead m00878 裁决 1/5）：只登记，不改底表
    _seg_ws = Counter()
    _gen_ws = Counter()
    _gen_int = Counter()
    for r in rows:
        _s = r[C_SEGMENT]
        if _s is not None and str(_s) != str(_s).strip():
            _seg_ws[str(_s)] += 1
        _g = r[C_GENERIC]
        if _g is not None and str(_g) != str(_g).strip():
            _gen_ws[str(_g)] += 1
        if str(_g or "").strip().isdigit():
            _gen_int[str(_g).strip()] += 1
    data_quality_registry = {
        "segment_whitespace_dirty": {
            "rows": sum(_seg_ws.values()), "variants": dict(_seg_ws.most_common()),
            "note": "前后空格脏值会凭空造桶（Lead m00878 裁决 1 的 239 辆差异来源）；"
                    "聚合前一律 strip，未 strip 口径作废",
        },
        "generic_name_whitespace_dirty": {
            "rows": sum(_gen_ws.values()), "distinct": len(_gen_ws),
            "variants": dict(_gen_ws.most_common()),
        },
        "generic_name_numeric_only": {
            "rows": sum(_gen_int.values()), "distinct": len(_gen_int),
            "values": dict(_gen_int.most_common()),
            "note": ("裸数字通用名称（001/09/11/009/007/7/6…）正是 AGENTS 2026-09-17 "
                     "跨品牌撞名事故的高危形态；本轮一律按原规则匹配、不特判"),
        },
        "action": "本轮只登记，不清洗底表",
    }

    # ---- 与 Lead 写入前全表指纹对拍（旧值 806 格 + 重复键陷阱）
    lead_fp = load_lead_fingerprint_check(rows, wb_sha, log=log)

    # ---- diff / 变更计划（先算，台账要引用其条数）
    plan = build_change_plan(rows, row_plan, header or [])

    # ---- 台账 JSON
    _cov = analytics["coverage"]
    _cov3 = {"row": _cov["row"]["pct"], "series": _cov["series"]["pct"],
             "current_board": _cov["current_board"]["pct"],
             "current_board_sales": _cov["current_board"]["sales_pct"],
             "self_consistent_ratio": _cov["self_consistent_ratio"]["pct"]}
    ledger = {
        "meta": {"task": "WS2 车系锚定 + 底表销量写入 + 联动指标", "snapshot_id": SNAPSHOT,
                 "captured_at": CAPTURED_AT, "month": CUR_MONTH, "unit": UNIT,
                 "dedup_rule": DEDUP_RULE, "source": SOURCE_TAG,
                 "contract": "Update/销量联动JSON契约_20261019.md v1.1",
                 "sales_source_kind": src_meta.get("source_kind"),
                 "sales_source_path": src_meta.get("source_path"),
                 "base_table_path": os.path.basename(WB_PATH),
                 "base_rows": len(rows), "base_table_sha256": wb_sha,
                 "base_table_columns": (NCOLS_NEW if not dry_run else (len(header) if header else None)),
                 "base_table_columns_pre": (len(header) if header else None),
                 "dry_run": bool(dry_run),
                 "batch_range": ([_batches[0], _batches[-1]] if _batches else []),
                 "batch_count": len(_batches),
                 "batch_range_note": (
                     ("末批 412 为**公示口径**（无正式发文字号，未正式公告）；"
                      "交付物封面/说明页引用批次范围时须标注" if _batches and _batches[-1] == 412
                      else "批次范围取自底表「批次」列实读")),
                 "suggest_enabled": use_suggest, "suggest_calls": matcher._sug_calls,
                 "suggest_failures": matcher._sug_fail[:50],
                 "match_channels": ["A_exact", "B_brand", "C_substr", "D_suggest", "A2_split",
                                    "multi_series_cell", "hold", "placeholder", "empty", "unmatched"],
                 "write_columns": {"35": "月销量(辆)", "36": "销量数据月份", "37": "销量更新日期",
                                   "39": "车系ID", "40": "销量覆盖状态"},
                 "note": ("多候选一律 HOLD（未匹配-命名待锚定），严禁硬填；"
                          "占位符通用名称不参与匹配（无车系身份）；未匹配一律留空，禁止写 0")},
        "summary": {"total_rows": len(rows), "matched_rows": matched_rows,
                    "matched_rows_channel": dict(channel_rows.most_common()),
                    "distinct_generic": len(gen_info),
                    "distinct_seriesid_anchored": match_stats["distinct_seriesid_anchored"],
                    "coverage": _cov3,
                    "status_counts": _cov["status_counts"],
                    "buckets": _cov["buckets"],
                    "status_rows_all": dict(status_rows),
                    "match_channel_counts": dict(channel_counts.most_common()),
                    "a2_base_channels": match_stats["a2_base_channels"],
                    "match_stats": {k: v for k, v in match_stats.items()},
                    "monthly_counts": analytics["meta"]["qc"]["monthly_counts"],
                    "union_series": match_stats["union_series"],
                    "ledger_rows": plan["ledger_rows"],
                    "ledger_per_cell": len(plan["per_cell_changes"]),
                    "ledger_aggregate": len(plan["aggregate_changes"]),
                    "new_col_cells": plan["new_col_cells"],
                    "ledger_rows_canonical": plan["canonical"]["ledger_rows"],
                    "ledger_per_cell_canonical": plan["canonical"]["per_cell"],
                    "ledger_aggregate_canonical": plan["canonical"]["aggregate"],
                    "new_col_cells_canonical": plan["canonical"]["new_col_cells"],
                    "ledger_canonical_note": plan["canonical"]["note"],
                    "ledger_rule": ("列 35/36/37 逐格登记；列 39/40（新增列初始填充）按取值类别聚合登记，"
                                    "逐行取值见 per_generic（Lead m00878 裁决 4）")},
        "per_generic": [{"generic": g, "rows": row_counts.get(g, 0),
                         "channel": v["channel"], "seriesid": v.get("sid"),
                         "seriesname": v.get("seriesname"), "candidates": v.get("cands", [])[:8],
                         "candidate_sids": v.get("sids", []), "trademark": v.get("trademark"),
                         "n_cands": v.get("n_cands"), "chan_before": v.get("chan_before"),
                         "note": v.get("note")}
                        for g, v in sorted(gen_info.items(),
                                           key=lambda kv: (-row_counts.get(kv[0], 0), kv[0]))],
        "pending_naming": pending,
        "pending_naming_detail": pending_detail,
        "unmatched_sample": [g for g, v in sorted(gen_info.items()) if v["channel"] == "unmatched"][:300],
        "placeholder_list": sorted(g for g, v in gen_info.items()
                                   if v["channel"] in ("placeholder", "empty")),
        "dedup_selfcheck": analytics["dedup_selfcheck"],
        "dedup_assertions": analytics["dedup_assertions"],
        "energy_flag_anomalies": erev_anom,
        "energy_flag_anomalies_summary": erev_anom_summary,
        "data_quality_registry": data_quality_registry,
        "lead_prewrite_crosscheck": lead_fp,
    }

    # ---- 写入 JSON（dry-run 与 apply 均写；均为非破坏性产物）
    write_json(LEDGER_OUT, ledger)
    write_json(ANALYTICS_OUT, analytics)
    anchor = build_anchor_candidates(rows, gen_info, monthly, batch_year, batch_month, log=log)
    anchor["meta"].update({"base_rows": len(rows), "base_table_sha256": wb_sha,
                           "dry_run": bool(dry_run)})
    write_json(ANCHOR_OUT, anchor)

    stats = {
        "dry_run": dry_run, "rows": len(rows), "matched_rows": matched_rows,
        "base_rows": len(rows), "base_table_sha256": wb_sha,
        "base_table_columns": (len(header) if header else None),
        "coverage": _cov3, "cov_detail": match_stats,
        "buckets": _cov["buckets"], "status_counts": _cov["status_counts"],
        "channels": dict(channel_counts.most_common()),
        "channels_rows": dict(channel_rows.most_common()),
        "changed_cells": plan["changed_cells"], "changed_by_field": plan["by_field"],
        "ledger_rows": plan["ledger_rows"], "ledger_per_cell": len(plan["per_cell_changes"]),
        "ledger_aggregate": len(plan["aggregate_changes"]), "new_col_cells": plan["new_col_cells"],
        "plan": plan,
    }
    log(f"[plan] 变更格数={plan['changed_cells']}  分列={plan['by_field']}")
    log(f"[ledg] 变更记录将 +{plan['ledger_rows']} 条 = 逐格(35/36/37) {len(plan['per_cell_changes'])} "
        f"+ 聚合(39/40) {len(plan['aggregate_changes'])}；39/40 实写 {plan['new_col_cells']} 格")
    log(f"[cov ] 行级={_cov3['row']}%  车系级={_cov3['series']}%  "
        f"当期在榜车系覆盖率={_cov3['current_board']}%（销量占比={_cov3['current_board_sales']}%）"
        f"  [自指比={_cov3['self_consistent_ratio']}%，非覆盖率]")
    log(f"[stat] 六类={_cov['status_counts']}  和={sum(_cov['status_counts'].values())}")
    log(f"[buck] 四类分桶={_cov['buckets']}")
    log(f"[chan] 单元格级={dict(channel_counts.most_common())}")
    log(f"[chan] 行级={dict(channel_rows.most_common())}")
    log(f"[dedup] all_pass={analytics['dedup_assertions']['all_pass']} "
        f"selfcheck_pass={analytics['dedup_selfcheck']['summary']['cases_pass']}"
        f"/{analytics['dedup_selfcheck']['summary']['cases_checked']}")

    if not dry_run:
        if int(getattr(matcher, "_sug_calls", 0) or 0) > 0:
            log(f"[WARN] 通道D已启用（suggest_calls={matcher._sug_calls}），"
                f"与验收基线不符：验收基线要求 --no-suggest（matched_rows 2185）")
        apply_to_workbook(rows, row_plan, plan, expect_sha=expect_sha, log=log)
        # 写入后工作簿 sha 已变 → 把 meta.base_table_sha256 更新为**落盘后的真实 sha**，
        # 保证「JSON 记录的 sha == 当前工作簿 sha」（Lead m01097 5 项 meta 对齐；消费方据此做一致性断言）
        post_sha = _wb_sha256()
        for _obj in (ledger, analytics):
            _obj["meta"]["base_table_sha256"] = post_sha
            _obj["meta"]["base_table_sha256_prewrite"] = wb_sha
            _obj["meta"]["base_table_sha256_note"] = (
                "**落盘后**工作簿 sha256（小写 hex）；base_table_sha256_prewrite 为本次运行读取时的 sha；"
                "apply 需带 --expect-sha <prewrite 值>")
        write_json(LEDGER_OUT, ledger)
        write_json(ANALYTICS_OUT, analytics)
        log(f"[sha ] 落盘后 sha256={post_sha}（已回写两个 JSON 的 meta.base_table_sha256）")

    stats["elapsed_sec"] = round(time.time() - t0, 1)
    stats.pop("plan", None)
    return stats, plan, ledger, analytics


def write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def build_change_plan(rows, row_plan, header):
    """逐格 diff：仅列 35/36/37/39/40。返回计划与统计。"""
    changes = []
    by_field = Counter()
    for idx, r in enumerate(rows):
        p = row_plan[idx]
        sid = p["sid"]
        new = {
            C_SALES: p["sales_new"] if p["sales_new"] else None,
            C_SMONTH: CUR_MONTH if p["sales_new"] else None,
            C_SDATE: CAPTURED_AT if p["sales_new"] else None,
            C_SERIESID: sid or None,
            C_STATUS: p["status"],
        }
        for ci, nv in new.items():
            ov = r[ci] if ci < len(r) else None
            if _same(ov, nv):
                continue
            changes.append({"row": idx + 2, "col": ci, "old": ov, "new": nv,
                            "field": _col_field(ci), "batch": p["batch"], "code": p["code"],
                            "generic": p["generic"], "sid": sid, "status": p["status"],
                            "channel": p["channel"], "sub": p["sub"]})
            by_field[_col_field(ci)] += 1

    # ---- 台账口径（Lead m00878 裁决 4）：
    #   列 35/36/37 = 逐格登记（真实「旧值→新值」）；列 39/40 = 新增列初始填充，聚合登记。
    per_cell = [c for c in changes if c["col"] in (C_SALES, C_SMONTH, C_SDATE)]
    new_cols = [c for c in changes if c["col"] in (C_SERIESID, C_STATUS)]
    agg = _aggregate_newcol_ledger(new_cols, per_cell, changes)
    # ---- 状态无关的「本应写入」台账口径（幂等重跑时 diff=0，但台账条数仍需可复述）
    n_sales = 0
    for idx, p in enumerate(row_plan):
        prior = None
        if idx < len(rows) and C_SALES < len(rows[idx]):
            prior = rows[idx][C_SALES]
        if p["sales_new"] or prior not in (None, ""):
            n_sales += 1
    n_sid = sum(1 for p in row_plan if p["sid"])
    canonical = {"per_cell": 3 * n_sales,
                 "aggregate": 8 if (n_sales or n_sid) else 0,
                 "new_col_cells": len(row_plan) + n_sid,
                 "ledger_rows": 3 * n_sales + (8 if (n_sales or n_sid) else 0),
                 "note": ("状态无关口径：按 row_plan 应写入的格数/台账条数（首次 apply 时的值）；"
                          "`changed_cells`/`ledger_rows` 是本次运行相对当前工作簿的增量，"
                          "幂等重跑时为 0（Lead m01097 5 项 meta 对齐）")}
    return {"changes": changes, "changed_cells": len(changes), "by_field": dict(by_field.most_common()),
            "per_cell_changes": per_cell, "new_col_cells": len(new_cols),
            "aggregate_changes": agg,
            "canonical": canonical,
            "ledger_rows": len(per_cell) + len(agg)}


def _aggregate_newcol_ledger(new_cols, per_cell, changes):
    """列 39/40（新增列初始填充）的聚合台账（Lead m00878 裁决 4）。

    幂等：仅当 39/40 确有格要写时才登记（重复运行不产生重复台账）。
    """
    if not new_cols:
        return []
    st = Counter(c["new"] for c in new_cols if c["col"] == C_STATUS)
    sid_cells = [c for c in new_cols if c["col"] == C_SERIESID]
    n_sid = sum(1 for c in sid_cells if c["new"] not in (None, ""))
    src_common = (f"{TRANSFER_TAG}；{SOURCE_TAG} {CUR_MONTH}；底表来源=易车202608（旧）→ 汽车之家202609（新）")
    out = []
    for k in STATUS_ENUM:
        out.append({"col": C_STATUS, "field": "销量覆盖状态", "old": None,
                    "new": f"{k} = {int(st.get(k, 0))} 行", "ctype": "新增列-聚合登记",
                    "batch": "（整表）", "code": "（整表）", "channel": "AGG",
                    "src": (f"{src_common}；新增列初始填充按取值类别聚合登记（不逐格）；"
                            f"类别={k}；行数={int(st.get(k, 0))}")})
    out.append({"col": C_SERIESID, "field": "车系ID", "old": None,
                "new": f"已填 {n_sid} 行；未匹配留空 {len(sid_cells) - n_sid} 行",
                "ctype": "新增列-聚合登记", "batch": "（整表）", "code": "（整表）", "channel": "AGG",
                "src": (f"{src_common}；新增列初始填充按覆盖聚合登记（不逐格）；"
                        f"车系ID 为可复现锚定键（seriesid ∈ 2026-09 新能源榜）")})
    out.append({"col": C_STATUS, "field": "销量覆盖状态/车系ID", "old": None,
                "new": (f"结构性说明：39/40 列为新增列初始填充，共 {len(new_cols)} 格；"
                        f"35/36/37 列逐格登记共 {len(per_cell)} 格"),
                "ctype": "结构性说明", "batch": "（整表）", "code": "（整表）", "channel": "AGG",
                "src": ("39/40 列逐行取值见工作簿本体与 audit-output/sales_link_20261019.json"
                        ".per_generic，可逐行复算；逐格台账不增加可核验信息（Lead m00878 裁决 4）；"
                        f"总改动格={len(changes)}，其中 35/36/37={len(per_cell)}（逐格）、"
                        f"39/40={len(new_cols)}（聚合 {len(out) + 1} 条）")})
    return out


def _col_field(ci):
    return {C_SALES: "月销量(辆)", C_SMONTH: "销量数据月份", C_SDATE: "销量更新日期",
            C_SERIESID: "车系ID", C_STATUS: "销量覆盖状态"}[ci]


def _same(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        if isinstance(a, str) and a.strip() == "" and b is None:
            return True
        if isinstance(b, str) and b.strip() == "" and a is None:
            return True
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 1e-9
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        try:
            return abs(float(a) - float(b)) < 1e-9
        except (TypeError, ValueError):
            return False
    return str(a).strip() == str(b).strip()


# --------------------------------------------------------------------------- 底表写入


def export_backup(rows, log=print):
    """旧值留档。**只在首次生成时写**：二次 apply 时工作簿已被写入，
    再生成会把「旧值」覆盖成「新值」（2026-10-19 实际发生过一次，靠 Lead 的
    写入前快照 `_lead_baseline_sales_20261019.json` 重建修正）。强制重写用 LS_FORCE_BACKUP=1。
    """
    if os.path.exists(BACKUP_OUT) and not os.environ.get("LS_FORCE_BACKUP"):
        try:
            _old = json.load(open(BACKUP_OUT, encoding="utf-8"))
            _m = _old.get("meta") or {}
            _n = int(_m.get("rows_with_prior_sales") or (_m.get("counts") or {}).get("prior_nonnull") or 0)
            _canon = bool(_m.get("canonical"))
        except Exception:                                          # noqa: BLE001
            _n, _canon = 0, False
        if _canon:
            log(f"[backup] 已存在 **canonical 留档**（Lead 合并版，prior_nonnull={_n}）→ 只读，直接返回")
            return None
        log(f"[backup] 已存在旧值留档（rows_with_prior_sales={_n}），**保留不覆盖**"
            f"（避免二次写入把旧值覆盖成新值）；强制重写请设 LS_FORCE_BACKUP=1")
        return None
    recs = [{"sheet_row": i + 2, "batch": r[C_BATCH], "code": r[C_CODE],
             "generic": str(r[C_GENERIC] or "").strip(),
             "prior_sales": r[C_SALES] if C_SALES < len(r) else None,
             "prior_month": r[C_SMONTH] if C_SMONTH < len(r) else None,
             "prior_date": r[C_SDATE] if C_SDATE < len(r) else None}
            for i, r in enumerate(rows)]
    n_nonempty = sum(1 for x in recs if x["prior_sales"] not in (None, ""))
    obj = {"meta": {"purpose": "底表「月销量(辆)/销量数据月份/销量更新日期」写入前旧值留档（行序）",
                    "prior_source": "易车车系销量榜（零售量/行业综合销量）2026-08",
                    "prior_month": "2026-08", "prior_update_date": "2026-09-16",
                    "new_source": SOURCE_TAG, "new_month": CUR_MONTH,
                    "captured_at": CAPTURED_AT, "base_table": os.path.basename(WB_PATH),
                    "total_rows": len(recs), "rows_with_prior_sales": n_nonempty,
                    "restore_hint": ("按 sheet_row 逐行回填 prior_sales/prior_month/prior_date "
                                     "（行序 zip 配对，勿按产品型号配对）")},
           "rows": recs}
    write_json(BACKUP_OUT, obj)
    log(f"[backup] {os.path.basename(BACKUP_OUT)}  rows={len(recs)}  "
        f"rows_with_prior_sales={n_nonempty}")
    return obj


def _wb_sha256(path=None):
    """工作簿 sha256（**小写 hex**，与 --expect-sha 比对口径一致）。"""
    h = hashlib.sha256()
    with open(path or WB_PATH, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


ARCHIVE_DIR = os.path.join(HERE, "archive")


def archive_workbook(log=print):
    """落盘前整本备份（Lead m00729②）：archive/v4.9.32_pre-sales_backup_<YYYYMMDD>.xlsx。"""
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    dst = os.path.join(ARCHIVE_DIR, "v4.9.32_pre-sales_backup_%s.xlsx" % time.strftime("%Y%m%d"))
    if os.path.exists(dst):
        base, ext = os.path.splitext(dst)
        n = 2
        while os.path.exists("%s_%d%s" % (base, n, ext)):
            n += 1
        dst = "%s_%d%s" % (base, n, ext)
    shutil.copy2(WB_PATH, dst)
    log(f"[archive] 整本备份 -> archive/{os.path.basename(dst)}  "
        f"sha256={_wb_sha256(dst)[:16]}…")
    return dst


def _cell_key(v):
    """单元格归一（用于整行指纹）：None→''，整数值浮点→int，其余 strip。"""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _row_fp(vals):
    """整行内容指纹（Lead m00925②：键非唯一，必须比对整行内容）。"""
    return hashlib.sha256("\x1f".join(_cell_key(v) for v in vals).encode("utf-8")).hexdigest()


def verify_row_identity(rows, wb_rows, limit=5, log=None):
    """落盘前行身份复核（Lead m00729④ + m00925② 修正）。

    本底表 `(批次, 产品型号)` **不是唯一键**（实测 `(410, None)` 出现 3 次：iX3/VLE/欧拉5，
    产品型号为空）。故规则为：按 `(批次, 产品型号)` 建桶 → 桶内按出现序与工作簿当前顺序逐一配对
    → **并对每行做整行内容指纹比对**；任一不等即返回不符项（不得仅位置对齐或仅键对齐）。
    """
    bad = []
    if len(wb_rows) != len(rows):
        return [{"reason": "行数不一致", "workbook_rows": len(wb_rows), "plan_rows": len(rows)}]

    def _buckets(seqs):
        d = defaultdict(list)
        for i, r in enumerate(seqs):
            d[(_as_int(r[0] if r else None), str((r[1] if r and len(r) > 1 else "") or "").strip())].append(i)
        return d

    ba, bb = _buckets(rows), _buckets(wb_rows)
    if set(ba) != set(bb):
        return [{"reason": "键集合不一致",
                 "only_plan": [list(k) for k in sorted(set(ba) - set(bb))[:5]],
                 "only_workbook": [list(k) for k in sorted(set(bb) - set(ba))[:5]]}]
    dup_keys = 0
    for k in sorted(ba):
        ia, ib = ba[k], bb[k]
        if len(ia) != len(ib):
            bad.append({"reason": "同键行数不一致", "key": list(k),
                        "plan_rows": len(ia), "workbook_rows": len(ib)})
            if len(bad) >= limit:
                return bad
        if len(ia) > 1:
            dup_keys += 1
        for pa, pb in zip(ia, ib):
            fa, fb = _row_fp(rows[pa]), _row_fp(wb_rows[pb])
            if fa != fb:
                diff = next((ci for ci in range(min(len(rows[pa]), len(wb_rows[pb])))
                             if _cell_key(rows[pa][ci]) != _cell_key(wb_rows[pb][ci])), None)
                bad.append({"sheet_row": pb + 2, "key": list(k),
                            "plan_row": pa + 2, "first_diff_col": diff,
                            "plan_value": (_cell_key(rows[pa][diff]) if diff is not None else None),
                            "workbook_value": (_cell_key(wb_rows[pb][diff]) if diff is not None else None),
                            "plan_fp": fa[:16], "workbook_fp": fb[:16]})
                if len(bad) >= limit:
                    return bad
    if log:
        log(f"[guard] 行身份复核：{len(rows)} 行、{len(ba)} 个 (批次,产品型号) 键、"
            f"其中重复键 {dup_keys} 个（桶内按出现序配对 + 整行指纹比对）")
    return bad


def apply_to_workbook(rows, row_plan, plan, expect_sha=None, log=print):
    # ---- 幂等短路：无任何格需改动时直接返回（不备份、不写、不追加台账）
    if not plan["changed_cells"]:
        log("[apply] 无改动格（幂等重跑）：跳过备份/写入/台账追加")
        return True

    # ---- 护具①：sha 守卫（落盘前，小写 hex）
    cur_sha = _wb_sha256()
    if expect_sha and cur_sha != expect_sha.lower():
        raise RuntimeError(f"[FATAL] 工作簿已被并发修改：现 {cur_sha} != 期望 {expect_sha.lower()}；中止，不落盘")
    log(f"[guard] 工作簿 sha256={cur_sha}（与 --expect-sha 一致）")

    # ---- 护具④：行身份复核（防位置式写入错位）
    _wb_chk = openpyxl.load_workbook(WB_PATH, read_only=True, data_only=True)
    _rows_chk = list(_wb_chk[SHEET].iter_rows(min_row=2, values_only=True))
    try:
        _wb_chk.close()
    except Exception:                                             # noqa: BLE001
        pass
    bad_id = verify_row_identity(rows, _rows_chk, log=log)
    if bad_id:
        raise RuntimeError("[FATAL] 行身份复核失败（(批次,产品型号) 桶 + 整行指纹），中止不落盘；前 %d 项：%s"
                           % (len(bad_id), json.dumps(bad_id[:5], ensure_ascii=False)))
    log("[guard] 行身份复核 PASS：全行 (批次,产品型号) 桶内出现序配对 + 整行内容指纹全等")

    # ---- 护具②：整本备份（落盘前）
    archive_workbook(log=log)

    export_backup(rows, log=log)

    wb = openpyxl.load_workbook(WB_PATH)
    ws = wb[SHEET]
    led = wb[LEDGER_SHEET]
    led_header = [c.value for c in led[1]]
    if [str(x).strip() for x in led_header] != ['序号', '批次', '产品型号', '动力类型', '字段',
                                                '变更类型', '附件1原值', '合并值', '数据来源']:
        raise RuntimeError(f"变更记录 列格式与预期不符：{led_header}")
    seq = 0
    for r in led.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
        if isinstance(r[0], (int, float)):
            seq = max(seq, int(r[0]))

    # 表头：尾插 2 列（不得插中间）
    hdr40 = ws.cell(row=1, column=C_SERIESID + 1)
    hdr41 = ws.cell(row=1, column=C_STATUS + 1)
    if hdr40.value in (None, "") or hdr41.value in (None, ""):
        src_style = ws.cell(row=1, column=C_SALES + 1)
        for c, name in ((hdr40, "车系ID"), (hdr41, "销量覆盖状态")):
            c.value = name
            c.font = src_style.font.copy()
            c.fill = src_style.fill.copy()
            c.border = src_style.border.copy()
            c.alignment = src_style.alignment.copy()
    if ws.cell(row=1, column=C_SERIESID + 1).value != "车系ID" or \
       ws.cell(row=1, column=C_STATUS + 1).value != "销量覆盖状态":
        raise RuntimeError("新增列表头写入失败")

    cells_written = 0
    for ch in plan["changes"]:
        rn, ci = ch["row"], ch["col"]
        cell = ws.cell(row=rn, column=ci + 1)
        cell.value = ch["new"]
        if ch["new"] in (None, ""):
            cell.fill = NO_FILL
        else:
            cell.fill = FILL_MEDIA
        cells_written += 1
    # 台账：35/36/37 逐格；39/40 聚合登记（Lead m00878 裁决 4）
    led_rows = []
    for ch in plan["per_cell_changes"]:
        seq += 1
        led_rows.append(_ledger_row(seq, ch))
    for ch in plan["aggregate_changes"]:
        seq += 1
        led_rows.append(_ledger_row(seq, ch))
    if led_rows:
        for row_vals in led_rows:
            led.append(row_vals)

    wb.save(WB_PATH)
    log(f"[apply] 写入格={cells_written}（35/36/37 逐格记台账={len(plan['per_cell_changes'])}；"
        f"39/40 = {plan['new_col_cells']} 格聚合为 {len(plan['aggregate_changes'])} 条）；"
        f"变更记录 +{len(led_rows)}（序号 {seq - len(led_rows) + 1}~{seq}）")
    if cells_written != plan["changed_cells"]:
        raise RuntimeError(f"自校验失败：写入格 {cells_written} != 计划改动格 {plan['changed_cells']}")
    if len(led_rows) != plan["ledger_rows"]:
        raise RuntimeError(f"自校验失败：台账条数 {len(led_rows)} != 计划 {plan['ledger_rows']}"
                           f"（逐格 {len(plan['per_cell_changes'])} + 聚合 {len(plan['aggregate_changes'])}）")
    log("[apply] 自校验 PASS：写入格 == 计划改动格；台账条数 == 逐格(35/36/37) + 聚合(39/40)")

    # 写后复核
    wb2 = openpyxl.load_workbook(WB_PATH, read_only=True, data_only=True)
    ws2 = wb2[SHEET]
    bad = 0
    nrows = 0
    for i, r in enumerate(ws2.iter_rows(min_row=2, values_only=True)):
        nrows += 1
        if len(r) < NCOLS_NEW:
            bad += 1
            continue
        p = row_plan[i]
        if not _same(r[C_SALES] if r[C_SALES] not in ("",) else None, p["sales_new"] or None):
            bad += 1
        if (r[C_SERIESID] or None) != (p["sid"] or None):
            bad += 1
        if r[C_STATUS] != p["status"]:
            bad += 1
    log(f"[verify] rows={nrows}  mismatches={bad}")
    if bad:
        raise RuntimeError(f"写后复核不一致：{bad} 处")
    try:
        wb2.close()
    except Exception:                                             # noqa: BLE001
        pass
    return True


def _ledger_row(seq, ch):
    old = ch["old"]
    new = ch["new"]
    field = ch["field"]
    if ch.get("ctype"):
        ctype = ch["ctype"]
    elif ch["col"] in (C_SERIESID, C_STATUS):
        ctype = "新增列"
    elif old in (None, "") and new not in (None, ""):
        ctype = "补充"
    elif old not in (None, "") and new in (None, ""):
        ctype = "清空"
    else:
        ctype = "口径替换"
    if ch.get("src"):
        src = ch["src"]
    else:
        src = (f"{TRANSFER_TAG}；{SOURCE_TAG} {CUR_MONTH}；底表来源=易车202608（旧）→ 汽车之家202609（新）")
        if field == "月销量(辆)":
            src += f"；车系={ch.get('seriesname') or ''}；车系ID={ch.get('sid') or ''}；匹配通道={ch.get('channel')}"
            if ch.get("sub"):
                src += f"；状态细分={ch['sub']}"
        elif field == "车系ID":
            src += f"；可复现锚定键；匹配通道={ch.get('channel')}"
        elif field == "销量覆盖状态":
            src += f"；枚举见口径规范 §2.3；渠道={ch.get('channel')}"
            if ch.get("sub"):
                src += f"；状态细分={ch['sub']}"
    return [seq, ch.get("batch"), ch.get("code"),
            "", field, ctype,
            "(空)" if old in (None, "") else old,
            "(空)" if new in (None, "") else new,
            src]


# --------------------------------------------------------------------------- CLI


def parse_args(argv):
    ap = argparse.ArgumentParser(
        description="WS2 车系锚定 + 底表销量写入 + 联动指标。"
                    "**默认（不传 --apply）为只读 dry-run：不修改工作簿，只写 audit-output/ 下的 JSON。**")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--dry-run", action="store_true",
                   help="只出统计/JSON，不写底表（默认；不带 --apply 即为此模式）")
    g.add_argument("--apply", action="store_true",
                   help="写底表（必须同时给 --expect-sha；含 sha 守卫/整本备份/行身份复核/台账自校验）")
    ap.add_argument("--expect-sha", default=None,
                    help="--apply 必填：工作簿当前 sha256（小写 hex）。不等即中止，绝不落盘")
    ap.add_argument("--source", default="auto",
                    choices=["auto", "official", "pagecache", "devcache"],
                    help="销量数据源优先级（默认 auto：official → pagecache → devcache）")
    ap.add_argument("--no-suggest", action="store_true", help="跳过汽车之家 suggest 兜底（通道 D）")
    ap.add_argument("--max-suggest", type=int, default=0, help="suggest 调用上限（0=不限，调试用）")
    ap.add_argument("--show-sha", action="store_true", help="只打印当前工作簿 sha256（小写 hex）后退出")
    return ap.parse_args(argv)


def main(argv=None):
    a = parse_args(argv if argv is not None else sys.argv[1:])
    if a.show_sha:
        print(_wb_sha256())
        return 0
    dry = not a.apply
    if not dry and not a.expect_sha:
        print("[FATAL] --apply 必须带 --expect-sha <sha256>（小写 hex）。"
              "当前 sha 可用 `python link_sales.py --show-sha` 取得。")
        return 2
    print("=" * 78)
    print(f"link_sales.py  {'DRY-RUN（只读，不写底表）' if dry else 'APPLY（写底表）'}  "
          f"source={a.source}  suggest={not a.no_suggest}")
    print(f"工作簿 sha256 = {_wb_sha256()}")
    print("=" * 78)
    stats, plan, ledger, analytics = run(dry_run=dry, source=a.source,
                                        use_suggest=not a.no_suggest,
                                        max_suggest=a.max_suggest,
                                        expect_sha=a.expect_sha)
    print("-" * 78)
    print(json.dumps(stats, ensure_ascii=False, indent=1))
    print("-" * 78)
    if dry:
        print("DRY-RUN 完成：未写底表。台账 -> %s" % os.path.basename(LEDGER_OUT))
        print("联动指标 -> %s" % os.path.basename(ANALYTICS_OUT))
        print("样例变更（前 8 条）：")
        for ch in plan["changes"][:8]:
            print("  ", {k: ch[k] for k in ("row", "field", "old", "new", "generic", "sid")})
    else:
        print("APPLY 完成。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
