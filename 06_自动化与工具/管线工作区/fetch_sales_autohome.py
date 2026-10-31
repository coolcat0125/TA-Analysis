#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fetch_sales_autohome.py — 汽车之家「行业综合销量」车系级月度榜采集器
====================================================================
口径锚点：Update/销量口径规范_20261019.md §1（主源 = 汽车之家行业综合销量）

主宇宙  : typeid=9  & subranktypeid=2305  → 新能源榜  （→ monthly）
参照宇宙: typeid=1  & subranktypeid=1     → 全车型总榜（→ reference_only，不入底表）

月份    : 2025-12 + 2026-01 … 2026-09（10 个月）

产出
----
  audit-output/sales_autohome_series_monthly_20261019.json   主快照（含 meta）
  audit-output/sales_autohome_raw_pages_20261019.json        原始分页归档（逐条全字段，未裁剪）
  _cache_sales_autohome.json                                 原始分页缓存（可复用）

字段保真说明
------------
  接口每条记录含 3 个深层 UI 对象：`rightinfo`(55%) / `shareinfo`(21%) / `rcmlist`(12%)，
  内容为 autohome:// 深链、埋点 eventid、分享图标等**无分析价值**的样板数据（占总体积 88%）。
  主快照保留**全部标量字段 + 从上述对象抽取的文本**，完整未裁剪记录另存 raw 归档，
  两者 seriesid/monthly 一一对应。这样主快照可入仓（几 MB），fidelity 仍可全回溯。

健壮性
------
  * 清空 HTTPS_PROXY/HTTP_PROXY/... 防代理劫持（同 media_fill.py）
  * 重试 ≥3 次 + 指数退避
  * 请求间隔 ≥0.6s（全局限速锁）
  * 翻页到 result.pagecount 为止
  * 每月条数校验落在 [330,460]（新能源榜），不达标该月自动重抓（最多 3 轮）
  * 本地缓存：第二次运行直接复用，不重抓（--force 可强制重抓）

已知坑（Lead 探测）：2026-05 第 3 页曾 WinError 10060 超时，导致该月仅取到 100 条。
本采集器通过「条数区间校验 + 该月整体重抓」闭环兜住。

用法
----
  python fetch_sales_autohome.py              # 抓取（优先缓存）
  python fetch_sales_autohome.py --force      # 强制重抓
  python fetch_sales_autohome.py --no-ref     # 跳过参照宇宙（调试用）
"""

import argparse
import json
import os
import ssl
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime

# ---------------------------------------------------------------- 环境：强制直连
for _k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy", "all_proxy"):
    os.environ.pop(_k, None)

HERE = os.path.dirname(os.path.abspath(__file__))
AUDIT = os.path.join(HERE, "audit-output")
CACHE_FILE = os.path.join(HERE, "_cache_sales_autohome.json")
OUT_FILE = os.path.join(AUDIT, "sales_autohome_series_monthly_20261019.json")
RAW_FILE = os.path.join(AUDIT, "sales_autohome_raw_pages_20261019.json")

# 深层 UI 样板对象（无分析价值，主快照中抽取文本后丢弃；完整记录留 raw 归档）
UI_OBJ_KEYS = ("rightinfo", "shareinfo", "rcmlist")


def slim_record(it):
    """主快照记录：全部标量字段 + 从 UI 样板对象抽取的文本。"""
    out = {}
    ui_texts = []
    for k, v in it.items():
        if k in UI_OBJ_KEYS:
            continue
        out[k] = v
    ri = it.get("rightinfo") or {}
    si = it.get("shareinfo") or {}
    for node in ((ri.get("priceinfo") or {}),):
        if node.get("title"):
            ui_texts.append(node["title"])
    if ri.get("ext"):
        ui_texts.append(str(ri["ext"]))
    for x in (it.get("rcmlist") or []):
        if isinstance(x, dict) and x.get("desc"):
            ui_texts.append(str(x["desc"]))
    # 分享扩展里带日期/榜单参数，属可核对元数据
    if si.get("shareext"):
        ui_texts.append(str(si["shareext"]))
    if ui_texts:
        out["_ui_texts"] = ui_texts
    return out


def _write_json(path, payload, indent=1):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=indent)
    os.replace(tmp, path)
    return os.path.getsize(path)


SN = "20261019"
CAPTURED_AT = "2026-10-19"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")
HEADERS = {
    "User-Agent": UA,
    "Referer": "https://www.autohome.com.cn/rank/",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

URL_TMPL = ("https://car-web-api.autohome.com.cn/car/rank/getList"
            "?from=28&pm=2&pluginversion=11.75.8&model=1&channel=0"
            "&pageindex={page}&pagesize=50&typeid={typeid}"
            "&subranktypeid={subranktypeid}&levelid=0&price=0-9000&date={date}")

MONTHS = ["2025-12"] + [f"2026-{m:02d}" for m in range(1, 10)]

# 懂车帝交叉校验（no captcha；rank_data_type=11 = 月度零售量，与主源同口径族）
DCD_URL = ("https://www.dongchedi.com/motor/pc/car/rank_data?aid=1839&app_name=auto_web_pc"
           "&city_name=%E5%8C%97%E4%BA%AC&count={count}&offset={offset}&month={ym}"
           "&rank_data_type=11&brand_id=0&price=0-9000&nation=0")
DCD_FILE = os.path.join(AUDIT, "_dcd_monthly_cache_20261019.json")

MAIN = {"typeid": 9, "subranktypeid": 2305, "name": "新能源榜"}
REF = {"typeid": 1, "subranktypeid": 1, "name": "全车型总榜"}

# 新能源榜每月条数合理区间（硬校验，见任务已知坑）
MAIN_MIN, MAIN_MAX = 330, 460
# 全车型榜仅作参照，给一个宽松下限
REF_MIN = 200
# 已缓存月份快速判据：条数落在此区间即视为完整来源，可直接复用
RESUME_MIN, RESUME_MAX = 200, 2000

MAX_RETRY = 4          # ≥3 次重试
MIN_INTERVAL = 0.65    # ≥0.6s 限速
MONTH_MAX_ROUNDS = 3   # 该月整体重抓轮次上限

_CTX = ssl.create_default_context()
_OPENER = urllib.request.build_opener(
    urllib.request.ProxyHandler({}),
    urllib.request.HTTPSHandler(context=_CTX),
)
_last_req = [0.0]


def _throttle():
    dt = time.time() - _last_req[0]
    if dt < MIN_INTERVAL:
        time.sleep(MIN_INTERVAL - dt)
    _last_req[0] = time.time()


def fetch_page(page, universe, date, log=None):
    """抓一页，带重试+指数退避。返回 (record_or_None, pagecount, err)"""
    url = URL_TMPL.format(page=page, typeid=universe["typeid"],
                          subranktypeid=universe["subranktypeid"], date=date)
    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        _throttle()
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with _OPENER.open(req, timeout=30) as r:
                raw = r.read().decode("utf-8", "ignore")
            d = json.loads(raw)
            if d.get("returncode") not in (0, None):
                raise RuntimeError(f"returncode={d.get('returncode')} msg={d.get('message')}")
            res = d.get("result") or {}
            lst = res.get("list")
            if not isinstance(lst, list):
                raise RuntimeError("result.list 缺失")
            rec = {"items": lst, "pagecount": int(res.get("pagecount") or 1),
                   "tip": res.get("saleranktip") or "", "scenetitle": res.get("scenetitle") or ""}
            if log is not None:
                log.append({"universe": universe["name"], "date": date, "page": page,
                            "ok": True, "n": len(lst), "attempt": attempt})
            return rec, rec["pagecount"], None
        except Exception as e:                                    # noqa: BLE001
            last_err = f"{type(e).__name__}: {e}"
            if log is not None:
                log.append({"universe": universe["name"], "date": date, "page": page,
                            "ok": False, "attempt": attempt, "error": last_err})
            if attempt < MAX_RETRY:
                time.sleep(min(2 ** (attempt - 1), 8))
    return None, None, last_err


def cached_month(page_cache):
    """若缓存已覆盖该月全部页，返回 (items, meta, n)；否则 None（不触碰网络）。"""
    if not page_cache or "pc" not in page_cache:
        return None
    pc = int(page_cache["pc"])
    items = []
    for p in range(1, pc + 1):
        rec = page_cache.get(str(p))
        if not rec or not isinstance(rec.get("items"), list):
            return None
        items.extend(rec["items"])
    n = len(items)
    if not (RESUME_MIN <= n <= RESUME_MAX):
        return None
    m = {"pagecount": pc, "pages_ok": pc,
         "tip": (page_cache.get("1") or {}).get("tip") or "",
         "scenetitle": (page_cache.get("1") or {}).get("scenetitle") or "",
         "from_cache": True, "fetched_pages": 0}
    return items, m, n


def fetch_month(universe, date, page_cache, log):
    """抓一个月；已有缓存页跳过，只补缺页。返回 (items, meta, err, n_fetched)。"""
    fetched = 0
    hit = cached_month(page_cache)
    if hit is not None:
        return hit[0], hit[1], None, 0

    need = [p for p in range(1, 20)
            if not (page_cache.get(str(p)) or {}).get("items")]
    for p in need:
        rec, _, e = fetch_page(p, universe, date, log)
        if rec is None:
            return None, {"pagecount": p - 1, "pages_ok": p - 1}, f"P{p} {e}", fetched
        page_cache[str(p)] = rec
        page_cache["pc"] = max(int(page_cache.get("pc") or 0), p)
        fetched += 1
        if p == 1 and int(rec.get("pagecount") or 1) < 1:
            break
    pagecount = max(1, int(page_cache.get("pc") or 1))
    items = []
    for p in range(1, pagecount + 1):
        rec = page_cache.get(str(p))
        if not rec or not isinstance(rec.get("items"), list):
            return None, {"pagecount": pagecount, "pages_ok": p - 1}, f"P{p} 缺页", fetched
        items.extend(rec["items"])
    first = page_cache.get("1") or {}
    meta = {"pagecount": pagecount, "pages_ok": pagecount,
            "tip": first.get("tip") or "", "scenetitle": first.get("scenetitle") or "",
            "from_cache": False, "fetched_pages": fetched}
    return items, meta, None, fetched


def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, encoding="utf-8") as f:
                c = json.load(f)
            if isinstance(c, dict) and "main" in c:
                _migrate_cache(c)
                return c
        except Exception as e:                                    # noqa: BLE001
            print(f"[warn] 缓存读取失败，忽略: {e}")
    return {"main": {}, "reference": {}, "log": []}


def _migrate_cache(c):
    """旧缓存无 'pc' 键（页数）→ 由已有页码回填，保证第二次运行零网络。"""
    for universe_key in ("main", "reference"):
        for date, pages in (c.get(universe_key) or {}).items():
            if not isinstance(pages, dict):
                continue
            nums = [int(k) for k in pages if str(k).isdigit()]
            if nums and int(pages.get("pc") or 0) < max(nums):
                pages["pc"] = max(nums)


def save_cache(cache):
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)
    os.replace(tmp, CACHE_FILE)


def collect(universe, cache, cache_key, force=False, min_n=0, max_n=None):
    """带条数校验+整月重抓的月份循环。返回 (monthly_dict, meta_dict)"""
    log = cache.setdefault("log", [])
    stored = cache.setdefault(cache_key, {})
    monthly, meta_out = {}, {}
    for date in MONTHS:
        got = None
        for rnd in range(1, MONTH_MAX_ROUNDS + 1):
            # 第一轮可复用已存分页（断点续传）；重抓轮次一律从零开始
            pages = {} if (force or rnd > 1) else dict(stored.get(date) or {})
            items, m, err, nf = fetch_month(universe, date, pages, log)
            if items is None:
                print(f"  [{universe['name']}] {date} 第{rnd}轮失败: {err}")
                continue
            n = len(items)
            hi = max_n if max_n is not None else 10 ** 9
            if min_n <= n <= hi:
                m["count"] = n
                stored[date] = pages
                # 仅在本轮真的抓过页时才落盘（缓存命中路径不写，否则 10 个月
                # 各写一次 ~47MB 缓存会拖成 100s+）
                if not m.get("from_cache"):
                    save_cache(cache)
                got = (items, m, n, nf)
                break
            print(f"  [{universe['name']}] {date} 第{rnd}轮条数 {n} 越界 "
                  f"[{min_n},{hi if max_n is not None else 'inf'}] → 重抓")
            stored.pop(date, None)
            save_cache(cache)
        if got is None:
            raise RuntimeError(f"{universe['name']} {date} 连续 {MONTH_MAX_ROUNDS} 轮均未达标")
        items, m, n, nf = got
        monthly[date] = items
        meta_out[date] = m
        tag = "缓存" if m.get("from_cache") else f"新抓 {nf} 页"
        print(f"  [{universe['name']}] {date}: {n} 条 / {m['pagecount']} 页  [{tag}]  OK")
    return monthly, meta_out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true", help="忽略缓存强制重抓")
    ap.add_argument("--no-ref", action="store_true", help="跳过参照宇宙")
    args = ap.parse_args()

    os.makedirs(AUDIT, exist_ok=True)
    cache = load_cache()
    if args.force:
        cache = {"main": {}, "reference": {}, "log": []}
        print("[info] --force：忽略并重建缓存")

    print(f"[1/3] 主宇宙 {MAIN['name']}  typeid={MAIN['typeid']}&subranktypeid={MAIN['subranktypeid']}")
    main_monthly, main_meta = collect(MAIN, cache, "main",
                                      force=args.force, min_n=MAIN_MIN, max_n=MAIN_MAX)

    ref_monthly, ref_meta = {}, {}
    if not args.no_ref:
        print(f"[2/3] 参照宇宙 {REF['name']}  typeid={REF['typeid']}&subranktypeid={REF['subranktypeid']}")
        ref_monthly, ref_meta = collect(REF, cache, "reference",
                                        force=args.force, min_n=REF_MIN)
    else:
        print("[2/3] 已跳过参照宇宙 (--no-ref)")

    # 并集车系数（主宇宙 10 个月）
    union = set()
    for _m, rows in main_monthly.items():
        for it in rows:
            union.add(str(it.get("seriesid")))

    payload = {
        "meta": {
            "source": "汽车之家 · 行业综合销量（车系级月度零售量）",
            "source_url_template": URL_TMPL,
            "scope_note": (
                "主宇宙 typeid=9&subranktypeid=2305（新能源榜）= 入库口径，与本项目"
                "「仅 NEV 乘用车」同宇宙；参照宇宙 typeid=1&subranktypeid=1（全车型总榜）"
                "仅用于大盘参照/交叉校验，不入底表主列。单位=辆/月，粒度=车系(series)。"
                "官方口径：* 数据源于行业综合销量，每月10日左右更新"
            ),
            "field_fidelity": (
                "主快照保留接口全部标量字段；接口的 rightinfo/shareinfo/rcmlist 三个深层 UI 样板"
                "对象（深链+埋点，占体积 88%）仅抽取文本进 _ui_texts，完整未裁剪记录见 "
                "audit-output/sales_autohome_raw_pages_20261019.json（同 seriesid/同月份，可全回溯）。"
            ),
            "raw_archive_file": os.path.basename(RAW_FILE),
            "months": MONTHS,
            "captured_at": CAPTURED_AT,
            "snapshot_id": SN,
            "main_universe": {"typeid": MAIN["typeid"], "subranktypeid": MAIN["subranktypeid"],
                              "scope": "primary", "name": MAIN["name"]},
            "reference_universe": {"typeid": REF["typeid"], "subranktypeid": REF["subranktypeid"],
                                   "scope": "reference_only", "name": REF["name"]},
            "count_guard": {"main": [MAIN_MIN, MAIN_MAX], "reference_min": REF_MIN},
            "fetch_params": {"pagesize": 50, "max_retry": MAX_RETRY,
                             "min_interval_sec": MIN_INTERVAL, "month_max_rounds": MONTH_MAX_ROUNDS},
            "per_month": main_meta,
            "reference_per_month": ref_meta,
            "series_union_count": len(union),
            "page_fetch_log": cache.get("log", []),
        },
        "monthly": {m: [slim_record(x) for x in rows] for m, rows in main_monthly.items()},
        "reference_only": {
            "全车型榜": {
                "clear_scope": "reference_only",
                "scope_note": "全车型销量总榜，仅供大盘参照/交叉校验，禁止进入底表主列",
                "universe": {"typeid": REF["typeid"], "subranktypeid": REF["subranktypeid"]},
                "monthly": {m: [slim_record(x) for x in rows] for m, rows in ref_monthly.items()},
                "per_month": ref_meta,
            }
        },
    }

    # 原始归档：逐条完整字段（供保真回溯）
    raw_payload = {
        "meta": {"snapshot_id": SN, "captured_at": CAPTURED_AT, "months": MONTHS,
                 "note": "接口原始返回，未裁剪；与 sales_autohome_series_monthly_*.json 按 seriesid 对齐",
                 "main_universe": MAIN, "reference_universe": REF},
        "monthly": main_monthly,
        "reference_only": {"全车型榜": {"monthly": ref_monthly}},
    }

    sz = _write_json(OUT_FILE, payload)
    szr = _write_json(RAW_FILE, raw_payload, indent=None)
    print(f"[3/3] 主快照 {OUT_FILE}  ({sz/1024/1024:.2f} MB)")
    print(f"      原始归档 {RAW_FILE}  ({szr/1024/1024:.2f} MB)")
    for date in MONTHS:
        print(f"    {date}: {main_meta[date]['count']:>4} 条")
    print(f"    并集车系数 = {len(union)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
