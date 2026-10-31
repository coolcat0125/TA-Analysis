# -*- coding: utf-8 -*-
"""按批次下载工信部公告发布件：①道路机动车辆生产企业及产品公告 ②减免车辆购置税目录

目录结构（在归档根目录下）：
  01_公告/<年>/第XXX批_YYYYMMDD_道路机动车辆生产企业及产品（第XXX批）.doc
  02_减免购置税/<年>/第XXX批_YYYYMMDD_减免车辆购置税的新能源汽车车型目录（第N批）.doc

用法：
  python download_batches.py 403-411        # 指定区间
  python download_batches.py 403 404 405    # 指定批次
  python download_batches.py --year 2026     # 按发布年
  python download_batches.py 403-411 --kinds 公告,减免购置税,车船税
"""
import argparse
import csv
import json
import os
import re
import ssl
import sys
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", "07_公告与减免购置税"))
CATALOG = os.path.join(HERE, "batch_catalog.json")
PROXY = "http://proxy.sgm.shanghaigm.com:8080"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

CTX = ssl._create_unverified_context()
OPENER = urllib.request.build_opener(
    urllib.request.ProxyHandler({"http": PROXY, "https": PROXY}),
    urllib.request.HTTPSHandler(context=CTX))

KIND_DIR = {"公告": "01_公告", "减免购置税": "02_减免购置税", "车船税": "03_车船税"}
ATTACH = re.compile(r"<a[^>]+href=[\"']([^\"']+\.(?:doc|docx|pdf|xls|xlsx|wps|wpt|et|ett|zip|rar))[\"'][^>]*>(.*?)</a>", re.I | re.S)


def classify(name):
    """兼容各年代附件命名：
    公告：2018 年前为「车辆生产企业及产品」，之后为「道路机动车辆生产企业及产品」
    购置税目录：2014-2023 为「免征车辆购置税」，2024 起为「减免车辆购置税」
    车船税目录：早期「节约能源 使用新能源汽车车型目录」，后加「享受车船税减免优惠的」
    """
    n = name.replace(" ", "").replace("\u3000", "")
    if "减免车辆购置税" in n or "免征车辆购置税" in n:
        return "减免购置税"
    if "车船税" in n or ("节约能源" in n and "车型目录" in n):
        return "车船税"
    # 第331-333批把公告与推荐目录合并成一个附件，名称含"目录"仍属公告
    if "生产企业及产品" in n:
        return "公告"
    return None


def clean(s):
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("&nbsp;", " ").replace("\u200b", "")
    s = re.sub(r"^\s*(?:附件\s*[:：]\s*)?\d+[.、]\s*", "", s)   # 去 "1." 序号
    s = re.sub(r"^\s*附件\s*[:：]\s*", "", s)                   # 去 "附件：" 前缀
    s = re.sub(r"[\\/:*?\"<>|]", "_", s)                        # Windows 非法字符
    s = re.sub(r"\s+", " ", s).strip()
    return s


ALLOWED_EXT = {".doc", ".docx", ".pdf", ".xls", ".xlsx", ".wps", ".wpt", ".et", ".ett", ".zip", ".rar"}


def ext_of(href):
    """真实扩展名：优先取下载接口的 filename 参数，其次取路径后缀"""
    p = urllib.parse.urlparse(href)
    q = urllib.parse.parse_qs(p.query)
    for k in ("filename", "fileName", "name", "file"):
        if k in q:
            e = os.path.splitext(q[k][0])[1].lower()
            if e in ALLOWED_EXT:
                return e
    e = os.path.splitext(p.path)[1].lower()
    return e if e in ALLOWED_EXT else ".doc"


def fetch_page(url, retries=3):
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with OPENER.open(req, timeout=45) as r:
                return r.read().decode("utf-8", "replace")
        except Exception:
            if i == retries - 1:
                raise
            time.sleep(2)


def download(url, dest, retries=2):
    # 老附件在老站可能只存在于其中一侧，404 时换域名再试一次
    cands = [url]
    if "www.miit-eidc.org.cn" in url:
        cands.append(url.replace("www.miit-eidc.org.cn", "www.miit.gov.cn"))
    elif "www.miit.gov.cn" in url and "/cms_files/" not in url:
        cands.append(url.replace("www.miit.gov.cn", "www.miit-eidc.org.cn"))
    last = None
    for cu in cands:
        # 附件名可能含中文，直接塞进请求头会触发 ascii 编码错误
        cu = urllib.parse.quote(cu, safe=":/?&=#%+,;@!$()*[]~'")
        for i in range(retries):
            try:
                req = urllib.request.Request(cu, headers=UA)
                with OPENER.open(req, timeout=180) as r:
                    data = r.read()
                if len(data) < 512:
                    raise ValueError(f"响应过小 {len(data)}B")
                tmp = dest + ".part"
                with open(tmp, "wb") as f:
                    f.write(data)
                os.replace(tmp, dest)
                return len(data)
            except Exception as e:
                last = e
                if i < retries - 1:
                    print(f"    重试 {i+1}/{retries}: {e}", file=sys.stderr)
                    time.sleep(3)
    raise last


def parse_batches(argv):
    out = []
    for a in argv:
        if re.fullmatch(r"\d+-\d+", a):
            s, e = a.split("-")
            out.extend(range(int(s), int(e) + 1))
        elif a.isdigit():
            out.append(int(a))
        else:
            sys.exit(f"无法识别的批次参数：{a}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("batches", nargs="*", help="批次区间 403-411 或单批 403 404")
    ap.add_argument("--year", type=int, help="按发布年份筛选")
    ap.add_argument("--kinds", default="公告,减免购置税", help="公告,减免购置税,车船税")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    kinds = [k.strip() for k in args.kinds.split(",") if k.strip()]
    cat = json.load(open(CATALOG, encoding="utf-8"))["items"]
    by_no = {x["batch"]: x for x in cat}

    if args.year:
        targets = [x["batch"] for x in cat if x["year"] == args.year]
    elif args.batches:
        targets = parse_batches(args.batches)
    else:
        sys.exit("请给批次区间或 --year")

    os.makedirs(os.path.join(ROOT, "00_索引"), exist_ok=True)
    log_rows = []
    other_rows = []
    for no in targets:
        rec = by_no.get(no)
        if not rec:
            print(f"[{no}] 目录中无此批次，跳过", file=sys.stderr)
            continue
        year, date = rec["year"], rec["date"].replace("-", "")
        print(f"[{no}] {rec['date']} 抓取页面…", file=sys.stderr)
        try:
            html = fetch_page(rec["url"])
        except Exception as e:
            print(f"[{no}] 页面失败 {e}", file=sys.stderr)
            log_rows.append([no, rec["date"], "页面失败", str(e)[:80], "", ""])
            continue

        items = []
        other = []
        for href, text in ATTACH.findall(html):
            name = clean(text)
            if not name:                      # 空锚点（冗余链接）直接跳过
                continue
            kind = classify(name)
            if not kind:
                other.append((name, href))
                continue
            if kind not in kinds:
                other.append((f"[{kind}未选]{name}", href))
                continue
            if href.startswith("//"):
                href = "https:" + href
            elif href.startswith("/"):
                # 相对路径相对的是页面所在站点(eidc)，不是 miit.gov.cn
                href = "https://www.miit-eidc.org.cn" + href
            items.append((kind, name, href))

        for name, href in other:
            other_rows.append([no, rec["date"], name, href])

        if not items:
            print(f"[{no}] 未解析到目标附件", file=sys.stderr)
            log_rows.append([no, rec["date"], "无附件", "", "", ""])
            continue

        for kind, name, href in items:
            d = os.path.join(ROOT, KIND_DIR[kind], str(year))
            os.makedirs(d, exist_ok=True)
            ext = ext_of(href)
            base = name
            while True:                             # 剥离标题里自带的扩展名，避免 .doc.docx
                e2 = os.path.splitext(base)[1].lower()
                if e2 in ALLOWED_EXT:
                    base = base[: -len(e2)]
                else:
                    break
            fname = f"第{no}批_{date}_{base}{ext}"
            dest = os.path.join(d, fname)
            if os.path.exists(dest) and os.path.getsize(dest) > 1024:
                print(f"    已存在 {fname} ({os.path.getsize(dest)/1048576:.1f} MB)", file=sys.stderr)
                log_rows.append([no, rec["date"], kind, fname, os.path.getsize(dest), "跳过(已存在)"])
                continue
            if args.dry_run:
                print(f"    [dry] {fname} <- {href}", file=sys.stderr)
                continue
            try:
                n = download(href, dest)
                print(f"    下载 {fname} ({n/1048576:.1f} MB)", file=sys.stderr)
                log_rows.append([no, rec["date"], kind, fname, n, "成功"])
            except Exception as e:
                print(f"    失败 {fname}: {e}", file=sys.stderr)
                log_rows.append([no, rec["date"], kind, fname, 0, f"失败 {str(e)[:60]}"])
            time.sleep(0.8)

    if log_rows and not args.dry_run:
        lg = os.path.join(ROOT, "00_索引", "下载日志.csv")
        new = not os.path.exists(lg)
        with open(lg, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["批次", "发布日期", "类别", "文件名", "字节数", "状态"])
            w.writerows(log_rows)
        print(f"日志追加 {len(log_rows)} 行 -> {lg}", file=sys.stderr)

    # 未归档附件（推荐目录、参数表、非本次类别等），供事后核查
    if other_rows:
        op = os.path.join(ROOT, "00_索引", "未归档附件.csv")
        new = not os.path.exists(op)
        with open(op, "a", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["批次", "发布日期", "附件名", "下载地址"])
            w.writerows(other_rows)
        print(f"未归档附件记录 {len(other_rows)} 条 -> {op}", file=sys.stderr)


if __name__ == "__main__":
    main()
