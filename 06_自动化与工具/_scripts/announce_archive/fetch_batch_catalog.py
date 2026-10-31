# -*- coding: utf-8 -*-
"""工信部装备工业发展中心「公告发布」栏目(col1691) 全量批次清单抓取

产出：batch_catalog.json
  每条 = {batch, date, url, title, tax_batch(减免购置税目录批次汉数字), year}
走企业代理（直连工信部超时）。
"""
import json
import os
import re
import ssl
import sys
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "batch_catalog.json")
PROXY = "http://proxy.sgm.shanghaigm.com:8080"
API = ("https://www.miit-eidc.org.cn/module/web/jpage/dataproxy.jsp"
       "?startrecord={}&endrecord={}&perpage=45")
FORM = ("col=1&webid=12&path=https%3A%2F%2Fwww.miit-eidc.org.cn%2F&columnid=1691"
        "&sourceContentType=1&unitid=4638"
        "&webname=%E5%B7%A5%E4%B8%9A%E5%92%8C%E4%BF%A1%E6%81%AF%E5%8C%96"
        "%E9%83%A8%E8%A3%85%E5%A4%87%E5%B7%A5%E4%B8%9A%E5%8F%91%E5%B1%95%E4%B8%AD%E5%BF%83"
        "&permissiontype=0").encode()
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

CTX = ssl._create_unverified_context()
OPENER = urllib.request.build_opener(
    urllib.request.ProxyHandler({"http": PROXY, "https": PROXY}),
    urllib.request.HTTPSHandler(context=CTX))

ENTRY = re.compile(r"href='([^']*?art_1691_(\d+)\.html)' title='([^']+)'")
DATE_URL = re.compile(r"/art/(\d{4})/(\d{1,2})/(\d{1,2})/")
GAZETTE = re.compile(r"第(\d+)批")
TAX = re.compile(r"减免车辆购置税的新能源汽车车型目录》[（(]第([^）)]+)批[）)]")


def han_to_int(s):
    """中文数字转阿拉伯数字，最大到 99"""
    d = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
         "六": 6, "七": 7, "八": 8, "九": 9}
    if s.isdigit():
        return int(s)
    if s == "十":
        return 10
    if "十" in s:
        a, _, b = s.partition("十")
        return (d.get(a, 1) if a else 1) * 10 + (d.get(b, 0) if b else 0)
    return d.get(s, -1)


def fetch(start, end):
    req = urllib.request.Request(API.format(start, end), data=FORM, headers=UA)
    with OPENER.open(req, timeout=45) as r:
        return r.read().decode("utf-8", "replace")


def main():
    records = {}
    total_page = None
    page = 1
    while True:
        start = (page - 1) * 45 + 1
        html = fetch(start, start + 44)
        if total_page is None:
            m = re.search(r"<totalpage>(\d+)</totalpage>", html)
            total_page = int(m.group(1)) if m else 99
            print(f"总页数 {total_page}", file=sys.stderr)
        ents = ENTRY.findall(html)
        if not ents:
            break
        for url, aid, title in ents:
            g = GAZETTE.search(title)
            md = DATE_URL.search(url)
            if not g or not md:
                continue
            no = int(g.group(1))
            t = TAX.search(title)
            han = t.group(1) if t else ""
            # 个别批次一次发布多批目录，如第402批"第二十五、二十六批"
            parts = [p for p in re.split(r"[、,，]", han) if p] if han else []
            rec = {
                "batch": no,
                "date": f"{md.group(1)}-{int(md.group(2)):02d}-{int(md.group(3)):02d}",
                "year": int(md.group(1)),
                "url": url,
                "title": title,
                "tax_han": han,
                "tax_batch": han_to_int(parts[0]) if parts else None,
                "tax_batch_all": [han_to_int(p) for p in parts],
                "is_draft": ("公示" in title or "征求意见" in title),
            }
            if no not in records or (not rec["is_draft"] and records[no]["is_draft"]):
                records[no] = rec
        print(f"page {page}/{total_page}: {len(ents)} 条，累计批次 {len(records)}", file=sys.stderr)
        if page >= total_page:
            break
        page += 1
        time.sleep(0.5)

    out = [records[k] for k in sorted(records)]
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({"generated": time.strftime("%Y-%m-%d %H:%M:%S"),
                   "count": len(out), "items": out}, f, ensure_ascii=False, indent=1)
    print(f"批次总数 {len(out)}，范围 {out[0]['batch']}~{out[-1]['batch']}", file=sys.stderr)
    print(f"写入 {OUT}", file=sys.stderr)


if __name__ == "__main__":
    main()
