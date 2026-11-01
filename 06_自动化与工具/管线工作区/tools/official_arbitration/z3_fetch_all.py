# -*- coding: utf-8 -*-
"""_z3_fetch_all.py — 390~403 批公告页 + 附件的健壮抓取（curl + 重试退避）。
   工信部站点不稳定（同一 URL 时快时慢、偶发 502/000），故一律「重试到 200 且体积达标」。
"""
import io
import json
import os
import re
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
CAT = os.path.join("..", "_scripts", "announce_archive", "batch_catalog.json")
RAW = os.path.join("公告批次扫描_411-412_20261017", "raw")
OUT = "audit-output/_z3_manifest.json"
BATCHES = [390, 391, 392, 393, 394, 395, 396, 399, 401, 402, 403]
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36"
LINK = re.compile(r'href="([^"]*(?:filemanager|attach)[^"]*\.(?:doc|docx|pdf))"[^>]*>(.*?)</a>', re.I | re.S)


def curl_get(url, out, min_size, tries=6):
    for i in range(tries):
        p = subprocess.run(["curl.exe", "-sS", "-L", "--max-time", "90", "--http1.1", "-A", UA,
                            "-o", out, "-w", "%{http_code}", url], capture_output=True)
        code = (p.stdout or b"").decode(errors="replace").strip()
        sz = os.path.getsize(out) if os.path.exists(out) else 0
        if code == "200" and sz >= min_size:
            return code, sz, i + 1
        time.sleep(3 + 4 * i)
    return code, sz, tries


def clean(s):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", s or "")).strip()


cat = json.load(io.open(CAT, encoding="utf-8"))
items = cat["items"] if isinstance(cat, dict) and "items" in cat else cat
bd = {int(x["batch"]): x for x in items}

man = {}
for b in BATCHES:
    e = bd.get(b)
    page = os.path.join(RAW, f"{b}_page.html")
    if os.path.exists(page):
        os.remove(page)          # 清掉此前 502 坏页
    code, sz, t = curl_get(e["url"], page, 12000)
    ok = (code == "200" and sz >= 12000)
    print(f"[{b}] 页面 http={code} size={sz:,} 尝试{t}  {'OK' if ok else 'FAIL'}")
    if not ok:
        man[b] = {"page": e["url"], "ok": False}
        continue
    html = io.open(page, encoding="utf-8", errors="replace").read()
    links, seen = [], set()
    for m in LINK.finditer(html):
        u = m.group(1)
        if u.startswith("//"):
            u = "https:" + u
        elif u.startswith("/"):
            u = "https://www.miit.gov.cn" + u
        if u in seen:
            continue
        seen.add(u)
        links.append({"url": u, "text": clean(m.group(2))[:90]})
    rec = {"page": e["url"], "date": e.get("date"), "tax_batch": e.get("tax_batch"), "ok": True, "docs": []}
    for idx, l in enumerate(links, 1):
        dst = os.path.join(RAW, f"{b}_doc{idx}.doc")
        if os.path.exists(dst) and os.path.getsize(dst) > 50000:
            code2, sz2, t2 = "cached", os.path.getsize(dst), 0
        else:
            code2, sz2, t2 = curl_get(l["url"], dst, 20000)
        rec["docs"].append({"idx": idx, "url": l["url"], "text": l["text"],
                            "file": os.path.basename(dst), "size": sz2, "http": code2})
        print(f"     doc{idx} {code2:>6} {sz2:>10,} B  {l['text'][:52]}")
    man[b] = rec

io.open(OUT, "w", encoding="utf-8", newline="\n").write(json.dumps(man, ensure_ascii=False, indent=1))
tot = sum(d["size"] for r in man.values() for d in r.get("docs", []))
print(f"\n清单 → {OUT}；附件合计 {tot:,} B")
