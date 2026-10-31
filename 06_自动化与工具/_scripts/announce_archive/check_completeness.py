# -*- coding: utf-8 -*-
"""归档完整性核查：对比官方批次目录与磁盘实有文件，输出缺口清单与统计表

用法：python check_completeness.py
产出：00_索引/完整性核查.md
"""
import json
import os
import re
import ssl
import sys
import urllib.request
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", "07_公告与减免购置税"))
CAT = os.path.join(HERE, "batch_catalog.json")
KIND_DIR = {"公告": "01_公告", "减免购置税": "02_减免购置税", "车船税": "03_车船税"}
OLE2 = bytes.fromhex("D0CF11E0A1B11AE1")
ZIP = b"PK\x03\x04"
PROXY = "http://proxy.sgm.shanghaigm.com:8080"
ATTACH = re.compile(r"<a[^>]+href=[\"']([^\"']+\.(?:doc|docx|pdf|xls|xlsx|zip|rar))[\"'][^>]*>(.*?)</a>", re.I | re.S)


def verify(p):
    """按扩展名校验文件头，返回 True/False"""
    ext = os.path.splitext(p)[1].lower()
    with open(p, "rb") as f:
        head = f.read(8)
    if ext in (".docx", ".xlsx", ".zip"):
        return head[:4] == ZIP
    if ext == ".pdf":
        return head[:4] == b"%PDF"
    # .doc/.xls/.wps 等均为 OLE2 复合文档
    return head == OLE2


def scan():
    """返回 {(批次,类别): [文件名...]}"""
    out = defaultdict(list)
    for kind, d in KIND_DIR.items():
        base = os.path.join(ROOT, d)
        if not os.path.isdir(base):
            continue
        for year in os.listdir(base):
            yp = os.path.join(base, year)
            if not os.path.isdir(yp):
                continue
            for f in os.listdir(yp):
                m = re.match(r"第(\d+)批_", f)
                if m:
                    out[(int(m.group(1)), kind)].append(f)
    return out


def main():
    cat = {x["batch"]: x for x in json.load(open(CAT, encoding="utf-8"))["items"]}
    have = scan()

    batch_no = sorted({k[0] for k in have})
    cnt = Counter(k[1] for k in have for _ in have[k])
    size = 0
    bad = []
    for kind, d in KIND_DIR.items():
        base = os.path.join(ROOT, d)
        if not os.path.isdir(base):
            continue
        for year in os.listdir(base):
            yp = os.path.join(base, year)
            if not os.path.isdir(yp):
                continue
            for f in os.listdir(yp):
                p = os.path.join(yp, f)
                if not os.path.isfile(p) or f.endswith(".csv"):
                    continue
                size += os.path.getsize(p)
                if not verify(p):
                    bad.append(p)

    lines = ["# 归档完整性核查", "",
             f"核查时间：{__import__('time').strftime('%Y-%m-%d %H:%M')}", ""]

    lines += ["## 一、规模", "",
              "| 项目 | 数值 |", "|---|---|",
              f"| 官方批次总数 | {len(cat)}（第{min(cat)}至第{max(cat)}批）|",
              f"| 已归档批次数 | {len(batch_no)} |",
              f"| 文件总数 | {sum(cnt.values())} |",
              f"| 总体积 | {size/1073741824:.2f} GB |",
              f"| 格式校验异常 | {len(bad)} |", ""]

    lines += ["## 二、分类持有", "", "| 类别 | 文件数 |", "|---|---|"]
    for k in KIND_DIR:
        lines.append(f"| {k} | {cnt.get(k,0)} |")
    lines.append("")

    # 按年代统计批次覆盖
    lines += ["## 三、批次覆盖（按发布年份）", "",
              "| 年份 | 官方批次数 | 公告 | 购置税目录 | 车船税目录 |", "|---|---|---|---|---|"]
    by_year = defaultdict(list)
    for no, c in cat.items():
        by_year[c["year"]].append(no)
    for y in sorted(by_year):
        ns = by_year[y]
        a = sum(1 for n in ns if (n, "公告") in have)
        b = sum(1 for n in ns if (n, "减免购置税") in have)
        d = sum(1 for n in ns if (n, "车船税") in have)
        lines.append(f"| {y} | {len(ns)} | {a} | {b} | {d} |")
    lines.append("")

    # 缺口
    miss_gg = [n for n in sorted(cat) if (n, "公告") not in have]
    lines += ["## 四、缺口", "", f"**缺公告的批次（{len(miss_gg)}）**：", ""]
    lines.append("、".join(f"第{n}批" for n in miss_gg) if miss_gg else "无")
    lines.append("")

    if bad:
        lines += ["## 五、校验异常文件", ""]
        for p in bad:
            lines.append(f"- {os.path.relpath(p, ROOT)}")
    else:
        lines += ["## 五、校验", "", f"全部 {sum(cnt.values())} 个文件通过 OLE2 文件头校验。", ""]

    out = os.path.join(ROOT, "00_索引", "完整性核查.md")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("\n".join(lines))
    print(f"\n-> {out}")


if __name__ == "__main__":
    main()
