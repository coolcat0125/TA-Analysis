# -*- coding: utf-8 -*-
"""生成归档索引：00_索引/批次总览.csv + 文件清单.csv（扫描磁盘实际文件）"""
import csv
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", "..", "07_公告与减免购置税"))
CAT = os.path.join(HERE, "batch_catalog.json")
KIND_DIR = {"01_公告": "公告", "02_减免购置税": "减免购置税", "03_车船税": "车船税"}


def scan():
    out = []
    for d, kind in KIND_DIR.items():
        base = os.path.join(ROOT, d)
        if not os.path.isdir(base):
            continue
        for year in sorted(os.listdir(base)):
            yp = os.path.join(base, year)
            if not os.path.isdir(yp):
                continue
            for f in sorted(os.listdir(yp)):
                p = os.path.join(yp, f)
                if not os.path.isfile(p):
                    continue
                out.append({
                    "年份": year, "类别": kind, "文件名": f,
                    "批次": f.split("_")[0].replace("第", "").replace("批", ""),
                    "发布日期": f.split("_")[1] if len(f.split("_")) > 1 else "",
                    "大小MB": round(os.path.getsize(p) / 1048576, 2),
                    "相对路径": os.path.relpath(p, ROOT),
                })
    return out


def main():
    cat = {x["batch"]: x for x in json.load(open(CAT, encoding="utf-8"))["items"]}
    files = scan()
    os.makedirs(os.path.join(ROOT, "00_索引"), exist_ok=True)

    # 1. 文件清单
    fp = os.path.join(ROOT, "00_索引", "文件清单.csv")
    with open(fp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["批次", "发布日期", "年份", "类别", "文件名", "大小MB", "相对路径"])
        w.writeheader()
        for r in sorted(files, key=lambda x: (int(x["批次"]), x["类别"])):
            w.writerow({k: r[k] for k in w.fieldnames})

    # 2. 批次总览（以官方目录为准，标注本地持有情况）
    tp = os.path.join(ROOT, "00_索引", "批次总览.csv")
    have = {}
    for r in files:
        have.setdefault(int(r["批次"]), {})[r["类别"]] = r["文件名"]
    with open(tp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["批次", "发布日期", "发布年份", "购置税目录批次",
                    "公告文件", "减免购置税文件", "车船税目录文件", "官方页面"])
        for no in sorted(cat):
            c = cat[no]
            w.writerow([no, c["date"], c["year"], c["tax_han"],
                        have.get(no, {}).get("公告", ""),
                        have.get(no, {}).get("减免购置税", ""),
                        have.get(no, {}).get("车船税", ""), c["url"]])

    print(f"文件清单 {len(files)} 条 -> {fp}")
    print(f"批次总览 {len(cat)} 批 -> {tp}")
    print(f"已归档批次 {len(have)} 个：{min(have)}~{max(have)}" if have else "无文件")


if __name__ == "__main__":
    main()
