# -*- coding: utf-8 -*-
"""extract_wps_doc.py · 工信部 WPS 壳 .doc 附件的文本恢复工具（v4.9.44 沉淀）
================================================================================
背景
----
工信部公告附件里有少数文件是「WPS 壳 + Word97 内核」：OLE 容器内同时存在
`WpsCustomData` 流与 `WordDocument` 流。这类文件 libr soffice 的常规转换路径
（txt / html / --infilter="MS Word 97" / 独立 profile）会长时间无产出后失败
（实测 403_big.doc 30.8MB，累计 >2,530 秒 CPU 仍失败），但文件本身并不损坏。

对策
----
绕过转换器，直接读 OLE：
  1. olefile 取 `WordDocument` 流；
  2. 依 FIB 的 `fcMin(0x18) / fcMac(0x1C)`（小端 uint32）截取正文区；
  3. 用 `utf-16-le` 解码（WPS 中文正文为 UTF-16LE；用 gb18030 会丢掉约 2/3 的汉字）；
  4. 按 `0x07`（Word 表格单元分隔符）切格，按 `0x0D` 切段落。

用法
----
  python extract_wps_doc.py <input.doc> [--out <dir>] [--records]

  --out      输出目录，默认与输入同目录下的 txt/ ；
  --records  额外按「车辆型号」形态重建结构化记录（车船税/购置税目录专用），
             输出 <input>_records.json。

产物
----
  <stem>_cells.txt    每行一个单元格（便于人工/正则复核）
  <stem>_text.txt     还原的段落行（0x07→制表符）
  <stem>_records.json 仅在 --records 时产出

说明
----
本工具只读原文件，不修改、不移动、不删除；不写底表。
单元格切分保留了原表格的列序，但**不做列语义映射**——列语义随目录（车船税/购置税）
与批次格式而变，必须由调用方按表头判断，切勿硬编码列下标。
"""
import argparse
import json
import os
import re
import struct
import sys

try:
    import olefile
except ImportError:  # pragma: no cover
    sys.exit("需要 olefile：pip install olefile")


def read_word_text(path):
    """返回 (text, meta)。text 为正文原始文本（未切格）。"""
    ole = olefile.OleFileIO(path)
    try:
        streams = ["/".join(s) for s in ole.listdir()]
        if not ole.exists("WordDocument"):
            raise RuntimeError("OLE 中无 WordDocument 流：%s" % streams)
        wd = ole.openstream("WordDocument").read()
    finally:
        ole.close()

    fc_min, fc_mac = struct.unpack_from("<II", wd, 0x18)
    if not (0 < fc_min < fc_mac <= len(wd)):
        # FIB 不合常规时退化为全文扫描，交由编码打分决定
        fc_min, fc_mac = 0x800, len(wd)
    raw = wd[fc_min:fc_mac]

    # 编码不必猜：按汉字数量打分，选最优
    best = None
    scores = {}
    for enc in ("utf-16-le", "gb18030", "cp936", "latin-1"):
        try:
            s = raw.decode(enc, "ignore")
        except Exception:
            continue
        cjk = sum(1 for ch in s if "\u4e00" <= ch <= "\u9fff")
        scores[enc] = cjk
        if best is None or cjk > best[0]:
            best = (cjk, enc, s)

    if best is None:
        raise RuntimeError("无法解码正文")
    cjk, enc, text = best
    meta = {"worddocument_bytes": len(wd), "fc_min": fc_min, "fc_mac": fc_mac,
            "encoding": enc, "cjk": cjk, "encoding_scores": scores, "streams": streams}
    return text, meta


def split_cells(text):
    """按 0x07 切单元格；段落符 0x0D 视为行边界，统一替换为空格。"""
    cells = [c.replace("\r", " ").replace("\x0b", " ").strip() for c in text.split("\x07")]
    return [c for c in cells if c != ""]


def rebuild_records(cells, code_re=None):
    """按「车辆型号」形态重建记录（车船税/购置税目录）。

    判据：某单元格形如车型号（默认 2 个大写字母 + 3 位以上数字 + 可选后缀），
    则以它为中心，向前取 企业名称/商标，向后取连续数值型字段。
    该重建只保证「型号 + 邻近字段」的对应关系，字段语义仍须按表头判定。
    """
    code_re = code_re or re.compile(r"^[A-Z]{2}\d{3,}[A-Z0-9]*(-[A-Z0-9]+)?$")
    num_re = re.compile(r"^\d+(\.\d+)?([±~-]\d+)?$")
    recs = []
    i = 0
    while i < len(cells):
        c = cells[i]
        if not code_re.match(c):
            i += 1
            continue
        oem = cells[i - 1] if i >= 1 else ""
        brand = cells[i - 2] if i >= 2 else ""
        rest = []
        j = i + 1
        while j < len(cells) and len(rest) < 10:
            if code_re.match(cells[j]):
                break
            rest.append(cells[j])
            j += 1
        recs.append({"code": c, "oem": oem, "brand": brand, "fields": rest})
        i = j
    return recs


def main(argv=None):
    ap = argparse.ArgumentParser(description="恢复 WPS 壳 Word97 .doc 的正文与表格单元格")
    ap.add_argument("input")
    ap.add_argument("--out", default=None)
    ap.add_argument("--records", action="store_true")
    a = ap.parse_args(argv)

    src = os.path.abspath(a.input)
    if not os.path.isfile(src):
        sys.exit("文件不存在：%s" % src)
    out_dir = a.out or os.path.join(os.path.dirname(src), "txt")
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(src))[0]

    print("输入 %s (%s B)" % (src, format(os.path.getsize(src), ",")))
    text, meta = read_word_text(src)
    print("WordDocument %s B  fcMin=%d fcMac=%d  编码=%s  CJK=%s"
          % (format(meta["worddocument_bytes"], ","), meta["fc_min"], meta["fc_mac"],
             meta["encoding"], format(meta["cjk"], ",")))
    print("编码打分：" + " ".join("%s=%s" % (k, format(v, ",")) for k, v in meta["encoding_scores"].items()))

    cells = split_cells(text)
    lines = []
    buf = []
    for ch in text:
        if ch == "\x07":
            buf.append("\t")
        elif ch == "\r":
            ln = "".join(buf).strip("\t\r\n ")
            if ln:
                lines.append(ln)
            buf = []
        elif ch >= " " or ch == "\t":
            buf.append(ch)
    if buf:
        ln = "".join(buf).strip()
        if ln:
            lines.append(ln)

    p_cells = os.path.join(out_dir, stem + "_cells.txt")
    p_text = os.path.join(out_dir, stem + "_text.txt")
    with open(p_cells, "w", encoding="utf-8") as f:
        f.write("\n".join(cells))
    with open(p_text, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("单元格 %s 个 → %s" % (format(len(cells), ","), p_cells))
    print("段落行 %s 行 → %s" % (format(len(lines), ","), p_text))

    if a.records:
        recs = rebuild_records(cells)
        p_rec = os.path.join(out_dir, stem + "_records.json")
        json.dump({"meta": meta, "count": len(recs), "records": recs},
                  open(p_rec, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("结构化记录 %s 条 → %s" % (format(len(recs), ","), p_rec))
        for r in recs[:5]:
            print("   ", r["oem"], "|", r["brand"], "|", r["code"], "|", r["fields"][:6])
    print("完成（未修改原文件）")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
