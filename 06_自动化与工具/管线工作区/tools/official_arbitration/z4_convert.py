# -*- coding: utf-8 -*-
"""_z4_convert.py — 把已下载的税目录 doc 转 UTF-8 txt（**先 doc3 后 doc2**，超大件按需）。
   doc3=减免车辆购置税目录（含完整车辆型号+产品名称+续航/整备/电池），体积小、优先级最高；
   doc2=车船税目录（同样含完整型号，个别批次 30MB，最后处理）。
"""
import glob
import os
import subprocess
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
RAW = os.path.join("公告批次扫描_411-412_20261017", "raw")
TXT = os.path.join(RAW, "txt")
os.makedirs(TXT, exist_ok=True)
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
MAXSZ = int(sys.argv[1]) if len(sys.argv) > 1 else 10_000_000

order = sorted(glob.glob(os.path.join(RAW, "*_doc3.doc")))
order += sorted(f for f in glob.glob(os.path.join(RAW, "*_doc2.doc")) if os.path.getsize(f) <= MAXSZ)
print(f"待转 {len(order)} 个（doc3 优先；doc2 仅 ≤{MAXSZ/1e6:.0f}MB）")
t0 = time.time()
for f in order:
    base = os.path.splitext(os.path.basename(f))[0]
    out = os.path.join(TXT, base + ".txt")
    if os.path.exists(out) and os.path.getsize(out) > 2000:
        print(f"  [skip] {base} {os.path.getsize(out):,} B")
        continue
    st = time.time()
    try:
        subprocess.run([SOFFICE, "--headless", "--convert-to", "txt:Text (encoded):UTF8",
                        "--outdir", TXT, f], capture_output=True, timeout=1800)
    except subprocess.TimeoutExpired:
        print(f"  [TIMEOUT] {base}")
        continue
    sz = os.path.getsize(out) if os.path.exists(out) else 0
    print(f"  [{'OK' if sz > 2000 else 'FAIL'}] {base}  {os.path.getsize(f):,} → {sz:,} B  {time.time()-st:.1f}s", flush=True)
print(f"\n合计 {time.time()-t0:.1f}s")
