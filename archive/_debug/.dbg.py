# -*- coding: utf-8 -*-
import re, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
p = r"D:\#AI\00 Project\17 TA Scan\data\nev-announcements\raw\b405\doc3.txt"
txt = open(p, encoding='utf-8', errors='replace').read()
txt = re.sub(r'\s+', '', txt)
MODEL_RE = re.compile(r'[A-Z]{1,4}\d[A-Z0-9]{2,}')
PROD_RE = re.compile(r'(?:换电式|插电式混合动力|燃料电池|纯电动|混合动力)[\u4e00-\u9fa5]{1,10}')
models = list(MODEL_RE.finditer(txt))
print("模型匹配数:", len(models))
for i, mm in enumerate(models[:4]):
    seg_end = models[i+1].start() if i+1 < len(models) else mm.end()+120
    seg = txt[mm.end():seg_end]
    print(f"\n--- [{i}] model={mm.group(0)} seg_len={len(seg)}")
    print("seg:", seg[:150])
    mp = PROD_RE.search(seg)
    print("PROD match:", mp.group(0) if mp else None)
    if mp:
        after = seg[mp.end():]
        m_digit = re.search(r'\d', after)
        print("digit at:", m_digit.start() if m_digit else None)
        if m_digit:
            print("tail:", after[m_digit.start():][:80])
