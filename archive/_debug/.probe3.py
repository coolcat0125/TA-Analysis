# -*- coding: utf-8 -*-
import os, re, io, sys, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
RAW = r"D:\#AI\00 Project\17 TA Scan\data\nev-announcements\raw"
for b in range(379, 409):
    p = os.path.join(RAW, f'b{b}', 'doc3.txt')
    if not os.path.exists(p): continue
    txt = open(p, encoding='utf-8', errors='replace').read()
    n7 = txt.count('\x07')
    cells = txt.split('\x07')
    print(f"b{b}: {len(txt):>6}字符  \\x07×{n7:>5}  split后{len(cells):>5}格  样例: {cells[10:16]}")
