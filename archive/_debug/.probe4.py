# -*- coding: utf-8 -*-
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
p = r"D:\#AI\00 Project\17 TA Scan\data\nev-announcements\raw\b405\doc3.txt"
txt = open(p, encoding='utf-8', errors='replace').read()
cells = [c.strip() for c in txt.split('\x07')]
# 找到第一个车型行附近打印30个cell
start = next(i for i, c in enumerate(cells) if c == 'BYD6480AMBEV5')
for i in range(start-4, start+40):
    print(f"{i:>5} | {cells[i]!r}")
print("...")
# 表头cell
hdr = next(i for i, c in enumerate(cells) if '车辆型号' in c)
print("表头附近:")
for i in range(hdr-6, hdr+14):
    print(f"{i:>5} | {cells[i]!r}")
