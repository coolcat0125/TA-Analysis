# -*- coding: utf-8 -*-
import openpyxl, io, sys, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE = r"D:\#AI\00 Project\17 TA Scan"
wb = openpyxl.load_workbook(BASE + r"\TA-Analysis\NEV公告参数汇总表_合并版（341~409批）.xlsx", read_only=True, data_only=True)
ws = wb['NEV公告参数汇总']
vals = collections.Counter()
by_batch = collections.defaultdict(collections.Counter)
samples = {}
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    if row[0] is None: continue
    b = int(row[0]) if isinstance(row[0], (int, float)) else 0
    v = row[10]
    s = repr(v)
    vals[s] += 1
    by_batch[b][s] += 1
    if s not in samples: samples[s] = (b, str(row[1])[:24])
wb.close()
print("购置税列全部取值分布:")
for v, c in vals.most_common(12):
    print(f"  {v:<12} {c:>5}  首见: 批{samples[v][0]} {samples[v][1]}")
print("\n402-404批取值:")
for b in [402, 403, 404]:
    print(f"  批{b}: {dict(by_batch[b])}")
