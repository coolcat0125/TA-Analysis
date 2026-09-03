# -*- coding: utf-8 -*-
import openpyxl, re, json, os, io, sys, random
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE = r"D:\#AI\00 Project\17 TA Scan"
REPO = BASE + r"\TA-Analysis"
RAW  = BASE + r"\data\nev-announcements\raw"

MODEL_RE = re.compile(r'[A-Z]{1,4}\d[A-Z0-9]{2,}')
ENT_RE = re.compile(r'([\u4e00-\u9fa5]{4,30}(?:有限公司|股份公司|工厂|集团有限公司|制造有限公司|汽车有限公司|工业有限公司))')
PROD_RE = re.compile(r'(?:换电式|插电式混合动力|燃料电池|纯电动|混合动力)[\u4e00-\u9fa5]{1,10}')
SEC_RE = re.compile(r'(一、纯电动汽车|二、插电式混合动力汽车|三、燃料电池汽车|四、)')

def first_num(v):
    m = re.match(r'(\d+(?:\.\d+)?)', str(v))
    return float(m.group(1)) if m else None

def parse_num_tail(tail):
    tail = tail.strip()
    m = re.search(r'[\u4e00-\u9fa5]', tail)
    if m: tail = tail[:m.start()]
    n = len(tail)
    if n < 8: return None
    def values_at(pos):
        res = {}
        md = re.match(r'\d+', tail[pos:])
        if not md: return []
        L = md.end()
        ends = set()
        for k in range(1, L+1):
            ends.add(pos+k)
            if pos+k < n and tail[pos+k] == '.':
                dm = re.match(r'\.(\d{1,3})', tail[pos+k:])
                if dm:
                    for dk in range(1, len(dm.group(1))+1):
                        ends.add(pos+k+1+dk)
        work = []
        seen = set()
        for e in sorted(ends):
            key = (e, tail[pos:e])
            if key not in seen:
                seen.add(key); work.append(key)
        queue = list(work); final = []
        while queue:
            e, r = queue.pop()
            final.append((e, r))
            if e < n and tail[e] == '±':
                m3 = re.match(r'±\d+(?:\.\d+)?', tail[e:])
                if m3:
                    c = (e+m3.end(), tail[pos:e+m3.end()])
                    if c not in seen: seen.add(c); queue.append(c)
            if e < n and tail[e] == '/':
                m4 = re.match(r'/(?:\d+(?:\.\d+)?)(?:±\d+(?:\.\d+)?)?', tail[e:])
                if m4:
                    c = (e+m4.end(), tail[pos:e+m4.end()])
                    if c not in seen: seen.add(c); queue.append(c)
            if e < n and tail[e] == '(':
                m5 = re.match(r'\((?:CLTC|WLTC|NEDC)\)', tail[e:])
                if m5:
                    c = (e+m5.end(), tail[pos:e+m5.end()])
                    if c not in seen: seen.add(c); queue.append(c)
        for e, r in final:
            num = first_num(r)
            if num is not None:
                res[(e, r)] = num
        return [(e, r, num) for (e, r), num in res.items()]

    sols = []
    def dfs(pos, vals):
        if len(sols) >= 30: return
        remain = 4 - len(vals)
        if n - pos < 2*remain: return
        if remain == 0:
            r_, w_, m_, e_ = vals
            if 30<=r_<=2000 and 400<=w_<=5000 and 50<=m_<=3000 and 5<=e_<=160:
                sols.append((n-pos, list(vals)))
            return
        for end, raw, num in values_at(pos):
            ok = (len(vals)==0 and 30<=num<=2000) or \
                 (len(vals)==1 and 400<=num<=5000) or \
                 (len(vals)==2 and 50<=num<=3000) or \
                 (len(vals)==3 and 5<=num<=160)
            if not ok: continue
            vals.append(num); dfs(end, vals); vals.pop()
    dfs(0, [])
    if not sols: return None
    sols.sort(key=lambda x: -x[0])
    return sols[0][1]

def parse_doc3(path):
    txt = open(path, encoding='utf-8', errors='replace').read()
    txt = re.sub(r'\s+', '', txt)
    sections = [(m.start(), m.group(0)) for m in SEC_RE.finditer(txt)]
    def sec_at(p):
        name = '纯电动'
        for s, nm in sections:
            if s <= p:
                name = '插电式' if nm.startswith('二、') else ('燃料电池' if nm.startswith('三、') else ('纯电动' if nm.startswith('一、') else name))
        return name
    entries = {}
    models = list(MODEL_RE.finditer(txt))
    for i, mm in enumerate(models):
        model = mm.group(0)
        seg = txt[mm.end(): models[i+1].start() if i+1 < len(models) else mm.end()+120]
        mp = PROD_RE.search(seg)
        if not mp: continue
        after = seg[mp.end():]
        mdig = re.search(r'\d', after)
        if not mdig: continue
        parsed = parse_num_tail(after[mdig.start():])
        if not parsed: continue
        if model not in entries:
            entries[model] = {'range':parsed[0], 'curb':parsed[1], 'mass':parsed[2], 'energy':parsed[3], 'section':sec_at(mm.start())}
    return entries

print("== 解析30批 doc3 ==")
catalog = {}
for b in range(379, 409):
    p = os.path.join(RAW, f'b{b}', 'doc3.txt')
    if not os.path.exists(p): continue
    ent = parse_doc3(p)
    catalog[b] = ent
print("各批次条目:", {b: len(catalog[b]) for b in sorted(catalog)})
print("合计:", sum(len(v) for v in catalog.values()))

# 验证: 已知否/是行命中
wb = openpyxl.load_workbook(REPO + r"\NEV公告参数汇总表_合并版（341~409批）.xlsx", read_only=True, data_only=True)
ws = wb['NEV公告参数汇总']
rows = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0 or row[0] is None: continue
    rows.append(list(row))
wb.close()
tp = fp = 0
for r in rows:
    b = int(r[0]) if isinstance(r[0],(int,float)) else 0
    model = str(r[1]).strip() if r[1] else ''
    if b in catalog and model:
        inc = model in catalog[b]
        tv = str(r[10]).strip()
        if tv == '否' and inc: fp += 1
        if tv == '是' and inc: tp += 1
print(f"已知'是'命中 {tp} / 已知'否'误命中 {fp}")

# 能量密度交叉验证
random.seed(7)
diffs = []
for r in rows:
    b = int(r[0]) if isinstance(r[0],(int,float)) else 0
    model = str(r[1]).strip() if r[1] else ''
    if b in catalog and model in catalog[b] and isinstance(r[16],(int,float)):
        e = catalog[b][model]
        calc = e['energy']*1000/e['mass']
        diffs.append((abs(calc-r[16])/r[16]*100, b, model, calc, r[16]))
diffs.sort()
if diffs:
    print(f"能量密度对比 n={len(diffs)} 中位偏差 {diffs[len(diffs)//2][0]:.2f}%  P90 {diffs[int(len(diffs)*0.9)][0]:.2f}%")
    print("最差5例:", [(d[1], d[2], round(d[3],1), d[4]) for d in diffs[-5:]])

json.dump({str(b): catalog[b] for b in catalog}, open(REPO+r'\.catalog_doc3.json','w',encoding='utf-8'), ensure_ascii=False)
print("catalog 缓存完成")
