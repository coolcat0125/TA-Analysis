# -*- coding: utf-8 -*-
import re, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
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
        work = []; seen = set()
        for e in sorted(ends):
            key = (e, tail[pos:e])
            if key not in seen: seen.add(key); work.append(key)
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
            if num is not None: res[(e, r)] = num
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
            ok = (len(vals)==0 and 30<=num<=2000) or (len(vals)==1 and 400<=num<=5000) or (len(vals)==2 and 50<=num<=3000) or (len(vals)==3 and 5<=num<=160)
            if not ok: continue
            vals.append(num); dfs(end, vals); vals.pop()
    dfs(0, [])
    if not sols: return None
    sols.sort(key=lambda x: -x[0])
    return sols[0][1]

print(parse_num_tail("501152034544.2"))
print(parse_num_tail("605(CLTC)1970490.669.07"))
print(parse_num_tail("510151034245.8"))
# 检查 b405 真实文本里模型段的实际 tail
p = r"D:\#AI\00 Project\17 TA Scan\data\nev-announcements\raw\b405\doc3.txt"
txt = re.sub(r'\s+', '', open(p, encoding='utf-8', errors='replace').read())
i = txt.find('BJ7000T669BEV')
print("b405 BJ7000T669BEV 段:", repr(txt[i+13:i+50]))
# b394 样例
p2 = r"D:\#AI\00 Project\17 TA Scan\data\nev-announcements\raw\b394\doc3.txt"
txt2 = re.sub(r'\s+', '', open(p2, encoding='utf-8', errors='replace').read())
print("b394 头400字符:", txt2[:400])
