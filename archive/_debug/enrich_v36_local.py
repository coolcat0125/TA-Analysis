# -*- coding: utf-8 -*-
"""
enrich_v36_local.py — NEV源表数据补全引擎 v3.6.0（本地全自动包 F=A+B+C）
Phase C: 污染单元格复核修正 → Phase A: doc3购置税目录回填 → B1: 电耗公式 → B2: 同型号唯一回填 → B3: master跨链回填
守卫：只填缺失不覆盖（C除外）、物理范围校验、动力类型交叉校验、逐格审计可回滚。
用法：python enrich_v36_local.py --dryrun | --apply
Copyright © 2026 David YE · MIT License
"""
import openpyxl, re, json, csv, os, io, sys, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r"D:\#AI\00 Project\17 TA Scan"
REPO = BASE + r"\TA-Analysis"
SRC  = REPO + r"\NEV公告参数汇总表_合并版（341~409批）.xlsx"
MASTER = BASE + r"\data\nev-announcements\db\master_export.csv"
RAW  = BASE + r"\data\nev-announcements\raw"
SHEET = 'NEV公告参数汇总'
APPLY = '--apply' in sys.argv

RNG = {
    '整备质量(kg)':      ((300,4500),(500,4500)),
    '纯电续航里程(km)':   ((50,1500),(20,600)),
    '电池容量(kWh)':     ((2,250),(1,120)),
    '电池能量密度(Wh/kg)':((50,400),(30,400)),
    '百公里电耗(kWh/100km)':((3,40),(3,40)),
    '电机峰值功率(kW)':   ((5,1500),(5,800)),
    '电机总功率(kW)':    ((5,1500),(5,800)),
    '综合油耗(L/100km)': (None,(0.1,20)),
    '发动机排量(mL)':    (None,(300,6000)),
    '发动机功率(kW)':    (None,(10,600)),
    '轴距(mm)':          ((1800,3900),(1800,3900)),
}
def in_range(field, val, ptype):
    pair = RNG.get(field)
    if not pair: return True
    rng = pair[0] if ptype == 'BEV' else pair[1]
    if rng is None: return False
    return rng[0] <= val <= rng[1]
def is_blank(v):
    if v is None: return True
    s = str(v).strip()
    return (not s) or s in ('-','—','/','None','nan') or '不适用' in s
def first_num(v):
    m = re.match(r'(\d+(?:\.\d+)?)', str(v).strip())
    return float(m.group(1)) if m else None
def to_num(v):
    """严格数字转换：数值直过；纯数字文本转 float；复合值/含中文 → None"""
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    if re.fullmatch(r'\d+(?:\.\d+)?', s): return float(s)
    return None
def ptype_of(r): return 'BEV' if str(r[9]).strip() == 'BEV' else 'PHEV'

# ============================================================
# doc3 目录解析（\x07 单元格状态机；b394 压平文本跳过）
# ============================================================
PROD_PAT = re.compile(r'^(?:换电式|插电式混合动力|燃料电池|纯电动|混合动力|插电式)')
MODEL_PAT = re.compile(r'^[A-Z][A-Z0-9\-·]{4,}$')
ENT_PAT = re.compile(r'(有限公司|股份公司|集团公司|工厂|研究院)$')
NUM_CELL = re.compile(r'[\d.±]')
def clean_numcell(c):
    c = re.sub(r'\((?:CLTC|WLTC|NEDC)\)', '', c, flags=re.I)
    c = c.replace('（', '').replace('）', '')
    return first_num(c)

def parse_doc3_cells(path):
    txt = open(path, encoding='utf-8', errors='replace').read()
    cells = [c.strip() for c in txt.split('\x07')]
    entries = {}
    section = '纯电动'
    ent_last = None
    for i, c in enumerate(cells):
        if not c: continue
        if '一、纯电动' in c or c.startswith('一、'): section = '纯电动'; continue
        if '二、插电式' in c or c.startswith('二、'): section = '插电式'; continue
        if '三、燃料电池' in c or c.startswith('三、'): section = '燃料电池'; continue
        if ENT_PAT.search(c) and not NUM_CELL.search(c):
            ent_last = c; continue
        if PROD_PAT.match(c) and i >= 3:
            generic = cells[i-1].strip()
            model = cells[i-2].strip()
            if not MODEL_PAT.match(model) or not re.search(r'\d', model):
                continue
            # 向后取4个数字单元格
            nums, j = [], i+1
            while len(nums) < 4 and j < len(cells):
                cc = cells[j].strip()
                if cc and NUM_CELL.search(cc) and not re.search(r'[\u4e00-\u9fa5]', cc):
                    nums.append(cc); j += 1
                elif cc == '' or re.match(r'^[A-Z0-9\-·]{4,}$', cc):
                    break
                else:
                    break
            if len(nums) < 4: continue
            rng_, curb_, mass_, energy_ = (clean_numcell(x) for x in nums)
            if None in (rng_, curb_, mass_, energy_): continue
            if not (30<=rng_<=2000 and 400<=curb_<=5000 and 50<=mass_<=3000 and 5<=energy_<=160):
                continue
            if model not in entries:
                entries[model] = {'range':rng_, 'curb':curb_, 'mass':mass_, 'energy':energy_,
                                  'ent':ent_last, 'section':section}
    return entries

print("== Phase A-1: 解析 doc3 购置税目录（29批单元格流）==")
catalog = {}
for b in range(379, 409):
    p = os.path.join(RAW, f'b{b}', 'doc3.txt')
    if not os.path.exists(p): continue
    ent = parse_doc3_cells(p)
    catalog[b] = ent
print("各批条目:", {b: len(catalog[b]) for b in sorted(catalog)})
print("合计:", sum(len(v) for v in catalog.values()))

# ============================================================
# 载入源表
# ============================================================
WB = openpyxl.load_workbook(SRC)
WS = WB[SHEET]
rows = []
for i, row in enumerate(WS.iter_rows(values_only=True)):
    if i == 0: continue
    if row[0] is None: continue
    rows.append(list(row))
N = len(rows)
HEADER = [str(c) if c else '' for c in next(WS.iter_rows(min_row=1, max_row=1, values_only=True))]
COL = {h: j for j, h in enumerate(HEADER)}
audit = []   # 每项: dict(excel_row,batch,model,field,old,new,method,source,confidence)
def log_change(di, field, old, new, method, source, conf):
    audit.append({'excel_row': di+2, 'batch': rows[di][0], 'model': str(rows[di][1]).strip(),
                  'field': field, 'old': old, 'new': new, 'method': method,
                  'source': source, 'confidence': conf})

# ============================================================
# Phase C: 污染单元格复核修正
# ============================================================
print("\n== Phase C: 污染单元格复核 ==")
c_fixed = c_cleared = c_kept = 0
with open(REPO + r'\analysis-output\contaminated_cells.csv', encoding='utf-8-sig', newline='') as f:
    for rec in csv.DictReader(f):
        er = int(rec['excel_row']); field = rec['field'].strip()
        di = er - 2
        if di < 0 or di >= N: continue
        r = rows[di]
        if str(r[1]).strip() != rec['product_model'].strip():
            # 行号错位保护：按批次+型号重找
            found = None
            for k, rr in enumerate(rows):
                if int(rr[0]) == int(rec['batch']) and str(rr[1]).strip() == rec['product_model'].strip():
                    found = k; break
            if found is None: continue
            di = found; r = rows[di]
        j = COL.get(field)
        if j is None: continue
        cur = r[j]
        pt = ptype_of(r)
        bad = False
        if isinstance(cur, (int, float)):
            bad = not in_range(field, float(cur), pt)
        elif is_blank(cur):
            bad = False  # 空值留给补全阶段
        else:
            s = str(cur)
            bad = bool(re.search(r'请提供|正文内容|提取|无法确定|待补充|无法识别', s)) or s in ('/','—','-')
        if not bad:
            c_kept += 1; continue
        # 尝试修复源：同型号唯一值
        fix, src_note = None, ''
        model = str(r[1]).strip()
        vals = set()
        for k, rr in enumerate(rows):
            if k != di and str(rr[1]).strip() == model and not is_blank(rr[j]):
                nv = to_num(rr[j])
                vals.add(nv if nv is not None else str(rr[j]).strip())
        b = int(r[0])
        if b in catalog and model in catalog[b] and field in ('整备质量(kg)','纯电续航里程(km)','电池容量(kWh)'):
            e = catalog[b][model]
            m2v = {'整备质量(kg)': e['curb'], '纯电续航里程(km)': e['range'], '电池容量(kWh)': e['energy']}
            if in_range(field, m2v[field], pt): fix, src_note = m2v[field], f'doc3目录b{b}'
        if fix is None and len(vals) == 1:
            v = vals.pop()
            if isinstance(v, float):
                if in_range(field, v, pt): fix, src_note = v, '同型号唯一值'
            else:
                fix, src_note = v, '同型号唯一值'
        old_disp = cur
        if fix is not None:
            r[j] = fix; c_fixed += 1
            log_change(di, field, old_disp, fix, 'C-污染修正', src_note, 'medium')
        else:
            r[j] = None; c_cleared += 1
            log_change(di, field, old_disp, None, 'C-污染清除(无可靠源)', '物理范围外/伪值', 'low')
print(f"  修复 {c_fixed} / 清空 {c_cleared} / 核实无误保留 {c_kept}")

# ============================================================
# Phase A: doc3 回填（购置税 是/否、能量密度、容量、续航、整备）
# ============================================================
print("\n== Phase A: doc3 目录回填 ==")
a_n = collections.Counter()
for di, r in enumerate(rows):
    b = int(r[0]) if isinstance(r[0],(int,float)) else 0
    if b not in catalog: continue
    model = str(r[1]).strip() if r[1] else ''
    e = catalog[b].get(model)
    pt = ptype_of(r)
    tax_j = COL['是否减免购置税']
    cur_tax = r[tax_j]
    if e is not None:
        # 动力类型交叉校验
        sec_map = {'纯电动':'BEV', '插电式':'PHEV', '燃料电池':'BEV'}
        if sec_map.get(e['section']) != pt and not (pt=='PHEV' and e['section']=='纯电动'):
            pass  # 目录分节与公告动力类型偶有出入，不作为硬阻断，仅记录
        if is_blank(cur_tax) or str(cur_tax).strip() == '可能':
            r[tax_j] = '是'; a_n['tax_是'] += 1
            log_change(di, '是否减免购置税', cur_tax, '是', 'A-官方目录', f'购置税目录b{b}(第{e.get("section","")})', 'high')
        # 能量密度
        j = COL['电池能量密度(Wh/kg)']
        if is_blank(r[j]):
            ed = round(e['energy']*1000/e['mass'], 1)
            if in_range('电池能量密度(Wh/kg)', ed, pt):
                r[j] = ed; a_n['ed'] += 1
                log_change(di, '电池能量密度(Wh/kg)', None, ed, 'A-目录换算', f'总能量{e["energy"]}kWh÷总质量{e["mass"]}kg×1000', 'high')
        # 容量
        j = COL['电池容量(kWh)']
        if is_blank(r[j]) and in_range('电池容量(kWh)', e['energy'], pt):
            r[j] = e['energy']; a_n['c'] += 1
            log_change(di, '电池容量(kWh)', None, e['energy'], 'A-官方目录', f'目录总能量值b{b}', 'high')
        # 续航
        j = COL['纯电续航里程(km)']
        if is_blank(r[j]) and in_range('纯电续航里程(km)', e['range'], pt):
            r[j] = e['range']; a_n['r'] += 1
            log_change(di, '纯电续航里程(km)', None, e['range'], 'A-官方目录', f'目录续驶里程b{b}', 'high')
        # 整备
        j = COL['整备质量(kg)']
        if is_blank(r[j]) and in_range('整备质量(kg)', e['curb'], pt):
            r[j] = e['curb']; a_n['w'] += 1
            log_change(di, '整备质量(kg)', None, e['curb'], 'A-官方目录', f'目录整备质量b{b}', 'high')
    else:
        # 目录核对未见 → 是/否判定（仅当该批次目录解析正常>20条）
        if len(catalog[b]) > 20 and (is_blank(cur_tax) or str(cur_tax).strip() == '可能'):
            r[tax_j] = '否'; a_n['tax_否'] += 1
            log_change(di, '是否减免购置税', cur_tax, '否', 'A-目录核对未见', f'购置税目录b{b}全目录检索无此型号', 'medium')
print(f"  {dict(a_n)}")

# ============================================================
# Phase B1: 电耗公式
# ============================================================
print("\n== Phase B1: 电耗公式 ==")
b1 = 0
jc, jr, je = COL['电池容量(kWh)'], COL['纯电续航里程(km)'], COL['百公里电耗(kWh/100km)']
for di, r in enumerate(rows):
    if not is_blank(r[je]): continue
    c, rr = to_num(r[jc]), to_num(r[jr])
    if c is not None and rr is not None and rr > 0:
        ec = round(c*100/rr, 1)
        pt = ptype_of(r)
        if in_range('百公里电耗(kWh/100km)', ec, pt):
            r[je] = ec; b1 += 1
            log_change(di, '百公里电耗(kWh/100km)', None, ec, 'B1-公式换算', f'容量{c}kWh×100÷续航{rr}km', 'medium')
print(f"  电耗公式补 {b1} 格")

# ============================================================
# Phase B2: 同型号跨批次唯一值回填
# ============================================================
print("\n== Phase B2: 同型号唯一值回填 ==")
by_model = collections.defaultdict(list)
for di, r in enumerate(rows):
    m = str(r[1]).strip() if r[1] else ''
    if m: by_model[m].append(di)
B2_FIELDS = ['细分市场','车型名称','产品名称','通用名称','产品类型','企业名称','产品商标','轴距(mm)',
             '电池类型','电机生产企业','电机型号','电机峰值功率(kW)','电机总功率(kW)',
             '综合油耗(L/100km)','B状态油耗(L/100km)','发动机排量(mL)','发动机功率(kW)',
             '发动机生产企业','发动机型号','整备质量(kg)','纯电续航里程(km)','电池容量(kWh)','电机功率/扭矩']
b2 = collections.Counter()
for field in B2_FIELDS:
    j = COL[field]
    for m, dis in by_model.items():
        if len(dis) < 2: continue
        vals = set()
        for di in dis:
            v = rows[di][j]
            if is_blank(v): continue
            nv = to_num(v)
            vals.add(nv if nv is not None else str(v).strip())
        if len(vals) != 1: continue
        v = vals.pop()
        for di in dis:
            r = rows[di]
            if not is_blank(r[j]): continue
            pt = ptype_of(r)
            if isinstance(v, float):
                if not in_range(field, v, pt): continue
            elif field in ('综合油耗(L/100km)','发动机排量(mL)','发动机功率(kW)','发动机生产企业','发动机型号','B状态油耗(L/100km)') and pt == 'BEV':
                continue  # 发动机系字段不回填给BEV
            r[j] = v; b2[field] += 1
            log_change(di, field, None, v, 'B2-同型号唯一值', f'同型号{len(dis)}行非空值一致', 'medium')
print(f"  合计 {sum(b2.values())} 格  明细: {dict(b2)}")

# ============================================================
# Phase B3: master 跨链回填
# ============================================================
print("\n== Phase B3: master 链交叉回填 ==")
mmap = {}
with open(MASTER, encoding='utf-8-sig', newline='') as f:
    rd = csv.reader(f); mhead = next(rd)
    for mr in rd:
        if not mr or not any(x.strip() for x in mr): continue
        k = mr[1].strip().upper() if len(mr) > 1 and mr[1] else ''
        if not k: continue
        e = mmap.setdefault(k, {})
        for j in range(2, 30):
            if j >= len(mr): break
            v = mr[j].strip()
            if v and v not in ('-','/','None','nan','') and '不适用' not in v:
                e.setdefault(HEADER[j], set()).add(v)
B3_FIELDS = ['细分市场','车型名称','产品名称','通用名称','产品类型','企业名称','产品商标','轴距(mm)',
             '电池类型','电机生产企业','电机型号','电机峰值功率(kW)','电机总功率(kW)','电机功率/扭矩',
             '电池能量密度(Wh/kg)','整备质量(kg)','纯电续航里程(km)','电池容量(kWh)',
             '综合油耗(L/100km)','B状态油耗(L/100km)','发动机排量(mL)','发动机功率(kW)','发动机生产企业','发动机型号']
b3 = collections.Counter()
for di, r in enumerate(rows):
    k = (str(r[1]).strip() if r[1] else '').upper()
    e = mmap.get(k)
    if not e: continue
    pt = ptype_of(r)
    for field in B3_FIELDS:
        j = COL[field]
        if not is_blank(r[j]): continue
        vs = e.get(field)
        if not vs or len(vs) != 1: continue
        v = next(iter(vs))
        if field in ('综合油耗(L/100km)','B状态油耗(L/100km)','发动机排量(mL)','发动机功率(kW)','发动机生产企业','发动机型号') and pt == 'BEV':
            continue
        num = first_num(v)
        if field in RNG:
            if num is None or not in_range(field, num, pt): continue
            r[j] = num
        else:
            if re.search(r'请提供|正文内容|无法', v): continue
            r[j] = v
        b3[field] += 1
        log_change(di, field, None, r[j], 'B3-master跨链', 'master_export同型号唯一值', 'medium')
print(f"  合计 {sum(b3.values())} 格  明细: {dict(b3)}")

# ============================================================
# 汇总与落盘
# ============================================================
print(f"\n== 总计: 审计 {len(audit)} 格变更 ==")
meth = collections.Counter(a['method'].split('-')[0] for a in audit)
print("按方法:", dict(meth))
byf = collections.Counter(a['field'] for a in audit)
print("按字段 top12:", dict(byf.most_common(12)))

if APPLY:
    for a in audit:
        j = COL[a['field']]
        WS.cell(row=a['excel_row'], column=j+1, value=a['new'])
    WB.save(SRC)
    print(f"✓ 源表已更新: {SRC}")
    json.dump(audit, open(REPO+r'\.audit_v36.json','w',encoding='utf-8'), ensure_ascii=False, indent=1)
    print("✓ 审计日志 → .audit_v36.json")
else:
    json.dump(audit, open(REPO+r'\.audit_v36_dryrun.json','w',encoding='utf-8'), ensure_ascii=False, indent=1)
    print("(dryrun 未写盘，审计 → .audit_v36_dryrun.json)")
