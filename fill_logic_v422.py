# -*- coding: utf-8 -*-
"""v4.2.2 表内逻辑推理补全（只补空、不覆盖、逐格台账、底纹标注）
====================================================================
阶段顺序（在 enrich_v5/fill_consensus_v5 通道约定基础上扩展）：
  0. 垃圾文本清洗 —— LLM 提取失败残留（"请提供…正文内容"等）置空（浅红 FFC7CE，类型=数据清洗更正）
  J. 企业名称推理 —— 商标/通用名称/型号代号 → 唯一企业（浅橙 FCE4D6，类型=企业名称推理）
  B. 同型号共识再跑 —— (产品型号,动力类型) 唯一非空共识（黄 FFF2CC，类型=同车型共识，沿用 v3.6）
  C. 企业+车型共识再跑 —— (企业,车型名称,动力类型) 唯一非空共识（绿 C6EFCE，类型=共识补全，沿用 v4.0）
  L. 电机功率解析 —— 「电机功率/扭矩」解析：单电机→总功率；前/后双电机→总功率=前+后、峰值=max（蓝 BDD7EE，类型=功率解析）
  K. 细分市场推理 —— K1 同通用名称/车型名称唯一细分市场回填（浅黄绿 E2EFDA，类型=通用名称共识）；
                      K2 轴距分级（车型级别定义表阈值，蓝 BDD7EE，类型=轴距分级）

用法：python fill_logic_v422.py
Copyright © 2026 David YE
"""
import os
import re
import sys
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SHEET = 'NEV公告参数汇总'

RED = openpyxl.styles.PatternFill('solid', fgColor='FFC7CE')
LIGHTORANGE = openpyxl.styles.PatternFill('solid', fgColor='FCE4D6')
YELLOW = openpyxl.styles.PatternFill('solid', fgColor='FFF2CC')
GREEN = openpyxl.styles.PatternFill('solid', fgColor='C6EFCE')
BLUE = openpyxl.styles.PatternFill('solid', fgColor='BDD7EE')
YELLOWGREEN = openpyxl.styles.PatternFill('solid', fgColor='E2EFDA')

NUM_RANGES = {'整备质量(kg)': (300, 4500), '轴距(mm)': (1800, 4200), '纯电续航里程(km)': (20, 1500),
              '电池容量(kWh)': (1, 250), '电池能量密度(Wh/kg)': (30, 400), '百公里电耗(kWh/100km)': (3, 40),
              '电机峰值功率(kW)': (5, 1500), '电机总功率(kW)': (5, 1500), '综合油耗(L/100km)': (0.1, 20),
              'B状态油耗(L/100km)': (0.1, 20), '发动机排量(mL)': (300, 6000), '发动机功率(kW)': (10, 600),
              '车长(mm)': (1800, 6500)}
EMPTY_PLACEHOLDER = {'', '/', '-', '—', '0', '未填报', '无', '无信息', '待查', '待核实',
                     '待确认', '未知', '？', '?', 'N/A', 'NA', 'None', '未提供'}
SEMANTIC_PLACEHOLDER = {'不适用', '不适用(BEV)', '不适用(PHEV)', '不适用(HEV)', '不适用(EREV)'}
BAD_TEXT = ('请提供', '无法', '无有效', '未找到', '无符合', 'unknown', 'n/a', '需人工',
            '暂缺', '待确认', '待核实', '略', '无信息', '未填写', '不祥')
DIRTY_WORDS = ('(', ')', '？', '?', '，', ',', '、', '；', ';', '×', '*')
GARBAGE_PAT = re.compile(r'请提供|请你|正文内容|以便我|无法直接|作为AI|对不起|抱歉|未提供具体')
NUM_RE = re.compile(r'^[\d.]+(\s*/\s*[\d.]+)*$')

MPV_KEYWORDS = ['MPV', 'GL8', '奥德赛', '艾力绅', '埃尔法', '威尔法', '赛那', '嘉际', '宋MAX',
                '传祺M', '腾势D', '极氪009', '岚图梦想家', '风行', '菱智', '瑞风', '大通G',
                'G10', 'G20', 'G50', 'G90', 'V级', '威霆', '图雅诺', '全顺', '依维柯', '欧胜',
                '星锐', '特顺', '福顺', '别克GL', '库斯途', '五菱宏光', '荣光', '征程', '凯捷',
                '长安之星', '金牛星', '欧诺', '睿行', 'V80', 'V90', '海狮', 'HQ9']


def norm(v):
    if v is None:
        return None
    s = str(v).replace('\u3000', '').strip()
    if s == '' or s in EMPTY_PLACEHOLDER or s in SEMANTIC_PLACEHOLDER:
        return None
    return s


def fillable(v):
    if v is None:
        return True
    s = str(v).replace('\u3000', '').strip()
    return s == '' or s in EMPTY_PLACEHOLDER


def cast_num(s):
    m = re.search(r'[\d.]+', str(s))
    if not m:
        return None
    f = float(m.group())
    return int(f) if f.is_integer() else f


def clean_txt(s):
    if s is None:
        return None
    s = str(s).strip()
    if not s or len(s) > 60:
        return None
    if any(b in s.lower() for b in BAD_TEXT) or any(d in s for d in DIRTY_WORDS):
        return None
    return s


def clean_num(s):
    return str(s).strip() if NUM_RE.match(str(s).strip()) else None


def in_range(field, v):
    f = cast_num(v)
    lo, hi = NUM_RANGES[field]
    return f is not None and lo <= f <= hi


def is_mpv(gname, pname):
    txt = (gname or '') + (pname or '')
    return any(k in txt for k in MPV_KEYWORDS)


def classify_seg(wb_val, pname, gname):
    """按车型级别定义表阈值分级（轴距 mm 为唯一分级依据）"""
    if not wb_val or not (2000 < wb_val < 4000):
        return None
    pname, gname = pname or '', gname or ''
    if '轿车' in pname:
        if wb_val < 2650:
            return 'Car-A'
        if wb_val < 2740:
            return 'Car-B'
        if wb_val < 2850:
            return 'Car-C'
        if wb_val < 3000:
            return 'Car-D'
        return 'Car-E'
    if is_mpv(gname, pname):
        return 'MPV-D' if wb_val >= 3000 else 'MPV-C'
    if '多用途' in pname:
        if wb_val < 2680:
            return 'SUV-B'
        if wb_val < 2850:
            return 'SUV-C'
        if wb_val < 3000:
            return 'SUV-D'
        return 'SUV-E'
    return None


def parse_motor_pt(v):
    """解析电机功率/扭矩：返回 ('dual', 前/主功率, 后/次功率) 或 ('single', 功率, None)。"""
    s = str(v).strip()
    m = re.search(r'F\s*[:：]\s*(\d+(?:\.\d+)?)\s*/\s*\d+(?:\.\d+)?\s+[Rr]\s*[:：]\s*(\d+(?:\.\d+)?)\s*/\s*\d+(?:\.\d+)?', s)
    if m:
        return ('dual', float(m.group(1)), float(m.group(2)))
    m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*k?W?\s*/\s*(\d+(?:\.\d+)?)\s*k?W?\s*$', s, re.I)
    if m and 'kW' in s.upper():
        return ('dual', float(m.group(1)), float(m.group(2)))
    m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*(?:k?W)?\s*(?:/\s*\d+(?:\.\d+)?)?\s*/?\s*$', s)
    if m:
        return ('single', float(m.group(1)), None)
    return None


def consensus_pass(ws, fi, rows, key_fields, fields, num_guard, color, ctype, note, stats, ledger):
    groups = defaultdict(lambda: defaultdict(set))
    gsize = defaultdict(int)
    for _, rec in rows:
        key = tuple(norm(rec.get(k)) for k in key_fields)
        if None in key:
            continue
        gsize[key] += 1
        for f in fields:
            v = norm(rec.get(f))
            if v is None:
                continue
            if f in num_guard and (not clean_num(v) or not in_range(f, v)):
                continue
            if f not in num_guard and not clean_txt(v):
                continue
            groups[key][f].add(v)
    filled = 0
    for r, rec in rows:
        key = tuple(norm(rec.get(k)) for k in key_fields)
        if None in key or gsize[key] < 2:
            continue
        for f in fields:
            if not fillable(rec.get(f)):
                continue
            s = groups[key].get(f)
            if s and len(s) == 1:
                val = next(iter(s))
                cell = ws.cell(r, fi[f] + 1)
                cell.value = cast_num(val) if f in num_guard else val
                cell.fill = color
                rec[f] = cell.value
                stats[f] += 1
                filled += 1
                ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), f, ctype,
                               '(空)', val, note.format(n=gsize[key])])
    return filled


def main():
    print(f'[load] {os.path.basename(SRC)}')
    wb = openpyxl.load_workbook(SRC)
    ws = wb[SHEET]
    hdr = [c.value for c in ws[1]]
    fi = {h: i for i, h in enumerate(hdr)}
    rows = []
    for r in range(2, ws.max_row + 1):
        rows.append((r, {h: ws.cell(r, fi[h] + 1).value for h in hdr}))

    cr = wb['变更记录']
    last_seq = 0
    for rr in range(cr.max_row, 1, -1):
        v = cr.cell(rr, 1).value
        if isinstance(v, int):
            last_seq = v
            break
    stats = defaultdict(int)
    ledgers = []

    # ---- 阶段0 垃圾文本清洗 ----
    print('[0] 垃圾文本清洗（LLM 提取失败残留）…')
    g_ledger = []
    for r, rec in rows:
        for h in hdr:
            v = rec.get(h)
            if v is None:
                continue
            s = str(v).strip()
            if s and GARBAGE_PAT.search(s) and h not in ('变更记录',):
                cell = ws.cell(r, fi[h] + 1)
                g_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), h,
                                 '数据清洗更正', s[:60], '(空)', 'LLM提取失败残留文本，置空防阻塞推理'])
                cell.value = None
                cell.fill = RED
                rec[h] = None
                stats['垃圾清洗'] += 1
    print(f'      置空 {len(g_ledger)} 格')
    ledgers.append((g_ledger, '数据清洗更正'))

    # ---- 阶段J 企业名称推理 ----
    print('[J] 企业名称推理（商标/车系/型号代号 → 唯一企业）…')
    ent_maps = {}
    ent_maps['产品商标'] = defaultdict(set)
    ent_maps['通用名称'] = defaultdict(set)
    ent_maps['型号代号'] = defaultdict(set)

    def code_of(model):
        m = re.match(r'^[A-Z]+', str(model or '').strip())
        return m.group() if m else None
    for _, rec in rows:
        e = clean_txt(norm(rec.get('企业名称')))
        if not e:
            continue
        for src, field in (('产品商标', '产品商标'), ('通用名称', '通用名称')):
            k = clean_txt(norm(rec.get(field)))
            if k:
                ent_maps[src][k].add(e)
        c = code_of(rec.get('产品型号'))
        if c:
            ent_maps['型号代号'][c].add(e)
    uniq = {src: {k: next(iter(v)) for k, v in m.items() if len(v) == 1} for src, m in ent_maps.items()}
    for src, m in uniq.items():
        print(f'      {src}→企业 唯一映射 {len(m)} 个')
    j_ledger = []
    for r, rec in rows:
        if not fillable(rec.get('企业名称')):
            continue
        for src, key in (('产品商标', clean_txt(norm(rec.get('产品商标')))),
                         ('通用名称', clean_txt(norm(rec.get('通用名称')))),
                         ('型号代号', code_of(rec.get('产品型号')))):
            if key and key in uniq[src]:
                e = uniq[src][key]
                cell = ws.cell(r, fi['企业名称'] + 1)
                cell.value = e
                cell.fill = LIGHTORANGE
                rec['企业名称'] = e
                stats['企业名称'] += 1
                j_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), '企业名称',
                                 '企业名称推理', '(空)', e, f'推断：{src}「{key}」全表唯一对应企业，非实测'])
                break
    print(f'      补 {len(j_ledger)} 格')
    ledgers.append((j_ledger, '企业名称推理'))

    # ---- 阶段B/C 共识再跑 ----
    print('[B] 同型号共识再跑…')
    num_f = set(NUM_RANGES) | {'整备质量(kg)', '纯电续航里程(km)', '电池容量(kWh)',
                               '电池能量密度(Wh/kg)', '百公里电耗(kWh/100km)', '综合油耗(L/100km)',
                               'B状态油耗(L/100km)', '发动机排量(mL)', '发动机功率(kW)'}
    txt_f = ['企业名称', '车型名称', '产品名称', '通用名称', '产品类型', '细分市场', '电池类型',
             '电机生产企业', '电机功率/扭矩', '电机型号', '发动机生产企业', '发动机型号']
    fields_bc = [f for f in txt_f if f in fi] + [f for f in num_f if f in fi]
    b_ledger = []
    n = consensus_pass(ws, fi, rows, ('产品型号', '动力类型'), fields_bc, num_f, YELLOW,
                       '同车型共识', '推断：(产品型号,动力类型)唯一共识(组n={n})，非实测', stats, b_ledger)
    print(f'      补 {n} 格')
    ledgers.append((b_ledger, '同车型共识'))

    print('[C] 企业+车型共识再跑…')
    c_ledger = []
    n = consensus_pass(ws, fi, rows, ('企业名称', '车型名称', '动力类型'), fields_bc, num_f, GREEN,
                       '共识补全', '推断：(企业,车型名称,动力类型)唯一共识(组n={n})，非实测', stats, c_ledger)
    print(f'      补 {n} 格')
    ledgers.append((c_ledger, '共识补全'))

    # ---- 阶段L 电机功率解析 ----
    print('[L] 电机功率/扭矩 解析（单电机→总功率；前/后双电机→总=和、峰值=max）…')
    l_ledger = []

    def fill_num(r, rec, field, val, note):
        lo, hi = NUM_RANGES[field]
        if not (lo <= val <= hi):
            return False
        cell = ws.cell(r, fi[field] + 1)
        cell.value = int(val) if float(val).is_integer() else val
        cell.fill = BLUE
        rec[field] = cell.value
        stats[field] += 1
        l_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), field,
                         '功率解析', '(空)', cell.value, note])
        return True
    for r, rec in rows:
        pt = norm(rec.get('电机功率/扭矩'))
        if not pt:
            continue
        parsed = parse_motor_pt(pt)
        if not parsed:
            continue
        kind, a, b = parsed
        if kind == 'single':
            if fillable(rec.get('电机总功率(kW)')):
                fill_num(r, rec, '电机总功率(kW)', a, f'解析自「电机功率/扭矩」{pt}（单电机总功率=该电机功率，蓝底纹）')
        else:
            if fillable(rec.get('电机总功率(kW)')):
                fill_num(r, rec, '电机总功率(kW)', a + b, f'解析自「电机功率/扭矩」前{int(a)}+后{int(b)}kW 求和（蓝底纹）')
            if fillable(rec.get('电机峰值功率(kW)')):
                fill_num(r, rec, '电机峰值功率(kW)', max(a, b), f'解析自「电机功率/扭矩」前{int(a)}/后{int(b)}kW 取大（蓝底纹）')
    print(f'      补 {len(l_ledger)} 格')
    ledgers.append((l_ledger, '功率解析'))

    # ---- 阶段K 细分市场推理 ----
    print('[K] 细分市场推理…')
    k_ledger = []
    # K1 同通用名称/车型名称唯一细分市场回填
    seg_maps = {'通用名称': defaultdict(set), '车型名称': defaultdict(set)}
    for _, rec in rows:
        seg = clean_txt(norm(rec.get('细分市场')))
        if not seg:
            continue
        for src in seg_maps:
            k = clean_txt(norm(rec.get(src)))
            if k:
                seg_maps[src][k].add(seg)
    seg_uniq = {src: {k: next(iter(v)) for k, v in m.items() if len(v) == 1} for src, m in seg_maps.items()}
    for r, rec in rows:
        if not fillable(rec.get('细分市场')):
            continue
        for src in ('通用名称', '车型名称'):
            k = clean_txt(norm(rec.get(src)))
            if k and k in seg_uniq[src]:
                seg = seg_uniq[src][k]
                cell = ws.cell(r, fi['细分市场'] + 1)
                cell.value = seg
                cell.fill = YELLOWGREEN
                rec['细分市场'] = seg
                stats['细分市场'] += 1
                k_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), '细分市场',
                                 '通用名称共识', '(空)', seg, f'推断：同{src}「{k}」已有唯一细分市场，回填（非实测）'])
                break
    n1 = len(k_ledger)
    print(f'      K1 历史/共识回填 {n1} 格')
    # K2 轴距分级
    wb_med = defaultdict(list)
    for _, rec in rows:
        g = norm(rec.get('通用名称'))
        w = cast_num(norm(rec.get('轴距(mm)')))
        if g and w and 2000 < w < 4000:
            wb_med[g].append(w)
    n2 = 0
    for r, rec in rows:
        if not fillable(rec.get('细分市场')):
            continue
        w = cast_num(norm(rec.get('轴距(mm)')))
        if w is None or not (2000 < w < 4000):
            ws_ = wb_med.get(norm(rec.get('通用名称')) or '')
            if ws_:
                w = sorted(ws_)[len(ws_) // 2]
        seg = classify_seg(w, norm(rec.get('产品名称')), norm(rec.get('通用名称')))
        if not seg:
            continue
        cell = ws.cell(r, fi['细分市场'] + 1)
        cell.value = seg
        cell.fill = BLUE
        rec['细分市场'] = seg
        stats['细分市场'] += 1
        n2 += 1
        basis = f'本行轴距{int(w)}' if cast_num(norm(rec.get('轴距(mm)'))) else f'同车系轴距中位数{int(w)}'
        k_ledger.append([rec['批次'], rec['产品型号'], str(rec['产品类型'] or ''), '细分市场',
                         '轴距分级', '(空)', seg, f'推断：{basis}+产品名称车身形式，按车型级别定义表分级（非实测）'])
    print(f'      K2 轴距分级 {n2} 格')
    ledgers.append((k_ledger, '细分市场推理'))

    # ---- 颜色说明图例（幂等） ----
    cs = wb['颜色说明']
    have = {str(row[1] or '') for row in cs.iter_rows(values_only=True) if len(row) > 1}
    for color_name, ctype, desc in (
            ('浅橙底纹', '企业名称推理',
             '产品商标/通用名称/型号代号在全表唯一对应一家企业时补空企业名称。推断依据：品牌注册与车系归属（非官方文件，详见变更记录）。'),
            ('蓝色底纹', '功率解析',
             '「电机功率/扭矩」列解析：单电机取该值作总功率；前/后双电机求和作总功率、取大作峰值功率（详见变更记录）。'),
            ('蓝色底纹', '轴距分级',
             '按「车型级别定义」表轴距阈值+产品名称车身形式分级补细分市场（详见变更记录）。')):
        if ctype not in have:
            cs.append([color_name, ctype, desc])

    total = sum(len(x[0]) for x in ledgers)
    # 保存前断言：任何单元格不得写入字面量 'None'
    bad = sum(1 for r in range(2, ws.max_row + 1) for c in ws[r]
              if isinstance(c.value, str) and c.value.strip() == 'None')
    if bad:
        raise SystemExit(f'发现 {bad} 个字面量 None 单元格，拒绝保存')
    for ledger, _ in ledgers:
        for row_vals in ledger:
            last_seq += 1
            cr.append([last_seq] + row_vals)
    print(f'\n台账追加 {total} 条（末序号 {last_seq}）')
    wb.save(SRC)
    print('已保存')

    print('\n===== 本轮补全统计 =====')
    for k, v in sorted(stats.items()):
        print(f'  {k}: +{v}')
    print(f'  合计 {sum(stats.values())} 格')


if __name__ == '__main__':
    sys.exit(main())
