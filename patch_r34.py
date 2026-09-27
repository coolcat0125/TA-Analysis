# -*- coding: utf-8 -*-
"""patch_r34.py — 向 build_report_v4925.py 注入 r3/r4 章节"""
p = 'build_report_v4925.py'
s = open(p, encoding='utf-8').read()

addition = '''

def _hhi(values):
    tot = sum(values)
    if not tot:
        return 0
    return sum(v / tot * 100 for v in values) ** 2 / 100


def build_r3(doc, rows):
    """第三章 驱动系统与能效（T24~T34）"""
    doc.add_heading('第三章 驱动系统与能效分析', level=1)
    stb = per_year_full(rows, 'BEV', '电机总功率/扭矩')
    stp = per_year_full(rows, 'PHEV/EREV', '电机总功率/扭矩')
    doc.add_heading('4.1 BEV电机功率趋势', level=2)
    add_table(doc, ['年份', '年均值', '同比增长率', '较2021年增幅'],
              [[y, fmt(stb[y]['mean']), '-' if i == 0 else f'{(stb[y]["mean"]/stb[ys[i-1]]["mean"]-1)*100:+.1f}%',
                f'{(stb[y]["mean"]/stb[ys[0]]["mean"]-1)*100:+.1f}%'] for i, (y, ys) in enumerate([(y, list(stb)) for y in stb])])
    doc.add_heading('4.2 PHEV电机功率趋势', level=2)
    add_table(doc, ['年份', '年均值', '同比增长率', '较2021年变化'],
              [[y, fmt(stp[y]['mean']), '-' if i == 0 else f'{(stp[y]["mean"]/stp[ys[i-1]]["mean"]-1)*100:+.1f}%',
                f'{(stp[y]["mean"]/stp[ys[0]]["mean"]-1)*100:+.1f}%'] for i, (y, ys) in enumerate([(y, list(stp)) for y in stp])])
    doc.add_heading('4.3 BEV vs PHEV电机功率对比', level=2)
    add_table(doc, ['年份', 'BEV均值', 'PHEV均值', '差值'],
              [[y, fmt(stb[y]['mean']), fmt(stp[y]['mean']), fmt(stb[y]['mean'] - stp[y]['mean'])] for y in sorted(stb) if y in stp])
    stcw_b = per_year_full(rows, 'BEV', '整备质量(kg)')
    stcw_p = per_year_full(rows, 'PHEV/EREV', '整备质量(kg)')
    doc.add_heading('4.4 整备质量趋势对比', level=2)
    yrs = sorted(set(stcw_b) & set(stcw_p))
    add_table(doc, ['年份', 'BEV均值', 'PHEV均值', '差值'],
              [[y, fmt(stcw_b[y]['mean']), fmt(stcw_p[y]['mean']), fmt(stcw_b[y]['mean'] - stcw_p[y]['mean'])] for y in yrs])

    def pq(pt):
        by = defaultdict(list)
        for r in rows:
            if r['PT'] == pt and r['year']:
                pw, cw = fnum(r.get('电机总功率/扭矩')), fnum(r.get('整备质量(kg)'))
                if pw and cw:
                    by[r['year']].append(pw / cw * 100)
        return {y: _stats_row(y, vs) for y, vs in sorted(by.items())}
    qp, qpb = pq('BEV'), pq('PHEV/EREV')
    doc.add_heading('4.5 功率质量比趋势', level=2)
    yrs = sorted(set(qp) & set(qpb))
    add_table(doc, ['年份', 'BEV(kW/100kg)', 'PHEV(kW/100kg)', '差值'],
              [[y, fmt(qp[y]['mean'], 2), fmt(qpb[y]['mean'], 2), fmt(qp[y]['mean'] - qpb[y]['mean'], 2)] for y in yrs])

    def ed(pt):
        by = defaultdict(list)
        for r in rows:
            if r['PT'] == pt and r['year']:
                rng, cap = fnum(r.get('纯电续航里程(km)')), fnum(r.get('电池容量(kWh)'))
                if rng and cap:
                    by[r['year']].append(rng / cap)
        return {y: _stats_row(y, vs) for y, vs in sorted(by.items())}
    eb = ed('BEV')
    doc.add_heading('4.6 整车能量密度趋势（BEV，续航/容量 km per kWh）', level=2)
    add_table(doc, ['年份', '均值(km/kWh)', '中位数', '样本数'],
              [[y, fmt(s['mean'], 2), fmt(s['median'], 2), s['n']] for y, s in eb.items()])
    sec = per_year_full(rows, 'BEV', '百公里电耗(kWh/100km)')
    doc.add_heading('4.7 BEV百公里电耗趋势', level=2)
    add_table(doc, ['年份', '均值', '中位数', 'P25'],
              [[y, fmt(s['mean'], 2), fmt(s['median'], 1), fmt(s['p25'], 1)] for y, s in sec.items()])

    def eff_by_bt(pt, bt_key):
        by = defaultdict(list)
        for r in rows:
            if r['PT'] == pt and r['year'] and bt_key in str(r.get('电池类型') or ''):
                v = fnum(r.get('百公里电耗(kWh/100km)'))
                if v:
                    by[r['year']].append(v)
        return {y: _stats_row(y, vs) for y, vs in sorted(by.items())}
    lfp, ncm = eff_by_bt('BEV', '磷酸铁锂'), eff_by_bt('BEV', '三元锂')
    yrs = sorted(set(lfp) & set(ncm))
    doc.add_heading('4.8 不同电池类型能效对比 (LFP vs NCM, BEV电耗)', level=2)
    add_table(doc, ['年份', 'LFP均值', 'NCM均值', '差值'],
              [[y, fmt(lfp[y]['mean'], 2), fmt(ncm[y]['mean'], 2), fmt(ncm[y]['mean'] - lfp[y]['mean'], 2)] for y in yrs])
    so = per_year_full(rows, 'PHEV/EREV', '综合油耗(L/100km)')
    sb = per_year_full(rows, 'PHEV/EREV', 'B状态油耗(L/100km)')
    doc.add_heading('4.9 PHEV综合油耗趋势', level=2)
    add_table(doc, ['年份', '均值', '中位数', '样本数'],
              [[y, fmt(s['mean'], 2), fmt(s['median'], 2), s['n']] for y, s in so.items()])
    doc.add_heading('4.10 PHEV B状态油耗趋势', level=2)
    add_table(doc, ['年份', '均值', '中位数', '样本数'],
              [[y, fmt(s['mean'], 2), fmt(s['median'], 2), s['n']] for y, s in sb.items()])

    def disp_bucket(pt, so_field):
        by = defaultdict(list)
        for r in rows:
            if r['PT'] == pt:
                dv, sv = fnum(r.get('发动机排量(mL)')), fnum(r.get(so_field))
                if dv and sv:
                    by['1.0-1.5L' if dv <= 1500 else ('1.5-2.0L' if dv <= 2000 else '2.0L以上')].append(sv)
        return by
    for so_field, lbl, num in (('综合油耗(L/100km)', '综合油耗', 11), ('B状态油耗(L/100km)', 'B状态油耗', 12)):
        by = disp_bucket('PHEV/EREV', so_field)
        doc.add_heading(f'4.{num} PHEV {lbl} 按排量区间', level=2)
        add_table(doc, ['排量区间', f'均值({lbl})', '中位数', '样本数'],
                  [[k, fmt(_stats_row(k, vs)['mean'], 2), fmt(_stats_row(k, vs)['median'], 2), len(vs)] for k, vs in sorted(by.items())])
    doc.add_heading('4.13 PHEV纯电续航里程趋势', level=2)
    add_table(doc, ['年份', '均值', '中位数', '样本数'],
              [[y, fmt(s['mean']), fmt(s['median'], 0), s['n']] for y, s in per_year_full(rows, 'PHEV/EREV', '纯电续航里程(km)').items()])
    byxy = defaultdict(lambda: defaultdict(list))
    for r in rows:
        if r['PT'] == 'PHEV/EREV' and r['year']:
            dv, sv = fnum(r.get('发动机排量(mL)')), fnum(r.get('综合油耗(L/100km)'))
            if dv and sv:
                byxy[r['year']]['1.0-1.5L' if dv <= 1500 else ('1.5-2.0L' if dv <= 2000 else '2.0L以上')].append(sv)
    doc.add_heading('4.14 PHEV综合油耗 - 排量×年份交叉分析', level=2)
    buckets = ['1.0-1.5L', '1.5-2.0L', '2.0L以上']
    add_table(doc, ['年份'] + buckets,
              [[y] + [fmt(_stats_row(y, byxy[y][k])['mean'], 2) if byxy[y].get(k) else '-' for k in buckets] for y in sorted(byxy)])


def build_r4(doc, rows):
    """第四章 供应商格局分析"""
    from collections import Counter
    doc.add_heading('第四章 供应商格局分析', level=1)

    def by_year_counter(pt, field):
        by = defaultdict(Counter)
        for r in rows:
            if r['PT'] == pt and r['year']:
                sup = str(r.get(field) or '').strip()
                sup = re.split(r'[,，;；/]', sup)[0].strip() if sup else ''
                if sup and sup not in ('未知', '不适用'):
                    by[r['year']][sup] += 1
        return by

    def rank_table(by_year, title):
        agg = Counter()
        for y in by_year:
            agg.update(by_year[y])
        doc.add_heading(title, level=2)
        add_table(doc, ['排名', '供应商', '累计配套款数'],
                  [[i + 1, k, v] for i, (k, v) in enumerate(agg.most_common(10))])

    def hhi_table(by_year, title):
        doc.add_heading(title, level=2)
        add_table(doc, ['年份', 'HHI', 'CR5'],
                  [[y, fmt(_hhi([v for _, v in by_year[y].most_common()]), 0),
                    fmt(sum(v for _, v in by_year[y].most_common()[:5]) / sum(by_year[y].values()) * 100, 1) + '%']
                   for y in sorted(by_year) if by_year[y] and sum(by_year[y].values())])

    bb = by_year_counter('BEV', '电芯供应商')
    rank_table(bb, '5.1 Top 10 电池供应商排名（BEV）')
    hhi_table(bb, '5.1b 电池供应商 HHI/CR5 趋势（BEV）')
    bm = by_year_counter('BEV', '电机生产企业')
    rank_table(bm, '5.2 Top 10 电机供应商排名（BEV）')
    hhi_table(bm, '5.2b 电机供应商 HHI/CR5 趋势（BEV）')
    combo = Counter()
    for r in rows:
        if r['PT'] == 'BEV':
            cell = str(r.get('电芯供应商') or '').split(',')[0].strip()
            mot = str(r.get('电机生产企业') or '').split(',')[0].strip()
            if cell and mot:
                combo[(cell, mot)] += 1
    doc.add_heading('5.6 Top 20 电池-电机供应商组合（BEV）', level=2)
    add_table(doc, ['排名', '电芯供应商', '电机供应商', '款数'],
              [[i + 1, k[0][:24], k[1][:24], v] for i, (k, v) in enumerate(combo.most_common(20))])
    pe = by_year_counter('PHEV/EREV', '发动机生产企业')
    rank_table(pe, '5.3 Top 10 发动机供应商排名（PHEV/EREV）')
    hhi_table(pe, '5.3b 发动机供应商 HHI/CR5 趋势（PHEV/EREV）')
'''
anchor = "def main():"
assert anchor in s
s = s.replace(anchor, addition + "\n\n" + anchor, 1)
old_stage = """    if 'r2' in st:
        build_r2(doc, rows)"""
new_stage = """    if 'r2' in st:
        build_r2(doc, rows)
    if 'r3' in st:
        build_r3(doc, rows)
    if 'r4' in st:
        build_r4(doc, rows)"""
assert old_stage in s
s = s.replace(old_stage, new_stage, 1)
open(p, 'w', encoding='utf-8').write(s)
print('r3/r4 injected')
