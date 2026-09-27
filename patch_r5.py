# -*- coding: utf-8 -*-
"""patch_r5.py — 注入 r5（第五~七章：细分市场/相关性/总结）"""
p = 'build_report_v4925.py'
s = open(p, encoding='utf-8').read()

addition = '''

def build_r5(doc, rows):
    """第五章 细分市场 + 第六章 相关性与增长 + 第七章 总结"""
    import statistics
    from collections import Counter
    doc.add_heading('第五章 细分市场深度分析', level=1)
    for pt, lbl in (('BEV', 'BEV'), ('PHEV/EREV', 'PHEV')):
        seg = defaultdict(list)
        for r in rows:
            if r['PT'] == pt:
                k = str(r.get('细分市场') or '未知')
                seg[k].append(r)
        doc.add_heading(f'6.{1 if pt=="BEV" else 2} {lbl}细分市场统计', level=2)
        add_table(doc, ['细分市场', '样本数', '续航均值(km)', '容量均值(kWh)', '整备均值(kg)'],
                  [[k, len(v),
                    fmt(sum(x for x in (fnum(r.get('纯电续航里程(km)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('纯电续航里程(km)')))),
                    fmt(sum(x for x in (fnum(r.get('电池容量(kWh)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('电池容量(kWh)')))),
                    fmt(sum(x for x in (fnum(r.get('整备质量(kg)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('整备质量(kg)'))))]
                   for k, v in sorted(seg.items(), key=lambda kv: -len(kv[1]))])
    # 6.3/6.4 BEV 2023 vs 2025 对比（按细分市场）
    seg_b = defaultdict(list)
    for r in rows:
        if r['PT'] == 'BEV' and r['year'] in (2023, 2025):
            seg_b[(str(r.get('细分市场') or '未知'), r['year'])].append(r)
    segs = sorted({k[0] for k in seg_b})
    doc.add_heading('6.3 BEV细分市场续航对比 - 2023年 vs 2025年', level=2)
    rows63 = []
    for sgm in segs:
        v23 = [fnum(r.get('纯电续航里程(km)')) for r in seg_b.get((sgm, 2023), [])]
        v25 = [fnum(r.get('纯电续航里程(km)')) for r in seg_b.get((sgm, 2025), [])]
        v23 = [x for x in v23 if x]; v25 = [x for x in v25 if x]
        if v23 or v25:
            rows63.append([sgm, fmt(sum(v23)/len(v23)) if v23 else '-', fmt(sum(v25)/len(v25)) if v25 else '-',
                           len(v23), len(v25)])
    add_table(doc, ['细分市场', '2023续航均值', '2025续航均值', '2023样本', '2025样本'], rows63)
    doc.add_heading('6.4 BEV细分市场容量对比 - 2023年 vs 2025年', level=2)
    rows64 = []
    for sgm in segs:
        v23 = [fnum(r.get('电池容量(kWh)')) for r in seg_b.get((sgm, 2023), [])]
        v25 = [fnum(r.get('电池容量(kWh)')) for r in seg_b.get((sgm, 2025), [])]
        v23 = [x for x in v23 if x]; v25 = [x for x in v25 if x]
        if v23 or v25:
            rows64.append([sgm, fmt(sum(v23)/len(v23)) if v23 else '-', fmt(sum(v25)/len(v25)) if v25 else '-',
                           len(v23), len(v25)])
    add_table(doc, ['细分市场', '2023容量均值', '2025容量均值', '2023样本', '2025样本'], rows64)
    doc.add_heading('6.5 PHEV细分市场续航对比', level=2)
    seg_p = defaultdict(list)
    for r in rows:
        if r['PT'] == 'PHEV/EREV':
            seg_p[str(r.get('细分市场') or '未知')].append(r)
    add_table(doc, ['细分市场', '样本数', '纯电续航均值(km)'],
              [[k, len(v), fmt(sum(x for x in (fnum(r.get('纯电续航里程(km)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('纯电续航里程(km)')))))]
               for k, v in sorted(seg_p.items(), key=lambda kv: -len(kv[1]))])
    doc.add_heading('6.6 整备质量按细分市场分布', level=2)
    seg_all = defaultdict(list)
    for r in rows:
        seg_all[str(r.get('细分市场') or '未知')].append(r)
    add_table(doc, ['细分市场', '样本数', '整备均值(kg)'],
              [[k, len(v), fmt(sum(x for x in (fnum(r.get('整备质量(kg)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('整备质量(kg)'))))]
               for k, v in sorted(seg_all.items(), key=lambda kv: -len(kv[1]))[:15]])

    doc.add_heading('第六章 相关性与增长分析', level=1)
    METRICS = [('纯电续航里程(km)', '续航'), ('电池容量(kWh)', '容量'), ('百公里电耗(kWh/100km)', '电耗'),
               ('电机总功率/扭矩', '总功率'), ('整备质量(kg)', '整备')]

    def pearson(xs, ys):
        n = len(xs)
        if n < 3:
            return None
        mx, my = sum(xs)/n, sum(ys)/n
        num = sum((x-mx)*(y-my) for x, y in zip(xs, ys))
        dx = (sum((x-mx)**2 for x in xs)) ** 0.5
        dy = (sum((y-my)**2 for y in ys)) ** 0.5
        return num/(dx*dy) if dx and dy else None

    for pt, lbl, num in (('BEV', 'BEV', 1), ('PHEV/EREV', 'PHEV', 2)):
        subset = [r for r in rows if r['PT'] == pt]
        vals = {label: [fnum(r.get(f)) for r in subset] for f, label in METRICS}
        keep = [i for i in range(len(subset)) if all(vals[label][i] is not None for _, label in METRICS)]
        doc.add_heading(f'7.{num} {lbl}关键指标相关性矩阵（完整样本 n={len(keep)}）', level=2)
        header = ['指标'] + [label for _, label in METRICS]
        matrix = []
        for _, rl in METRICS:
            row = [rl]
            for _, cl in METRICS:
                r12 = pearson([vals[rl][i] for i in keep], [vals[cl][i] for i in keep])
                row.append(fmt(r12, 2) if r12 is not None else '-')
            matrix.append(row)
        add_table(doc, header, matrix)
        # 同比增长率汇总（核心指标年均值 YoY）
        doc.add_heading(f'7.{num+2} {lbl}同比增长率汇总', level=2)
        st_all = per_year_full(rows, pt, '纯电续航里程(km)')
        sc = per_year_full(rows, pt, '电池容量(kWh)')
        sp = per_year_full(rows, pt, '电机总功率/扭矩')
        yoy = []
        years = sorted(set(st_all) & set(sc) & set(sp))
        for i, y in enumerate(years):
            if i == 0:
                continue
            y0 = years[i-1]
            yoy.append([y,
                        fmt((st_all[y]['mean']/st_all[y0]['mean']-1)*100, 1) + '%',
                        fmt((sc[y]['mean']/sc[y0]['mean']-1)*100, 1) + '%',
                        fmt((sp[y]['mean']/sp[y0]['mean']-1)*100, 1) + '%'])
        add_table(doc, ['年份', '续航YoY', '容量YoY', '总功率YoY'], yoy)
        # 年度综合指标
        doc.add_heading(f'7.{num+4} {lbl}年度综合指标', level=2)
        add_table(doc, ['年份', '续航均值', '容量均值', '总功率均值', '样本数'],
                  [[y, fmt(st_all[y]['mean']), fmt(sc[y]['mean']), fmt(sp[y]['mean']), st_all[y]['n']] for y in years])
        # 描述性统计
        doc.add_heading(f'7.{num+6} {lbl}关键指标描述性统计（全期）', level=2)
        desc = []
        for f, label in METRICS:
            vs = sorted(x for x in (fnum(r.get(f)) for r in subset) if x is not None)
            if not vs:
                continue
            n = len(vs)
            desc.append([label, n, fmt(sum(vs)/n, 1), fmt(statistics.median(vs), 1),
                         fmt(vs[int(n*0.05)], 1), fmt(vs[int(n*0.95)-1], 1), fmt(vs[0], 1), fmt(vs[-1], 1)])
        add_table(doc, ['指标', '样本', '均值', '中位数', 'P5', 'P95', '最小', '最大'], desc)

    doc.add_heading('第七章 总结与展望', level=1)
    doc.add_heading('8.1 核心发现汇总', level=2)
    doc.add_paragraph('[待人工校订：基于上述全量统计的核心发现叙述]')
    doc.add_heading('8.2 BEV技术演进总结', level=2)
    doc.add_paragraph('[待人工校订]')
    doc.add_heading('8.3 PHEV技术演进总结', level=2)
    doc.add_paragraph('[待人工校订]')
    doc.add_heading('8.4 供应商格局演变总结', level=2)
    doc.add_paragraph('[待人工校订]')
    doc.add_heading('8.5 市场趋势展望', level=2)
    doc.add_paragraph('[待人工校订]')
'''
anchor = "def main():"
assert anchor in s
s = s.replace(anchor, addition + "\n\n" + anchor, 1)
old_stage = """    if 'r4' in st:
        build_r4(doc, rows)"""
new_stage = """    if 'r4' in st:
        build_r4(doc, rows)
    if 'r5' in st:
        build_r5(doc, rows)"""
assert old_stage in s
s = s.replace(old_stage, new_stage, 1)
open(p, 'w', encoding='utf-8').write(s)
print('r5 injected')
