# -*- coding: utf-8 -*-
"""build_report_v4925.py — 行业分析报告 docx 重建生成器（v4.9.25 口径）

数据源：NEV公告参数汇总表_合并版（341~410批）.xlsx（5,392×38）+ 批次时间表（公告年月→年份）
结构：沿用旧报告骨架（audit-output/old_report_docx_snapshot_20260927.json，63 表/75 标题）
口径：全部表格由本脚本从权威底表重算（表值可复算、无手抄）；叙述段数字留 [待人工校订] 标记

分期：r1=执行摘要+第一章（T00~T07）；r2=第二章电池（T08~T23）；r3=第三~七章。
用法：python build_report_v4925.py --out <docx路径> [--stages r1,r2,r3]
"""
import argparse
import json
import os
import re
from collections import defaultdict, Counter

from docx.oxml.ns import qn

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')

import os


def load_data():
    """返回 rows(list of dict) + 批次→公告年 映射。动力类型归一（泄漏值→BEV/PHEV）。"""
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}

    ts = wb['批次时间表']
    batch_year = {}
    for r in ts.iter_rows(min_row=2, values_only=True):
        if r[0] is not None:
            ym = str(r[1] or '')[:4]
            if ym.isdigit():
                batch_year[str(r[0])] = int(ym)

    def pt_norm(v):
        s = str(v or '')
        if '纯电' in s or 'BEV' in s:
            return 'BEV'
        if '插电' in s or '增程' in s or 'PHEV' in s:
            return 'PHEV/EREV'
        return None

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = {'批次': str(r[hi['批次']] or '').strip(), 'year': batch_year.get(str(r[hi['批次']] or '').strip())}
        for f in hdr:
            d[f] = r[hi[f]] if hi.get(f) is not None else None
        d['PT'] = pt_norm(r[hi['动力类型']])
        rows.append(d)
    wb.close()
    return rows


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', ''))
    except Exception:
        return None


def per_year(rows, pt, field, agg='mean'):
    by = defaultdict(list)
    for r in rows:
        if r['PT'] == pt and r['year']:
            v = fnum(r.get(field))
            if v is not None:
                by[r['year']].append(v)
    out = {}
    for y, vs in sorted(by.items()):
        vs.sort()
        n = len(vs)
        mean = sum(vs) / n
        med = vs[n // 2] if n % 2 else (vs[n // 2 - 1] + vs[n // 2]) / 2
        p25 = vs[int(n * 0.25)]
        out[y] = {'mean': mean, 'median': med, 'p25': p25, 'n': n}
    return out


def add_table(doc, header, rows_data):
    t = doc.add_table(rows=1 + len(rows_data), cols=len(header))
    t.style = 'Table Grid'
    rows_data = [list(rd[:len(header)]) + [''] * (len(header) - len(rd)) for rd in rows_data]
    for j, h in enumerate(header):
        c = t.rows[0].cells[j]
        c.text = str(h)
        for p in c.paragraphs:
            for run in p.runs:
                run.font.bold = True
    for i, rd in enumerate(rows_data, 1):
        for j, v in enumerate(rd):
            t.rows[i].cells[j].text = str(v)
    return t


def set_font(run, name='微软雅黑', size=10.5):
    run.font.name = name
    run.font.size = Pt(size)
    run._element.rPr.rFonts.set(qn('rFonts:eastAsia'), name)


def build_r1(doc, rows):
    """执行摘要 + 第一章（T00 概览 / T03 年度批次 / T05 BEV占比 / T06-T07 细分市场）"""
    from collections import Counter
    doc.add_heading('执行摘要', level=1)
    doc.add_heading('报告数据概览', level=2)
    n_total = len(rows)
    bev = sum(1 for r in rows if r['PT'] == 'BEV')
    phev = sum(1 for r in rows if r['PT'] == 'PHEV/EREV')
    years = sorted({r['year'] for r in rows if r['year']})
    brands = len({str(r.get('品牌') or '') for r in rows if r.get('品牌')})
    ents = len({str(r.get('企业名称') or '') for r in rows if r.get('企业名称')})
    T00 = [
        ['总记录数', f'{n_total:,}', '第341-411批公告车型（乘用车纯净口径）'],
        ['BEV车型数', f'{bev:,}', f'纯电动车型，占比{bev/n_total*100:.1f}%'],
        ['PHEV/EREV车型数', f'{phev:,}', f'插电混动/增程车型，占比{phev/n_total*100:.1f}%'],
        ['年份范围', f'{years[0]}-{years[-1]}', f'{len(years)}年数据跨度'],
        ['品牌数', f'{brands}', '涵盖主流及新兴品牌'],
        ['企业数', f'{ents}', '生产企业覆盖面'],
    ]
    doc.add_paragraph('以下概览数据由权威底表全量重算（v4.9.25 口径）。')
    add_table(doc, ['维度', '数值', '说明'], T00)

    doc.add_heading('核心指标演变: 2021年 vs 2025年', level=2)
    doc.add_paragraph('[待人工校订：2021 vs 2025 对比表在 r2 补齐 per-year 统计后回填]')
    add_table(doc, ['指标', '2021年', '2025年', '变化值'], [['（r2 回填）', '-', '-', '-']])

    doc.add_heading('六大核心发现', level=2)
    doc.add_paragraph('[待人工校订：核心发现叙述基于 r2 全量统计后归纳]')

    doc.add_heading('第一章 市场全景分析', level=1)
    doc.add_heading('2.1 数据集概述', level=2)
    doc.add_paragraph('本报告数据集为工信部《道路机动车辆生产企业及产品》公告第 341~411 批新能源乘用车纯净口径，共 5,392 款，覆盖 2021-2026 年公告窗口。')
    add_table(doc, ['维度', '数值', '说明'], T00)

    doc.add_heading('2.2 动力类型分布', level=2)
    doc.add_paragraph(f'BEV {bev:,} 款（{bev/n_total*100:.1f}%）、PHEV/EREV {phev:,} 款（{phev/n_total*100:.1f}%）。')
    doc.add_heading('2.3 年度公告趋势', level=2)
    by_year = Counter(r['year'] for r in rows if r['year'])
    by_pt = Counter((r['year'], r['PT']) for r in rows if r['year'])
    T03 = []
    T05 = []
    prev = None
    for y in sorted(by_year):
        tot = by_year[y]
        b = by_pt.get((y, 'BEV'), 0)
        p = by_pt.get((y, 'PHEV/EREV'), 0)
        T03.append([y, b, p, tot])
        T05.append([y, b, p, f'{b/tot*100:.1f}%'])
    add_table(doc, ['年份', 'BEV', 'PHEV/EREV', '合计'], T03)
    doc.add_heading('2.3.1 年度公告量与同比增长', level=3)
    T04 = []
    for i, row in enumerate(T03):
        y, b, p, tot = row
        grow = f'{tot - T03[i-1][3]:+d}' if i > 0 else '-'
        rate = f'{(tot - T03[i-1][3]) / T03[i-1][3] * 100:+.1f}%' if i > 0 else '-'
        T04.append([y, tot, grow, rate])
    add_table(doc, ['年份', '公告数量', '同比增长', '增长率'], T04)
    doc.add_heading('2.4 BEV/PHEV占比变化', level=2)
    add_table(doc, ['年份', 'BEV数量', 'PHEV/EREV数量', 'BEV占比'], T05)
    doc.add_heading('2.5 细分市场分布 - BEV', level=2)
    for ptname, tbl in (('BEV', 'T06'), ('PHEV/EREV', 'T07'),):
        seg = Counter(str(r.get('细分市场') or '未知') for r in rows if r['PT'] == ptname)
        data = [[k, v, f'{v/sum(seg.values())*100:.1f}%'] for k, v in seg.most_common(15)]
        doc.add_heading(f'2.{"5" if ptname=="BEV" else "6"} 细分市场分布 - {ptname}', level=2)
        add_table(doc, ['细分市场', '样本数', '占比'], data)




def _stats_row(y, vs):
    vs = sorted(vs)
    n = len(vs)
    med = vs[n // 2] if n % 2 else (vs[n // 2 - 1] + vs[n // 2]) / 2
    return {'mean': sum(vs) / n, 'median': med, 'p25': vs[int(n * 0.25)], 'n': n}


def per_year_full(rows, pt, field):
    by = defaultdict(list)
    for r in rows:
        if r['PT'] == pt and r['year']:
            v = fnum(r.get(field))
            if v is not None:
                by[r['year']].append(v)
    return {y: _stats_row(y, vs) for y, vs in sorted(by.items())}


def fmt(v, nd=1):
    return f'{v:,.{nd}f}' if nd else f'{v:,.0f}'


def build_r2(doc, rows):
    doc.add_heading('第二章 动力电池深度分析', level=1)
    ptmap = [('BEV', 'BEV'), ('PHEV/EREV', 'PHEV')]

    # T08/T09 续航趋势（均值/中位数/P25/样本）
    for pt, lbl in ptmap:
        st = per_year_full(rows, pt, '纯电续航里程(km)')
        doc.add_heading(f'3.{1 if pt=="BEV" else 2} {lbl}续航里程趋势', level=2)
        add_table(doc, ['年份', '均值', '中位数', 'P25', '样本'],
                  [[y, fmt(s['mean']), fmt(s['median'], 0), fmt(s['p25'], 0), s['n']] for y, s in st.items()])
    # T10 对比
    stb, stp = per_year_full(rows, 'BEV', '纯电续航里程(km)'), per_year_full(rows, 'PHEV/EREV', '纯电续航里程(km)')
    doc.add_heading('3.3 BEV vs PHEV续航对比', level=2)
    add_table(doc, ['年份', 'BEV均值', 'PHEV均值', '差值'],
              [[y, fmt(stb[y]['mean']), fmt(stp[y]['mean']), fmt(stb[y]['mean'] - stp[y]['mean'])] for y in sorted(stb) if y in stp])
    # T11-T13 容量
    stb2, stp2 = per_year_full(rows, 'BEV', '电池容量(kWh)'), per_year_full(rows, 'PHEV/EREV', '电池容量(kWh)')
    doc.add_heading('3.4 BEV电池容量趋势', level=2)
    add_table(doc, ['年份', '年均值', '同比增长率', '较2021年增幅'],
              [[y, fmt(stb2[y]['mean']), '-' if i == 0 else f'{(stb2[y]["mean"]/stb2[ys[i-1]]["mean"]-1)*100:+.1f}%',
                f'{(stb2[y]["mean"]/stb2[ys[0]]["mean"]-1)*100:+.1f}%'] for i, (y, ys) in enumerate([(y, list(stb2)) for y in stb2])])
    doc.add_heading('3.5 PHEV电池容量趋势', level=2)
    add_table(doc, ['年份', '年均值', '同比增长率', '较2021年变化'],
              [[y, fmt(stp2[y]['mean']), '-' if i == 0 else f'{(stp2[y]["mean"]/stp2[ys[i-1]]["mean"]-1)*100:+.1f}%',
                f'{(stp2[y]["mean"]/stp2[ys[0]]["mean"]-1)*100:+.1f}%'] for i, (y, ys) in enumerate([(y, list(stp2)) for y in stp2])])
    doc.add_heading('3.6 BEV vs PHEV电池容量对比', level=2)
    add_table(doc, ['年份', 'BEV均值', 'PHEV均值', '差值'],
              [[y, fmt(stb2[y]['mean']), fmt(stp2[y]['mean']), fmt(stb2[y]['mean'] - stp2[y]['mean'])] for y in sorted(stb2) if y in stp2])
    # T14-T16 电池类型趋势
    def bt_by_year(pt):
        by = defaultdict(Counter)
        tot = Counter()
        for r in rows:
            if r['PT'] == pt and r['year']:
                bt = str(r.get('电池类型') or '').strip()
                bt = bt.replace('电池', '')
                if bt and bt not in ('未知',):
                    by[r['year']][bt] += 1
                    tot[r['year']] += 1
        return by, tot
    for pt, lbl, topn in (('BEV', 'BEV', 6), ('PHEV/EREV', 'PHEV', 5)):
        by, tot = bt_by_year(pt)
        years = sorted(tot)
        tops = Counter()
        for y in years:
            for bt, n in by[y].most_common(topn):
                tops[bt] += n
        keep = [bt for bt, _ in tops.most_common(topn)]
        doc.add_heading(f'3.{7 if pt=="BEV" else 8} 电池类型占比趋势 - {lbl}', level=2)
        add_table(doc, ['电池类型'] + [str(y) + '%' for y in years],
                  [[bt] + [f'{by[y].get(bt,0)/tot[y]*100:.1f}%' if tot[y] else '-' for y in years] for bt in keep])
    doc.add_heading('3.9 电池类型数量趋势 - BEV', level=2)
    by, tot = bt_by_year('BEV')
    add_table(doc, ['电池类型'] + [str(y) for y in sorted(tot)],
              [[bt] + [by[y].get(bt, 0) for y in sorted(tot)] for bt in sorted({b for y in by for b in by[y]})])
    # T17 能量密度
    st_ed = per_year_full(rows, 'BEV', '电池能量密度(Wh/kg)')
    doc.add_heading('3.10 BEV电池能量密度趋势', level=2)
    add_table(doc, ['年份', '均值', '中位数', '样本数'],
              [[y, fmt(s['mean'], 2), fmt(s['median'], 1), s['n']] for y, s in st_ed.items()])
    # T18-T21 分布
    def dist(rows, pt, field, edges, labels):
        vs = [fnum(r.get(field)) for r in rows if r['PT'] == pt]
        vs = [v for v in vs if v is not None]
        cnt = Counter()
        for v in vs:
            for i, (lo, hi_) in enumerate(zip(edges, edges[1:])):
                if lo <= v < hi_:
                    cnt[i] += 1
                    break
            else:
                cnt[len(edges) - 2] += 1
        total = sum(cnt.values())
        cum = 0
        out = []
        for i in range(len(edges) - 1):
            cum += cnt[i]
            out.append([labels[i], cnt[i], f'{cnt[i]/total*100:.1f}%' if total else '-', f'{cum/total*100:.1f}%' if total else '-'])
        return out
    doc.add_heading('3.11 BEV续航里程分布', level=2)
    add_table(doc, ['续航区间', '车型数', '占比', '累计占比'],
              dist(rows, 'BEV', '纯电续航里程(km)', [100, 150, 200, 250, 300, 350, 400, 450, 500, 600, 700, 800],
                   ['100-150km', '150-200km', '200-250km', '250-300km', '300-350km', '350-400km', '400-450km', '450-500km', '500-600km', '600-700km', '700-800km', '800km以上']))
    doc.add_heading('3.12 PHEV续航里程分布', level=2)
    add_table(doc, ['续航区间', '车型数', '占比', '累计占比'],
              dist(rows, 'PHEV/EREV', '纯电续航里程(km)', [25, 50, 75, 100, 125, 150, 200],
                   ['25-50km', '50-75km', '75-100km', '100-125km', '125-150km', '150-200km', '200km以上']))
    doc.add_heading('3.13 BEV电池容量分布', level=2)
    add_table(doc, ['容量区间', '车型数', '占比'],
              dist(rows, 'BEV', '电池容量(kWh)', [10, 20, 30, 40, 50, 60, 70, 80, 100],
                   ['10-20kWh', '20-30kWh', '30-40kWh', '40-50kWh', '50-60kWh', '60-70kWh', '70-80kWh', '80-100kWh'])[:8])
    doc.add_heading('3.14 PHEV电池容量分布', level=2)
    add_table(doc, ['容量区间', '车型数', '占比'],
              dist(rows, 'PHEV/EREV', '电池容量(kWh)', [0, 10, 20, 30, 40, 50],
                   ['0-10kWh', '10-20kWh', '20-30kWh', '30-40kWh', '40-50kWh', '50kWh以上'])[:6])
    # T22-T23 象限
    for pt, lbl in (('BEV', 'BEV'), ('PHEV/EREV', 'PHEV')):
        rng = [fnum(r.get('纯电续航里程(km)')) for r in rows if r['PT'] == pt]
        cap = [fnum(r.get('电池容量(kWh)')) for r in rows if r['PT'] == pt]
        rng = [v for v in rng if v]; cap = [v for v in cap if v]
        if not rng or not cap:
            continue
        import statistics
        med_r = statistics.median(rng); med_c = statistics.median(cap)
        q = defaultdict(list)
        for r in rows:
            if r['PT'] != pt: continue
            rv, cv = fnum(r.get('纯电续航里程(km)')), fnum(r.get('电池容量(kWh)'))
            if rv is None or cv is None: continue
            qr = '高续航' if rv >= med_r else '低续航'
            qc = '高容量' if cv >= med_c else '低容量'
            q[f'Q{("高续航高容量 高续航低容量 低续航高容量 低续航低容量".split().index(qr+qc))+1}_{qr}{qc}'].append((rv, cv))
        rows_q = []
        for name in sorted(q):
            xs = q[name]
            rows_q.append([name, len(xs), f'{len(xs)/sum(len(v) for v in q.values())*100:.1f}%',
                           fmt(sum(x[0] for x in xs)/len(xs)), fmt(sum(x[1] for x in xs)/len(xs))])
        doc.add_heading(f'3.15 技术路线矩阵 - {lbl}象限分析' if pt == 'BEV' else f'3.15b 技术路线矩阵 - {lbl}象限分析', level=2)
        add_table(doc, ['象限', '数量', '占比', '平均续航(km)', '平均容量(kWh)'], rows_q)




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
                    fmt(sum(x for x in (fnum(r.get('纯电续航里程(km)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('纯电续航里程(km)'))))),
                    fmt(sum(x for x in (fnum(r.get('电池容量(kWh)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('电池容量(kWh)'))))),
                    fmt(sum(x for x in (fnum(r.get('整备质量(kg)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('整备质量(kg)')))))]
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
              [[k, len(v), fmt(sum(x for x in (fnum(r.get('整备质量(kg)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('整备质量(kg)')))))]
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', required=True)
    ap.add_argument('--stages', default='r1')
    args = ap.parse_args()
    rows = load_data()
    doc = Document()
    st = args.stages.split(',')
    if 'r1' in st:
        build_r1(doc, rows)
    if 'r2' in st:
        build_r2(doc, rows)
    if 'r3' in st:
        build_r3(doc, rows)
    if 'r4' in st:
        build_r4(doc, rows)
    if 'r5' in st:
        build_r5(doc, rows)
    doc.save(args.out)
    print(f'saved: {args.out} (stages={args.stages}, rows={len(rows)})')


if __name__ == '__main__':
    main()
