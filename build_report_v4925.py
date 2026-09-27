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
    doc.save(args.out)
    print(f'saved: {args.out} (stages={args.stages}, rows={len(rows)})')


if __name__ == '__main__':
    main()
