#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把《NEV公告车型行业分析_数据底表.xlsx》抽取为统一 JSON，供三款候选看板共用。"""
import json, math, os, sys
import openpyxl

SRC = os.path.join(os.path.dirname(__file__), '..', 'NEV公告车型行业分析_数据底表.xlsx')
OUT = os.path.join(os.path.dirname(__file__), 'base_data.json')

wb = openpyxl.load_workbook(SRC, data_only=True)

def rows(sheet):
    ws = wb[sheet]
    return [list(r) for r in ws.iter_rows(values_only=True)]

def num(v):
    if v is None: return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else v
    try:
        return float(str(v).replace('%', '').replace(',', ''))
    except: return None

def table(sheet, header_row, first_col_label=None):
    """从 header_row（0 基）读取表头，直到空行为止。"""
    rs = rows(sheet)
    hdr = rs[header_row]
    out = []
    for r in rs[header_row + 1:]:
        if r[0] is None: break
        out.append(r)
    return hdr, out

def yearly_series(sheet, header_row=2):
    hdr, body = table(sheet, header_row)
    years = [int(num(r[0])) for r in body]
    cols = {}
    for i, h in enumerate(hdr[1:], start=1):
        if h is None: continue
        cols[str(h)] = [num(r[i]) for r in body]
    return {'years': years, 'cols': cols}

D = {}
D['meta'] = {
    'title': 'NEV公告车型行业分析',
    'baseline': 'v4.3.6 · 4,613 款纯新能源乘用车 · 第341~410批（2021–2026）',
    'source': '工信部《道路机动车辆生产企业及产品公告》',
    'note': '全部指标为均值口径；供应商分析基于电芯/电机有值子集（电芯覆盖29.5%）；410批为公示占位计入2026',
}

# 执行摘要
hdr, body = table('执行摘要', 3)
D['summary'] = [{'k': r[0], 'v': num(r[1]), 'u': r[2], 'note': r[3]} for r in body]

# 年度分布
D['yearly'] = yearly_series('年度分布', 3)

# 续航
D['range_bev'] = yearly_series('BEV续航分析')
D['range_phev'] = yearly_series('PHEV续航分析')

# 年度综合指标（续航/容量/功率/质量/整车密度）
D['annual'] = yearly_series('年度综合指标')
D['capacity'] = yearly_series('电池容量分析')
D['power'] = yearly_series('电机功率分析')
D['mass'] = yearly_series('整备质量分析')
D['pw_ratio'] = yearly_series('功率质量比')
D['veh_density'] = yearly_series('整车能量密度')
D['cell_density'] = yearly_series('电池能量密度')
D['efficiency'] = yearly_series('能效分析', 3)
D['phev_fuel'] = yearly_series('PHEV油耗分析')

# 电池类型趋势
D['bt_bev'] = yearly_series('电池类型趋势-BEV', 3)
D['bt_phev'] = yearly_series('电池类型趋势-PHEV', 3)

# 分布（按年度桶）
def dist(sheet):
    hdr, body = table(sheet, 3)
    order_key = lambda h: (0 if h.startswith('<') else 2 if h.startswith('≥') else 1,
                           num(h.replace('<', '').replace('≥', '').split('-')[0]) or 0)
    buckets = [h for h in hdr[1:] if h]
    idx = {h: i + 1 for i, h in enumerate(hdr[1:]) if h}
    buckets.sort(key=order_key)
    return {'years': [int(num(r[0])) for r in body], 'buckets': buckets,
            'data': {b: [num(r[idx[b]]) for r in body] for b in buckets}}
D['dist_range_bev'] = dist('续航分布-BEV')
D['dist_range_phev'] = dist('续航分布-PHEV')
D['dist_cap_bev'] = dist('BEV容量分布')
D['dist_cap_phev'] = dist('PHEV容量分布')

# 技术路线矩阵
hdr, body = table('技术路线矩阵', 3)
D['quadrant'] = {'note': rows('技术路线矩阵')[2][0], 'items': [{'name': r[0], 'n': num(r[1]), 'pct': num(r[2])} for r in body]}

# 供应商
D['hhi'] = yearly_series('供应商集中度', 3)
D['cr_batt'] = yearly_series('电池供应商CR')
D['cr_motor'] = yearly_series('电机供应商CR')
def rank(sheet):
    hdr, body = table(sheet, 2)
    return [{'name': r[1], 'n': num(r[2])} for r in body]
D['rank_batt'] = rank('电池供应商排名')
D['rank_motor'] = rank('电机供应商排名')
hdr, body = table('供应商组合', 2)
D['combo'] = [{'name': r[0], 'n': num(r[1])} for r in body]
def matrix(sheet):
    hdr, body = table(sheet, 2)
    cols = [h for h in hdr[1:] if h]
    return {'rows': [r[0] for r in body], 'cols': cols,
            'data': [[num(r[i + 1]) or 0 for i in range(len(cols))] for r in body]}
D['brand_batt'] = matrix('品牌-电池供应商')
D['brand_motor'] = matrix('品牌-电机供应商')

# 细分市场
def seg(sheet):
    hdr, body = table(sheet, 2)
    return [{'seg': r[0], 'n': num(r[1]), 'range': num(r[2]), 'cap': num(r[3]), 'power': num(r[4]), 'mass': num(r[5])} for r in body]
D['seg_bev'] = seg('细分市场-BEV')
D['seg_phev'] = seg('细分市场-PHEV')
def seg_cmp(sheet):
    hdr, body = table(sheet, 2)
    return [{'seg': r[0], 'y2023': num(r[1]), 'y2025': num(r[2]), 'chg': num(r[3])} for r in body]
D['seg_range_cmp'] = seg_cmp('细分市场续航对比')
D['seg_cap_cmp'] = seg_cmp('细分市场容量对比')
hdr, body = table('质量按细分市场', 2)
D['seg_mass'] = [{'seg': r[0], 'mass': num(r[1])} for r in body]

# 相关性
def corr(sheet):
    hdr, body = table(sheet, 2)
    labels = [r[0] for r in body]
    return {'labels': labels, 'data': [[num(r[i + 1]) for i in range(len(labels))] for r in body]}
D['corr_bev'] = corr('相关性-BEV')
D['corr_phev'] = corr('相关性-PHEV')

# 同比
def yoy(sheet):
    hdr, body = table(sheet, 2)
    years = [int(num(h)) for h in hdr[1:] if h is not None]
    return {'years': years, 'rows': [{'name': r[0], 'vals': [num(r[i + 1]) for i in range(len(years))]} for r in body]}
D['yoy_bev'] = yoy('同比增长-BEV')
D['yoy_phev'] = yoy('同比增长-PHEV')

# 关键指标
def kpi(sheet):
    hdr, body = table(sheet, 2)
    return [{'name': r[0], 'mean': num(r[1]), 'median': num(r[2]), 'n': num(r[3])} for r in body]
D['kpi_bev'] = kpi('关键指标汇总-BEV')
D['kpi_phev'] = kpi('关键指标汇总-PHEV')

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(D, f, ensure_ascii=False, separators=(',', ':'))
print('written', OUT, os.path.getsize(OUT), 'bytes')
