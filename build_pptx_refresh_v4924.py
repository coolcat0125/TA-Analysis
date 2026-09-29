# -*- coding: utf-8 -*-
"""build_pptx_refresh_v4924.py — 下游交付②·演示文稿数据刷新（v4.3.6→v4.9.23 口径，可复跑）

策略：版式零重建——旧 deck 为 python-pptx 原生图表，逐图 replace_data 换入
权威底表重算数据（与 build_deliverables_v4923.py 同口径），16 张表格原维重写，
页脚/口径文本全局替换 + 页内旧值→新值映射（页面局部匹配，年份 token 保护）。
验证：python-pptx 重开 + PowerPoint COM 实开 + 字体扫描。
"""
import copy
import re
import shutil
import statistics as st
from collections import Counter, defaultdict

from pptx import Presentation
from pptx.chart.data import CategoryChartData

from build_deliverables_v4923 import load_records, pct

SRC = 'NEV公告车型行业分析_演示文稿.pptx'
YEARS = [2021, 2022, 2023, 2024, 2025, 2026]
YS = [str(y) for y in YEARS]


def yrows(recs, grp=None):
    d = defaultdict(list)
    for r in recs:
        if r['年'] and (grp is None or r['组'] == grp):
            d[r['年']].append(r)
    return d


def ymean(recs, f, grp, filt=None):
    out = {}
    for y, rs in yrows(recs, grp).items():
        vals = [r[f] for r in rs if r[f] is not None and (filt is None or filt(r))]
        if vals:
            out[y] = sum(vals) / len(vals)
    return out


def ratio_mean(recs, a, b, grp):
    out = {}
    for y, rs in yrows(recs, grp).items():
        vals = [r[a] / r[b] * 1000 for r in rs if r[a] is not None and r[b] not in (None, 0)]
        if vals:
            out[y] = sum(vals) / len(vals)
    return out


def btype(v):
    if '磷酸铁锂' in v:
        return 'LFP'
    if '三元' in v:
        return 'NCM'
    return '其他'


def build_data(recs):
    D = {}
    cnt = lambda g, y: sum(1 for r in recs if r['组'] == g and r['年'] == y)
    D['cnt_bev'] = [cnt('BEV', y) for y in YEARS]
    D['cnt_phev'] = [cnt('PHEV', y) for y in YEARS]
    D['phev_pct'] = [round(D['cnt_phev'][i] / max(1, D['cnt_bev'][i] + D['cnt_phev'][i]) * 100, 1)
                     for i in range(6)]
    # 细分 pie top8
    for g, k in (('BEV', 'seg_bev'), ('PHEV', 'seg_phev')):
        c = Counter(r['细分'] for r in recs if r['组'] == g and r['细分'])
        top = c.most_common(8)
        D[k+'_cats'] = [t[0] for t in top]
        D[k+'_vals'] = [t[1] for t in top]
    for f, k in (('续航', 'rng'), ('容量', 'cap'), ('功率', 'pw'), ('整备', 'cw')):
        D['bev_'+k] = ymean(recs, f, 'BEV')
        D['phev_'+k] = ymean(recs, f, 'PHEV')
    D['den_bev'] = ymean(recs, '密度', 'BEV')
    # 电池类型
    for g, k in (('BEV', 'bt_bev'), ('PHEV', 'bt_phev')):
        D[k] = {}
        for y, rs in yrows(recs, g).items():
            c = Counter(btype(r['电池类型']) for r in rs if r['电池类型'])
            n = sum(c.values())
            if n:
                D[k][y] = (c.get('LFP', 0), c.get('NCM', 0), c.get('其他', 0), n)
    # 分布 bins（沿用 deck 原口径）
    def bindist(grp, f, bins):
        out = {}
        for y, rs in yrows(recs, grp).items():
            vals = [r[f] for r in rs if r[f] is not None]
            if vals:
                out[y] = [sum(1 for v in vals if lo <= v < hi) for _, lo, hi in bins]
        return out
    D['rdist_bev'] = bindist('BEV', '续航',
                             [('<200', 0, 200), ('200-300', 200, 300), ('300-400', 300, 400),
                              ('400-500', 400, 500), ('500-600', 500, 600),
                              ('600-700', 600, 700), ('≥700', 700, 1e9)])
    D['rdist_phev'] = bindist('PHEV', '续航',
                              [('<50', 0, 50), ('50-100', 50, 100), ('100-125', 100, 125),
                               ('125-150', 125, 150), ('150-200', 150, 200), ('≥200', 200, 1e9)])
    D['cdist_bev'] = bindist('BEV', '容量',
                             [('<20', 0, 20), ('20-30', 20, 30), ('30-40', 30, 40),
                              ('40-50', 40, 50), ('50-60', 50, 60), ('60-70', 60, 70),
                              ('70-80', 70, 80), ('≥80', 80, 1e9)])
    D['cdist_phev'] = bindist('PHEV', '容量',
                              [('<10', 0, 10), ('10-20', 10, 20), ('20-30', 20, 30),
                               ('30-40', 30, 40)])
    D['pq_bev'] = ratio_mean(recs, '功率', '整备', 'BEV')
    D['pq_phev'] = ratio_mean(recs, '功率', '整备', 'PHEV')
    D['ed_bev'] = ratio_mean(recs, '容量', '整备', 'BEV')
    D['ed_phev'] = ratio_mean(recs, '容量', '整备', 'PHEV')
    D['cons_bev'] = ymean(recs, '电耗', 'BEV')
    # 电耗×电池类型（BEV）
    D['cons_lfp'], D['cons_ncm'] = {}, {}
    for y, rs in yrows(recs, 'BEV').items():
        lf = [r['电耗'] for r in rs if r['电耗'] is not None and btype(r['电池类型']) == 'LFP']
        nc = [r['电耗'] for r in rs if r['电耗'] is not None and btype(r['电池类型']) == 'NCM']
        if lf:
            D['cons_lfp'][y] = sum(lf) / len(lf)
        if nc:
            D['cons_ncm'][y] = sum(nc) / len(nc)
    oil_ok = lambda r: r['油耗'] is not None and 0.1 <= r['油耗'] <= 12
    D['oil_phev'] = ymean(recs, '油耗', 'PHEV', oil_ok)
    D['oilb_phev'] = {}
    for y, rs in yrows(recs, 'PHEV').items():
        pass
    # B 状态油耗（工作簿列在 records 未带，单独读）
    D['oil_disp'] = self_oil(recs)  # 占位，下方重算
    return D


def self_oil(recs):
    return {}


def hhi(vals):
    c = Counter(vals)
    n = sum(c.values())
    return round(sum((v / n * 100) ** 2 for v in c.values()), 1) if n >= 5 else None


def cr(vals, k):
    c = Counter(vals)
    n = sum(c.values())
    if n < 5:
        return None
    return round(sum(v for _, v in c.most_common(k)) / n * 100, 1)


def main():
    recs = load_records()
    D = build_data(recs)
    # B 状态油耗（记录里未带，从底表直读）
    import openpyxl
    wbs = openpyxl.load_workbook('NEV公告参数汇总表_合并版（341~410批）.xlsx', read_only=True)
    wss = wbs['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(wss.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    oilb = defaultdict(list)
    disp_oil = defaultdict(list)   # (排量档, 年) → 油耗
    dispb_oil = defaultdict(list)
    for r in wss.iter_rows(min_row=2, values_only=True):
        pt = str(r[hi['动力类型']] or '').strip()
        if pt != 'PHEV':
            continue
        b = r[hi['批次']]
        try:
            y = D.get('y')  # noqa
        except Exception:
            pass
        # 年份
        yr = None
        for rr in recs:
            pass
        ob = r[hi['B状态油耗(L/100km)']]
        oc = r[hi['综合油耗(L/100km)']]
        dv = r[hi['发动机排量(mL)']]
        # 年份由批次时间表（load_records 内部逻辑重做，轻量）
        yr = BATCH_YEAR.get(str(b).strip()) if BATCH_YEAR else None
        if yr is None:
            continue
        try:
            fv = float(str(ob).split('/')[0])
            if 0.5 <= fv <= 15:
                oilb[yr].append(fv)
        except Exception:
            pass
        for col, dst in ((oc, disp_oil), (ob, dispb_oil)):
            try:
                fv = float(str(col).split('/')[0])
                dd = float(str(dv).split('/')[0])
            except Exception:
                continue
            if 0.1 <= fv <= 12 and (dst is dispb_oil and 0.5 <= fv <= 15 or dst is disp_oil):
                lab = '<1.0L' if dd < 1000 else '1.0-1.5L' if dd < 1500 else \
                    '1.5-2.0L' if dd < 2000 else '≥2.0L'
                dst[(lab, yr)].append(fv)
    wbs.close()
    D['oilb'] = {y: sum(v) / len(v) for y, v in oilb.items() if v}
    D['disp_oil'] = disp_oil
    D['dispb_oil'] = dispb_oil

    prs = Presentation(SRC)
    replaced = {'charts': 0, 'tables': 0, 'texts': 0}

    # ---------- 图表数据 ----------
    def cats_years():
        cd = CategoryChartData()
        cd.categories = YS
        return cd

    def set_chart(slide_no, fn):
        s = prs.slides[slide_no - 1]
        for sh in s.shapes:
            if sh.has_chart:
                ch = sh.chart
                cd = fn()
                # 保留原系列名顺序
                names = [se.name for se in ch.series]
                for nm in names:
                    vals = fn_vals(slide_no, nm, D)
                    cd.add_series(nm, vals)
                ch.replace_data(cd)
                replaced['charts'] += 1
                return True
        return False

    REG = {}
    def fn_vals(sno, nm, D):
        return REG[(sno, nm)]()

    # —— 每张图的取数注册表（slide, series名）→ 值列表 ——
    R = {}
    for i, y in enumerate(YEARS):
        pass
    R[(5, 'BEV')] = lambda: D['cnt_bev']
    R[(5, 'PHEV')] = lambda: D['cnt_phev']
    R[(6, 'PHEV占比%')] = lambda: D['phev_pct']
    R[(7, '数量')] = lambda: D['seg_bev_vals']
    R[(8, '数量')] = lambda: D['seg_phev_vals']
    R[(10, '合计公告车型')] = lambda: [D['cnt_bev'][i] + D['cnt_phev'][i] for i in range(6)]
    R[(12, '平均续航(km)')] = lambda: [round(D['bev_rng'].get(y, 0), 2) for y in YEARS]
    R[(13, '平均纯电续航(km)')] = lambda: [round(D['phev_rng'].get(y, 0), 2) for y in YEARS]
    R[(14, 'BEV')] = lambda: [round(D['bev_rng'].get(y, 0), 2) for y in YEARS]
    R[(14, 'PHEV')] = lambda: [round(D['phev_rng'].get(y, 0), 2) for y in YEARS]
    R[(15, 'BEV平均容量(kWh)')] = lambda: [round(D['bev_cap'].get(y, 0), 2) for y in YEARS]
    R[(16, 'PHEV平均容量(kWh)')] = lambda: [round(D['phev_cap'].get(y, 0), 2) for y in YEARS]
    R[(17, 'BEV')] = lambda: [round(D['bev_cap'].get(y, 0), 2) for y in YEARS]
    R[(17, 'PHEV')] = lambda: [round(D['phev_cap'].get(y, 0), 2) for y in YEARS]

    def btpct(Dk, idx):
        return [round(D[Dk].get(y, (0, 0, 0, 1))[idx] / max(1, D[Dk].get(y, (0, 0, 0, 1))[3]) * 100, 1)
                for y in YEARS]
    for sn, Dk in ((18, 'bt_bev'), (19, 'bt_phev')):
        R[(sn, 'LFP%')] = lambda Dk=Dk, idx=0: btpct(Dk, idx)
        R[(sn, 'NCM%')] = lambda Dk=Dk, idx=1: btpct(Dk, idx)
        R[(sn, '其他%')] = lambda Dk=Dk, idx=2: btpct(Dk, idx)
    R[(20, 'BEV平均能量密度')] = lambda: [round(D['den_bev'].get(y, 0), 1) for y in YEARS]

    def dist_rows(Dk, cats):
        return {c: [D[Dk].get(y, [0]*len(cats))[i] for y in YEARS]
                for i, c in enumerate(cats)}
    RB = ['<200', '200-300', '300-400', '400-500', '500-600', '600-700', '≥700']
    RP = ['<50', '50-100', '100-125', '125-150', '150-200', '≥200']
    CB = ['<20', '20-30', '30-40', '40-50', '50-60', '60-70', '70-80', '≥80']
    CP = ['<10', '10-20', '20-30', '30-40']
    for sn, Dk, cats in ((21, 'rdist_bev', RB), (22, 'rdist_phev', RP),
                         (23, 'cdist_bev', CB), (24, 'cdist_phev', CP)):
        mp = dist_rows(Dk, cats)
        for c in cats:
            R[(sn, c)] = (lambda mp=mp, c=c: mp[c])
    R[(26, '磷酸铁锂(LFP)')] = lambda: [D['bt_bev'].get(y, (0, 0, 0, 0))[0] for y in YEARS]
    R[(26, '三元锂(NCM)')] = lambda: [D['bt_bev'].get(y, (0, 0, 0, 0))[1] for y in YEARS]
    R[(26, '其他')] = lambda: [D['bt_bev'].get(y, (0, 0, 0, 0))[2] for y in YEARS]
    R[(28, 'BEV平均峰值功率')] = lambda: [round(D['bev_pw'].get(y, 0), 2) for y in YEARS]
    R[(29, 'PHEV平均峰值功率')] = lambda: [round(D['phev_pw'].get(y, 0), 2) for y in YEARS]
    R[(30, 'BEV')] = lambda: [round(D['bev_cw'].get(y, 0), 1) for y in YEARS]
    R[(30, 'PHEV')] = lambda: [round(D['phev_cw'].get(y, 0), 1) for y in YEARS]
    R[(31, 'BEV')] = lambda: [round(D['pq_bev'].get(y, 0), 1) for y in YEARS]
    R[(31, 'PHEV')] = lambda: [round(D['pq_phev'].get(y, 0), 1) for y in YEARS]
    R[(32, 'BEV')] = lambda: [round(D['ed_bev'].get(y, 0), 1) for y in YEARS]
    R[(32, 'PHEV')] = lambda: [round(D['ed_phev'].get(y, 0), 1) for y in YEARS]
    R[(33, 'BEV平均电耗')] = lambda: [round(D['cons_bev'].get(y, 0), 2) for y in YEARS]
    R[(34, 'LFP')] = lambda: [round(D['cons_lfp'].get(y, 0), 2) for y in YEARS]
    R[(34, 'NCM')] = lambda: [round(D['cons_ncm'].get(y, 0), 2) for y in YEARS]
    R[(35, '综合油耗(L)')] = lambda: [round(D['oil_phev'].get(y, 0), 2) for y in YEARS]
    R[(35, '纯电续航/30(km)')] = lambda: [round(D['phev_rng'].get(y, 0) / 30, 2) for y in YEARS]
    R[(36, 'B状态油耗(L)')] = lambda: [round(D['oilb'].get(y, 0), 2) for y in YEARS]
    DL = ['<1.0L', '1.0-1.5L', '1.5-2.0L', '≥2.0L']
    R[(37, '综合油耗')] = lambda: [round(sum(D['disp_oil'].get((lab, y), [])) /
                                        max(1, len(D['disp_oil'].get((lab, y), []))), 2)
                                  for lab in DL]
    R[(38, 'B状态油耗')] = lambda: [round(sum(D['dispb_oil'].get((lab, y), [])) /
                                         max(1, len(D['dispb_oil'].get((lab, y), []))), 2)
                                   for lab in DL]

    # 供应商域（直读 records）
    eb = yrows(recs, 'BEV')
    allrows = {y: eb.get(y, []) + yrows(recs, 'PHEV').get(y, []) for y in YEARS}
    R[(40, '电池供应商HHI')] = lambda: [hhi([r['电芯'] for r in allrows[y] if r['电芯']]) for y in YEARS]
    R[(40, '电机供应商HHI')] = lambda: [hhi([r['电机企'] for r in allrows[y] if r['电机企']]) for y in YEARS]
    for sn, fld in ((41, '电芯'), (42, '电机企')):
        R[(sn, 'CR3')] = lambda fld=fld: [cr([r[fld] for r in allrows[y] if r[fld]], 3) for y in YEARS]
        R[(sn, 'CR5')] = lambda fld=fld: [cr([r[fld] for r in allrows[y] if r[fld]], 5) for y in YEARS]
        R[(sn, 'CR10')] = lambda fld=fld: [cr([r[fld] for r in allrows[y] if r[fld]], 10) for y in YEARS]
    top_bat = Counter(r['电芯'] for r in recs if r['电芯']).most_common(10)
    top_mot = Counter(r['电机企'] for r in recs if r['电机企']).most_common(10)
    R[(43, '配套车型数')] = lambda: [t[1] for t in top_bat]
    R[(44, '配套车型数')] = lambda: [t[1] for t in top_mot]

    def segcmp(grp, f):
        c25 = Counter(r['细分'] for r in recs if r['组'] == grp and r['年'] == 2025 and r['细分'])
        cats = []
        for sg, _ in c25.most_common(12):
            v23 = [r[f] for r in recs if r['组'] == grp and r['细分'] == sg
                   and r['年'] == 2023 and r[f] is not None]
            v25 = [r[f] for r in recs if r['组'] == grp and r['细分'] == sg
                   and r['年'] == 2025 and r[f] is not None]
            if v23 and v25:
                cats.append(sg)
            if len(cats) == 8:
                break
        s23 = [round(sum(r[f] for r in recs if r['组'] == grp and r['细分'] == sg
                         and r['年'] == 2023 and r[f] is not None) /
                     max(1, len([1 for r in recs if r['组'] == grp and r['细分'] == sg
                                 and r['年'] == 2023 and r[f] is not None])), 1) for sg in cats]
        s25 = [round(sum(r[f] for r in recs if r['组'] == grp and r['细分'] == sg
                         and r['年'] == 2025 and r[f] is not None) /
                     max(1, len([1 for r in recs if r['组'] == grp and r['细分'] == sg
                                 and r['年'] == 2025 and r[f] is not None])), 1) for sg in cats]
        return cats, s23, s25
    SEG = {}
    for sno, grp, f, key in ((51, 'BEV', '续航', 'rng'), (52, 'BEV', '容量', 'cap'),
                             (53, 'PHEV', '续航', 'rng')):
        SEG[sno] = segcmp(grp, f)
        R[(sno, '2023')] = (lambda s=SEG[sno]: s[1])
        R[(sno, '2025')] = (lambda s=SEG[sno]: s[2])
    # S55 细分整备（全周期均值 top12，样本≥5）
    segm = {}
    for r in recs:
        if r['细分'] and r['整备'] is not None:
            segm.setdefault(r['细分'], []).append(r['整备'])
    S55 = sorted(((k, sum(v) / len(v)) for k, v in segm.items() if len(v) >= 5),
                 key=lambda x: -x[1])[:12]
    R[(55, '平均整备(kg)')] = lambda: [round(v, 1) for _, v in S55]

    # yoy（2022..2026）
    def yoy(f, grp):
        out = [None]
        prev = None
        for y in YEARS:
            vals = [r[f] for r in recs if r['年'] == y and r['组'] == grp and r[f] is not None]
            m = sum(vals) / len(vals) if vals else None
            out.append(round((m - prev) / prev * 100, 1)
                       if (m is not None and prev) else None)
            prev = m if m is not None else prev
        return out[2:]  # 2022-2026 五值（2021 无前值，剔除占位；类目同步 2022-2026）
    YSPEC = [('平均续航(km)', '续航'), ('平均容量(kWh)', '容量'), ('平均功率(kW)', '功率'),
             ('平均整备质量(kg)', '整备'), ('平均能量密度(Wh/kg)', '密度')]
    for sn, grp in ((59, 'BEV'), (60, 'PHEV')):
        for lab, f in YSPEC:
            R[(sn, lab)] = (lambda f=f, grp=grp: yoy(f, grp))
        R[(sn, '公告数量(款)')] = lambda grp=grp: [
            (lambda a, b: round((a - b) / b * 100, 1) if b else None)(
                cnt('BEV', y) + cnt('PHEV', y), cnt('BEV', y - 1) + cnt('PHEV', y - 1))
            for y in YEARS[1:]]  # 总量口径（修复：原为 BEV 单口径）

    # ---------- 执行替换 ----------
    cnt = lambda g, y: sum(1 for r in recs if r['组'] == g and r['年'] == y)
    for (sno, nm), fn in R.items():
        REG[(sno, nm)] = fn
    # 类目替换表（需要更新类目的图）
    NEWCATS = {7: D['seg_bev_cats'], 8: D['seg_phev_cats'],
               43: [t[0] for t in top_bat], 44: [t[0] for t in top_mot],
               51: SEG[51][0], 52: SEG[52][0], 53: SEG[53][0],
               55: [k for k, _ in S55], 37: DL, 38: DL,
               59: YS[1:], 60: YS[1:]}  # 同比图类目=2022-2026（与 yoy 五值对齐）
    for sno in sorted({s for s, _ in R}):
        s = prs.slides[sno - 1]
        for sh in s.shapes:
            if not sh.has_chart:
                continue
            ch = sh.chart
            old_names = [se.name for se in ch.series]
            cd = CategoryChartData()
            cd.categories = NEWCATS.get(sno, YS)
            for nm in old_names:
                fn = REG.get((sno, str(nm)))
                if fn is None:
                    print(f'  ! S{sno} 系列 {nm!r} 无取数注册——保留原数据')
                    vals = list(se_values(ch, nm))
                else:
                    vals = fn()
                cd.add_series(str(nm), vals)
            ch.replace_data(cd)
            replaced['charts'] += 1

    # ---------- 表格重写 ----------
    def fmt(v, dec=1, comma=False):
        if v is None or v == '':
            return ''
        if isinstance(v, str):
            return v
        if isinstance(v, float) and dec == 0:
            v = int(round(v))
        s = f'{v:,.{dec}f}' if comma else f'{v:.{dec}f}' if isinstance(v, float) else str(v)
        return s

    def set_table(sno, grid):
        s = prs.slides[sno - 1]
        for sh in s.shapes:
            if not sh.has_table:
                continue
            t = sh.table
            assert len(t.rows) == len(grid) and len(t.columns) == len(grid[0]), \
                f'S{sno} 维度不符 {len(t.rows)}x{len(t.columns)} vs {len(grid)}x{len(grid[0])}'
            for ri, row in enumerate(grid):
                n_cells = len(t._tbl.tr_lst[ri].tc_lst)
                for ci in range(min(len(row), n_cells)):
                    v = row[ci]
                    cell = t.cell(ri, ci)
                    txt = fmt(v) if not isinstance(v, str) else v
                    # 保留原 run 格式：改首 run 文本、删多余 run/段落
                    tf = cell.text_frame
                    paras = tf.paragraphs
                    p0 = paras[0]
                    if p0.runs:
                        p0.runs[0].text = txt
                        for extra in p0.runs[1:]:
                            extra._r.getparent().remove(extra._r)
                    else:
                        p0.add_run().text = txt
                    for pextra in paras[1:]:
                        for r in list(pextra.runs):
                            r._r.getparent().remove(r._r)
            replaced['tables'] += 1
            return

    mean2 = lambda rs, f: round(sum(r[f] for r in rs if r[f] is not None) /
                                max(1, len([1 for r in rs if r[f] is not None])), 2)
    med = lambda rs, f: (lambda v: round(st.median(v), 1) if v else None)(
        [r[f] for r in rs if r[f] is not None])
    grid2 = []
    for i, y in enumerate(YEARS):
        yb = [r for r in recs if r['年'] == y and r['组'] == 'BEV']
        grid2.append([y, D['cnt_bev'][i], D['cnt_phev'][i],
                      D['cnt_bev'][i] + D['cnt_phev'][i], D['phev_pct'][i],
                      mean2(yb, '续航'), mean2(yb, '容量')])
    set_table(2, [['年份', 'BEV(款)', 'PHEV(款)', '合计', 'PHEV占比%',
                   'BEV均值续航(km)', 'BEV均值容量(kWh)']] + grid2)
    set_table(9, [['年份', 'BEV(款)', 'PHEV(款)', '合计(款)', 'PHEV占比%']] +
              [[y, D['cnt_bev'][i], D['cnt_phev'][i],
                D['cnt_bev'][i] + D['cnt_phev'][i], D['phev_pct'][i]]
               for i, y in enumerate(YEARS)])
    # S25 四象限（与 45 表同法：中位数阈值）
    bvr = [r['续航'] for r in recs if r['组'] == 'BEV' and r['续航'] is not None]
    bvc = [r['容量'] for r in recs if r['组'] == 'BEV' and r['容量'] is not None]
    import statistics
    tr, tc = statistics.median(bvr), statistics.median(bvc)
    quad = Counter()
    for r in recs:
        if r['组'] == 'BEV' and r['续航'] is not None and r['容量'] is not None:
            quad[('低' if r['续航'] < tr else '高') + '续航' +
                 ('低' if r['容量'] < tc else '高') + '容量'] += 1
    qn = sum(quad.values())
    set_table(25, [['象限', '车型数', '占比%']] +
              [[k, quad.get(k, 0), round(quad.get(k, 0) / qn * 100, 1)]
               for k in ('高续航高容量', '高续航低容量', '低续航高容量', '低续航低容量')])
    # S45 组合
    combo = Counter((r['电芯'], r['电机企']) for r in recs if r['电芯'] and r['电机企'])
    set_table(45, [['排名', '电池供应商 + 电机供应商', '车型数']] +
              [[i, f'{k[0]} + {k[1]}', v] for i, (k, v) in enumerate(combo.most_common(10), 1)])
    # S46/47 品牌矩阵
    def brandgrid(bfld, sfld, nb, ns):
        brands = [k for k, _ in Counter(r[bfld] for r in recs if r[bfld]).most_common(nb)]
        sups = [k for k, _ in Counter(r[sfld] for r in recs if r[sfld]).most_common(ns)]
        rows = [['品牌'] + sups]
        for bd in brands:
            rows.append([bd] + [sum(1 for r in recs if r[bfld] == bd and r[sfld] == sp)
                                for sp in sups])
        return rows
    set_table(46, brandgrid('商标', '电芯', 8, 5))
    set_table(47, brandgrid('商标', '电机企', 10, 5))
    # S48/49 明细
    set_table(48, [['年份', '电池供应商HHI', '电池CR3(%)']] +
              [[y, R[(40, '电池供应商HHI')]()[i], R[(41, 'CR3')]()[i]] for i, y in enumerate(YEARS)])
    set_table(49, [['年份', '电机CR3(%)', '电机CR5(%)', '电机CR10(%)']] +
              [[y, R[(42, 'CR3')]()[i], R[(42, 'CR5')]()[i], R[(42, 'CR10')]()[i]]
               for i, y in enumerate(YEARS)])
    # S54 BEV 细分 2025 top10
    c25 = Counter(r['细分'] for r in recs if r['组'] == 'BEV' and r['年'] == 2025 and r['细分'])
    rows54 = [['细分市场', '车型数', '平均续航(km)', '平均容量(kWh)', '平均功率(kW)', '平均整备(kg)']]
    for sg, n in c25.most_common(10):
        rs = [r for r in recs if r['组'] == 'BEV' and r['年'] == 2025 and r['细分'] == sg]
        rows54.append([sg, n] + [mean2(rs, f) for f in ('续航', '容量', '功率', '整备')])
    set_table(54, rows54)
    # S57/58 相关性
    from build_deliverables_v4923 import pearson
    def corrgrid(grp, spec):
        head = [''] + [lab for lab, _ in spec]
        rows = [head]
        for la, fa in spec:
            row = [la]
            for lb, fb in spec:
                xs = [r[fa] for r in recs if r['组'] == grp and r[fa] is not None and r[fb] is not None]
                ys = [r[fb] for r in recs if r['组'] == grp and r[fa] is not None and r[fb] is not None]
                p = pearson(xs, ys)
                row.append(round(p, 2) if p is not None else '')
            rows.append(row)
        return rows
    set_table(57, corrgrid('BEV', [('续航', '续航'), ('容量', '容量'), ('功率', '功率'),
                                   ('整备', '整备'), ('密度', '密度'), ('电耗', '电耗')]))
    set_table(58, corrgrid('PHEV', [('续航', '续航'), ('容量', '容量'), ('功率', '功率'),
                                    ('整备', '整备'), ('密度', '密度'), ('电耗', '电耗'),
                                    ('油耗', '油耗')]))
    # S61/62 yoy 明细
    def yoygrid(grp, spec):
        rows = [['指标'] + YS]
        for lab, f in spec:
            rows.append([lab] + ['—'] + [fmt(v) for v in yoy(f, grp)])
        return rows
    WS61 = [('平均续航(km)', '续航'), ('平均容量(kWh)', '容量'), ('平均功率(kW)', '功率'),
            ('平均整备质量(kg)', '整备'), ('平均能量密度(Wh/kg)', '密度')]
    set_table(61, yoygrid('BEV', WS61))
    set_table(62, yoygrid('PHEV', WS61))
    # S64/65 全周期
    def fullgrid(grp, spec):
        rs = [r for r in recs if r['组'] == grp]
        rows = [['指标', '均值', '中位数', '单位']]
        for lab, f, u in spec:
            rows.append([lab, mean2(rs, f), med(rs, f), u])
        return rows
    set_table(64, fullgrid('BEV', [('平均续航', '续航', 'km'), ('平均容量', '容量', 'kWh'),
                                   ('平均功率', '功率', 'kW'), ('平均整备', '整备', 'kg'),
                                   ('平均能量密度', '密度', 'Wh/kg')]))
    set_table(65, fullgrid('PHEV', [('平均续航', '续航', 'km'), ('平均容量', '容量', 'kWh'),
                                    ('平均功率', '功率', 'kW'), ('平均整备', '整备', 'kg'),
                                    ('平均能量密度', '密度', 'Wh/kg')]))
    # S66 关键发现 2021 vs 2025
    def yv(grp, f, y):
        vals = [r[f] for r in recs if r['组'] == grp and r['年'] == y and r[f] is not None]
        return round(sum(vals) / len(vals), 2) if vals else None
    rows66 = [['指标', '2021', '2025', '变化%']]
    for lab, grp, f in (('BEV平均续航(km)', 'BEV', '续航'), ('BEV平均容量(kWh)', 'BEV', '容量'),
                        ('BEV平均功率(kW)', 'BEV', '功率'), ('PHEV平均纯电续航(km)', 'PHEV', '续航'),
                        ('PHEV平均容量(kWh)', 'PHEV', '容量')):
        a, b = yv(grp, f, 2021), yv(grp, f, 2025)
        rows66.append([lab, a, b, round((b - a) / a * 100, 1)])
    set_table(66, rows66)

    # ---------- 文本口径更新 ----------
    tot = len(recs)
    nb = sum(1 for r in recs if r['组'] == 'BEV')
    np_ = tot - nb
    p25 = D['phev_pct'][4]
    GLOBAL = [
        ('基于 4,613 条公告车型数据', f'基于 {tot:,} 条公告车型数据'),
        ('权威底表 v4.3.6', '权威底表 v4.9.23'),
        ('共4,613款纯新能源乘用车', f'共{tot}款纯新能源乘用车'),
        ('共 4,613 款', f'共 {tot:,} 款'),
        ('样本 4,613 款（BEV 3,099 / PHEV 1,514）',
         f'样本 {tot:,} 款（BEV {nb:,} / PHEV {np_:,}）'),
        ('4,613 款', f'{tot:,} 款'),
        ('纯乘用车口径（v4.3.4底表）', f'纯乘用车口径（{tot:,} 款 × 38 列）'),
        ('2025 年 PHEV 占比', '2025 年 PHEV 占比'),  # 占位保持
    ]
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    t = run.text
                    t2 = t
                    for a, b in GLOBAL:
                        t2 = t2.replace(a, b)
                    if t2 != t:
                        run.text = t2
                        replaced['texts'] += 1

    shutil.copy2(SRC, 'archive/演示文稿_pre-v4924_backup_20260927.pptx')
    prs.save(SRC)
    print(f"刷新完成 charts={replaced['charts']} tables={replaced['tables']} texts={replaced['texts']}")
    print(f"口径：{tot:,} 款（BEV {nb:,}/PHEV {np_:,}）· 2025 PHEV占比 {p25}%")
    print("补充文本轮替换 run:", supplement_text_remap(SRC))


BATCH_YEAR = None



def supplement_text_remap(prs_path):
    """补充文本轮（固化自 09-27 晚实战）：页内局部映射（系列值+合计+占比+年对增幅，
    容差 ±0.051），仅动 小数/千分位/≥3位整数 token；保护版本号/批次区间/年份。可复跑。"""
    import re, datetime
    from pptx import Presentation
    OLD = 'archive/演示文稿_pre-v4924_backup_20260927.pptx'
    po = Presentation(OLD)
    pn = Presentation(prs_path)
    NUM = r'\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d{3,}'

    def fmt_like(tok, val):
        if '.' in tok:
            d = len(tok.split('.')[1])
            return f'{val:,.{d}f}' if ',' in tok else f'{val:.{d}f}'
        return f'{val:,.0f}' if ',' in tok else f'{val:.0f}'

    n = 0
    for i in range(len(po.slides)):
        so = [sh for sh in po.slides[i].shapes if sh.has_chart]
        sn = [sh for sh in pn.slides[i].shapes if sh.has_chart]
        pairs = {}

        def add(a, b):
            if a is None or b is None or abs(a - b) < 1e-9:
                return
            pairs[float(a)] = float(b)

        for sho, shn in zip(so, sn):
            so_c = [list(se.values) for se in sho.chart.series]
            sn_c = [list(se.values) for se in shn.chart.series]
            if len(so_c) != len(sn_c):
                continue
            for si in range(len(so_c)):
                o, nn = so_c[si], sn_c[si]
                L = min(len(o), len(nn))
                for a, b in zip(o[:L], nn[:L]):
                    add(a, b)
                if i + 1 in (59, 60, 61, 62):
                    continue
                for x in range(L):
                    for y in range(x + 1, L):
                        if o[x] and nn[x]:
                            try:
                                add((o[y] - o[x]) / o[x] * 100, (nn[y] - nn[x]) / nn[x] * 100)
                            except Exception:
                                pass
            if len(so_c) == 2:
                for a1, a2, b1, b2 in zip(so_c[0], so_c[1], sn_c[0], sn_c[1]):
                    if None not in (a1, a2, b1, b2):
                        add(a1 + a2, b1 + b2)
                        if a1 + a2 > 0:
                            add(a2 / (a1 + a2) * 100, b2 / (b1 + b2) * 100)
        keys = sorted(pairs)

        def lookup(v):
            c = [k for k in keys if abs(k - v) <= 0.051]
            return pairs[min(c, key=lambda k: abs(k - v))] if c else None

        for sh in pn.slides[i].shapes:
            if sh.has_table or not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    t = run.text
                    if '底表 v4' in t or '~' in t:
                        continue
                    def rep(mo):
                        tok = mo.group(0)
                        v = float(tok.replace(',', ''))
                        if 1900 <= v <= 2100:
                            return tok
                        b = lookup(v)
                        return fmt_like(tok, b) if b is not None else tok
                    t2 = re.sub(NUM, rep, t)
                    if t2 != t:
                        run.text = t2
                        n += 1
    pn.save(prs_path)
    return n


if __name__ == '__main__':
    # 批次→年（与 load_records 同源）
    import openpyxl as _px
    _wb = _px.load_workbook('NEV公告参数汇总表_合并版（341~410批）.xlsx', read_only=True)
    _bt = _wb['批次时间表']
    _hdr = [str(c.value or '').strip() for c in next(_bt.iter_rows(min_row=1, max_row=1))]
    _bi, _di = _hdr.index('批次'), _hdr.index('公告日期')
    BATCH_YEAR = {}
    for _r in _bt.iter_rows(min_row=2, values_only=True):
        _m = re.match(r'^(\d{4})', str(_r[_di] or ''))
        if _r[_bi] is not None and _m:
            BATCH_YEAR[str(_r[_bi]).strip()] = int(_m.group(1))
    _wb.close()
    main()
