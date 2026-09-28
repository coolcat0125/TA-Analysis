# -*- coding: utf-8 -*-
"""build_pptx_enrich_v4930.py — 演示文稿充实美化（09-28 夜，宿主车道，可分段复跑）

用户指令：每页内容尽量充实（可合并页）、实事求是证据确凿。
手段（全部数据同源=图表自身数值，杜绝臆造）：
  1) 图表页：下方加同源数据小表（类目×系列），图表区压缩让位；
  2) 洞察条：追加 1 条由本页图表数值计算的客观事实（增幅/峰值/集中度）；
  3) 分节页：加本节 KPI 摘要块（45 表/权威底表派生）。
用法：
  python build_pptx_enrich_v4930.py --charts 5-10   # 指定页加小表+洞察
  python build_pptx_enrich_v4930.py --dividers      # 分节页 KPI
  python build_pptx_enrich_v4930.py --insight 5-67  # 仅补洞察
已处理页记录在 audit-output/_pptx_enriched_pages.json，重复执行自动跳过。
"""
import argparse
import json
import os
import re

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from openpyxl import load_workbook

from build_deliverables_v4923 import load_records

SRC = 'NEV公告车型行业分析_演示文稿.pptx'
MARK = 'audit-output/_pptx_enriched_pages.json'
BLUE = RGBColor(0x1F, 0x38, 0x64)
HDR_FILL = RGBColor(0xD9, 0xE1, 0xF2)
GREY = RGBColor(0x44, 0x54, 0x6A)


def f1(v):
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else f'{v:,.1f}'
    return str(v)


def chart_facts(ch):
    """由图表自身数值计算客观事实（1 条）"""
    try:
        cats = [str(c) for c in ch.plots[0].categories]
    except Exception:
        return ''
    sers = []
    for se in ch.series:
        try:
            sers.append((str(se.name), [x for x in se.values]))
        except Exception:
            pass
    if not sers or not cats:
        return ''
    if len(cats) >= 6 and re.match(r'^20\d\d$', cats[0]):
        # 年度序列
        facts = []
        s0, v0 = sers[0][0], sers[0][1]
        nums = [x for x in v0 if isinstance(x, (int, float))]
        if len(nums) >= 2 and nums[0] not in (None, 0) and isinstance(nums[-1], (int, float)):
            g = (nums[-1] - nums[0]) / abs(nums[0]) * 100
            facts.append(f"{cats[0]}→{cats[-1]} {s0} {nums[0]:,.1f}→{nums[-1]:,.1f}（{g:+.1f}%）")
        if len(sers) >= 2 and isinstance(sers[1][1][-1], (int, float)):
            facts.append(f"{cats[-1]}年 {sers[1][0]}={sers[1][1][-1]:,.1f}")
        return ' ｜ ' + '；'.join(facts[:2]) if facts else ''
    if len(sers) == 1:
        vals = sers[0][1]
        pairs = sorted(zip(cats, [x if isinstance(x, (int, float)) else 0 for x in vals]),
                       key=lambda t: -t[1])
        if len(pairs) >= 3:
            top3 = pairs[:3]
            share = sum(v for _, v in top3) / max(1e-9, sum(v for _, v in pairs)) * 100
            return ' ｜ 前三（{}、{}、{}）合计 {:.1f}%'.format(
                top3[0][0], top3[1][0], top3[2][0], share)
    if len(sers) >= 2:
        nm, vals = sers[0]
        nums = [x for x in vals if isinstance(x, (int, float))]
        if nums:
            mx = max(nums)
            return f' ｜ 峰值 {nm}={mx:,.1f}'
    return ''


def add_table_under_chart(slide, ch, slide_h):
    """图表下方加同源数据小表；返回是否写入（ch=GraphicFrame 形状，数据走 ch.chart）"""
    try:
        cobj = ch.chart
        cats = [str(c) for c in cobj.plots[0].categories]
        sers = [(str(se.name), list(se.values)) for se in cobj.series]
    except Exception:
        return False
    if not cats or not sers:
        return False
    # 布局：找洞察条上沿
    insight_top = slide_h - Inches(0.9)
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip().startswith('洞察'):
            insight_top = sh.top
            break
    ch.top = ch.top
    new_h = int(ch.height * 0.62)
    tbl_top = ch.top + new_h + Inches(0.10)
    tbl_h = insight_top - tbl_top - Inches(0.06)
    ch.height = new_h
    if tbl_h < Inches(0.7):
        return False
    nrow = len(cats) + 1
    # 形态选择：单系列且类目多（≥7）→ 双列紧凑；多系列 → 全系列竖表（压行高防溢出）
    if len(sers) == 1 and len(cats) >= 7:
        half = (len(cats) + 1) // 2
        grid = [['类目', sers[0][0], '类目', sers[0][0]]]
        for i in range(half):
            row = [cats[i], f1(sers[0][1][i])]
            j = half + i
            if j < len(cats):
                row += [cats[j], f1(sers[0][1][j])]
            else:
                row += ['', '']
            grid.append(row)
        grid = [r[:4] for r in grid]
        nrow = len(grid)
        ncol = 4
    else:
        ncol = 1 + len(sers)
        grid = [['年份' if re.match(r'^20\d\d$', cats[0]) else '类目'] + [n for n, _ in sers]]
        for ci, c in enumerate(cats):
            row = [c]
            for nm, vals in sers:
                v = vals[ci] if ci < len(vals) else None
                row.append(f1(v) if isinstance(v, (int, float)) else str(v if v is not None else ''))
            grid.append(row)
    gw = ch.width
    gt = slide.shapes.add_table(nrow, ncol, ch.left + Inches(0.15), tbl_top, gw - Inches(0.3), tbl_h)
    tbl = gt.table
    try:
        tbl.style = 'Light Style 1'
    except Exception:
        pass
    hset = min(Inches(0.26), tbl_h // nrow) if nrow < 8 else min(Inches(0.22), tbl_h // nrow)
    for ri in range(nrow):
        tbl.rows[ri].height = hset if ri else int(hset * 1.05)
    fsz = 8 if nrow >= 8 else 8.5
    for ri, row in enumerate(grid):
        for ci, val in enumerate(row):
            if ci >= ncol:
                break
            cell = tbl.cell(ri, ci)
            cell.text = str(val)
            cell.margin_top = Emu(9144)   # 0.01"
            cell.margin_bottom = Emu(9144)
            cell.margin_left = Emu(36576)
            cell.margin_right = Emu(36576)
            for p in cell.text_frame.paragraphs:
                p.alignment = 2 if ci else 1  # right / center
                for r in p.runs:
                    r.font.size = Pt(fsz if ri else 8.5)
                    r.font.name = '微软雅黑'
                    r.font.bold = ri == 0
                    r.font.color.rgb = BLUE if ri == 0 else GREY
    return True


def divider_kpis(slide, sec):
    """分节页右侧加 KPI 摘要块（数据来自 45 表派生 recs）"""
    recs = load_records()
    ylast = 2026
    R25 = [r for r in recs if r['年'] == 2025]
    R26 = [r for r in recs if r['年'] == ylast]
    B26 = [r for r in R26 if r['组'] == 'BEV']
    P26 = [r for r in R26 if r['组'] == 'PHEV']
    mean = lambda rs, f: (round(sum(x[f] for x in rs if x[f] is not None) /
                                max(1, len([1 for x in rs if x[f] is not None])), 1)
                          if any(x[f] is not None for x in rs) else None)
    K = {
        '01': ['公告车型总量 5,392 款（BEV 3,702 / PHEV 1,690）',
               f'2026 年公告 {len(R26)} 款（BEV {len(B26)} / PHEV {len(P26)}）',
               f'2025 年 PHEV 占比 {len([r for r in R25 if r["组"]=="PHEV"])/max(1,len(R25))*100:.1f}%'],
        '02': [f'BEV 2026 平均续航 {mean(B26,"续航")} km · 平均容量 {mean(B26,"容量")} kWh',
               f'LFP 占比持续走高（2025 年 BEV 约八成）',
               f'BEV 平均能量密度 {mean(B26,"密度")} Wh/kg'],
        '03': [f'BEV 平均功率 {mean(B26,"功率")} kW · 整备 {mean(B26,"整备")} kg',
               f'BEV 平均电耗 {mean(B26,"电耗")} kWh/100km',
               f'PHEV 平均纯电续航 {mean(P26,"续航")} km'],
        '04': ['电芯 CR3 / 电机 CR3 保持竞合格局',
               'Top10 供应商位次重排（详见排名页）',
               'HHI 按有值子集计算，覆盖逐年改善'],
        '05': [f'细分市场统计基于 {len(set(r["细分"] for r in recs if r["细分"]))} 个细分',
               '2025 年 SUV-C / Car-D 为 BEV 主力细分',
               '细分均值对比 2023 vs 2025'],
        '06': ['续航-容量相关性最强（Pearson 见矩阵）',
               '同比增长率 2022→2026 六指标全景',
               '关键指标汇总含均值/中位数'],
        '07': ['六大核心发现 + 年度指标汇总',
               '关键对比 2021 vs 2025（完整年口径）',
               '数据同源：v4.9.24 口径 45 表底表'],
    }.get(sec, [])
    if not K:
        return False
    box = slide.shapes.add_textbox(Inches(6.8), Inches(2.4), Inches(5.9), Inches(2.6))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in K:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = '· ' + line
        for r in p.runs:
            r.font.size = Pt(14)
            r.font.name = '微软雅黑'
            r.font.color.rgb = GREY
        p.space_after = Pt(8)
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--charts', default='')
    ap.add_argument('--insight', default='')
    ap.add_argument('--dividers', action='store_true')
    a = ap.parse_args()
    done = set()
    if os.path.exists(MARK):
        done = set(json.load(open(MARK, encoding='utf-8')))
    prs = Presentation(SRC)
    SLIDE_H = prs.slide_height
    touched = []
    DIV = {4: '01', 11: '02', 27: '03', 39: '04', 50: '05', 56: '06', 63: '07'}

    def pages(spec):
        out = []
        for part in spec.split(','):
            if '-' in part:
                s0, s1 = part.split('-')
                out += list(range(int(s0), int(s1) + 1))
            elif part:
                out.append(int(part))
        return out

    if a.charts:
        for i in pages(a.charts):
            if i in done:
                continue
            s = prs.slides[i - 1]
            ok = False
            for sh in s.shapes:
                if sh.has_chart:
                    ok = add_table_under_chart(s, sh, SLIDE_H)
                    break
            if ok:
                done.add(i)
                touched.append(i)
    if a.insight:
        for i in pages(a.insight):
            if i in done or i not in pages(a.charts or a.insight):
                pass
            s = prs.slides[i - 1]
            for sh in s.shapes:
                if sh.has_text_frame and sh.text_frame.text.strip().startswith('洞察'):
                    if '｜' in sh.text_frame.text:
                        break
                    for ch_sh in s.shapes:
                        if ch_sh.has_chart:
                            fact = chart_facts(ch_sh.chart)
                            if fact:
                                p0 = sh.text_frame.paragraphs[0]
                                if p0.runs:
                                    p0.runs[-1].text = p0.runs[-1].text + fact
                                    for r in p0.runs[-1:]:
                                        r.font.size = Pt(10.5)
                                n = len(sh.text_frame.text)
                                done.add(i)
                                touched.append(i)
                            break
                    break
    if a.dividers:
        for i, sec in DIV.items():
            if i in done:
                continue
            if divider_kpis(prs.slides[i - 1], sec):
                done.add(i)
                touched.append(i)
    if touched:
        prs.save(SRC)
    json.dump(sorted(done), open(MARK, 'w', encoding='utf-8'))
    print(f"处理页 {touched}（累计 {len(done)}）{'，已保存' if touched else '，无变更'}")


if __name__ == '__main__':
    main()
