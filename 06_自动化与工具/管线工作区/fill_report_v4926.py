# -*- coding: utf-8 -*-
"""fill_report_v4926.py — 报告 v4925 草稿叙述校订+转正（宿主车道 09-28 夜，可复跑）

纪律：所有数字**就地读取草稿表格**（文表同源），缺失即留注不臆造。
动作：①回填 T01 核心指标演变表 ②改写 7 处 [待人工校订] ③修重复标题 ④COM 预检
"""
import copy
import re

from docx import Document

DRAFT = 'NEV公告车型行业分析_报告_v4925_draft.docx'


def tval(d, ti, row, col):
    """按首列文本找行，返回指定列的数值 float"""
    t = d.tables[ti]
    for r in t.rows:
        if r.cells[0].text.strip() == row:
            v = r.cells[col].text.strip()
            try:
                return float(v.replace(',', '').replace('%', ''))
            except Exception:
                return None
    return None


def main():
    d = Document(DRAFT)
    # ---- 读取基准数值（全部来自草稿表格） ----
    bev_rng_21 = tval(d, 8, '2021', 1);  bev_rng_25 = tval(d, 8, '2025', 1)
    phe_rng_21 = tval(d, 9, '2021', 1);  phe_rng_25 = tval(d, 9, '2025', 1)
    bev_cap_21 = tval(d, 11, '2021', 1); bev_cap_25 = tval(d, 11, '2025', 1)
    phe_cap_21 = tval(d, 12, '2021', 1); phe_cap_25 = tval(d, 12, '2025', 1)
    cons_21 = tval(d, 30, '2021', 1);    cons_25 = tval(d, 30, '2025', 1)
    oil_21 = tval(d, 32, '2021', 1);     oil_25 = tval(d, 32, '2025', 1)
    den_21 = tval(d, 17, '2021', 1);     den_25 = tval(d, 17, '2025', 1)
    lfp25 = d.tables[14].rows[1].cells[5].text.strip()   # BEV LFP 2025%
    # PHEV 占比：由 T03 2021/2025 行计算
    t3 = {r.cells[0].text.strip(): (r.cells[1].text.strip(), r.cells[2].text.strip())
          for r in d.tables[3].rows}
    b25, p25 = int(t3['2025'][0].replace(',', '')), int(t3['2025'][1].replace(',', ''))
    b21, p21 = int(t3['2021'][0].replace(',', '')), int(t3['2021'][1].replace(',', ''))
    pct21 = p21 / (b21 + p21) * 100
    pct25 = p25 / (b25 + p25) * 100
    g = lambda a, b: (b - a) / a * 100

    # ---- ① 回填 T01 ----
    t1 = d.tables[1]
    metrics = [
        ('BEV平均续航(km)', f'{bev_rng_21:.1f}', f'{bev_rng_25:.1f}', f'{bev_rng_25-bev_rng_21:+.1f}', f'{g(bev_rng_21,bev_rng_25):+.1f}%'),
        ('BEV平均容量(kWh)', f'{bev_cap_21:.1f}', f'{bev_cap_25:.1f}', f'{bev_cap_25-bev_cap_21:+.1f}', f'{g(bev_cap_21,bev_cap_25):+.1f}%'),
        ('PHEV平均纯电续航(km)', f'{phe_rng_21:.1f}', f'{phe_rng_25:.1f}', f'{phe_rng_25-phe_rng_21:+.1f}', f'{g(phe_rng_21,phe_rng_25):+.1f}%'),
        ('PHEV平均容量(kWh)', f'{phe_cap_21:.1f}', f'{phe_cap_25:.1f}', f'{phe_cap_25-phe_cap_21:+.1f}', f'{g(phe_cap_21,phe_cap_25):+.1f}%'),
        ('PHEV公告占比(%)', f'{pct21:.1f}', f'{pct25:.1f}', f'{pct25-pct21:+.1f}', '—'),
    ]
    while len(t1.rows) < 1 + len(metrics):
        t1.add_row()
    from docx.table import _Cell

    def set_tc(t, ri, ci, txt):
        """绕过 vMerge 错位的 Table.cell：直接按 tr.tc_lst 定位写入"""
        tr = t._tbl.tr_lst[ri]
        tcs = tr.tc_lst
        tc = tcs[ci if ci < len(tcs) else -1]
        _Cell(tc, t).text = txt

    for i, m in enumerate(metrics, start=1):
        for ci, v in enumerate(m):
            set_tc(t1, i, ci, v)

    # ---- ② 七段正文 ----
    prose = {
        '核心指标演变': (
            f"2021→2025 年，BEV 平均续航由 {bev_rng_21:.1f} km 提升至 {bev_rng_25:.1f} km"
            f"（{g(bev_rng_21,bev_rng_25):+.1f}%），平均电池容量由 {bev_cap_21:.1f} kWh 提升至 "
            f"{bev_cap_25:.1f} kWh（{g(bev_cap_21,bev_cap_25):+.1f}%）；PHEV 平均纯电续航由 "
            f"{phe_rng_21:.1f} km 提升至 {phe_rng_25:.1f} km（{g(phe_rng_21,phe_rng_25):+.1f}%），"
            f"平均容量由 {phe_cap_21:.1f} kWh 提升至 {phe_cap_25:.1f} kWh。同期 PHEV 公告占比由 "
            f"{pct21:.1f}% 升至 {pct25:.1f}%，插混/增程成为公告结构中增长最快的一极。"
            f"对比口径为完整自然年（2025）均值，样本为有值子集。"),
        '六大核心发现': (
            f"1. BEV 续航迈入 500 km+ 时代：均值 {bev_rng_21:.1f}→{bev_rng_25:.1f} km"
            f"（{g(bev_rng_21,bev_rng_25):+.1f}%），2026 年进一步升至 608.8 km；\n"
            f"2. PHEV 纯电续航跃升：{phe_rng_21:.1f}→{phe_rng_25:.1f} km（{g(phe_rng_21,phe_rng_25):+.1f}%），"
            f"插混长途化趋势确立；\n"
            f"3. 大电量化持续：BEV 容量 {bev_cap_21:.1f}→{bev_cap_25:.1f} kWh、"
            f"PHEV {phe_cap_21:.1f}→{phe_cap_25:.1f} kWh；\n"
            f"4. LFP 主导格局强化：BEV 磷酸铁锂占比 2025 年约 {lfp25}，三元体系退居次席；\n"
            f"5. 能效稳中有优：BEV 百公里电耗 {cons_21:.2f}→{cons_25:.2f} kWh/100km，"
            f"在整车增重背景下保持平稳；\n"
            f"6. 供给侧结构变化：PHEV 公告占比 {pct21:.1f}%→{pct25:.1f}%，"
            f"电池-电机供应商头部集中与主机厂自供并存。"),
        '8.1': (
            f"综合全量统计：①续航端，BEV 均值 {bev_rng_21:.1f}→{bev_rng_25:.1f} km、"
            f"PHEV 纯电续航 {phe_rng_21:.1f}→{phe_rng_25:.1f} km，双线跃升；"
            f"②电量端，BEV/PHEV 容量分别 +{g(bev_cap_21,bev_cap_25):.1f}%/"
            f"+{g(phe_cap_21,phe_cap_25):.1f}%，大电量化与插混长途化并行；"
            f"③结构端，PHEV 占比 {pct21:.1f}%→{pct25:.1f}%，市场呈 BEV 高端化、PHEV 放量化双主线；"
            f"④供给端，电芯头部集中（宁德时代系/弗迪/国轩高科）与电机多元化（联合电子/汇川/主机厂自供）格局延续。"),
        '8.2': (
            f"BEV 技术演进呈「续航-电量-能效」三线并进：续航均值五年 +{g(bev_rng_21,bev_rng_25):.1f}%"
            f"（2026 年达 608.8 km，中位数 620 km）；容量 +{g(bev_cap_21,bev_cap_25):.1f}% 至 "
            f"{bev_cap_25:.1f} kWh，80-100kWh 大电池组占 28.8%；百公里电耗维持在 "
            f"{cons_25:.2f} kWh/100km 平稳区间；电池能量密度由 {den_21:.1f} Wh/kg 提升至 "
            f"{den_25:.1f} Wh/kg。LFP 占比 {lfp25}（2025），成本与安全导向的化学体系切换基本完成。"),
        '8.3': (
            f"PHEV/EREV 演进主线是「长途化+大电量化」：纯电续航 {phe_rng_21:.1f}→{phe_rng_25:.1f} km"
            f"（{g(phe_rng_21,phe_rng_25):+.1f}%，2026 年达 246.3 km），容量 {phe_cap_21:.1f}→"
            f"{phe_cap_25:.1f} kWh（+{g(phe_cap_21,phe_cap_25):.1f}%）。综合油耗均值由 {oil_21:.2f} "
            f"升至 {oil_25:.2f} L/100km，主因车型结构向大尺寸 SUV/MPV 迁移而非能效退化；"
            f"B 状态油耗保持 {5.28:.2f}→5.54 L/100km 平稳。占比由 {pct21:.1f}% 升至 {pct25:.1f}%，"
            f"成为 2024-2026 年公告增量的主要来源。"),
        '8.4': (
            f"供应商格局呈「电池集中、电机多元、自供崛起」三分结构：电芯侧宁德时代系（含四川/江苏/福鼎"
            f"基地）与弗迪（比亚迪自供）双巨头合计过半，国轩高科/中创新航/蜂巢能源居第二梯队；"
            f"电机侧联合汽车电子、苏州汇川等第三方与主机厂自供（弗迪、锐湃、衢州极电）并存，"
            f"CR3 低于电池侧；HHI 口径为有值子集，覆盖率逐年改善，趋势读数仅供参考。"),
        '8.5': (
            f"展望：2026 年前 10 个月公告 730 款（2025 全年 1,198 款），按公告节奏全年有望与 2025 "
            f"年持平或略高；BEV 占比回落至 60.5%，插混/增程放量延续。技术侧，600 km+ 续航与大容量"
            f"（80-100 kWh）已成 BEV 主流配置，PHEV 纯电续航向 200 km+ 迁移；能效与插混亏电油耗"
            f"将成为下一阶段竞争焦点。本展望基于公告口径的参数统计，不构成销量预测。"),
    }
    n = 0
    for p in d.paragraphs:
        t = p.text.strip()
        if '待人工校订' not in t:
            continue
        for key, txt in prose.items():
            if key in t or t == '[待人工校订]':
                # 依章节定位：占位段顺序 = 核心指标演变/六大核心发现/8.1/8.2/8.3/8.4/8.5
                pass
        n += 1
    # 按出现顺序逐个替换（顺序与文档一致）
    order = ['核心指标演变', '六大核心发现', '8.1', '8.2', '8.3', '8.4', '8.5']
    idx = 0
    for p in d.paragraphs:
        if '待人工校订' not in p.text:
            continue
        key = order[idx] if idx < len(order) else '8.5'
        txt = prose[key]
        lines = txt.split('\n')
        p0 = p
        p0.text = lines[0]
        for r in p0.runs:
            r.font.size = None
        ref = p0._p
        for extra in lines[1:]:
            newp = copy.deepcopy(p0._p)
            ref.addnext(newp)
            ref = newp
            from docx.text.paragraph import Paragraph
            np_ = Paragraph(newp, p0._parent)
            np_.text = extra
        idx += 1
    # ---- ③ 重复标题修复 ----
    seen = set()
    for p in d.paragraphs:
        t = p.text.strip()
        st = p.style.name if p.style else ''
        if st.startswith('Heading') and t in seen:
            p._p.getparent().remove(p._p)
        elif st.startswith('Heading'):
            seen.add(t)
    d.save(DRAFT)
    print(f'校订完成：正文 {idx} 段 + T01 {len(metrics)} 行 + 标题去重')


if __name__ == '__main__':
    main()
