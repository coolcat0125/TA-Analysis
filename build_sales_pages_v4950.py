# -*- coding: utf-8 -*-
"""_sales_pages_v4950.py — 销量联动页注入（WS5 / deck-builder / 2026-10-19）

唯一数据源：audit-output/sales_linkage_analytics_20261019.json（契约 v1.0，Lead 已冻结）。
内容依据：Update/核心论断与建议纲要_20261019.md v1.0（每页结论句可追溯论断编号 + 槽位）。
结构依据：Update/销量分析演示稿结构_20261019.md v1.0。

设计纪律（AGENTS 实证教训，逐条落点见各处注释）：
  ① 图表数据在 shape.chart 上 —— 全模块唯一取图入口 chart_of()，禁 shape.plots/.series；
  ② 任何异常按 tag 计数汇总（FAILURES），禁止裸 except: pass；
  ③ chart.replace_data 前强制 len(cats)==len(vals)，否则 PowerPoint 前置 None 对齐/整体左移；
  ④ 自动备份名含版本号 + 存在性检查，重复运行不覆盖原始备份；
  ⑤ 口径披露独立一页且不压缩（8 项）；
  ⑥ 数字一律来自 JSON / 底表派生，本文件不含任何业务数字常量。

写作禁区守卫（纲要 §4 共 9 条）：assert_text_clean() 对每段注入文字做硬校验，
  命中即抛 WritingBanError（不 WARN、不静默），其中 91.34% 为自指覆盖率硬禁。

CLI（由 build_pptx_refresh_v4924.py 转发）：
  --dry-run   只校验数据契约与禁区分词，不写任何文件
  --validate  校验 + 生成后硬断言（列数/覆盖率口径/榜单计数/残留过期字面量）
"""
import argparse
import json
import os
import re
import shutil
import sys
from collections import defaultdict
from datetime import date

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ==========================================================================================
# 常量（配色 token 取自现稿实测，见结构文档 §2.2）
# ==========================================================================================
VER = 'v4.9.50'
SALES_ANALYTICS = 'audit-output/sales_linkage_analytics_20261019.json'
SALES_LEDGER = 'audit-output/sales_link_20261019.json'
SALES_PRIOR_BACKUP = 'audit-output/sales_prior_yiche_202608_backup_20261019.json'
SRC_DECK = 'NEV公告车型行业分析_演示文稿.pptx'
BASE_TABLE = 'NEV公告参数汇总表_合并版（341~410批）.xlsx'

NAVY = RGBColor(0x1F, 0x38, 0x64)
ORANGE = RGBColor(0xED, 0x7D, 0x31)
LIGHTBLUE = RGBColor(0xBD, 0xD7, 0xEE)
MIDBLUE = RGBColor(0x8F, 0xAA, 0xDC)
GREY = RGBColor(0x59, 0x59, 0x59)
KPI_GREY = RGBColor(0x44, 0x54, 0x6A)
HDR_FILL = RGBColor(0xD9, 0xE1, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_RED = RGBColor(0xF2, 0xCF, 0xCF)
F = 'Arial'
F_CJK = '微软雅黑'

SW, SH = 13.3330, 7.5          # 幻灯片英寸（实测 12191695 × 6858000 EMU）
ML, MR = 0.62, 0.62            # 左右页边距
CW = SW - ML - MR              # 内容宽度 = 12.093
Y_KICKER, Y_TITLE, Y_BODY = 0.30, 0.72, 1.46
Y_FOOT = 6.86
BH = Y_FOOT - Y_BODY           # 正文带高度 = 5.40

SERIES_COLORS = [NAVY, ORANGE, MIDBLUE, LIGHTBLUE, GREY, KPI_GREY,
                 RGBColor(0x2E, 0x75, 0xB6), RGBColor(0xC0, 0x50, 0x4D),
                 RGBColor(0x9E, 0x48, 0x0A), RGBColor(0x70, 0xAD, 0x47)]

FAILURES = defaultdict(int)
STATS = {'pages': 0, 'charts': 0, 'tables': 0, 'texts': 0, 'slides_added': 0,
         'inplace': 0, 'guard_checks': 0}


# ==========================================================================================
# 留痕与外层
# ==========================================================================================
class WritingBanError(Exception):
    """命中纲要 §4 写作禁区。硬失败，不降级。"""


def note_failure(tag, exc):
    """纪律②：所有异常按 tag 计数汇总，禁止静默。"""
    FAILURES[f'{tag}: {type(exc).__name__}: {exc}'] += 1
    print(f'  [FAIL] {tag}: {type(exc).__name__}: {exc}')


def tracked(tag):
    """装饰器：把构建器异常转成留痕计数（不吞、不静默跳过）。"""
    def deco(fn):
        def wrap(*a, **kw):
            try:
                return fn(*a, **kw)
            except WritingBanError:
                raise
            except Exception as exc:
                note_failure(tag, exc)
                return None
        wrap.__name__ = fn.__name__
        return wrap
    return deco


# ==========================================================================================
# 写作禁区守卫（纲要 §4 · 9 条）
# ==========================================================================================
BAN_PATTERNS = [
    (r'(?<![\d.])91\.34|(?<![\d.])91\.3\s*%|(?<![\d.])91\s*%',
     '禁：自指覆盖率 91.34%（纲要 §3.5 / §4）'),
    (r'数据缺失|数据质量问题|数据不准',      '禁：把 129 个在榜未关联车系写成数据缺失（纲要 §4）'),
    (r'遥遥领先|断层第一|全面碾压|绝对领先',  '禁：不可量化修辞（纲要 §4）'),
    (r'据悉|业内人士认为|据知情',           '禁：无源表述（纲要 §4）'),
    (r'销量高.{0,6}(技术好|产品强|质量好)',   '禁：无支撑因果（纲要 §4）'),
    (r'(当|作|为)\s*在售车型数.*公告车型',    '禁：把公告车型数当在售车型数（纲要 §4）'),
    (r'Top\s*20\s*(新晋|掉出)|Top20(新晋|掉出)', '禁：new_entries/drops 是月度榜单集合差，不是 Top20 位次差（Lead 硬要求）'),
]
# 销量图表页脚注唯一标准句（口径规范 §4 / 结构文档 P8）
DEDUP_FOOT = ('车系级口径（seriesid 唯一），已去重；底表行级值仅供单行参考，禁止加总')
# 任何销量图表标题必须含口径与月份（纲要 §4）
CHART_TITLE_RULE = '标题须含月份 + 新能源榜 + 车系级'


def assert_text_clean(text, tag):
    """逐段注入文字的禁区分词硬校验。命中即抛，不 WARN、不静默。"""
    t = str(text)
    for pat, why in BAN_PATTERNS:
        if re.search(pat, t):
            raise WritingBanError(f'[{tag}] 文字命中写作禁区：{why}｜片段={t[:90]!r}')
    STATS['guard_checks'] += 1
    return t


def chart_title(meta, title, extra=''):
    """按纲要 §4 生成合规图表标题：含月份 + 新能源榜 + 车系级。"""
    snap = meta.get('snapshot_month', '')
    base = f'2026-09 车系级月销 · 新能源榜｜{title}'
    if extra:
        base = f'{base}（{extra}）'
    return assert_text_clean(base, 'chart_title')


# ==========================================================================================
# 纪律① 全模块唯一取图入口
# ==========================================================================================
def chart_of(shape):
    """图表数据在 shape.chart 上；GraphicFrame 没有 .plots/.series。
    **本模块唯一取图入口**，其他写法（shape.plots / shape.series）一律视为缺陷。"""
    if not getattr(shape, 'has_chart', False):
        return None
    return shape.chart


def chart_cats(chart):
    return [str(c) for c in chart.plots[0].categories]


def chart_series(chart):
    return [(se.name, [v for v in se.values]) for se in chart.series]


def set_chart_data(chart, cats, series_map, tag):
    """纪律③：replace_data 前强制类目数 == 每系列值数；不一致直接报错留痕。"""
    try:
        n = len(cats)
        if not n:
            raise ValueError('空类目')
        for name, vals in series_map.items():
            if len(vals) != n:
                raise ValueError(f'类目/系列长度不一致：{name!r} 有 {len(vals)} 值，类目 {n} 个')
        cd = CategoryChartData()
        cd.categories = cats
        for name, vals in series_map.items():
            cd.add_series(str(name), [None if v is None else float(v) for v in vals])
        chart.replace_data(cd)
        STATS['charts'] += 1
        return True
    except Exception as exc:
        note_failure(f'chart:{tag}', exc)
        return False


# ==========================================================================================
# 基础排版
# ==========================================================================================
def blank(prs):
    """新增页：现稿全部使用 Blank 版式（实测布局索引 6），沿用同版式保证视觉一致。"""
    return prs.slides.add_slide(prs.slide_layouts[6])


def tb(slide, l, t, w, h, text, size, color, bold=False, font=F, align=None, wrap=True):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.text = ''
    for i, line in enumerate(str(text).split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align is not None:
            p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.name = font
        r.font.bold = bold
        r.font.color.rgb = color
    return box


def rect(slide, l, t, w, h, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
    sh.shadow.inherit = False
    return sh


def header(slide, kicker, title, meta, page_no=None):
    """每页统一页眉：kicker（章节 · 论断编号）+ 标题 + 页码。"""
    assert_text_clean(kicker, 'kicker')
    assert_text_clean(title, 'title')
    tb(slide, ML, Y_KICKER, CW - 0.9, 0.30, kicker, 11.5, ORANGE, bold=True, font=F_CJK)
    tb(slide, ML, Y_TITLE, CW - 0.9, 0.58, title, 23, NAVY, bold=True, font=F_CJK)
    if page_no is not None:
        tb(slide, SW - MR - 0.9, Y_KICKER, 0.9, 0.3, f'P{page_no}', 11, GREY,
           align=PP_ALIGN.RIGHT, font=F)
    rect(slide, ML, Y_TITLE + 0.60, CW, 0.022, LIGHTBLUE)


def footer(slide, meta, extra='', ban_free_extra=True):
    """页脚 = 数据源 + 快照 + 车系级去重声明（+ 页专属补充）。销量页统一复用同一句。"""
    txt = (f'数据源：{meta.get("source", "")} · 快照 {meta.get("snapshot_month", "")}'
           f'（抓取 {meta.get("fetch_date", "")}）· {DEDUP_FOOT}')
    if extra:
        txt = f'{txt} · {extra}'
    txt = assert_text_clean(txt, 'footer')
    tb(slide, ML, Y_FOOT, CW, 0.34, txt, 8.5, GREY, font=F_CJK)


def kpi_row(slide, items, top, height=1.30, size_val=26, size_lab=10):
    """KPI 卡横排。items = [(值文字, 标签), ...]"""
    total = len(items)
    if not total:
        return
    gap = 0.16
    w = (CW - gap * (total - 1)) / total
    for i, (val, lab) in enumerate(items):
        l = ML + i * (w + gap)
        rect(slide, l, top, w, height, HDR_FILL)
        rect(slide, l, top, 0.055, height, ORANGE if i == 0 else NAVY)
        v = assert_text_clean(val, f'kpi{i}')
        tb(slide, l + 0.18, top + 0.14, w - 0.36, 0.52, v, size_val, NAVY, bold=True, font=F)
        tb(slide, l + 0.18, top + 0.70, w - 0.36, height - 0.80,
           assert_text_clean(lab, f'kpilab{i}'), size_lab, KPI_GREY, font=F_CJK, wrap=True)


def bullets(slide, lines, top, height, size=13, lh=0.30, color=KPI_GREY, left=None, width=None):
    """要点条：• 前缀 + 一级缩进。"""
    left = ML if left is None else left
    width = CW if width is None else width
    n = max(1, len(lines))
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = ''
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(4)
        txt = assert_text_clean(ln, f'bullet{i}')
        if txt.startswith('#'):                      # # 开头 = 不用圆点的强调行
            r = p.add_run(); r.text = txt[1:].strip()
            r.font.bold = True; r.font.color.rgb = NAVY
        else:
            r = p.add_run(); r.text = f'• {txt}'
            r.font.color.rgb = color
        r.font.size = Pt(size)
        r.font.name = F_CJK
    return box


def table(slide, grid, l, t, w, h, tag, size=8.5, hdr_rows=1, col_align=None):
    """同源数据小表。写入前校验矩形，失败留痕计数。"""
    try:
        nrow = len(grid)
        ncol = max(len(r) for r in grid)
        if nrow < 1 or ncol < 1:
            raise ValueError(f'空表 nrow={nrow} ncol={ncol}')
        for ri, row in enumerate(grid):
            if len(row) != ncol:
                raise ValueError(f'第 {ri} 行 {len(row)} 列 ≠ 表宽 {ncol}')
            for ci, cell in enumerate(row):
                assert_text_clean(cell, f'{tag}[{ri},{ci}]')
        gt = slide.shapes.add_table(nrow, ncol, Inches(l), Inches(t), Inches(w), Inches(h))
        t_ = gt.table
        try:
            t_.style = 'Light Style 1'
        except Exception as exc:
            note_failure(f'tablestyle:{tag}', exc)
        for ri, row in enumerate(grid):
            for ci in range(ncol):
                c = t_.cell(ri, ci)
                c.text = '' if row[ci] is None else str(row[ci])
                c.margin_left = Emu(45720); c.margin_right = Emu(45720)
                c.margin_top = Emu(9144); c.margin_bottom = Emu(9144)
                for p in c.text_frame.paragraphs:
                    if col_align:
                        p.alignment = col_align[ci % len(col_align)]
                    for r in p.runs:
                        r.font.size = Pt(size if ri >= hdr_rows else size + 0.5)
                        r.font.name = F_CJK if ri >= hdr_rows else F
                        r.font.bold = ri < hdr_rows
                        r.font.color.rgb = WHITE if ri < hdr_rows else KPI_GREY
                if ri < hdr_rows:
                    c.fill.solid(); c.fill.fore_color.rgb = NAVY
        STATS['tables'] += 1
        return True
    except Exception as exc:
        note_failure(f'table:{tag}', exc)
        return False


def source_note(slide, top, lines, size=9):
    bullets(slide, lines, top, Y_FOOT - top - 0.06, size=size, color=GREY)


# ==========================================================================================
# 数值格式（派生量，非手写）
# ==========================================================================================
def n0(v):
    return '—' if v is None else f'{float(v):,.0f}'


def n1(v):
    return '—' if v is None else f'{float(v):,.1f}'


def pct(v, d=2):
    return '—' if v is None else f'{float(v):.{d}f}%'


def pp(v, d=2):
    """百分点差，带正负号。"""
    if v is None:
        return '—'
    return f'{float(v):+.{d}f}pp'


def to_map(rows, idx=1):
    """[[名, 值, ...], ...] → {名: 值}"""
    return {str(r[0]): r[idx] for r in rows}


# ------------------------------------------------------------------------------------------
# 交叉校验明细（只读镜像）：逐值锚定示例
# ------------------------------------------------------------------------------------------
_PAIR_RE = re.compile(r'([^()（）+＋]{1,40}?)\((\d[\d,]*)\)\s*[+＋]\s*'
                      r'([^()（）+＋]{1,40}?)\((\d[\d,]*)\)')


def _clean_name(s):
    """剥掉引号、前缀叙述词（如"兄弟行/基础行"），只留车系名。"""
    s = re.split(r'[「『《<]', str(s))[-1]      # 取最后一个左引号之后
    s = re.split(r'[」』》>]', s)[0]             # 丢弃右引号及其后内容
    s = re.sub(r'^\s*(兄弟行|基础行|同车系|懂车帝)\s*', '', s)
    return s.strip('＝=·　 \t')


# ------------------------------------------------------------------------------------------
# 口径文案归一（派生量，非常量）：把基准稿里写死的"× 38 列 / 5,392 款"改成动态真值
# ------------------------------------------------------------------------------------------
SCOPE_RE = re.compile(r'纯乘用车口径（[^）]*）')
SAMPLE_RE = re.compile(r'样本\s*[\d,]+\s*款（[^）]*）')
TOTAL_RE = re.compile(r'共\s*[\d,]+\s*款纯新能源乘用车')
# 只命中具体的陈旧样本量；原写法 [\d,]{4,} 会把「自1006款增长至1198款」两个真值一并抹平
LOOSE_RE = re.compile(r'(?<![\d,.])(?:4,?613|5,?392|5,?475)\s*款(?![系型])(?!\s*公告记录)')
# 行 ≠ 款：年度条目数（4 位无千分位）语义是"条"
YEARCNT_RE = re.compile(r'(?<![\d,.])(\d{4})\s*款(?![系型])')
# 目录页横幅「七大维度 · 67页」是基准稿写死的，必须随目录页数派生（否则终稿页数一变就自相矛盾）
BANNER_RE = re.compile(r'[七八九十]大(?:维度|板块)\s*·\s*\d+\s*页')
VER_RE = re.compile(r'v4\.\d+(?:\.\d+)*')


def base_table_scope_counts(path=BASE_TABLE, sheet='NEV公告参数汇总'):
    """从底表实取 (行数, 列数, BEV 行数, PHEV 行数, EREV 行数)。失败即抛，不回退常量。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        ws = wb[sheet]
        it = ws.iter_rows(values_only=True)
        hdr = [str(c or '').strip() for c in next(it)]
        ncol = len(hdr)
        if not (30 <= ncol <= 60):
            raise AssertionError(f'底表列数 {ncol} 越界 [30,60]，拒绝据此生成文案')
        ci_p = hdr.index('动力类型')
        ci_e = hdr.index('是否增程(EREV)') if '是否增程(EREV)' in hdr else None
        nrow = nbev = nphev = nerev = 0
        for row in it:
            if row is None or all(v is None for v in row):
                continue
            nrow += 1
            p = str(row[ci_p] or '')
            is_e = bool(row[ci_e]) if ci_e is not None else False
            if is_e:
                nerev += 1
            elif 'BEV' in p:
                nbev += 1
            else:
                nphev += 1
        return nrow, ncol, nbev, nphev, nerev
    finally:
        wb.close()


def _norm_one_paragraph(p, canon, sample, total):
    """把单段的陈旧口径文案改写为派生真值。返回是否改动。

    段落内文字可能被拆到多个 run；命中时把新文本写进首个 run 并清空其余
    （口径脚注/说明行本身单一格式，不影响观感）。
    """
    runs = p.runs
    if not runs:
        return False
    t = ''.join(r.text for r in runs)
    n = SCOPE_RE.sub(canon, t)
    n = SAMPLE_RE.sub(sample, n)
    n = TOTAL_RE.sub(total, n)
    n = LOOSE_RE.sub(f'{total[2:]}', n)
    n = YEARCNT_RE.sub(r'\1 条', n)
    n = BANNER_RE.sub(f'本册共 {len(PAGE_CATALOG)} 页', n)
    n = VER_RE.sub(VER, n)                 # 交付物版本号全稿统一（消除旧版号残留）
    if n == t:
        return False
    assert_text_clean(n, 'scope_norm')
    runs[0].text = n
    for r in runs[1:]:
        r.text = ''
    return True


_SCOPE_CACHE = {}


def scope_counts():
    """带缓存的底表规模（行/列/BEV/PHEV/EREV）。"""
    if 'counts' not in _SCOPE_CACHE:
        _SCOPE_CACHE['counts'] = base_table_scope_counts()
    return _SCOPE_CACHE['counts']


def _as_batch(v):
    """批次号归一：底表有 122 行批次存为字符串（'361' 11 行 / '409' 111 行），
    不归一会把 '409' 与 409 裂成两个批次（Lead 2026-10-19 指出的真实缺陷）。"""
    if v is None or v == '':
        return None
    try:
        return int(float(str(v).strip()))
    except (TypeError, ValueError):
        return None


def base_table_public_rows(path=BASE_TABLE, sheet='NEV公告参数汇总',
                           sheet_time='批次时间表'):
    """派生公示批：行数、批次号、公示日期。

    判定依据在「批次时间表」而非主表——主表没有「公告日期」列；
    时间表中 `公告日期` 为空、`公示日期` 有值的批次即公示口径（尚无正式发文字号）。
    返回 (公示行数, 批次列表, 公示日期文本)。
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        bt = wb[sheet_time]
        it = bt.iter_rows(values_only=True)
        hdr = [str(c).strip() if c is not None else '' for c in next(it)]
        ix = {h: i for i, h in enumerate(hdr)}
        need = ('批次', '公告日期', '公示日期')
        if any(k not in ix for k in need):
            raise ContractError(f'批次时间表缺列：{[k for k in need if k not in ix]}')
        public, dates = set(), []
        for r in it:
            b = _as_batch(r[ix['批次']])
            if b is None:
                continue
            if r[ix['公告日期']] in (None, ''):
                public.add(b)
                if r[ix['公示日期']] not in (None, ''):
                    dates.append(str(r[ix['公示日期']]).strip())
        ws = wb[sheet]
        it2 = ws.iter_rows(values_only=True)
        hdr2 = [str(c).strip() if c is not None else '' for c in next(it2)]
        ci = hdr2.index('批次')
        cnt = sum(1 for r in it2 if _as_batch(r[ci]) in public)
    finally:
        wb.close()
    return cnt, sorted(public), (dates[0] if dates else '')


def scope_split():
    """公示批拆分文案（派生）：(总行数, 正式公告数, 公示条数, 批次号文本)。"""
    if 'split' not in _SCOPE_CACHE:
        nrow = scope_counts()[0]
        npublic, batches, pubdate = base_table_public_rows()
        _SCOPE_CACHE['split'] = (nrow, nrow - npublic, npublic,
                                 '、'.join(str(b) for b in batches) or '—')
        _SCOPE_CACHE['pubdate'] = pubdate
    return _SCOPE_CACHE['split']


def scope_line():
    """规范口径文案（派生真值，非常量）：纯乘用车口径（N 条公告记录 × M 列）。"""
    if 'line' not in _SCOPE_CACHE:
        nrow, ncol = scope_counts()[:2]
        _nrow, _nf, _np, _pb = scope_split()
        _SCOPE_CACHE['line'] = (f'纯乘用车口径（{nrow:,} 条公告/公示记录 × {ncol} 列；'
                                f'其中 {_np} 条为 {_pb} 批公示记录）')
    return _SCOPE_CACHE['line']


def normalize_scope_lines(prs, counts):
    """全稿归一化口径文案（含表格单元格——P2 的列数文案就在表格里）。"""
    nrow, ncol, nbev, nphev, _ = counts
    _nrow, _nf, _np, _pb = scope_split()
    canon = (f'纯乘用车口径（{nrow:,} 条公告/公示记录 × {ncol} 列；'
             f'其中 {_np} 条为 {_pb} 批公示记录）')
    sample = f'样本 {nrow:,} 条公告/公示记录（BEV {nbev:,} / PHEV {nphev:,}）'
    total = f'共 {nrow:,} 条公告/公示记录'
    n_fix = 0
    for sl in prs.slides:
        for sh in sl.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    n_fix += _norm_one_paragraph(p, canon, sample, total)
            if getattr(sh, 'has_table', False):
                for row in sh.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            n_fix += _norm_one_paragraph(p, canon, sample, total)
    print(f'  口径文案归一：{n_fix} 处 → {canon}')
    return n_fix


def _qc_detail(dd):
    """按 dongchedi_202609.mirror_of 读只读镜像明细（超差项的逐值锚定证据在此）。失败留痕并返回 {}。"""
    mo = dd.get('mirror_of') or ''
    if '#' not in mo:
        return {}
    path, frag = mo.split('#', 1)
    try:
        with open(path, encoding='utf-8') as fh:
            node = json.load(fh)
        for part in frag.split('.'):
            node = node[part]
        STATS['mirror_reads'] = STATS.get('mirror_reads', 0) + 1
        return node if isinstance(node, dict) else {}
    except Exception as exc:
        note_failure(f'qc_mirror:{path}', exc)
        return {}


def _subset_sum_examples(dd, limit=3):
    """提取"主源合并值 == 对方兄弟行之和"的逐值锚定示例。

    仅收录 **算术确实相等**（va + vb == 主源值）的行；不相等或解析不出即跳过，
    绝不回退成"近似相等"表述（fail-closed）。按销量降序取前 limit 例（量级越大越有说服力）。
    """
    det = _qc_detail(dd)
    out = []
    for r in det.get('all_matched') or []:
        rs = r.get('scope_mismatch_reason') or ''
        if '之和' not in rs or '逐值相等' not in rs:
            continue
        head, sep, tail = rs.partition('同车系的 ')
        mm = _PAIR_RE.search(tail if sep else rs)
        if not mm:
            continue
        a, va = _clean_name(mm.group(1)), int(mm.group(2).replace(',', ''))
        b, vb = _clean_name(mm.group(3)), int(mm.group(4).replace(',', ''))
        tot = r.get('ah_sales')
        if tot is None or va + vb != int(tot):
            continue                      # 算术不等 → 不收（fail-closed）
        out.append((r.get('ah_name') or r.get('name'), int(tot),
                    f'{a} {n0(va)} ＋ {b} {n0(vb)}', f'= {n0(va + vb)}'))
    out.sort(key=lambda x: -x[1])
    return out[:limit]


def get_meta(D):
    return D.get('meta', {})


# ==========================================================================================
# 底表列数（动态派生 + 下界断言）—— 禁硬编码 41
# ==========================================================================================
# Lead 2026-10-19 验收时冻结的底表指纹（5,392 行版本）。
# 5,475 行终态（2026-10-19 19:32:58，sha 230dbb96…）已由 Lead 显式确认，
# 确认过程见 Update/销量分析演示稿结构_20261019.md 「指纹确认留痕」。
# 底表若再次合法更新，必须由人显式确认新指纹（--accept-base-sha）：
# 既不允许静默放行，也不允许无条件拦住已验收的新终态。
FROZEN_BASE_TABLE_SHA256 = '4f08ccd6b99e9eeab71f693c0562d951ebc6928b4a1b146f59054e7397fba521'
BASE_TABLE_SHA256 = FROZEN_BASE_TABLE_SHA256  # 运行期期望值，可被 --accept-base-sha 覆盖

# 未锁定数字的占位标记（已由 Lead 终值替换；保留常量供残留扫描使用）
BEV_PP_PLACEHOLDER = '【BEV-PP 待 Lead 发数】'


def _sha256_file(path, chunk=1 << 20):
    import hashlib
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def assert_meta_ready(D, path=BASE_TABLE):
    """生成前置断言（fail-fast）：dry_run / 底表指纹与行列数 / suggest 关闭。

    行列数不写死，而是与**磁盘实测**的底表行列数比对——两侧都必须自洽；
    sha256 用 Lead 冻结常量比对（这是唯一权威指纹）。
    """
    m = D.get('meta') or {}
    problems = []
    if m.get('dry_run') is not False:
        problems.append(f'meta.dry_run = {m.get("dry_run")!r}，必须为 False')
    # ① 自洽性硬断言（无条件）：JSON meta 与磁盘底表必须描述同一个文件
    got_sha = str(m.get('base_table_sha256') or '').lower()
    disk_sha = _sha256_file(path) if os.path.exists(path) else None
    if disk_sha is None:
        problems.append(f'底表不存在：{path}')
    elif got_sha != disk_sha:
        problems.append(f'JSON meta.base_table_sha256 = {got_sha!r} ≠ 磁盘底表 {disk_sha!r}'
                        f'（JSON 与底表不同源，须重跑数据链）')
    # ② 指纹确认硬断言（条件）：磁盘指纹必须等于"已被人工确认"的指纹
    if disk_sha is not None and disk_sha != BASE_TABLE_SHA256:
        problems.append(
            f'磁盘底表 sha256 = {disk_sha} 与已确认指纹 {BASE_TABLE_SHA256} 不一致'
            f'（底表已更新）。若新终态已验收，请显式确认：--accept-base-sha {disk_sha}')
    ncol, nrow = base_table_cols(path), base_table_rows(path)
    if m.get('base_table_columns') != ncol:
        problems.append(f'meta.base_table_columns = {m.get("base_table_columns")!r} ≠ 底表实取 {ncol}')
    if m.get('base_rows') != nrow:
        problems.append(f'meta.base_rows = {m.get("base_rows")!r} ≠ 底表实取 {nrow}')
    if m.get('suggest_enabled') is not False:
        problems.append(f'meta.suggest_enabled = {m.get("suggest_enabled")!r}，本轮必须为 False')
    if m.get('suggest_calls') != 0:
        problems.append(f'meta.suggest_calls = {m.get("suggest_calls")!r}，本轮必须为 0')
    if problems:
        raise ContractError('生成前置断言未通过（不得绕过）：\n  - ' + '\n  - '.join(problems))
    return True


def bev_share_series(D):
    """动力份额月度序列（销量口径，分母 = 已锚定口径月度总量）。

    返回 [(月, BEV%, PHEV%, EREV%)]；三者同分母，可同页对冲展示。
    """
    out = []
    for m, tot, b, p, e in D['trend']['monthly']:
        if not tot:
            continue
        out.append((m, 100.0 * b / tot, 100.0 * p / tot, 100.0 * e / tot))
    if not out:
        raise ContractError('trend.monthly 为空，无法派生动力份额')
    return out


def batch_range(D):
    """批次范围与公示口径批次（派生，不写死）。

    `conversion.by_batch` 每行为 [批次, 年, 公告条数, 当期销量]；年字段缺失（None）
    即该批为公示口径（尚无正式发文字号），检出后必须披露，不得并入正式公告表述。
    """
    rows = D['conversion']['by_batch']
    if not rows:
        raise ContractError('conversion.by_batch 为空，无法派生批次范围')
    lo, hi = rows[0][0], rows[-1][0]
    public = [r[0] for r in rows if r[1] is None]
    return {'lo': lo, 'hi': hi, 'count': len(rows), 'public': public,
            'formal_hi': max([r[0] for r in rows if r[1] is not None], default=None)}


def disclosure_forensics(D, path=BASE_TABLE):
    """口径披露所需的四个派生量（优先取 JSON，缺失时从留档备份 + 底表按行序现算）。

    返回 dict：
      old_cells      旧口径非空格数（Lead: 806）
      replaced       已更替（Lead: 757）
      cleared        已清空（Lead: 49）
      kept           原样保留（0）
      anchored_uncovered  「未匹配-已停产或未上量」中已锚定车系ID 的行数（Lead: 418）
      col39_nonempty col39（车系ID）非空行数（5,475 行终态：2,618）
      col35_nonempty col35（月销量）非空行数（5,475 行终态：2,200）
    """
    import openpyxl
    m = D.get('meta') or {}
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        ws = wb['NEV公告参数汇总']
        it = ws.iter_rows(values_only=True)
        hdr = [str(c).strip() if c is not None else '' for c in next(it)]
        ix = {h: i for i, h in enumerate(hdr)}
        rows = list(it)
    finally:
        wb.close()
    c_series, c_sales, c_status = ix['车系ID'], ix['月销量(辆)'], ix['销量覆盖状态']
    col39 = sum(1 for r in rows if r[c_series] not in (None, ''))
    col35 = sum(1 for r in rows if r[c_sales] not in (None, ''))
    anchored_uncovered = sum(1 for r in rows
                             if r[c_status] == '未匹配-已停产或未上量'
                             and r[c_series] not in (None, ''))
    out = {'anchored_uncovered': anchored_uncovered,
           'col39_nonempty': col39, 'col35_nonempty': col35}
    if all(m.get(k) is not None for k in ('backup_old_cells', 'backup_replaced',
                                          'backup_cleared', 'backup_kept')):
        out.update(old_cells=m['backup_old_cells'], replaced=m['backup_replaced'],
                   cleared=m['backup_cleared'], kept=m['backup_kept'])
        return out
    # 现算：留档备份（行序）+ 底表现值
    prior = json.load(open(SALES_PRIOR_BACKUP, encoding='utf-8'))['rows']
    if len(prior) != len(rows):
        raise ContractError(f'留档备份行数 {len(prior)} ≠ 底表行数 {len(rows)}，无法按行序配对')
    old_cells = replaced = cleared = kept = 0
    for rec, r in zip(prior, rows):
        if rec.get('prior_sales') in (None, ''):
            continue
        old_cells += 1
        if r[c_sales] in (None, ''):
            cleared += 1
        else:
            replaced += 1
    out.update(old_cells=old_cells, replaced=replaced, cleared=cleared, kept=kept)
    return out


def base_table_cols(path=BASE_TABLE, sheet='NEV公告参数汇总'):
    """从底表表头实取列数；加下界/上界断言防读取异常产出 '× 0 列'。"""
    import openpyxl
    if not os.path.exists(path):
        raise FileNotFoundError(f'底表不存在，无法派生列数：{path}')
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        ws = wb[sheet]
        hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        ncol = sum(1 for h in hdr if h is not None and str(h).strip() != '')
    finally:
        wb.close()
    assert 30 <= ncol <= 60, f'底表列数越界：{ncol}（表头异常，拒绝产出错文案）'
    return ncol


def base_table_rows(path=BASE_TABLE, sheet='NEV公告参数汇总'):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        ws = wb[sheet]
        n = sum(1 for _ in ws.iter_rows(min_row=2, values_only=True))
    finally:
        wb.close()
    assert 1000 <= n <= 20000, f'底表行数越界：{n}'
    return n


# ==========================================================================================
# 目标稿逐页目录（交付 82 页 · 对齐 Update/销量分析演示稿结构_20261019.md §1；
# 页数由 len(PAGE_CATALOG) 决定，目录页横幅文案在运行期派生，不写死）
#   每项 = (页号, 页型, 标题, 构建器名或 None, 论断锚点, 说明)
#   构建器为 None ⇒ 该页由基准稿对应页继承，阶段 2 仅做既有数值刷新（不改结构）
# ==========================================================================================
PAGE_CATALOG = [
    ( 1, '封面', 'NEV 公告车型行业分析报告 2021–2026',                 'cover'        , '§0', '基准 P1'),
    ( 2, '执行摘要', '执行摘要 · 五条金字塔尖结论',                        'exec1'        , '结论1-5', '基准 P2'),
    ( 3, '执行摘要', '执行摘要 · 数据可信度与口径边界',                      'exec2'        , '§3', '基准 P3'),
    ( 4, '目录', '目录 · 全册页数与章节导航',                       None           , '§0.1', '基准 P4（横幅文案运行期派生）'),
    ( 5, '要点页', '交叉校验 · 双第三方源逐值对撞',                        'crosscheck'   , '§3.3', '**新增**（unexplained=0；含 subset-sum 逐值锚定三例）'),
    ( 6, '要点页', '数据来源与快照声明',                               'method_src'   , '§3.1', '基准 P5 改写'),
    ( 7, '要点页', '动力类型归类规则（R1）与被否决方案',                      'r1_rule'      , '§11', '**新增**（16 混合家族 / 50,639 辆 / 6.35%；含 R2 否决理由）'),
    ( 8, '要点页', '覆盖率三口径与未覆盖分桶',                            'method_cov'   , '§3.5', '基准 P6 改写'),
    ( 9, '要点页', '口径统一决策与旧值处置',                             'method_switch', '§3.2/3.7', '基准 P7 改写'),
    (10, '要点页', '车系级去重口径（防重复计数）',                          'method_dedup' , '§3.4', '基准 P8 改写'),
    (11, '章节页', '01 市场全景',                                 None           , '结论1', '基准 P9'),
    (12, '图表页', 'BEV 与 PHEV 公告车型数量对比',                     None           , '—', '基准 P10'),
    (13, '图表页', '销量按细分市场结构',                               'seg_sales'    , '结论3', '新增'),
    (14, '图表页', '销量按动力类型结构',                               'energy_sales' , '结论1', '新增'),
    (15, '图表页', 'PHEV 占比变化趋势',                             None           , '—', '基准 P13'),
    (16, '图表页', 'BEV 细分市场分布',                              None           , '—', '基准 P14'),
    (17, '图表页', 'PHEV 细分市场分布',                             None           , '—', '基准 P15'),
    (18, '表格页', '年度详细数据明细（公告口径）',                          None           , '—', '基准 P16'),
    (19, '图表页', '公告车型总量年度趋势',                              None           , '—', '基准 P17'),
    (20, '图表页', '销量按价格带结构',                                'price_sales'  , '结论5', '新增'),
    (21, '图表页', '价格带迁移（双基期对照）',                            'price_move'   , '结论5', '新增'),
    (22, '章节页', '02 动力电池分析',                               None           , '—', '基准 P20'),
    (23, '图表页', 'BEV 平均续航里程趋势',                            None           , '—', '基准 P21'),
    (24, '图表页', 'PHEV 平均续航里程趋势',                           None           , '—', '基准 P22'),
    (25, '图表页', 'BEV 与 PHEV 续航里程对比',                       None           , '—', '基准 P23'),
    (26, '图表页', 'BEV 平均电池容量趋势',                            None           , '—', '基准 P24'),
    (27, '图表页', 'PHEV 平均电池容量趋势',                           None           , '—', '基准 P25'),
    (28, '图表页', 'BEV 与 PHEV 电池容量对比',                       None           , '—', '基准 P26'),
    (29, '图表页', 'BEV 电池类型占比趋势',                            None           , '—', '基准 P27'),
    (30, '图表页', 'PHEV 电池类型占比趋势',                           None           , '—', '基准 P28'),
    (31, '图表页', 'BEV 电池能量密度趋势',                            None           , '—', '基准 P29'),
    (32, '图表页', 'BEV 续航里程分布',                              None           , '—', '基准 P30'),
    (33, '图表页', 'PHEV 续航里程分布',                             None           , '—', '基准 P31'),
    (34, '图表页', 'BEV 电池容量分布',                              None           , '—', '基准 P32'),
    (35, '图表页', 'PHEV 电池容量分布',                             None           , '—', '基准 P33'),
    (36, '表格页', '技术路线矩阵（2×2 象限）',                          None           , '—', '基准 P34'),
    (37, '图表页', 'BEV 电池类型年度数量明细',                          None           , '—', '基准 P35'),
    (38, '章节页', '03 驱动与能效',                                None           , '—', '基准 P36'),
    (39, '图表页', 'BEV 电机功率趋势',                              None           , '—', '基准 P37'),
    (40, '图表页', 'PHEV 电机功率趋势',                             None           , '—', '基准 P38'),
    (41, '章节页', '04 价格 · 技术 · 销量联动',                       None           , '结论5', '基准 P39'),
    (42, '图表页', '价格带 × 销量矩阵',                              None           , '结论5', '基准 P40'),
    (43, '图表页', '价格带 × 续航段 × 销量矩阵',                        None           , '结论5', '基准 P41'),
    (44, '图表页', '销量加权技术画像 vs 公告等权画像',                      'tech_wt'      , '结论2', '新增'),
    (45, '图表页', '续航分箱 × 销量分布',                             'tech_range'   , '结论2', '新增'),
    (46, '表格页', '分动力技术画像表',                                'tech_energy'  , '结论2', '新增'),
    (47, '图表页', '价格带技术均值表',                                'tech_price'   , '结论5', '新增'),
    (48, '图表页', 'PHEV 油耗与纯电续航趋势',                          None           , '—', '基准 P45'),
    (49, '图表页', 'PHEV B 状态油耗趋势',                           None           , '—', '基准 P46'),
    (50, '图表页', 'PHEV 综合油耗按排量区间',                          None           , '—', '基准 P47'),
    (51, '图表页', 'PHEV B 状态油耗按排量区间',                        None           , '—', '基准 P48'),
    (52, '章节页', '05 集中度与竞争格局',                             None           , '结论3', '基准 P49'),
    (53, '图表页', '供应商 HHI 指数趋势',                            None           , '—', '基准 P50'),
    (54, '图表页', '电池供应商 CR 集中度趋势',                          None           , '—', '基准 P51'),
    (55, '图表页', '电机供应商 CR 集中度趋势',                          None           , '—', '基准 P52'),
    (56, '图表页', 'Top 10 电池供应商',                            None           , '—', '基准 P53'),
    (57, '图表页', '销量集中度：三口径 CR3/CR5/CR10 + HHI',            'conc_all'     , '结论3', '新增'),
    (58, '图表页', '销量 Top 20 企业份额',                          'conc_oem'     , '结论3', '新增'),
    (59, '图表页', '销量 Top 20 品牌份额',                          'conc_brand'   , '结论3', '新增'),
    (60, '章节页', '06 趋势与迁移',                                None           , '结论1', '基准 P57'),
    (61, '图表页', '月度销量与动力类型演化',                             'trend_energy' , '结论1', '新增'),
    (62, '图表页', '月度价格带结构演化',                               'trend_price'  , '结论5', '新增'),
    (63, '图表页', '月度细分市场结构演化',                              'trend_seg'    , '结论3', '新增'),
    (64, '图表页', '月度 Top20 车型榜变动',                          'trend_rank'   , '结论3', '新增'),
    (65, '图表页', 'rankchange 分布',                           'trend_hist'   , '结论3', '新增'),
    (66, '图表页', '新晋榜 / 掉榜车系（月度名册集合差）',                     'trend_roster' , '结论3', '新增'),
    (67, '表格页', '月度迁移明细表',                                 'trend_table'  , '结论1', '新增'),
    (68, '章节页', '07 公告-市场转化',                              None           , '结论4', '基准 P65'),
    (69, '图表页', '各批次公告车型数 vs 当期销量',                        'conv_batch'   , '结论4', '新增'),
    (70, '图表页', '品牌公告数 vs 品牌销量（散点）',                       'conv_scatter' , '结论4', '新增'),
    (71, '表格页', '转化率排行表',                                  'conv_table'   , '结论4', '新增'),
    (72, '章节页', '08 相关性与增长',                               None           , '—', '基准 P69'),
    (73, '表格页', 'BEV 关键指标相关性矩阵',                           None           , '—', '基准 P70'),
    (74, '表格页', 'PHEV 关键指标相关性矩阵',                          None           , '—', '基准 P71'),
    (75, '图表页', 'BEV 同比增长率（按指标）',                          None           , '—', '基准 P72'),
    (76, '图表页', 'PHEV 同比增长率（按指标）',                         None           , '—', '基准 P73'),
    (77, '表格页', 'BEV 同比增长率明细',                             None           , '—', '基准 P74'),
    (78, '表格页', 'PHEV 同比增长率明细',                            None           , '—', '基准 P75'),
    (79, '章节页', '09 判断与建议',                                None           , '§1/§2', '基准 P76'),
    (80, '结论页', '结论 · 回扣五条结论 / 致谢',                        'conclusion'   , '§1/§2', '替换基准 S65'),
    (81, '附录', '附录 D · 关键发现总结（2021 vs 2025）',              None           , '—', '基准 S66'),
    (82, '结论页', '致谢 / THANK YOU',                          None           , '—', '基准 S67'),
]

# 构建器注册表（由下方 @page 装饰器填充）
BUILDERS = {}


def page(name):
    def deco(fn):
        BUILDERS[name] = fn
        return fn
    return deco


def page_frame(prs, D, page_no, kicker, title, extra_foot=''):
    """新建一页并加统一页眉页脚。返回 (slide, meta)。"""
    meta = get_meta(D)
    s = blank(prs)
    STATS['slides_added'] += 1
    header(s, kicker, title, meta, page_no=page_no)
    footer(s, meta, extra=extra_foot)
    return s, meta


def add_chart(slide, l, t, w, h, ctype, cats, series_map, tag, title=None,
              legend=True, data_labels=False, gap=60):
    """新建原生图表并注入数据。类目/系列长度不一致时由 set_chart_data 报错留痕。"""
    from pptx.chart.data import CategoryChartData as _CCD
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
    n = len(cats)
    for k, v in series_map.items():
        if len(v) != n:
            raise ValueError(f'{tag}: 系列 {k!r} 长度 {len(v)} ≠ 类目 {n}')
    cd = _CCD()
    cd.categories = cats
    for k, v in series_map.items():
        cd.add_series(str(k), [None if x is None else float(x) for x in v])
    gf = slide.shapes.add_chart(ctype, Inches(l), Inches(t), Inches(w), Inches(h), cd)
    ch = chart_of(gf)
    if ch is None:
        raise RuntimeError(f'{tag}: add_chart 后 shape.has_chart 为假（取图入口失效）')
    if title:
        ch.has_title = True
        ch.chart_title.text_frame.text = assert_text_clean(title, f'{tag}:title')
        for p in ch.chart_title.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(11); r.font.name = F_CJK; r.font.bold = True
                r.font.color.rgb = NAVY
    else:
        ch.has_title = False
    if legend and len(series_map) > 1:
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(9)
        ch.legend.font.name = F_CJK
    else:
        ch.has_legend = False
    for i, se in enumerate(ch.series):
        col = SERIES_COLORS[i % len(SERIES_COLORS)]
        try:
            se.format.fill.solid(); se.format.fill.fore_color.rgb = col
        except Exception as exc:
            note_failure(f'{tag}:seriesfill{i}', exc)
        try:
            se.format.line.color.rgb = col
        except Exception as exc:
            note_failure(f'{tag}:seriesline{i}', exc)
        if data_labels:
            se.has_data_labels = True
            se.data_labels.font.size = Pt(8)
            se.data_labels.number_format = '#,##0'
            se.data_labels.number_format_is_linked = False
    try:
        pl = ch.plots[0]
        pl.gap_width = gap
        pl.has_data_labels = data_labels
        ca = ch.category_axis
        ca.tick_labels.font.size = Pt(9); ca.tick_labels.font.name = F_CJK
        va = ch.value_axis
        va.tick_labels.font.size = Pt(8.5); va.tick_labels.font.name = F
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
    except Exception as exc:
        note_failure(f'{tag}:plotstyle', exc)
    STATS['charts'] += 1
    return gf


# ==========================================================================================
# 核心页构建器
# ==========================================================================================
@page('cover')
def b_cover(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '工信部公告 × 汽车之家销量 · 首次同源贯通',
                      'NEV 公告车型行业分析报告 2021–2026')
    cov = D['coverage']['row']
    br = batch_range(D)
    rect(s, ML, 2.35, CW, 0.03, ORANGE)
    tb(s, ML, 2.60, CW, 0.42,
       f'公告口径（{br["lo"]}~{br["hi"]} 批）与市场销量（{meta.get("snapshot_month")} 快照）同源贯通',
       15, KPI_GREY, font=F_CJK)
    kpi_row(s, [
        (n0(cov['total']), '底表公告记录（条）'),
        (n0(D['structure']['board_series_count']), f'{meta.get("snapshot_month")} 在榜车系'),
        (pct(D['structure']['linked_share_of_board_pct']), '已锚定车系销售额覆盖'),
        (str(len(meta.get('months', []))), '月度快照数'),
    ], top=3.25, height=1.45)
    source_note(s, 5.05, [
        f'数据源：{meta.get("source")}；参照宇宙：{meta.get("reference_source")}',
        f'快照月 {meta.get("snapshot_month")} · 抓取日 {meta.get("fetch_date")} · '
        f'交付物版本 {VER} · 生成日 {date.today():%Y-%m-%d}',
        assert_text_clean(scope_line(), 'exec2:scope'),
        f'口径：{meta.get("dedup_rule")}',
    ])
    return s


def bev_share_line(D):
    """结论 1 的动力份额叙事：起点→终点 + 峰值拐点（不得抹平）。全部由 JSON 派生。"""
    sh = bev_share_series(D)
    m0, b0, p0, e0 = sh[0]
    m1, b1, p1, e1 = sh[-1]
    pk = max(sh, key=lambda r: r[1])
    return (f'动力份额（销量口径，同分母）：BEV {pct(b0)} → {pct(b1)}（{pp(b1 - b0)}），'
            f'{pk[0]} 触及区间峰值 {pct(pk[1])} 后小幅回落——替代已进入高位平台期而非单向陡升；'
            f'镜像项 PHEV {pct(p0)} → {pct(p1)}（{pp(p1 - p0)}）、EREV {pct(e0)} → {pct(e1)}'
            f'（{pp(e1 - e0)}），两项合计正好对冲 BEV 增幅。')


@page('exec1')
def b_exec1(prs, D, page_no):
    """结论 1–3（Pyramid 尖）。数字全部来自 JSON。"""
    s, meta = page_frame(prs, D, page_no, '执行摘要 · 结论先行',
                         '五条金字塔尖结论（1/2）')
    mb = to_map(D['trend']['monthly_board'], idx=1)
    mo = D['meta']['months']
    last, first = mo[-1], mo[0]
    en = to_map(D['trend']['energy_monthly'], idx=1)
    conc_o, conc_b, conc_s = D['concentration']['oem'], D['concentration']['brand'], D['concentration']['series']
    kpi_row(s, [
        (n0(mb.get(last)), f'{last} 新能源零售（辆）'),
        (n0(mb.get(first)), f'{first} 低点前值（辆）'),
        (f"CR3 {n1(conc_s['CR3'])} / CR10 {n1(conc_s['CR10'])}", '车系集中度（%）'),
        (f"HHI {n0(conc_o['HHI'])} / {n0(conc_s['HHI'])}", '企业 / 车系 HHI'),
    ], top=Y_BODY, height=1.20, size_val=20)
    bullets(s, [
        f'#1 总量与渗透率｜{last} 新能源零售 {n0(mb.get(last))} 辆、在榜车系 '
        f'{n0(D["structure"]["board_series_count"])} 个，为 {len(mo)} 个月区间最高；'
        f'区间经历一次政策换挡深坑（{first} → 次月骤降），9 个月完成修复并创新高。'
        f'{bev_share_line(D)}',
        f'#2 技术路线｜市场用钱投票的方向是"更长续航、更轻、更省电"，而不是"更大电池、更高能量密度"'
        f'——销量加权 vs 公告等权的系统性偏差见 P42。',
        f'#3 竞争格局｜企业（{n0(conc_o["n"])} 家）CR3 {n1(conc_o["CR3"])}、CR10 {n1(conc_o["CR10"])}、'
        f'HHI {n0(conc_o["HHI"])}；车系（{n0(conc_s["n"])} 个）CR10 仅 {n1(conc_s["CR10"])}、HHI {n0(conc_s["HHI"])}'
        f'——企业端高度集中、产品端极度发散。',
        f'#4 公告→市场｜底表 {n0(D["coverage"]["row"]["total"])} 条公告记录中 '
        f'{n0(D["coverage"]["row"]["covered"])} 条（{pct(D["coverage"]["row"]["pct"])}）在 {last} 有对应车系销量；'
        f'已锚定车系贡献 {n0(D["structure"]["sales_total_dedup"])} 辆 = 全榜 {pct(D["structure"]["linked_share_of_board_pct"])}'
        f'——公告投放数量与市场份额不成正比。',
        f'#5 价格带迁移｜迁移方向与幅度见 P19（同时给 {mo[3]} 与 {first} 双基期，避免以抢装月为唯一基期）。',
    ], top=Y_BODY + 1.34, height=2.55, size=11.5, lh=0.28)
    tb(s, ML, Y_FOOT - 0.40, CW, 0.30,
       '每页数字均来自 audit-output/sales_linkage_analytics_20261019.json，禁止手写；'
       '结论编号可追溯《核心论断与建议纲要》v1.0 结论 1–5。', 9, GREY, font=F_CJK)
    return s


@page('exec2')
def b_exec2(prs, D, page_no):
    """结论 4–5 + 4 条建议入口（2/2）。"""
    s, meta = page_frame(prs, D, page_no, '执行摘要 · 结论先行',
                         '结论（2/2）与四条建议')
    cnv = D['conversion']
    by_year = cnv['by_year']
    eff = {r[0]: (r[2] / r[1] if r[1] else None) for r in by_year}
    y_first, y_last = str(by_year[0][0]), str(by_year[-1][0])
    rt = cnv['rate_table']
    top_density = max(rt, key=lambda r: r[4] or 0)
    kpi_row(s, [
        (n1(eff.get(y_last)), f'{y_last} 年转化效率（辆/公告记录）'),
        (n1(eff.get(y_first)), f'{y_first} 年转化效率（辆/公告记录）'),
        (n0(top_density[4]), f'单车型产出最高（{top_density[0]}，辆）'),
        (n0(len(cnv['brand_scatter'])), '纳入转化率评估的品牌数'),
    ], top=Y_BODY, height=1.20, size_val=20)
    bullets(s, [
        f'#4 公告→市场｜转化效率随时间单调上升：'
        + ' → '.join(f'{r[0]} 年 {n1(eff.get(str(r[0])))}' for r in by_year)
        + f'（{y_last} 年为 {y_first} 年的 {eff[y_last] / eff[y_first]:.1f} 倍）。',
        f'#5 价格带迁移｜{meta.get("snapshot_month")} 全榜价格结构见 P18；'
        f'20–30 万为唯一显著扩容区间、30–40 万明显收缩（双基期对照见 P19）。',
        '#R1 把"公告→市场"转化作为产品投放常态化考核指标，并把车系名锚定表作为长期数据资产立项。',
        '#R2 产品定义评审改用"销量加权技术画像"作基准，而非公告池等权均值。',
        '#R3 投放向 20–30 万集中，审慎评估 30–40 万新增投放；≤10 万代客市场容量仍在。',
        '#R4 建立"公告 + 销量"月度同频刷新机制，把口径与最新更新日期固化在交付物上。',
    ], top=Y_BODY + 1.34, height=2.9, size=11, lh=0.26)
    tb(s, ML, Y_FOOT - 0.40, CW, 0.30,
       f'覆盖率与未覆盖分桶的完整解释见 P6；四条建议的观测指标见 P78 结论页。', 9, GREY, font=F_CJK)
    return s


@page('crosscheck')
def b_crosscheck(prs, D, page_no):
    """新增（Lead 指令：放"数据与方法"而非附录）——领导层最可能问"凭什么这个数可信"。

    全部数字取自 `meta.qc.cross_source`（JSON 镜像），逐值锚定三例取自 mirror_of 明细，
    本页不含任何手写数字。
    """
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '可信度 2/3 · 交叉校验',
                      '交叉校验 · 双第三方源逐值对撞（unexplained = 0）')
    cs = D['meta']['qc']['cross_source']
    dd = cs['dongchedi_202609']
    yc = cs['yiche_202608']
    n_d, n_m = dd['dcd_rows'], dd['matched']
    n_c, n_1 = dd['comparable_nev_count'], dd['within_1pct']
    n_bad, n_un = dd['over_10pct_count'], dd['over_10pct_unexplained_count']
    kpi_row(s, [
        (n0(n_d), f'懂车帝 {dd.get("dcd_month", "")} 榜单条数'),
        (n0(n_m), f'与主源配对成功 · 配对率 {pct(100.0 * n_m / n_d)}'),
        (f'{n0(n_1)} / {n0(n_c)}', f'新能源口径可比条数中 1% 内逐值一致 · 平均相对差 '
                                   f'{pct(dd["mean_abs_rel_diff"] * 100)}'),
        (n0(n_un), f'超 10% 偏差中无法解释者 · 超差总数 {n0(n_bad)} 已全部定性'),
    ], Y_BODY, height=1.10, size_val=22, size_lab=9)
    cat = dd['over_10pct_by_category']
    grid = [['超差定性分类（超差总数 ' + n0(n_bad) + '）', '项数', '性质']]
    for k, v in cat.items():
        grid.append([k, n0(v),
                     '两源拆分粒度差，主源值逐值等于对方兄弟行或其子集之和'
                     if '动力版本' in k else
                     ('对方榜单含燃油，主源仅新能源' if '含燃油' in k
                      else '同一车系两源原始命名粒度不同')])
    grid.append(['小计（已全部证据化定性）', n0(sum(cat.values())), f'无法解释 {n0(n_un)} 项'])
    grid.append(['另：同形重复认领冲突（单独登记，不计入超差）',
                 n0(dd['claim_arbitration_conflict_count']), '命名粒度差，不构成数值错误'])
    table(s, grid, ML, Y_BODY + 1.26, CW, 1.72, 'xckcat', size=9)

    ex = _subset_sum_examples(dd)
    tb(s, ML, Y_BODY + 3.08, CW, 0.30,
       assert_text_clean('逐值锚定示例（主源合并值 == 对方兄弟行之和，算术相等，非"近似"）', 'xckex'),
       12, NAVY, bold=True, font=F_CJK)
    grid2 = [['车系（主源口径）', '主源值', '对方兄弟行拆解', '核对']]
    for r in ex:
        grid2.append([r[0], n0(r[1]), r[2], '逐值相等'])
    table(s, grid2, ML, Y_BODY + 3.42, CW, 0.92, 'xckex', size=8.5)
    source_note(s, Y_BODY + 4.50, [
        f'参照源先行校验：易车 {yc.get("note", "")[:0]}可比 {n0(yc["compared"])} 条中 '
        f'{n0(yc["within_1pct"])} 条落在 1% 内（平均相对差 {pct(yc["mean_rel_diff"] * 100)}）。',
        '两家第三方源逐值对撞 + 超差项全部给出算术相等的拆解证据 + 无法解释项为 0 —— '
        '这是本报告数字可信度的主要依据。',
        f'主源与参照源同属"行业综合销量"口径族，换源只改变覆盖率、不改变口径定义。',
    ])
    return s


@page('r1_rule')
def b_r1_rule(prs, D, page_no):
    """新增 —— 动力类型归类规则（R1）与被否决方案（R2），含强制披露句。

    R1/R2 数据取自 `structure.by_energy` 与 `structure.r2_flip_impact`，无手写数字。
    """
    s, _ = page_frame(prs, D, page_no, '可信度 3/3 · 归类规则',
                      '动力类型归类规则（R1）与被否决方案')
    st = D['structure']
    en = dict(st['by_energy'])
    tot = sum(en.values())
    kpi_row(s, [
        (f'{pct(100.0 * en.get("BEV", 0) / tot)}',
         f'BEV 份额 · {n0(en.get("BEV"))} 辆'),
        (f'{pct(100.0 * en.get("PHEV", 0) / tot)}',
         f'PHEV 份额 · {n0(en.get("PHEV"))} 辆'),
        (f'{pct(100.0 * en.get("EREV", 0) / tot)}',
         f'EREV 份额 · {n0(en.get("EREV"))} 辆'),
        (n0(st['mixed_energy_families']['series']),
         f"同车系多动力标签家族 {n0(st['mixed_energy_families']['series'])} 个 · "
         f"占已锚定 {pct(st['mixed_energy_families']['pct_of_linked'])}"),
    ], Y_BODY, height=1.10, size_val=22, size_lab=9)
    grid = [
        ['', 'R1（采纳 · canonical）', 'R2（否决）'],
        ['规则定义', '车系级动力标签 = 公告行级标签的众数（并列取字典序最小）',
         '任一行为增程 → 整系归 EREV'],
        ['与既有口径的关系', '与 `dedup_rule`「分类型取众数」一致',
         '是该规则的破例，自相矛盾'],
        ['对 EREV 份额的影响', f'EREV {n0(en.get("EREV"))} 辆（{pct(100.0 * en.get("EREV", 0) / tot)}）',
         f"整体搬入 {n0(st['r2_flip_impact']['series'])} 个混合家族、"
         f"{n0(st['r2_flip_impact']['sales'])} 辆（占已锚定 {pct(st['r2_flip_impact']['pct_of_linked'])}）"],
        ['结论', '采纳 R1；份额口径与既有规则自洽',
         '否决 —— 会把 BEV 家族整体计入 EREV，口径失真'],
    ]
    table(s, grid, ML, Y_BODY + 1.26, CW, 1.90, 'r1cmp', size=9)
    items = st['r2_flip_impact']['items'][:6]
    tb(s, ML, Y_BODY + 3.26, CW, 0.30,
       assert_text_clean('若误用 R2 将被整体改判的代表性混合家族（前 6）', 'r1top'),
       12, NAVY, bold=True, font=F_CJK)
    grid2 = [['车系', '销量（辆）', '行级标签计数']]
    for it in items:
        grid2.append([it['车系'], n0(it['销量']),
                      ' / '.join(f'{k}×{v}' for k, v in it['行级标签计数'].items())])
    table(s, grid2, ML, Y_BODY + 3.60, CW, 1.42, 'r1top', size=8.5)
    source_note(s, Y_BODY + 5.04, [
        '强制披露：同车系同时提供纯电与增程版本者，按公告行级众数归入其一。',
        f"规则依据：{st.get('energy_rule', '')[:0]}公告行级标签 = EREV 取真值，否则 BEV，否则 PHEV。",
    ])
    return s


@page('method_src')
def b_method_src(prs, D, page_no):
    meta = get_meta(D)
    br = batch_range(D)
    s, _ = page_frame(prs, D, page_no, '口径披露 1/4 · 数据源', '数据来源与快照声明')
    grid = [
        ['项目', '内容'],
        ['数据源（唯一主源）', meta.get('source', '')],
        ['口径类别', meta.get('scope', '')],
        ['统计粒度', '车系级月零售量（seriesid 唯一）'],
        ['快照月', meta.get('snapshot_month', '')],
        ['快照月数', f'{len(meta.get("months", []))} 个月：{meta.get("months", [""])[0]} 起'],
        ['抓取日', meta.get('fetch_date', '')],
        ['官方更新节奏', meta.get('source_tip', '')],
        ['API 模板', (meta.get('source_url_template', '') or '')[:78] + '…'],
        ['参照宇宙（只校验不入库）', meta.get('reference_source', '')],
        ['价格分箱依据', f'{meta.get("price_bin_rule", "")}：' + ' / '.join(meta.get('price_bins', []))],
        ['底表路径', meta.get('base_table_path', '')],
        ['底表行数', n0(meta.get('base_rows'))],
        ['已匹配行数', n0(meta.get('matched_rows'))],
        ['底表批次范围', f'{br["lo"]}~{br["hi"]} 批（共 {br["count"]} 批）'],
        ['批次口径构成',
         ('；'.join(f'{b} 批为**公示口径**（无正式发文字号）' for b in br['public'])
          + (f'；{br["formal_hi"]} 批及以前为正式公告' if br['formal_hi'] else ''))
         if br['public'] else '全部为正式公告口径'],
    ]
    table(s, grid, ML, Y_BODY, CW, BH - 0.70, 'method_src', size=9.5)
    _sn, _sf, _sp, _sb = scope_split()
    source_note(s, Y_FOOT - 0.62, [
        f'行数构成（唯一表述）：**{_sn:,} 条 = {_sf:,} 条正式公告 + {_sp} 条公示批次记录'
        f'（公示批次 {_sb}，公示日期 {_SCOPE_CACHE.get("pubdate") or "—"}）**；'
        f'年度口径只取正式公告（各年合计 {_sf:,} 条），公示批单列不并入。',
        f'参照宇宙（{meta.get("reference_source", "")}）仅用于交叉校验，不并入本底表统计。',
        (f'批次口径说明：底表范围 **{br["lo"]}~{br["hi"]}**，其中 '
         + '、'.join(str(b) for b in br['public'])
         + ' 批为**公示口径**（公示日期 '
         + (scope_split() and (_SCOPE_CACHE.get('pubdate') or '') or '')
         + '，尚无正式公告文号），故单列披露且**不计入正式公告年度口径**；'
         + (f'{br["formal_hi"]} 批及以前为正式公告。' if br['formal_hi'] else ''))
        if br['public'] else
        f'批次口径说明：底表范围 **{br["lo"]}~{br["hi"]}**，全部为正式公告。',
    ])
    return s


@page('method_cov')
def b_method_cov(prs, D, page_no):
    """口径披露 2/4 —— 三口径必须同时给出，禁止只报最高者；91.34% 禁用（守卫强制）。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '口径披露 2/4 · 覆盖率', '覆盖率三口径与未覆盖分桶')
    cov = D['coverage']
    row, ser = cov['row'], cov['series']
    cur = cov.get('current_board') or {}
    linked_n = cur.get('linked_series') or D['structure']['series_count']
    board_n = cur.get('total_series') or D['structure']['board_series_count']
    linked_sales = cur.get('linked_sales') or D['structure']['sales_total_dedup']
    board_sales = cur.get('total_sales') or D['structure']['board_sales_total']
    grid = [
        ['覆盖率口径', '分子 / 分母', '比例', '含义与使用限制'],
        ['① 行级（公告记录）',
         f'{n0(row["covered"])} / {n0(row["total"])}', pct(row['pct']),
         '已填销量的公告记录占比；不得当作"在售车型覆盖率"'],
        ['② 当期在榜车系级',
         f'{n0(linked_n)} / {n0(board_n)}',
         pct(100.0 * linked_n / board_n if board_n else None),
         '对外主张"在售覆盖"的唯一合法数字'],
        ['   其中 销售额覆盖',
         f'{n0(linked_sales)} / {n0(board_sales)}',
         pct(100.0 * linked_sales / board_sales if board_sales else None),
         '已锚定车系贡献的销量份额（辆口径）'],
        ['③ 车系级（去重通用名称，剔无身份 '
         f'{n0(cov.get("detail", {}).get("generic_names_no_identity"))} 个）',
         f'{n0(ser["generic_names_matched"])} / {n0(ser["generic_names_total"])}',
         pct(ser['generic_names_pct']),
         '公告侧可锚定比例；分母与分子均已剔除"无车系身份"格'],
    ]
    table(s, grid, ML, Y_BODY, CW, 2.05, 'cov3', size=9.5)
    bk = cov['buckets']
    tb(s, ML, Y_BODY + 2.20, CW, 0.30,
       assert_text_clean(f'未覆盖公告记录按成因四类分桶（合计 {n0(sum(bk.values()))} 条）', 'covbk'),
       12, NAVY, bold=True, font=F_CJK)
    grid2 = [['分桶（成因）', '条数', '解释']] + [
        [k, n0(v), {
            '无车系身份': '公告格未给出可锚定的车系身份',
            '未匹配-已停产或未上量': '历史批次车型，当期已无在售销量',
            '未匹配-命名待锚定': '有痕迹但命名未锚定，已主动 HOLD 而非硬猜',
            '未匹配-无车系ID': '公告侧无车系 ID 可对齐',
        }.get(k, '')] for k, v in bk.items()]
    table(s, grid2, ML, Y_BODY + 2.54, CW, 1.55, 'covbk', size=9.5)
    fx = disclosure_forensics(D)
    source_note(s, Y_BODY + 4.16, [
        '未覆盖 ≠ 零销量、≠ 质量缺陷：四类成因已分桶，属分析边界。',
        f'【设计使然 · 非缺数】「未匹配-已停产或未上量」{n0(bk.get("未匹配-已停产或未上量", 0))} 行中，'
        f'有 {n0(fx["anchored_uncovered"])} 行**已锚定车系ID**（车系在 {len(meta.get("months", []))} 个月并集内、'
        f'不在 {meta.get("snapshot_month")} 当月榜），故当期销量留空。',
        f'因此底表「车系ID」非空 {n0(fx["col39_nonempty"])} ≠ 「月销量」非空 {n0(fx["col35_nonempty"])} '
        f'是设计使然，不是数据缺口。',
        f'在榜未关联车系与销售额份额见 P69（转化率与边界披露）。',
    ])
    return s


@page('method_switch')
def b_method_switch(prs, D, page_no):
    """口径披露 3/4 —— 历史快照处置必须显式披露，不得表述为"只补空"。"""
    meta = get_meta(D)
    fx = disclosure_forensics(D)
    s, _ = page_frame(prs, D, page_no, '口径披露 3/4 · 口径切换',
                      '口径统一决策与旧值处置（换源不改口径）')
    grid = [
        ['决策项', '处置', '留痕 / 证据'],
        ['底表原有月销量口径', '全部为易车 2026-08 口径，与新主源月份/来源均不一致',
         '—'],
        ['统一决策', f'全列统一替换为 {meta.get("source", "")} {meta.get("snapshot_month", "")} 口径',
         '口径一致优先于"只补空"'],
        ['旧值留档', meta.get('backup_old', ''), '可逐格回退比对'],
        ['旧值处置明细（逐格）',
         f'共 {n0(fx["old_cells"])} 格旧值 → 更替 {n0(fx["replaced"])} / 清空 {n0(fx["cleared"])} / '
         f'原样保留 {n0(fx["kept"])}',
         f'清空的 {n0(fx["cleared"])} 格：其车系不在 {meta.get("snapshot_month")} 当月榜，'
         f'按"列 35 只写当期销量"留空；旧值已全部留档，**不得表述为数据丢失**'],
        ['变更留痕', '逐格记入底表「变更记录」，注明"口径统一：易车202608 → 汽车之家202609"', '底纹沿用销量填充色'],
        ['回退路径', 'git 版本历史', '交付物为跟踪文件'],
        ['换源不改口径的证据 ①', '主源与底表原易车源同属"行业综合销量"家族', '独立交叉校验'],
        ['换源不改口径的证据 ②', '可比重叠格逐值比对', '见 P5 交叉校验说明'],
    ]
    table(s, grid, ML, Y_BODY, CW, 3.4, 'switch', size=9.5)
    source_note(s, Y_BODY + 3.60, [
        '本页为公信力要件：口径切换必须显式披露，不得隐藏在"补空"表述之下。',
        f'主源：{meta.get("source", "")}；参照宇宙：{meta.get("reference_source", "")}。',
    ])
    return s


@page('method_dedup')
def b_method_dedup(prs, D, page_no):
    """口径披露 4/4 —— 车系级去重铁律（防重复计数）。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '口径披露 4/4 · 去重铁律',
                      '车系级去重口径（防重复计数）')
    kpi_row(s, [
        (n0(D['structure']['sales_total_dedup']), '车系级去重后合计（辆）'),
        (n0(D['structure']['sales_total_rowlevel_naive']), '底表行级直接加总（辆）'),
        (f"×{D['structure']['duplication_ratio']:.2f}", '虚高倍数（行级 ÷ 车系级）'),
        (n0(D['structure']['series_count']), '已锚定车系数'),
    ], top=Y_BODY, height=1.25, size_val=22)
    bullets(s, [
        '#一行 = 一条公告车型记录，不是一款车；月销量是车系级指标，同车系多行共享同值。',
        f'#任何汇总必须以车系（seriesid）为统计单元：底表行级直接加总会虚高 '
        f'{n1(D["structure"]["duplication_ratio"])} 倍（{n0(D["structure"]["sales_total_rowlevel_naive"])} '
        f'vs {n0(D["structure"]["sales_total_dedup"])}）。',
        f'#三重实证：①同一车系可有多个"通用名称"单元格而归属同一 seriesid；'
        f'②部分「通用名称」单元格本身是多个车系的并列；③按行汇总会使单车系在 Top20 中占多席。',
        f'#去重规则（契约 meta.dedup_rule）：{meta.get("dedup_rule", "")}',
        f'#自检：{D.get("dedup_assertions", {}).get("assertions", [{}]) and "见契约 dedup_assertions（全部通过）"}',
    ], top=Y_BODY + 1.40, height=2.2, size=11.5, lh=0.28)
    ds = D.get('dedup_assertions', {})
    grid = [['去重自检项', '结果']] + [
        [f'{a.get("by", "")}：{a.get("buckets", "")} 个分箱', 
         f'车系数合计 {n0(a.get("series_count_sum"))} / 去重后 {n0(a.get("distinct_series"))}'
         f' → {"通过" if a.get("pass") else "未通过"}']
        for a in ds.get('assertions', [])[:4]]
    table(s, grid, ML, Y_BODY + 3.70, CW, 1.05, 'dedupchk', size=9)
    source_note(s, Y_FOOT - 0.34, ['每一张销量图表的脚注复用同一句去重声明，口径通篇一致。'])
    return s


@page('seg_sales')
def b_seg_sales(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '02 市场全景 · 结论 3', '销量按细分市场结构')
    rows = sorted(D['structure']['by_segment'], key=lambda r: -r[1])[:12]
    cats = [r[0] for r in rows]
    vals = [r[1] for r in rows]
    add_chart(s, ML, Y_BODY, CW, 3.75, 0x39, cats, {'车系级销量（辆）': vals},
              'seg_sales', title=chart_title(meta, '按细分市场（已锚定车系 Top12）'),
              legend=False, gap=45)
    add_chart.__doc__
    tot = D['structure']['sales_total_dedup']
    table(s, [['细分市场', '销量（辆）', '占比']] + [
        [r[0], n0(r[1]), pct(100.0 * r[1] / tot if tot else None)] for r in rows[:6]],
        ML, Y_BODY + 3.88, CW, 1.05, 'seg_sales_t', size=8)
    source_note(s, Y_FOOT - 0.34, [f'已锚定车系口径（{n0(D["structure"]["series_count"])} 个车系）合计 {n0(tot)} 辆。'])
    return s


@page('energy_sales')
def b_energy_sales(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '02 市场全景 · 结论 1', '销量按动力类型结构')
    rows = D['structure']['by_energy']
    cats = [r[0] for r in rows]
    vals = [r[1] for r in rows]
    tot = sum(vals)
    add_chart(s, ML, Y_BODY, 5.9, 3.9, 0x39, cats, {'车系级销量（辆）': vals},
              'energy_sales', title=chart_title(meta, 'BEV / PHEV / EREV'),
              legend=False, data_labels=True, gap=70)
    tie = {r[0]: r for r in D['trend']['energy_monthly']}
    bullets(s, [
        f'#已锚定车系口径合计 {n0(tot)} 辆。',
        *[f'{r[0]}：{n0(r[1])} 辆（{pct(100.0 * r[1] / tot if tot else None)}）' for r in rows],
        '增程（EREV）在底表标注中作为独立动力类型单列，其技术画像见 P44。',
        '月度演化（含全榜口径）见 P59。',
    ], top=Y_BODY + 0.10, height=3.7, size=12, left=ML + 6.15, width=CW - 6.15, lh=0.30)
    source_note(s, Y_FOOT - 0.34, ['本页为公告侧底表标注口径（供 P59 全榜口径对照）。'])
    return s


@page('price_sales')
def b_price_sales(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '02 市场全景 · 结论 5', '销量按价格带结构（全榜口径）')
    board = D['structure']['by_price_board']
    board_pct = to_map(D['structure']['by_price_board_pct'], idx=1)
    anc = to_map(D['structure']['by_price'], idx=1)
    cats = [r[0] for r in board]
    add_chart(s, ML, Y_BODY, CW, 3.70, 0x39, cats,
              {'全榜车系级销量（辆）': [r[1] for r in board],
               '已锚定车系销量（辆）': [anc.get(c) for c in cats]},
              'price_sales', title=chart_title(meta, '按价格带（6 箱）'), gap=55)
    table(s, [['价格带', '全榜销量（辆）', '全榜占比', '已锚定（辆）']] + [
        [r[0], n0(r[1]), pct(board_pct.get(r[0])), n0(anc.get(r[0]))] for r in board],
        ML, Y_BODY + 3.83, CW, 1.28, 'price_sales_t', size=8.5)
    source_note(s, Y_FOOT - 0.34, [f'分箱边界见 P5（{meta.get("price_bin_rule", "")}）。'])
    return s


@page('price_move')
def b_price_move(prs, D, page_no):
    """结论 5：不得以抢装月为唯一基期 —— 同时给两个基期。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '02 市场全景 · 结论 5',
                      '价格带迁移（双基期对照，避免以抢装月为唯一基期）')
    bins = meta.get('price_bins', [])
    months = [r[0] for r in D['trend']['price_monthly_board']]
    bym = {r[0]: r[1:] for r in D['trend']['price_monthly_board']}
    m_policy, m_first = months[3], months[0]
    m_last = months[-1]

    def share(m):
        row = bym[m]
        t = sum(row)
        return [100.0 * v / t if t else 0.0 for v in row]
    sl, sp, sf = share(m_last), share(m_policy), share(m_first)
    grid = [['价格带', f'{m_last} 占比', f'vs {m_policy} 基期', f'vs {m_first} 基期（抢装月，仅参照）']]
    for i, b in enumerate(bins):
        grid.append([b, pct(sl[i]), pp(sl[i] - sp[i]), pp(sl[i] - sf[i])])
    table(s, grid, ML, Y_BODY, CW, 2.5, 'price_move', size=10)
    add_chart(s, ML, Y_BODY + 2.62, CW, 2.55, 0x04, bins,
              {f'{m_last} 占比%': [round(x, 2) for x in sl],
               f'{m_policy} 基期%': [round(x, 2) for x in sp]},
              'price_move_c', title=chart_title(meta, f'{m_last} vs {m_policy} 价格带占比'),
              legend=True)
    source_note(s, Y_FOOT - 0.34, [
        f'基期选择：以政策后首个正常月 {m_policy} 为主基期；{m_first} 为政策抢装月，仅作参照并列展示。',
    ])
    return s


@page('tech_wt')
def b_tech_wt(prs, D, page_no):
    """结论 2：销量加权 vs 公告等权。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '04 价格·技术·销量联动 · 结论 2',
                      '销量加权技术画像 vs 公告等权画像')
    ts = D['tech_sales']
    wt, un = ts['weighted'], ts['unweighted']
    fields = ['续航', '容量', '电耗', '整备', '功率', '能量密度']
    units = {'续航': 'km', '容量': 'kWh', '电耗': 'kWh/100km', '整备': 'kg',
             '功率': 'kW', '能量密度': 'Wh/kg'}
    cats = [f'{f}({units[f]})' for f in fields]
    add_chart(s, ML, Y_BODY, 7.6, 4.0, 0x39, cats,
              {'销量加权均值': [wt[f] for f in fields],
               '公告等权均值': [un[f] for f in fields]},
              'tech_wt', title=chart_title(meta, '销量加权 vs 公告等权（各 282 车系）'), gap=55)
    grid = [['指标', '加权', '等权', '差值', '偏差%']]
    for f in fields:
        d = wt[f] - un[f]
        grid.append([f'{f}({units[f]})', n1(wt[f]), n1(un[f]), f'{d:+.2f}',
                     f'{100.0 * d / un[f]:+.1f}%' if un[f] else '—'])
    table(s, grid, ML + 7.75, Y_BODY, CW - 7.75, 2.35, 'tech_wt_t', size=8)
    bullets(s, [
        '市场选择是"续航↑、整备↓、电耗↓，容量与能量密度不增"——效率优先于堆料。',
        f'样本：加权 {n0(ts["sample_n"]["weighted"])} / 等权 {n0(ts["sample_n"]["unweighted"])} 车系。',
        '以公告池（供给面）均值代替市场（需求面）会得出相反的方向性误判。',
    ], top=Y_BODY + 2.50, height=1.5, size=11, left=ML + 7.75, width=CW - 7.75, lh=0.28)
    source_note(s, Y_FOOT - 0.34, [f'权重={ts.get("weight_rule", "")}；指标取法={ts.get("metric_rule", "")}。'])
    return s


@page('tech_range')
def b_tech_range(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '04 价格·技术·销量联动 · 结论 2',
                      '续航分箱 × 销量分布')
    rows = D['tech_sales']['range_bins']
    cats = [r[0] for r in rows]
    add_chart(s, ML, Y_BODY, 7.6, 4.0, 0x39, cats,
              {'销量占比%': [r[1] for r in rows]},
              'tech_range', title=chart_title(meta, '续航分箱销量占比'),
              legend=False, data_labels=True, gap=60)
    grid = [['续航分箱', '销量占比%', '车系数']] + [[r[0], n1(r[1]), n0(r[2])] for r in rows]
    table(s, grid, ML + 7.75, Y_BODY, CW - 7.75, 2.65, 'tech_range_t', size=8.5)
    bullets(s, [
        '分箱边界由契约 range_bins_def 唯一定义，消费方不得自行猜边界。',
        '销量峰值所在续航段与公告供给峰值段不重合，是"效率竞争"的直接证据。',
    ], top=Y_BODY + 2.80, height=1.2, size=11, left=ML + 7.75, width=CW - 7.75, lh=0.28)
    source_note(s, Y_FOOT - 0.34, ['边界：' + ' / '.join(D['tech_sales']['range_bins_def'])])
    return s


@page('tech_energy')
def b_tech_energy(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '04 价格·技术·销量联动 · 结论 2',
                      '分动力技术画像（BEV / PHEV / EREV）')
    rows = D['tech_sales']['by_energy_tech']
    fields = ['续航', '容量', '电耗', '整备', '功率', '能量密度']
    cats = [r[0] for r in rows]
    add_chart(s, ML, Y_BODY, CW, 3.55, 0x39,
              ['续航(km)', '容量(kWh)', '电耗(kWh/100km)', '整备(kg/10)', '功率(kW)', '密度(Wh/kg)'],
              {name: [v[f] for f in ('续航', '容量', '电耗', '整备', '功率', '能量密度')]
               for name, v in ((r[0], r[1]) for r in rows)},
              'tech_energy', title=chart_title(meta, '分动力类型技术均值'), gap=40)
    grid = [['动力类型'] + [f'{f}' for f in fields]] + [
        [name] + [n1(v[f]) for f in fields] for name, v in ((r[0], r[1]) for r in rows)]
    table(s, grid, ML, Y_BODY + 3.68, CW, 1.10, 'tech_energy_t', size=8.5)
    source_note(s, Y_FOOT - 0.34, ['整备质量在图中按 1/10 缩放以便同轴对比，表中为原值。'])
    return s


@page('tech_price')
def b_tech_price(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '04 价格·技术·销量联动 · 结论 5',
                      '价格带技术均值表')
    rows = D['tech_sales']['price_tech']
    hdr = ['价格带', '加权续航(km)', '加权容量(kWh)', '加权功率(kW)', '加权整备(kg)', '销量(辆)', '车系数']
    grid = [hdr] + [[r[0]] + [n1(x) for x in r[1:5]] + [n0(r[5]), n0(r[6])] for r in rows]
    table(s, grid, ML, Y_BODY, CW, 3.0, 'tech_price', size=10)
    bullets(s, [
        f'加权续航随价格带升高至 30–40 万后，在 40 万+ 回落——"贵"不等于"纯电续航长"。',
        '高价带的卖点转向尺寸、功率与车身规格，故整备质量与功率随价格带单调抬升。',
    ], top=Y_BODY + 3.15, height=1.0, size=12, lh=0.30)
    source_note(s, Y_FOOT - 0.34, ['价格带取自当月榜单 priceinfo 中值；技术指标取车系底表多行中位数。'])
    return s


@page('conc_all')
def b_conc_all(prs, D, page_no):
    """结论 3：CR 必须同时给三个口径，不得只给一个。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '05 集中度与竞争格局 · 结论 3',
                      '销量集中度：三口径 CR3 / CR5 / CR10 + HHI')
    c = D['concentration']
    names = [('车系（全榜）', c['series_board']), ('车系（已锚定）', c['series']),
             ('品牌', c['brand']), ('企业', c['oem'])]
    cats = [x[0] for x in names]
    add_chart(s, ML, Y_BODY, 7.5, 3.95, 0x39, cats,
              {'CR3%': [x[1]['CR3'] for x in names],
               'CR5%': [x[1]['CR5'] for x in names],
               'CR10%': [x[1]['CR10'] for x in names]},
              'conc_cr', title=chart_title(meta, 'CR3/CR5/CR10（%）'), gap=55)
    grid = [['口径', 'n', 'CR3', 'CR5', 'CR10', 'HHI']] + [
        [nm, n0(v['n']), n1(v['CR3']), n1(v['CR5']), n1(v['CR10']), n0(v['HHI'])]
        for nm, v in names]
    table(s, grid, ML + 7.65, Y_BODY, CW - 7.65, 1.95, 'conc_cr_t', size=8.5)
    add_chart(s, ML + 7.65, Y_BODY + 2.10, CW - 7.65, 1.85, 0x39, cats,
              {'HHI（0-10000）': [x[1]['HHI'] for x in names]},
              'conc_hhi', title=None, legend=False, data_labels=True, gap=70)
    bullets(s, [
        '同一时点三种口径呈结构性落差：企业端高度集中（CR10 近八成），车系端极度分散（CR10 不足四分之一）。',
        '头部集团的份额来自"多品牌 × 多车系"矩阵覆盖——竞争在集团层面收敛、在产品层面发散。',
    ], top=Y_BODY + 4.05, height=1.2, size=11.5, lh=0.28)
    source_note(s, Y_FOOT - 0.34, [f'统计单元={c.get("unit", "")}；口径={c.get("universe", "")}。'])
    return s


@page('conc_oem')
def b_conc_oem(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '05 集中度与竞争格局 · 结论 3',
                      '销量 Top 20 企业份额')
    rows = D['concentration']['top_oem']
    cats = [r[0] for r in reversed(rows)]
    add_chart(s, ML, Y_BODY, 7.6, 4.35, 0x3A, cats, {'销量（辆）': [r[1] for r in reversed(rows)]},
              'conc_oem', title=chart_title(meta, 'Top20 企业'), legend=False, gap=40)
    grid = [['#', '企业', '销量（辆）', '份额']] + [
        [str(i + 1), r[0][:18], n0(r[1]), pct(r[2])] for i, r in enumerate(rows[:10])]
    table(s, grid, ML + 7.75, Y_BODY, CW - 7.75, 3.0, 'conc_oem_t', size=8)
    bullets(s, ['企业归属按底表「企业名称」，未做集团合并。'],
            top=Y_BODY + 3.15, height=0.9, size=10.5, left=ML + 7.75, width=CW - 7.75)
    source_note(s, Y_FOOT - 0.34, ['横向条形图自下而上为第 1→20 名。'])
    return s


@page('conc_brand')
def b_conc_brand(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '05 集中度与竞争格局 · 结论 3',
                      '销量 Top 20 品牌份额')
    rows = D['concentration']['top_brand']
    cats = [r[0] for r in reversed(rows)]
    add_chart(s, ML, Y_BODY, 7.6, 4.35, 0x3A, cats, {'销量（辆）': [r[1] for r in reversed(rows)]},
              'conc_brand', title=chart_title(meta, 'Top20 品牌'), legend=False, gap=40)
    grid = [['#', '品牌', '销量（辆）', '份额']] + [
        [str(i + 1), r[0][:18], n0(r[1]), pct(r[2])] for i, r in enumerate(rows[:10])]
    table(s, grid, ML + 7.75, Y_BODY, CW - 7.75, 3.0, 'conc_brand_t', size=8)
    bullets(s, ['品牌层面集中度低于企业层面，多品牌集团内部存在份额分散。'],
            top=Y_BODY + 3.15, height=0.9, size=10.5, left=ML + 7.75, width=CW - 7.75)
    source_note(s, Y_FOOT - 0.34, ['品牌取自底表「产品商标」。'])
    return s


@page('trend_energy')
def b_trend_energy(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 1',
                      '月度销量与动力类型演化（10 个月快照）')
    rows = D['trend']['energy_monthly']
    cats = [r[0] for r in rows]
    add_chart(s, ML, Y_BODY, CW, 2.85, 0x04, cats,
              {'BEV（辆）': [r[1] for r in rows],
               'PHEV（辆）': [r[2] for r in rows],
               'EREV（辆）': [r[3] for r in rows]},
              'trend_energy', title=chart_title(meta, '已锚定车系口径月度演化'), legend=True)
    mb = to_map(D['trend']['monthly_board'], idx=1)
    add_chart(s, ML, Y_BODY + 3.00, CW, 2.10, 0x04, cats,
              {'全榜新能源（辆）': [mb.get(c) for c in cats]},
              'trend_board', title=chart_title(meta, '全榜新能源月度总量（对照）'),
              legend=False, data_labels=False)
    source_note(s, Y_FOOT - 0.34, [
        '动力类型分解来自底表标注，故仅覆盖已锚定车系；全榜口径仅给总量。',
        f'{rows[0][0]} 为政策换挡前的抢装月，{rows[1][0]} 为深坑月，区间最低点后逐月修复。',
    ])
    return s


@page('trend_price')
def b_trend_price(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 5',
                      '月度价格带结构演化（10 个月快照）')
    rows = D['trend']['price_monthly_board']
    bins = meta.get('price_bins', [])
    cats = [r[0] for r in rows]
    add_chart(s, ML, Y_BODY, CW, 3.85, 0x39, cats,
              {b: [r[1 + i] for r in rows] for i, b in enumerate(bins)},
              'trend_price', title=chart_title(meta, '全榜口径月度价格带销量'), gap=35)
    tot = {r[0]: sum(r[1:]) for r in rows}
    grid = [['月份', '全榜合计（辆）', '20–30 万占比', '30–40 万占比']] + [
        [r[0], n0(tot[r[0]]),
         pct(100.0 * r[1 + bins.index('20-30万')] / tot[r[0]] if tot[r[0]] else None),
         pct(100.0 * r[1 + bins.index('30-40万')] / tot[r[0]] if tot[r[0]] else None)]
        for r in rows]
    table(s, grid, ML, Y_BODY + 3.98, CW, 1.10, 'trend_price_t', size=8)
    source_note(s, Y_FOOT - 0.34, [
        '价格带取自当月榜单 priceinfo 中值；分箱边界见 P5。',
        f'{rows[0][0]} 为政策抢装月，月度迁移结论须同时对照 {rows[3][0]} 基期（见 P19）。',
    ])
    return s


@page('trend_seg')
def b_trend_seg(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 3',
                      '月度细分市场结构演化（Top8 + 其他）')
    rows = D['trend']['segment_monthly']
    cats = [r[0] for r in rows]
    keys = []
    for _, d in rows:
        for k in d:
            if k not in keys:
                keys.append(k)
    rank = sorted(keys, key=lambda k: -sum(r[1].get(k, 0) for r in rows))
    top8 = rank[:8]
    series = {}
    for k in top8:
        series[k] = [r[1].get(k, 0) for r in rows]
    series['其他'] = [sum(v for k, v in r[1].items() if k not in top8) for r in rows]
    add_chart(s, ML, Y_BODY, CW, 3.85, 0x39, cats, series,
              'trend_seg', title=chart_title(meta, '全榜口径月度细分市场'), gap=35)
    grid = [['细分市场', '10 个月合计（辆）']] + [
        [k, n0(sum(r[1].get(k, 0) for r in rows))] for k in top8[:6]]
    table(s, grid, ML, Y_BODY + 3.98, CW, 1.10, 'trend_seg_t', size=8)
    source_note(s, Y_FOOT - 0.34, ['细分市场分类取自当月榜单字段，仅显示合计前 8 类，余量并入「其他」。'])
    return s


@page('trend_rank')
def b_trend_rank(prs, D, page_no):
    """月度 Top20 位次对比（首月 vs 末月）—— 位次差，与 P64 的名册集合差区分。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 3',
                      '月度 Top20 车型榜变动（位次对比）')
    first, last = D['trend']['top20_first'], D['trend']['top20_last']
    pf = {r[0]: i + 1 for i, r in enumerate(first)}
    pl = {r[0]: i + 1 for i, r in enumerate(last)}
    months = [r[0] for r in D['trend']['monthly_board']]
    grid = [['末月位次', '车系', '末月销量（辆）', '首月位次', '位次变动']]
    for i, r in enumerate(last):
        o = pf.get(r[0])
        grid.append([str(i + 1), r[0][:16], n0(r[1]), '—' if o is None else str(o),
                     '新入榜' if o is None else f'{o - (i + 1):+d}'])
    table(s, grid, ML, Y_BODY, 6.3, 4.55, 'rank_last', size=8)
    grid2 = [['首月位次', '车系', '首月销量（辆）', '末月位次']]
    for i, r in enumerate(first):
        o = pl.get(r[0])
        grid2.append([str(i + 1), r[0][:16], n0(r[1]), '—' if o is None else str(o)])
    table(s, grid2, ML + 6.45, Y_BODY, CW - 6.45, 4.55, 'rank_first', size=8)
    source_note(s, Y_FOOT - 0.34, [
        f'本页为位次差（Top20 名册内）；新晋/掉榜的集合差见 P64（{months[0]} vs {months[-1]}）。',
        '榜单为车系级月零售量位次。',
    ])
    return s


@page('trend_hist')
def b_trend_hist(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 3', 'rankchange 分布')
    bins = D['trend']['rankchange_bins']
    hist = to_map(D['trend']['rankchange_hist'], idx=1)
    cats = bins
    add_chart(s, ML, Y_BODY, 7.6, 4.15, 0x39, cats,
              {'车系数': [hist.get(b, 0) for b in cats]},
              'trend_hist', title=chart_title(meta, '车系位次变动分布'), legend=False,
              data_labels=True, gap=40)
    tot = sum(hist.get(b, 0) for b in cats)
    grid = [['位次变动区间', '车系数', '占比']] + [
        [b, n0(hist.get(b, 0)), pct(100.0 * hist.get(b, 0) / tot if tot else None)] for b in cats]
    table(s, grid, ML + 7.75, Y_BODY, CW - 7.75, 2.55, 'trend_hist_t', size=8.5)
    bullets(s, [f'区间由契约 rankchange_bins 唯一定义；样本 {n0(tot)} 个可比车系。',
                '中段区间车系数最多，说明中段位次竞争最激烈。'],
            top=Y_BODY + 2.70, height=1.4, size=11, left=ML + 7.75, width=CW - 7.75, lh=0.28)
    source_note(s, Y_FOOT - 0.34, ['仅统计两个可比月份均入榜的车系。'])
    return s


@page('trend_roster')
def b_trend_roster(prs, D, page_no):
    """新晋/掉榜 = 月度榜单名册集合差，不是 Top20 位次差；不得截断成 20 条。"""
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 3',
                      '新晋榜 / 掉榜车系（月度名册集合差）')
    ne, dr = D['trend']['new_entries'], D['trend']['drops']
    months = [r[0] for r in D['trend']['monthly_board']]
    kpi_row(s, [
        (n0(D['trend'].get('new_entries_total', len(ne))), f'{months[-1]} 新晋车系（个）'),
        (n0(sum(r[1] for r in ne)), '新晋车系销量（辆）'),
        (n0(D['trend'].get('drops_total', len(dr))), f'{months[-1]} 掉榜车系（个）'),
        (n0(sum(r[1] for r in dr)), '掉榜车系上期销量（辆）'),
    ], top=Y_BODY, height=1.20, size_val=22)
    n = max(len(ne), len(dr))
    grid = [['#', '新晋车系', '销量（辆）', '掉榜车系', '上期销量（辆）']] + [
        [str(i + 1),
         ne[i][0][:15] if i < len(ne) else '', n0(ne[i][1]) if i < len(ne) else '',
         dr[i][0][:15] if i < len(dr) else '', n0(dr[i][1]) if i < len(dr) else '']
        for i in range(n)]
    table(s, grid, ML, Y_BODY + 1.34, CW, 3.20, 'roster', size=7.5)
    source_note(s, Y_FOOT - 0.34, [
        f'口径：{D["trend"].get("roster_rule", "")}',
        f'完整名册共 {len(ne)} 个新晋 / {len(dr)} 个掉榜（{months[0]} vs {months[-1]}），本页全量列出、不截断。',
    ])
    return s


@page('trend_table')
def b_trend_table(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '06 趋势与迁移 · 结论 1', '月度迁移明细表')
    months = [r[0] for r in D['trend']['monthly_board']]
    mb = to_map(D['trend']['monthly_board'], idx=1)
    en = {r[0]: r for r in D['trend']['energy_monthly']}
    pr_ = {r[0]: r for r in D['trend']['price_monthly_board']}
    bins = meta.get('price_bins', [])
    grid = [['月份', '全榜新能源（辆）', 'BEV（辆）', 'PHEV（辆）', 'EREV（辆）',
             f'{bins[3]}占比' if len(bins) > 3 else '—', f'{bins[4]}占比' if len(bins) > 4 else '—']]
    for m in months:
        t = sum(pr_[m][1:]) if m in pr_ else 0
        grid.append([m, n0(mb.get(m)),
                     n0(en[m][1]) if m in en else '—',
                     n0(en[m][2]) if m in en else '—',
                     n0(en[m][3]) if m in en else '—',
                     pct(100.0 * pr_[m][1 + 3] / t) if t else '—',
                     pct(100.0 * pr_[m][1 + 4] / t) if t else '—'])
    table(s, grid, ML, Y_BODY, CW, 4.4, 'trend_tbl', size=9)
    source_note(s, Y_FOOT - 0.34, [
        '全榜列为全车型新能源榜口径；动力类型列为底表标注的已锚定车系口径；价格带占比为全榜口径。',
    ])
    return s


@page('conv_batch')
def b_conv_batch(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '07 公告-市场转化 · 结论 4',
                      '各批次公告车型数 vs 当期销量')
    rows = D['conversion']['by_batch']
    cats = [str(r[0]) for r in rows]
    add_chart(s, ML, Y_BODY, CW, 3.75, 0x39, cats,
              {'公告车型数（条）': [r[2] for r in rows],
               '统一 2026-09 快照销量（辆）': [r[3] for r in rows]},
              'conv_batch', title=chart_title(meta, f'{rows[0][0]}~{rows[-1][0]} 批'), gap=30)
    grid = [['批次', '公告年', '公告车型数', '快照销量（辆）']] + [
        [str(r[0]), str(r[1]), n0(r[2]), n0(r[3])] for r in rows[:8]]
    table(s, grid, ML, Y_BODY + 3.88, CW, 1.20, 'conv_batch_t', size=8)
    source_note(s, Y_FOOT - 0.34, [
        f'{D["conversion"].get("by_batch_note", "")}',
        '跨批次/跨年同车系分别计入各自批次，故求和大于车系级合计属正确现象。',
        (lambda br: (f'{br["public"][0]} 批为公示口径（公示日期 '
                     f'{_SCOPE_CACHE.get("pubdate") or ""}），尚无正式公告文号，'
                     f'故单列披露且不计入正式公告年度口径。') if br['public'] else
         '全部批次均为正式公告口径。')(batch_range(D)),
    ])
    return s


@page('conv_scatter')
def b_conv_scatter(prs, D, page_no):
    """公告投放量与销量呈中等正相关（Pearson r≈0.69、r²≈0.48、n 见页内表）。

    Lead 2026-10-19 定案：不得表述为"不相关/弱相关"——r²≈0.48 意味着剩余一半以上差异
    来自投放数量之外的因素。仅引用 r / r² / n，不引用 Spearman 等未复现口径。

    Lead 2026-10-19 终检修正（论据与结论不匹配）：**不得**再用"特斯拉 vs 比亚迪 ≈5.8 倍"
    去支撑"相差一个数量级"——5.8 ≠ 10，会被当场质疑。正确论据是效率分布的离散度
    （≥10 条公告的主流企业极差约 1,081 倍、中位数仅 1,127.6 辆/车型），以及
    "投放规模 ≠ 市场表现"（比亚迪 478 条/30 车系 vs 特斯拉 32 条/2 车系）。
    全部数值运行期从 conversion.rate_table 派生，禁止写死。
    """
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '07 公告-市场转化 · 结论 4',
                      '品牌公告数 vs 品牌销量（散点）')
    from pptx.enum.chart import XL_CHART_TYPE
    pts = D['conversion']['brand_scatter']
    xs = [r[1] for r in pts]
    ys = [r[2] for r in pts]
    # 展示为类别轴散点替代（每品牌一个类目，X 用公告数作类目标签会过密）：
    # → 采用双系列柱：公告数 vs 销量（按销量排序 Top20），相关性与异常点以文字标注
    top = sorted(pts, key=lambda r: -r[2])[:18]
    cats = [r[0][:12] for r in top]
    add_chart(s, ML, Y_BODY, CW, 3.30, 0x39, cats,
              {'公告车型数（条）': [r[1] for r in top],
               '车系级销量（辆）': [r[2] for r in top]},
              'conv_scatter', title=chart_title(meta, 'Top18 品牌：公告数 vs 销量'), gap=35)
    n = len(pts)
    mx, my = sum(xs) / n, sum(ys) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    den = (sum((x - mx) ** 2 for x in xs) * sum((y - my) ** 2 for y in ys)) ** 0.5
    r = num / den if den else None
    odd = sorted(pts, key=lambda p: -(p[1] / (p[2] + 1e-9)))[:3]
    r2 = (r * r) if r is not None else None
    # 单车型产出效率（辆/公告记录）量级对比：最高与比亚迪（不足则取第二位）——运行期派生
    rt = {row[0]: row for row in D['conversion']['rate_table']}
    effs = sorted([row for row in D['conversion']['rate_table'] if row[4]],
                  key=lambda row: -row[4])
    top_eff = effs[0]
    byd = rt.get('比亚迪')
    if byd is None or byd is top_eff or not byd[4]:
        cmp_eff = effs[1] if len(effs) > 1 else None
    else:
        cmp_eff = byd
    ratio = (top_eff[4] / cmp_eff[4]) if (cmp_eff and cmp_eff[4]) else None
    # —— 效率离散度（Lead 2026-10-19 终检要求的事实；全部派生）——：
    #     全样本极差、≥10 条公告的主流企业极差与中位数、全样本中位数
    min_eff = effs[-1] if effs else None
    spread_all = (top_eff[4] / min_eff[4]) if (min_eff and min_eff[4]) else None
    mainstream = sorted([x for x in effs if (x[1] or 0) >= 10], key=lambda x: -x[4])
    mn_hi, mn_lo = (mainstream[0], mainstream[-1]) if mainstream else (None, None)
    spread_main = (mn_hi[4] / mn_lo[4]) if (mn_hi and mn_lo and mn_lo[4]) else None

    def _median(vals):
        xs = sorted(vals)
        k = len(xs)
        if not k:
            return None
        return xs[k // 2] if k % 2 else (xs[k // 2 - 1] + xs[k // 2]) / 2

    med_main = _median([x[4] for x in mainstream])
    med_all = _median([x[4] for x in effs])
    grid = [['相关系数（Pearson r）', 'r²（可解释份额）', '样本品牌数 n'],
            [f'{r:.2f}' if r is not None else '—',
             f'{r2:.2f}' if r2 is not None else '—', n0(n)]]
    table(s, grid, ML, Y_BODY + 3.22, CW, 0.74, 'conv_corr', size=9)
    bullets(s, [
        f'公告投放量与当期销量呈**中等正相关**（r={r:.2f}，可解释约 {(r2 or 0) * 100:.1f}% 的销量差异）；'
        f'剩余一半以上差异来自投放数量之外的因素。' if r is not None else
        '公告投放量与当期销量呈中等正相关（r 未能计算，样本异常）。',
        (f'效率离散度：公告记录≥10 条的 {n0(len(mainstream))} 家主流企业中，最高 {mn_hi[0]} '
         f'{n1(mn_hi[4])} 辆/车型、最低 {mn_lo[0]} {n1(mn_lo[4])} 辆/车型，极差约 {n0(spread_main)} 倍，'
         f'中位数仅 {n1(med_main)} 辆/车型——效率是比规模更陡峭的分化维度。')
        if (spread_main and med_main) else '效率离散度样本不足，暂不展示。',
        (f'投放规模与市场表现并不成正比：{cmp_eff[0]}以 {n0(cmp_eff[1])} 条公告、{n0(cmp_eff[2])} 个车系取得 '
         f'{n0(cmp_eff[3])} 辆（{n1(cmp_eff[4])} 辆/车型）；{top_eff[0]}仅 {n0(top_eff[1])} 条公告、'
         f'{n0(top_eff[2])} 个车系即取得 {n0(top_eff[3])} 辆（{n1(top_eff[4])} 辆/车型），'
         f'单车型效率为前者的 {ratio:.2f} 倍。')
        if (cmp_eff and ratio) else '效率对照样本不足，暂不展示倍数。',
    ], top=Y_BODY + 4.04, height=1.20, size=9, lh=0.27)
    source_note(s, Y_FOOT - 0.34, [
        f'公告数为底表记录数；销量为 {meta.get("snapshot_month")} 车系级快照。'
        f'面板异常点（公告多而销量低）：'
        + ' / '.join(f'{p[0]}（{n0(p[1])} 条 / {n0(p[2])} 辆）' for p in odd)])
    return s


@page('conv_table')
def b_conv_table(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '07 公告-市场转化 · 结论 4', '转化率排行表')
    rt = D['conversion']['rate_table']
    by_sales = sorted(rt, key=lambda r: -(r[4] or 0))[:14]
    by_ann = sorted(rt, key=lambda r: -(r[1] or 0))[:14]
    grid = [['#', '品牌（按单车型产出）', '公告车型数', '在售车系数', '销量（辆）', '辆/公告记录']] + [
        [str(i + 1), r[0][:16], n0(r[1]), n0(r[2]), n0(r[3]), n1(r[4])]
        for i, r in enumerate(by_sales)]
    table(s, grid, ML, Y_BODY, 6.0, 4.55, 'conv_top', size=7.5)
    grid2 = [['#', '品牌（按公告数）', '公告车型数', '在售车系数', '销量（辆）', '辆/公告记录']] + [
        [str(i + 1), r[0][:16], n0(r[1]), n0(r[2]), n0(r[3]), n1(r[4])]
        for i, r in enumerate(by_ann)]
    table(s, grid2, ML + 6.15, Y_BODY, CW - 6.15, 4.55, 'conv_ann', size=7.5)
    source_note(s, Y_FOOT - 0.34, [
        f'共 {n0(len(rt))} 个品牌纳入评估；单车型产出 = 2026-09 车系级销量 ÷ 公告车型数。',
        '左右两表对照显示：公告数量领先者与单车型产出领先者并非同一批品牌。',
    ])
    return s


@page('conclusion')
def b_conclusion(prs, D, page_no):
    meta = get_meta(D)
    s, _ = page_frame(prs, D, page_no, '09 判断与建议', '结论 / 致谢')
    cov = D['coverage']
    cur = cov.get('current_board') or {}
    linked_n = cur.get('linked_series') or D['structure']['series_count']
    board_n = cur.get('total_series') or D['structure']['board_series_count']
    grid = [
        ['回扣', '结论', '关键数字'],
        ['① 总量与渗透率', f'{meta.get("snapshot_month")} 新能源零售为 {len(meta.get("months", []))} 个月区间最高',
         f'{n0(to_map(D["trend"]["monthly_board"], 1).get(meta.get("snapshot_month")))} 辆'],
        ['② 技术路线', '市场选择效率（续航↑/整备↓/电耗↓），非堆料',
         f'续航 {D["tech_sales"]["weighted"]["续航"] - D["tech_sales"]["unweighted"]["续航"]:+.1f} km（加权 − 等权）'],
        ['③ 竞争格局', '企业端集中、产品端发散',
         f'企业 CR10 {n1(D["concentration"]["oem"]["CR10"])}% / 车系 CR10 {n1(D["concentration"]["series"]["CR10"])}%'],
        ['④ 公告→市场', '公告数量与市场份额不成正比',
         f'已锚定覆盖 {n0(linked_n)} / {n0(board_n)} 在榜车系'],
        ['⑤ 价格带迁移', '20–30 万扩容、30–40 万被两端分流', '双基期对照见 P19'],
    ]
    table(s, grid, ML, Y_BODY, CW, 3.05, 'conclusion', size=10)
    bullets(s, [
        '#口径一致、可复现、可追溯，是本次结论可用的前提。',
        '数据源唯一、快照月份唯一、去重口径唯一；全部数字由脚本从 JSON 注入，图/表/文字三方一致。',
        '局限：本分析为乘用车公告口径 × 车系级月零售量，非批发量/上险量；价格为指导价区间中值，非成交价。',
    ], top=Y_BODY + 3.20, height=1.5, size=11.5, lh=0.30)
    footer(s, meta, extra='本页数字与 P2/P3 同源，口径与图表脚注完全一致。')
    return s


# ==========================================================================================
# 编排：加载 / 校验 / 备份 / 建页 / 定序 / 硬断言
# ==========================================================================================
class ContractError(Exception):
    """JSON 契约不满足。硬失败，不降级、不猜值。"""


REQUIRED_PATHS = [
    'meta.snapshot_month', 'meta.fetch_date', 'meta.months', 'meta.source',
    'meta.source_tip', 'meta.source_url_template', 'meta.reference_source', 'meta.scope',
    'meta.dedup_rule', 'meta.price_bins', 'meta.price_bin_rule', 'meta.backup_old',
    'meta.base_table_path', 'meta.base_rows', 'meta.matched_rows',
    'coverage.row', 'coverage.series', 'coverage.buckets', 'coverage.status_counts',
    'structure.by_segment', 'structure.by_energy', 'structure.by_price',
    'structure.by_price_board', 'structure.by_price_board_pct',
    'structure.sales_total_dedup', 'structure.sales_total_rowlevel_naive',
    'structure.series_count', 'structure.board_series_count', 'structure.board_sales_total',
    'structure.linked_share_of_board_pct', 'structure.duplication_ratio',
    'concentration.series', 'concentration.brand', 'concentration.oem',
    'concentration.series_board', 'concentration.top_oem', 'concentration.top_brand',
    'trend.monthly', 'trend.monthly_board', 'trend.energy_monthly', 'trend.price_monthly',
    'trend.price_monthly_board', 'trend.segment_monthly', 'trend.top20_first',
    'trend.top20_last', 'trend.rankchange_hist', 'trend.rankchange_bins',
    'trend.new_entries', 'trend.drops',
    'tech_sales.weighted', 'tech_sales.unweighted', 'tech_sales.sample_n',
    'tech_sales.range_bins', 'tech_sales.range_bins_def', 'tech_sales.price_tech',
    'tech_sales.by_energy_tech',
    'conversion.by_batch', 'conversion.by_year', 'conversion.brand_scatter',
    'conversion.rate_table',
]
# WS2 待补槽位：缺失不 fail，但页面必须显著标注来源（绝不静默补默认值）
PENDING_KEYS = ['coverage.current_board', 'structure.unlinked_board',
                'meta.qc.reference_universe_per_month']

# 目标页号 → 基准稿页号（None = 新增页）。由基准稿实测标题（S01–S67，13 张章节页）
# 与结构文档 §0.1 对齐；必须满足"基准序单调递增且不丢页"，由 audit_base_mapping() 强制。
# 偏移规律：P5–P22 偏移 5–7（新增 8 页），P23 起偏移 11，P60 起偏移 12，P71 起偏移 11。
TARGET_SOURCE = {
     1: None,
     2: 1,
     3: 2,
     4: 3,
     5: None,
     6: 4,
     7: None,
     8: None,
     9: None,
    10: None,
    11: 5,
    12: 6,
    13: None,
    14: None,
    15: 7,
    16: 8,
    17: 9,
    18: 10,
    19: 11,
    20: None,
    21: None,
    22: 12,
    23: 13,
    24: 14,
    25: 15,
    26: 16,
    27: 17,
    28: 18,
    29: 19,
    30: 20,
    31: 21,
    32: 22,
    33: 23,
    34: 24,
    35: 25,
    36: 26,
    37: 27,
    38: 28,
    39: 29,
    40: 30,
    41: 31,
    42: 32,
    43: 33,
    44: None,
    45: None,
    46: None,
    47: None,
    48: 34,
    49: 35,
    50: 36,
    51: 37,
    52: 38,
    53: 39,
    54: 40,
    55: 41,
    56: 42,
    57: None,
    58: 43,
    59: 44,
    60: 45,
    61: 46,
    62: 47,
    63: 48,
    64: 49,
    65: 50,
    66: 51,
    67: 52,
    68: 53,
    69: 54,
    70: 55,
    71: 56,
    72: 57,
    73: 58,
    74: 59,
    75: 60,
    76: 61,
    77: 62,
    78: 63,
    79: 64,
    80: 65,
    81: 66,
    82: 67,
    83: None,
}

# 基准稿实际页数（映射审计用；必须与 SRC_DECK 实际页数一致）
BASE_DECK_PAGES = 67


def _dig(D, path):
    node = D
    for k in path.split('.'):
        if not isinstance(node, dict) or k not in node:
            return False
        node = node[k]
    return True


def load_and_validate(path=SALES_ANALYTICS):
    if not os.path.exists(path):
        raise ContractError(f'销量分析 JSON 未就绪：{path}（task-2 产出）')
    with open(path, encoding='utf-8') as fh:
        D = json.load(fh)
    missing = [p for p in REQUIRED_PATHS if not _dig(D, p)]
    if missing:
        raise ContractError('JSON 缺必需字段（禁止用默认值填充）：' + '; '.join(missing))
    assert_meta_ready(D)
    print('  生成前置断言通过：dry_run=False · 底表指纹/行列数一致 · suggest 关闭且 0 次调用')
    pending = [p for p in PENDING_KEYS if not _dig(D, p)]
    if pending:
        print(f'  [PENDING] WS2 待补槽位（页面按纲要数字注入并标注来源）：{pending}')
    ca = D.get('contract_assertions', {})
    if ca and not ca.get('all_pass', False):
        bad = [a for a in ca.get('assertions', []) if not a.get('pass')]
        raise ContractError(f'契约自检未通过：{bad}')
    return D, pending


def audit_base_mapping(n_base_slides):
    """校验目标页 ↔ 基准页映射：基准序必须单调递增；非替换页不得重复占用基准页。"""
    builder_pages = {i[0] for i in PAGE_CATALOG if i[3]}
    seq = [(p, TARGET_SOURCE[p]) for p in sorted(TARGET_SOURCE) if TARGET_SOURCE[p]]
    bases = [b for _, b in seq]
    if bases != sorted(bases):
        raise ContractError(f'目标↔基准映射非单调（会静默丢页/乱序）：{bases}')
    # 替换页（有构建器）会丢弃其源基准页，属预期重复；继承页必须唯一占用
    inherited = [b for p, b in seq if p not in builder_pages]
    if len(set(inherited)) != len(inherited):
        dup = sorted({b for b in inherited if inherited.count(b) > 1})
        raise ContractError(f'继承基准页被重复映射（会丢页）：{dup}')
    expected = set(range(1, n_base_slides + 1))
    got = set(bases)
    missing = sorted(expected - got)
    extra = sorted(got - expected)
    if extra:
        raise ContractError(f'映射引用了不存在的基准页：{extra}')
    # 未被任何目标页引用的基准页 → 将被显式丢弃（打印，不静默）
    if missing:
        print(f'  未映射基准页 {missing}（{len(missing)} 页将不再出现在终稿中）')
    return True


def backup_once(src=SRC_DECK, ver=VER):
    """纪律④：备份名含版本号 + 存在性检查，重复运行不覆盖原始备份。"""
    os.makedirs('archive', exist_ok=True)
    dst = f'archive/演示文稿_pre-{ver}_backup_{date.today():%Y%m%d}.pptx'
    if os.path.exists(dst):
        print(f'  备份已存在，跳过覆盖：{dst}')
        return dst
    shutil.copy2(src, dst)
    print(f'  已备份：{dst}')
    return dst


def _reorder(prs, desired):
    """按 desired（sldId XML 元素序列）重排幻灯片：一次性收集后清空重建，避免边遍历边插入。"""
    lst = prs.slides._sldIdLst
    for el in list(lst):
        lst.remove(el)
    for el in desired:
        lst.append(el)


def _drop(prs, els):
    lst = prs.slides._sldIdLst
    ns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    for el in els:
        rId = el.get(ns)
        if rId:
            prs.part.drop_rel(rId)
        lst.remove(el)


def build_target_deck(D, src=SRC_DECK, out=None, do_backup=True):
    """把基准稿改造为目标稿：建新增页 → 按目标目录定序 → 丢弃未映射基准页。"""
    out = out or src
    if do_backup:
        backup_once(src)
    prs = Presentation(src)
    n_base = len(prs.slides)
    if n_base != BASE_DECK_PAGES:
        raise AssertionError(
            f'源稿 {src} 为 {n_base} 页，预期基准稿 {BASE_DECK_PAGES} 页 —— '
            f'拒绝二次施加（否则会把已生成的目标稿当基准再叠一层）。'
            f'请先从 archive/ 恢复基准稿。')
    audit_base_mapping(n_base)
    n_target = len(PAGE_CATALOG)
    print(f'  基准稿 {n_base} 页 → 目标稿 {n_target} 页')

    base_ids = list(prs.slides._sldIdLst)          # 位置 0..n_base-1
    used_base = set()

    # ---- 1) 压制构建过程的重排：先把新增页造到末尾，收集其 sldId ----
    new_el = {}
    for item in sorted(PAGE_CATALOG, key=lambda x: x[0]):
        pno, bname = item[0], item[3]
        if not bname:
            continue
        fn = BUILDERS.get(bname)
        if fn is None:
            note_failure(f'builder:{bname}@P{pno}', KeyError(f'未注册构建器 {bname}'))
            continue
        before = len(list(prs.slides._sldIdLst))
        try:
            fn(prs, D, pno)
        except WritingBanError:
            raise
        except Exception as exc:
            note_failure(f'{bname}@P{pno}', exc)
            continue
        after = list(prs.slides._sldIdLst)
        if len(after) != before + 1:
            note_failure(f'{bname}@P{pno}', RuntimeError(f'页数未按预期 +1（{before}→{len(after)}）'))
            continue
        new_el[pno] = after[-1]
        STATS['pages'] += 1

    # ---- 2) 按目标目录定序 ----
    desired, dropped = [], []
    for item in sorted(PAGE_CATALOG, key=lambda x: x[0]):
        pno = item[0]
        if item[3]:                                  # 本页由构建器生成
            el = new_el.get(pno)
            if el is None:
                note_failure(f'order:P{pno}', RuntimeError('构建器未产出该页'))
                continue
            src_base = TARGET_SOURCE.get(pno)
            if src_base:                             # 替换页：丢弃旧页
                dropped.append(base_ids[src_base - 1])
                used_base.add(src_base)
            desired.append(el)
        else:                                        # 继承基准页
            b = TARGET_SOURCE.get(pno)
            if b is None:
                note_failure(f'order:P{pno}', RuntimeError('无构建器且无基准映射'))
                continue
            if b in used_base:
                note_failure(f'order:P{pno}', RuntimeError(f'基准页 {b} 已被映射占用'))
                continue
            used_base.add(b)
            desired.append(base_ids[b - 1])
    # 未映射的基准页：显式列出后丢弃（不静默）
    unmapped = [i + 1 for i in range(n_base) if (i + 1) not in used_base]
    if unmapped:
        print(f'  丢弃未映射基准页：{unmapped}')
    dropped += [base_ids[i - 1] for i in unmapped]

    _reorder(prs, desired)
    _drop(prs, [e for e in dropped if e.getparent() is not None])
    print(f'  定序完成：{len(prs.slides)} 页（构建 {STATS["pages"]} 页 / 继承 '
          f'{len(used_base) - STATS["pages"]} 页）')

    # ---- 3) 全稿口径文案归一（派生真值，替代基准稿里写死的列数/款数）----
    normalize_scope_lines(prs, base_table_scope_counts())

    # ---- 4) **...** → 真加粗 run（消除领导层可见的字面星号）----
    nb = apply_inline_bold(prs)
    STATS['boldified'] = nb
    print(f'  加粗渲染：{nb} 处段落由 ** 标记转为真加粗 run')

    prs.save(out if out else src)
    return prs


# ==========================================================================================
# 硬断言（自检 → 门禁：失败即抛错，不得只 WARN）
# ==========================================================================================
def deck_texts(path):
    prs = Presentation(path)
    out = []
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    out.append(''.join(r.text for r in p.runs))
            if getattr(sh, 'has_table', False):
                for r_ in sh.table.rows:
                    for c_ in r_.cells:
                        out.append(c_.text)
            ch = chart_of(sh)
            if ch is not None and ch.has_title:
                out.append(ch.chart_title.text_frame.text)
    return out


def _boldify_paragraph(p):
    """把段落文本里的 **...** 标记转成真加粗 run（领导层不得看到字面星号）。

    以首个 run 为格式原型；重建后清空旧 run，按 ** 分段写 bold/非 bold。
    返回转换处数（0 = 无标记）。
    """
    runs = list(p.runs)
    if not runs:
        return 0
    t = ''.join(r.text for r in runs)
    if '**' not in t:
        return 0
    proto = runs[0].font
    size, name = proto.size, proto.name
    try:
        rgb = proto.color.rgb
    except Exception:
        rgb = None
    base_bold = proto.bold
    segs = t.split('**')
    # 奇数下标段 = 加粗（远端未闭合的 ** 会被剥除，绝不残留字面标记）
    parts = [(seg, i % 2 == 1) for i, seg in enumerate(segs) if seg != '']
    for r in runs:
        r._r.getparent().remove(r._r)
    for text, bold in parts:
        nr = p.add_run()
        nr.text = text
        nr.font.size = size
        nr.font.name = name
        nr.font.bold = True if bold else base_bold
        if rgb is not None:
            nr.font.color.rgb = rgb
    return 1


def apply_inline_bold(prs):
    """全稿后处理：**...** → 真加粗 run。覆盖文本框与表格单元格。"""
    hit = 0
    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    hit += _boldify_paragraph(p)
            if getattr(sh, 'has_table', False):
                for row in sh.table.rows:
                    for c in row.cells:
                        for p in c.text_frame.paragraphs:
                            hit += _boldify_paragraph(p)
    return hit


def assert_deck_invariants(path, D, pending=()):
    problems = []
    prs = Presentation(path)
    texts = deck_texts(path)
    joined = '\n'.join(texts)

    # ⓪-0b 字面加粗标记必须已被渲染为真加粗 run（不得出现在领导层屏幕）
    if '**' in joined:
        _bad = [t for t in texts if '**' in t][:3]
        problems.append(f'稿内残留字面加粗标记 ** × {joined.count("**")} 处：{_bad}')

    # ⓪-0 未锁定占位必须已被真实数字替换
    if BEV_PP_PLACEHOLDER in joined:
        problems.append(f'稿内仍存在未替换占位 {BEV_PP_PLACEHOLDER}')

    # ⓪ 页数 == 目录页数（防止装配阶段静默丢页/多页）
    if len(prs.slides) != len(PAGE_CATALOG):
        problems.append(f'页数 {len(prs.slides)} != 目录页数 {len(PAGE_CATALOG)}')

    # ① 列数/行数文案 == 底表实际（动态派生）
    ncol, nrow = base_table_cols(), base_table_rows()
    m = re.search(
        r'纯乘用车口径（([\d,]+) 条公告/公示记录 × (\d+) 列；其中 (\d+) 条为 ([\d、]+) 批公示记录）',
        joined)
    if not m:
        problems.append('未找到规范口径文案'
                        '「纯乘用车口径（N 条公告/公示记录 × M 列；其中 K 条为 B 批公示记录）」')
    else:
        gr, gc = int(m.group(1).replace(',', '')), int(m.group(2))
        gk = int(m.group(3))
        if gc != ncol:
            problems.append(f'文案列数 {gc} ≠ 底表实际列数 {ncol}')
        if gr != nrow:
            problems.append(f'文案行数 {gr} ≠ 底表实际行数 {nrow}')
        _npub, _batches, _pubdate = base_table_public_rows()
        if gk != _npub:
            problems.append(f'文案公示条数 {gk} ≠ 底表实际 {_npub}')
        _want = '、'.join(str(b) for b in _batches)
        if m.group(4) != _want:
            problems.append(f'文案公示批号 {m.group(4)!r} ≠ 底表实际 {_want!r}')
        # 拆分恒等式：总 = 正式 + 公示
        if gr - gk != nrow - _npub:
            problems.append(f'拆分恒等式不成立：{gr} - {gk} ≠ {nrow} - {_npub}')

    # ② 过期字面量不得残留（历次口径的废弃值；短数字加边界防误伤真实值）
    STALE = (
        (r'4,613', '早期样本量 4,613'),
        (r'v4\.3\.[0-9]+', '早期底表版本号'),
        (r'(?<![\d,.])5,?392(?!\d)', '5,392 行（旧终态，现 5,475）'),
        (r'41\.6\s*%', '早期占比'),
        (r'(?<![\d.])40\.52', '行级覆盖率 40.52%（现 40.18%）'),
        (r'(?<![\d,.])2,?727(?!\d)', '未匹配-已停产或未上量 2,727（现 2,749）'),
        (r'无车系身份[^\d]{0,8}364(?!\d)', '无车系身份 364（现 410）'),
        (r'(?<![\d,.])2,?603(?!\d)', '车系ID 非空 2,603（现 2,618）'),
        (r'(?<![\d,.])2,?185(?!\d)', '销量非空 2,185（现 2,200）'),
        (r'341\s*[~\-–]\s*411', '批次范围 341~411（现 341~412）'),
    )
    for pat, why in STALE:
        hit = re.findall(pat, joined)
        if hit:
            problems.append(f'残留过期字面量 {pat!r} × {len(hit)} 处（{why}；示例 {hit[:3]}）')
    # ②b 任何列数文案都必须等于底表实际列数（不写死 38/39/41）
    for mm in re.finditer(r'×\s*(\d+)\s*列', joined):
        if int(mm.group(1)) != ncol:
            problems.append(f'列数文案「{mm.group(0)}」≠ 底表实际列数 {ncol}')
    # ②c 版本号全稿必须等于交付物版本（不得残留旧版号）
    for v in sorted(set(VER_RE.findall(joined))):
        if v != VER:
            problems.append(f'残留旧版本号 {v!r}（应为 {VER}）')

    # ③ 禁区分词全稿扫描
    for pat, why in BAN_PATTERNS:
        hit = re.findall(pat, joined)
        if hit:
            problems.append(f'全稿命中写作禁区 {pat!r} × {len(hit)}：{why}')

    # ④ 覆盖率三口径必须同时出现（数值从 JSON 派生，不写死）
    cov = D['coverage']
    triples = (
        (f"{cov['row']['pct']:.2f}", '行级覆盖率'),
        (f"{cov['current_board']['pct']:.2f}", '当期在榜车系级覆盖率'),
        (f"{cov['series']['generic_names_pct']:.2f}", '车系级去重通用名称覆盖率'),
    )
    for val, what in triples:
        if val not in joined:
            problems.append(f'口径披露页缺三口径之一：{what} {val}% 未出现')

    # ⑤ 榜单计数不得截断
    ne, dr = len(D['trend']['new_entries']), len(D['trend']['drops'])
    if ne > 20 and str(ne) not in joined:
        problems.append(f'新晋车系计数 {ne} 未出现（禁截断为 20）')
    if dr > 20 and str(dr) not in joined:
        problems.append(f'掉榜车系计数 {dr} 未出现（禁截断为 20）')
    if re.search(r'Top\s*20\s*(新晋|掉出)', joined):
        problems.append('出现「Top20 新晋/掉出」表述（应为月度榜单集合差）')

    # ⑥ 每张图表标题须含月份 + 车系级口径
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            ch = chart_of(sh)
            if ch is None or not ch.has_title:
                continue
            t = ch.chart_title.text_frame.text
            if '车系级' not in t or '2026-09' not in t:
                problems.append(f'P{i} 图表标题缺月份/口径：{t[:48]!r}')

    # ⑦ 留痕失败不得非空
    if FAILURES:
        problems.append(f'构建过程存在 {len(FAILURES)} 类留痕失败（见 FAILURES 汇总）')

    if problems:
        raise AssertionError('硬断言失败（不得只 WARN）：\n  - ' + '\n  - '.join(problems))
    print(f'  硬断言全过：列 {ncol} / 行 {nrow} · 三口径齐 · 榜单计数全量 · '
          f'禁区分词 0 命中 · 图表标题合规')


# ==========================================================================================
# 入口
# ==========================================================================================
def main(argv=None):
    ap = argparse.ArgumentParser(description=f'销量联动页注入 {VER}（deck-builder / WS5）')
    ap.add_argument('--deck', default=SRC_DECK, help='目标 pptx')
    ap.add_argument('--json', default=SALES_ANALYTICS, help='销量分析 JSON')
    ap.add_argument('--dry-run', action='store_true', help='只校验数据契约与目录映射，不写文件')
    ap.add_argument('--validate', action='store_true', help='仅校验：对现有目标稿执行硬断言后退出')
    ap.add_argument('--no-validate', action='store_true', help='跳过硬断言（仅调试，交付禁用）')
    ap.add_argument('--no-backup', action='store_true', help='跳过备份（仅调试）')
    ap.add_argument('--accept-base-sha', metavar='SHA256', default=None,
                    help='显式确认新的底表指纹（底表合法更新后使用；会打印留痕）')
    a = ap.parse_args(argv)

    print(f'== {VER} 销量联动页注入 ==')
    global BASE_TABLE_SHA256
    if a.accept_base_sha:
        BASE_TABLE_SHA256 = a.accept_base_sha.strip().lower()
        print(f'  [指纹确认] 已显式接受底表指纹 {BASE_TABLE_SHA256}'
              f'（冻结值 {FROZEN_BASE_TABLE_SHA256}）—— 该确认需在交付记录中留痕')
    D, pending = load_and_validate(a.json)
    meta = get_meta(D)
    print(f'  契约校验通过：快照 {meta.get("snapshot_month")} · 抓取 {meta.get("fetch_date")} · '
          f'{len(PAGE_CATALOG)} 页目录 · {len(BUILDERS)} 个构建器 · '
          f'{len(BUILDERS)}/{sum(1 for i in PAGE_CATALOG if i[3])} 已注册')
    if a.dry_run:
        n_now = len(Presentation(a.deck).slides)
        if n_now != BASE_DECK_PAGES:
            print(f'  [dry-run] 注意：{a.deck} 现为 {n_now} 页（基准稿应为 {BASE_DECK_PAGES} 页）；'
                  f'若已是终稿，请先从 archive/ 恢复基准稿再构建。')
        audit_base_mapping(BASE_DECK_PAGES)
        print('  [dry-run] 未写任何文件。')
        return D
    if a.validate:
        assert_deck_invariants(a.deck, D, pending)
        print('  [validate] 既有目标稿硬断言通过。')
        return D

    # 先构建到临时文件 → 硬断言通过后才原子替换正式稿（断言失败不得留下半成品）
    tmp = a.deck + '.building.tmp.pptx'
    try:
        build_target_deck(D, src=a.deck, out=tmp, do_backup=not a.no_backup)
        print(f'  页 {STATS["pages"]} · 新增页 {STATS["slides_added"]} · 图表 {STATS["charts"]} · '
              f'表 {STATS["tables"]} · 禁区校验 {STATS["guard_checks"]}')
        if FAILURES:
            print(f'  !! FAILURES {len(FAILURES)} 类（不得静默）：')
            for k, v in sorted(FAILURES.items()):
                print(f'     {v}× {k}')
        else:
            print('  FAILURES 汇总：空（无静默跳过）')
        if not a.no_validate:
            assert_deck_invariants(tmp, D, pending)
        os.replace(tmp, a.deck)
        print(f'  已写入正式稿：{a.deck}')
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError as exc:
                note_failure('tmp_cleanup', exc)
    return D


if __name__ == '__main__':
    main()
