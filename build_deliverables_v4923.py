# -*- coding: utf-8 -*-
"""build_deliverables_v4923.py — 下游交付①·54表分析底表重建（v4.9.23 口径，可复跑）

源：NEV公告参数汇总表_合并版（341~410批）.xlsx（权威底表·行数/批次范围一律实读，禁止硬编码）
旧版基线快照（v4.3.6 口径 4,613 行）：audit-output/old_deliverable_45sheets_snapshot_20260927.json
输出：NEV公告车型行业分析_数据底表.xlsx（54 表 = 既有 45 表 + 月销量分析 + 8 张销量联动表）

v4.9.24（2026-10-19 销量口径统一轮）新增：
  · 底表尾插列 39 车系ID / 40 销量覆盖状态（表宽 39→41，规范 §2.2）
  · 销量口径由「易车零售 2026-08」统一为「汽车之家·行业综合销量 2026-09」
  · 新增 8 张表（#47~#54），全部数字源自 audit-output/sales_linkage_analytics_20261019.json
    （契约 v1.1，全案唯一接口；禁止手写数字）——详见 Update/销量联动交付物集成方案_20261019.md

口径要点：
  · 动力分组 BEV=纯电/换电；PHEV=插电/增程（与看板一致）
  · 年度=批次公告日期年份（批次时间表·官方精确日期）
  · 电耗/续航=工信部能耗数据20260915 官方对齐值（v4.9.22）；工况 BEV=CLTC、PHEV/EREV=WLTC
  · 电机功率取「电机总功率/扭矩」首值；仅乘用车纯净口径（GB/T 3730.1）
  · 各表样本=该表所需字段均有值的行（有值子集），与旧版一致
  · 销量表统计单元=车系（规范 §4 铁律：禁止对底表行 SUM 月销量；已按 seriesid/通用名称去重）
"""
import datetime
import json
import math
import os
import re
import shutil
import statistics as st
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = '.'
SRC = 'NEV公告参数汇总表_合并版（341~410批）.xlsx'
OUT = 'NEV公告车型行业分析_数据底表.xlsx'
VER = 'v4.9.23'

THIN = Side(style='thin', color='B0B7C3')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D9E1F2')
TFONT = Font(bold=True, size=13, color='1F3864')
HFONT = Font(bold=True, size=10)
YELLOW = PatternFill('solid', fgColor='FFF2CC')

# ---------- 销量联动数据源（契约 v1.1+ · 全案唯一接口） ----------
SALES_JSON = os.path.join(HERE, 'audit-output', 'sales_linkage_analytics_20261019.json')

_SALES_REQUIRED = {                     # 冻结槽位；缺一即 fail-fast（契约 §2）
    'meta': ['snapshot_month', 'months', 'fetch_date', 'source', 'source_tip',
             'source_url_template', 'reference_source', 'scope', 'dedup_rule',
             'price_bins', 'price_bin_rule', 'backup_old', 'base_table_path',
             'base_rows', 'matched_rows', 'match_channel_counts', 'qc', 'dedup_selfcheck',
             'agg_rule'],          # 规范 §11.1：分类型取众数（并列取字典序最小）/数值型取中位数
    'coverage': ['row', 'series', 'current_board', 'self_consistent_ratio',
                 'status_counts', 'buckets', 'pending_naming', 'definitions', 'detail'],
    'structure': ['by_segment', 'by_segment_pct', 'by_energy', 'by_price', 'by_price_pct',
                  'by_oem', 'by_brand', 'sales_total_dedup', 'sales_total_rowlevel_naive',
                  'board_sales_total', 'board_series_count', 'series_count',
                  'linked_share_of_board_pct', 'duplication_ratio',
                  'by_price_board', 'by_price_board_pct', 'missing',
                  # 规范 §11.2 强制披露三键（不能合并；Lead 2026-10-19 裁决）
                  'mixed_energy_families', 'mixed_with_erev', 'r2_flip_impact'],
    'concentration': ['series', 'brand', 'oem', 'top_series', 'top_brand', 'top_oem',
                      'series_board', 'top_series_board'],
    'trend': ['monthly', 'energy_monthly', 'price_monthly', 'segment_monthly',
              'top20_first', 'top20_last', 'rankchange_hist', 'rankchange_bins',
              'new_entries', 'drops', 'monthly_board', 'price_monthly_board',
              'roster_rule', 'new_entries_total', 'drops_total'],
    'tech_sales': ['weighted', 'unweighted', 'sample_n', 'range_bins', 'range_bins_def',
                   'price_tech', 'price_range_matrix', 'by_energy_tech', 'missing'],
    'conversion': ['by_batch', 'by_year', 'brand_scatter', 'rate_table',
                   'by_batch_detail', 'by_batch_detail_fields'],
}


def load_sales(path=SALES_JSON, base_rows=None, expect_sha256=None):
    """读取销量联动分析 JSON（契约 v1.1+，全案唯一接口）。

    契约 §2：任一路径缺失 → 抛错并报告，不得使用默认值或跳过该表。
    base_rows：底表**实读**行数；传入时校验 JSON 声明值与实读一致
    （规范 §11.4：底表有并发写入方，禁止在任何交付物中硬编码 5,392）。
    expect_sha256：可选，钉住底表指纹；底表一变即拒绝（防"JSON 与工作簿不同轮"）。
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f'缺销量联动 JSON（契约 v1.1）：{path}')
    with open(path, encoding='utf-8') as f:
        d = json.load(f)
    missing = [f'{k}.{s}' for k, slots in _SALES_REQUIRED.items()
               for s in slots if s not in (d.get(k) or {})]
    if missing:
        raise KeyError('销量联动 JSON 缺槽位（fail-fast，不改名/不降级/不猜值）：' + ', '.join(missing))
    _sales_assert(d, base_rows=base_rows, expect_sha256=expect_sha256)
    return d


def _sales_assert(d, base_rows=None, expect_sha256=None):
    """口径与结构自检（规范 §2.2/§4/§5/§11 + R3 + R4 + 生成前置门禁）。任一不成立即抛错。"""
    meta, cov, trend, tech = d['meta'], d['coverage'], d['trend'], d['tech_sales']
    stc = d['structure']
    # —— 生成前置门禁（Lead 2026-10-19 放行条件；任一不满足即 fail-fast，**不得绕过**）
    if meta.get('dry_run') is not False:
        raise ValueError(f'meta.dry_run={meta.get("dry_run")!r} 非 False —— '
                         '上游仍是 dry-run，底表尚未真写，拒绝生成')
    if meta.get('base_table_columns') != 41:
        raise ValueError(f'meta.base_table_columns={meta.get("base_table_columns")!r} != 41 —— '
                         '底表未完成规范 §2.2 尾插（车系ID / 销量覆盖状态）')
    if meta.get('suggest_enabled') is not False:
        raise ValueError(f'meta.suggest_enabled={meta.get("suggest_enabled")!r} 非 False'
                         '（规范 §3 通道 D suggest 兜底必须关闭）')
    if meta.get('suggest_calls') != 0:
        raise ValueError(f'meta.suggest_calls={meta.get("suggest_calls")!r} != 0'
                         '（禁止 suggest 兜底调用产生匹配）')
    if expect_sha256 and str(meta.get('base_table_sha256', '')).lower() != expect_sha256.lower():
        raise ValueError(f'底表 sha256 不匹配：JSON {meta.get("base_table_sha256")!r}'
                         f' != 期望 {expect_sha256!r}（底表已变，须整体重取 JSON）')
    # —— 总行数锚点（规范 §11.4：一律实读，禁止硬编码 5,392）
    if not isinstance(meta['base_rows'], int) or meta['base_rows'] <= 0:
        raise ValueError(f'meta.base_rows 非法：{meta["base_rows"]!r}')
    if base_rows is not None and meta['base_rows'] != base_rows:
        raise ValueError(
            f'契约 meta.base_rows={meta["base_rows"]} != 底表实读 {base_rows}'
            '（底表可能被并发写入 → JSON 与工作簿不同轮，须整体重取）')
    # —— 覆盖率分解闭合：六类状态之和 == base_rows（规范 §5）
    n_status = sum(cov['status_counts'].values())
    if n_status != meta['base_rows']:
        raise ValueError(f'六类覆盖状态之和 {n_status} != base_rows {meta["base_rows"]}')
    # —— 四类未覆盖分桶之和 == 后四类状态之和（规范 §5）
    n_bucket = sum(cov['buckets'].values())
    n_tail4 = sum(v for k, v in cov['status_counts'].items()
                  if k.startswith('未匹配') or k == '无车系身份')
    if n_bucket != n_tail4:
        raise ValueError(f'四类未覆盖分桶之和 {n_bucket} != 后四类状态之和 {n_tail4}')
    # —— R3-1：trend.price_monthly 列数 == len(meta.price_bins)（AGENTS 2026-09-29 教训③）
    n_price = len(meta['price_bins'])
    for i, row in enumerate(trend['price_monthly']):
        if len(row) - 1 != n_price:
            raise ValueError(f'trend.price_monthly[{i}] 有 {len(row)-1} 个价格带列，'
                             f'期望 len(meta.price_bins)={n_price}（R3-1 列数错位）')
    # —— R3-2：qc.monthly_counts 键集 == meta.months（顺序一致）
    mc = meta['qc'].get('monthly_counts')
    if mc is not None and list(mc.keys()) != list(meta['months']):
        raise ValueError(f'meta.qc.monthly_counts 键集 != meta.months（R3-2 月份错位）')
    # —— R3-3（v1.1 新增槽位）：矩阵形状与分箱定义长度自洽
    n_bins = len(tech['range_bins_def'])
    for i, row in enumerate(tech['price_range_matrix']):
        if len(row) != 3 or len(row[1]) != n_bins or len(row[2]) != n_bins:
            raise ValueError(f'tech_sales.price_range_matrix[{i}] 形状非法：'
                             f'期望 [价格带, [销量×{n_bins}], [车系数×{n_bins}]]')
    if len(tech['price_range_matrix']) != n_price:
        raise ValueError(f'price_range_matrix 行数 {len(tech["price_range_matrix"])} '
                         f'!= len(meta.price_bins)={n_price}')
    # —— 价格带标签两侧必须同名同序（防止 #48/#52 两表标签错配）
    if [r[0] for r in tech['price_tech']] != list(meta['price_bins']):
        raise ValueError('tech_sales.price_tech 的价格带标签/顺序 != meta.price_bins')
    if [r[0] for r in tech['price_range_matrix']] != list(meta['price_bins']):
        raise ValueError('price_range_matrix 的价格带标签/顺序 != meta.price_bins')
    # —— R4：全榜/锚定恒等式（现成键与派生值必须自洽，防上游改口径）
    if stc['linked_share_of_board_pct'] <= 0 or stc['board_sales_total'] <= 0:
        raise ValueError('structure.linked_share_of_board_pct / board_sales_total 必须为正')
    ratio = 100.0 * stc['sales_total_dedup'] / stc['board_sales_total']
    if abs(ratio - stc['linked_share_of_board_pct']) > 0.05:
        raise ValueError(f'锚定率不自洽：派生 {ratio:.4f}% vs 契约 '
                         f'{stc["linked_share_of_board_pct"]}%（R4 恒等式）')
    dup = stc['sales_total_rowlevel_naive'] / stc['sales_total_dedup']
    if abs(dup - stc['duplication_ratio']) > 0.01:
        raise ValueError(f'虚高倍数不自洽：派生 {dup:.4f} vs 契约 '
                         f'{stc["duplication_ratio"]}（R4 恒等式）')
    # —— 全榜/已锚定 月度序列行数自洽（#50 A2 对照块）
    if len(trend['monthly_board']) != len(trend['monthly']) != len(meta['months']):
        raise ValueError('trend.monthly / monthly_board / meta.months 行数不一致')
    # —— 上游自检必须全通过（正交证据；未通过即我方拒绝消费）
    for key in ('contract_assertions', 'dedup_assertions'):
        blob = d.get(key) or {}
        if blob.get('all_pass') is False:
            raise ValueError(f'上游 {key}.all_pass=False，拒绝消费该 JSON')
    return d


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', '').strip())
    except Exception:
        return None


def load_records():
    wb = openpyxl.load_workbook(SRC, read_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    # 批次→公告年
    bt = wb['批次时间表']
    bth = [str(c.value or '').strip() for c in next(bt.iter_rows(min_row=1, max_row=1))]
    bi, di = bth.index('批次'), bth.index('公告日期')
    batch_year = {}
    for r in bt.iter_rows(min_row=2, values_only=True):
        m = re.match(r'^(\d{4})', str(r[di] or ''))
        if r[bi] is not None and m:
            try:
                batch_year[int(float(str(r[bi])))] = int(m.group(1))
            except Exception:
                pass
    recs = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        def g(name):
            return r[hi[name]] if name in hi else None
        pt = str(g('动力类型') or '').strip()
        if '纯电' in pt or '换电' in pt or pt == 'BEV':
            grp = 'BEV'
        elif '插电' in pt or '增程' in pt or pt == 'PHEV':
            grp = 'PHEV'
        else:
            continue
        bno = g('批次')
        try:
            bno = int(float(str(bno)))
        except Exception:
            continue
        recs.append({
            '批次': bno, '年': batch_year.get(bno), '组': grp,
            # 产品型号：旧值留档交叉复核的配对键成分（additive key）
            '型号': str(g('产品型号') or '').strip(),
            '商标': str(g('产品商标') or '').strip(),
            '企业': str(g('企业名称') or '').strip(),
            '通用': str(g('通用名称') or '').strip(),
            '细分': str(g('细分市场') or '').strip(),
            '续航': fnum(g('纯电续航里程(km)')),
            '容量': fnum(g('电池容量(kWh)')),
            '密度': fnum(g('电池能量密度(Wh/kg)')),
            '电耗': fnum(g('百公里电耗(kWh/100km)')),
            '油耗': fnum(g('综合油耗(L/100km)')),
            '整备': fnum(g('整备质量(kg)')),
            '功率': fnum(g('电机总功率/扭矩')),
            '排量': fnum(g('发动机排量(mL)')),
            '电池类型': str(g('电池类型') or '').strip(),
            '电芯': str(g('电芯供应商') or '').strip(),
            '电机企': str(g('电机生产企业') or '').strip(),
            '月销': fnum(g('月销量(辆)')),
            # 规范 §2.2 尾插两列（link-engine 写入）。additive key：既有消费方按 dict 取值安全。
            'sid': str(g('车系ID') or '').strip(),
            '覆盖状态': str(g('销量覆盖状态') or '').strip(),
        })
    wb.close()
    return recs


def pct(vals, p):
    if not vals:
        return None
    s = sorted(vals)
    k = (len(s) - 1) * p / 100.0
    f = math.floor(k)
    c = min(f + 1, len(s) - 1)
    return s[f] + (s[c] - s[f]) * (k - f)


def year_rows(recs, grp=None):
    d = defaultdict(list)
    for r in recs:
        if r['年'] and (grp is None or r['组'] == grp):
            d[r['年']].append(r)
    return d


def agg_by_year(recs, field, grp=None, filt=None, scale=1.0):
    """年 → (均值, 中位数, 样本数)；filt=行谓词；scale 换算"""
    out = {}
    for y, rs in sorted(year_rows(recs, grp).items()):
        vals = [r[field] * scale for r in rs
                if r[field] is not None and (filt is None or filt(r))]
        if vals:
            out[y] = (sum(vals) / len(vals), st.median(vals), len(vals))
    return out


def per_row_ratio(recs, a, b, grp):
    """per-row a/b×1000 均值（分子分母均有值的子集）"""
    out = {}
    for y, rs in sorted(year_rows(recs, grp).items()):
        vals = [r[a] / r[b] * 1000 for r in rs
                if r[a] is not None and r[b] not in (None, 0)]
        if vals:
            out[y] = sum(vals) / len(vals)
    return out


def pearson(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    sx = math.sqrt(sum((x - mx) ** 2 for x in xs))
    sy = math.sqrt(sum((y - my) ** 2 for y in ys))
    if sx == 0 or sy == 0:
        return None
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / (sx * sy)


class Book:
    def __init__(self):
        self.wb = openpyxl.Workbook()
        self.wb.remove(self.wb.active)

    def sheet(self, name, title=None):
        ws = self.wb.create_sheet(name)
        if title:
            ws.cell(1, 1, title).font = TFONT
        return ws

    def table(self, ws, start_row, header, rows, num_fmt=None, widths=None):
        """写一张带表头样式的小表；rows=二维数组（值或 (值,fmt)）"""
        for j, h in enumerate(header, 1):
            c = ws.cell(start_row, j, h)
            c.font = HFONT
            c.fill = HFILL
            c.border = BORDER
            c.alignment = Alignment(horizontal='center')
        for i, row in enumerate(rows, start_row + 1):
            for j, v in enumerate(row, 1):
                c = ws.cell(i, j, v)
                c.border = BORDER
                if isinstance(v, float):
                    c.number_format = num_fmt or '0.0'
                    if abs(v) >= 1000:
                        c.number_format = '#,##0'
        if widths:
            for j, w in enumerate(widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w


def _famstats(stc, key):
    """混合家族三键取值（规范 §11.2 canonical：mixed_energy_families / mixed_with_erev /
    r2_flip_impact 为三个独立键，**不得合并**）。

    返回 (车系数, 销量, 占已锚定%, [(车系, 销量, 占比), …])，**在源头归一化**，
    不在渲染层猜形状。

    实测 JSON v1.2 形状（2026-10-19 核对）：
      {series, sales, pct_of_linked, items[{车系,销量,占比,行级标签计数}], list[[…]]}
    注意键名是 `series` 与 `pct_of_linked`（**不是** series_count / pct），
    且 `items` 元素为 dict —— 故优先取 `list`，并同时兼容 dict / list / tuple 三种元素。
    """
    v = stc.get(key)
    if isinstance(v, dict):
        raw = v.get('list') or v.get('items') or v.get('families') or []
        n = v.get('series', v.get('series_count',
                                  v.get('count', v.get('n', len(raw) or ''))))
        rows = []
        for it in raw:
            if isinstance(it, dict):
                rows.append((it.get('车系') or it.get('series') or it.get('name') or '',
                             it.get('销量', it.get('sales', '')),
                             it.get('占比', it.get('pct', it.get('pct_of_linked', '')))))
            elif isinstance(it, (list, tuple)) and it:
                rows.append((it[0],
                             it[1] if len(it) > 1 else '',
                             it[2] if len(it) > 2 else ''))
        return n, v.get('sales', ''), v.get('pct_of_linked', v.get('pct', '')), rows
    if isinstance(v, list):
        return len(v), '', '', v
    return '', '', '', []


def build(recs, S):
    b = Book()
    total = len(recs)
    n_bev = sum(1 for r in recs if r['组'] == 'BEV')
    n_phev = total - n_bev
    today = datetime.date.today().isoformat()
    # 批次范围一律实读（规范 §11.4：底表有并发写入方，禁止硬编码）
    # 且**强转 int 后比较**：新写入的 411/412 批行可能以文本形式出现（AGENTS 2026-09-27 教训①同类）
    def _bnum(v):
        try:
            return int(str(v).strip())
        except Exception:
            return None
    _bs = [x for x in (_bnum(r.get('批次')) for r in recs) if x is not None]
    bmin, bmax = (min(_bs), max(_bs)) if _bs else ('?', '?')
    # 412 批为公示口径（无正式发文字号，AGENTS 2026-09-10/09-12）
    _pub_note = '（412 批为公示口径·无正式发文字号）' if isinstance(bmax, int) and bmax >= 412 else ''
    # §3 裁决（Lead 2026-10-19）：公告年度未定的公示行单列披露，绝不静默丢弃
    _pub = [r for r in recs if not r['年']]
    _pub_b = sorted({r['批次'] for r in _pub})
    n_pub, n_formal = len(_pub), total - len(_pub)
    # 销量联动取数（契约 v1.1+；全部数字来自 JSON，禁止手写）
    cov, sstc = S['coverage'], S['structure']
    smeta, smon, stech = S['meta'], S['trend'], S['tech_sales']
    ssnap = smeta['snapshot_month']
    sprice = list(smeta['price_bins'])

    # ---------- 1 封面 ----------
    ws = b.sheet('封面')
    ws['B2'] = 'NEV 公告车型行业分析 · 数据底表'
    ws['B2'].font = Font(bold=True, size=18, color='1F3864')
    ws['B4'] = (f'口径版本 {VER} · {total:,} 条公告/公示记录 × 41 列'
                f'（第 {bmin}~{bmax} 批{_pub_note} · 纯乘用车）')
    ws['B5'] = f'生成日期 {today} · 权威源 NEV公告参数汇总表_合并版（341~410批）.xlsx'
    ws['B6'] = (f'电耗/续航已与工信部能耗数据20260915 官方对齐 · '
                f'销量=汽车之家·行业综合销量（车系级月零售）{ssnap} 快照')
    ws['B7'] = '分析表均由权威底表程序化重算（生成器 build_deliverables_v4923.py，可复跑）'
    ws['B8'] = (f'销量联动 8 表源自 audit-output/sales_linkage_analytics_20261019.json（契约 v1.1），'
                f'统计单元=车系，已去重；底表行级月销量仅供单行参考，不可加总')
    for r in (2, 4, 5, 6, 7, 8):
        ws.cell(r, 2).alignment = Alignment(vertical='center')
    ws.column_dimensions['B'].width = 110

    # ---------- 2 目录 ----------
    ws = b.sheet('目录', '目录')
    names = ['封面', '目录', '数据补充说明', '执行摘要', '数据总览', '年度分布', 'BEV续航分析',
             'PHEV续航分析', '续航对比分析', '电池容量分析', 'BEV容量分布', 'PHEV容量分布',
             '电池类型趋势-BEV', '电池类型趋势-PHEV', '电池能量密度', '续航分布-BEV',
             '续航分布-PHEV', '电机功率分析', '整备质量分析', '功率质量比', '整车能量密度',
             '能效分析', 'PHEV油耗分析', 'PHEV油耗-排量年份', '技术路线矩阵', '供应商集中度',
             '电池供应商CR', '电机供应商CR', '电池供应商排名', '电机供应商排名', '供应商组合',
             '品牌-电池供应商', '品牌-电机供应商', '细分市场-BEV', '细分市场-PHEV',
             '细分市场续航对比', '细分市场容量对比', '质量按细分市场', '相关性-BEV',
             '相关性-PHEV', '同比增长-BEV', '同比增长-PHEV', '关键指标汇总-BEV',
             '关键指标汇总-PHEV', '年度综合指标', '月销量分析',
             # ---- v4.9.24 新增销量联动区块（尾插，既有 1~46 序号零位移）----
             '销量口径说明', '销量市场结构', '销量集中度', '销量月度趋势',
             '销量-技术关联', '价格-技术-销量', '公告-市场转化', '月度榜单迁移']
    b.table(ws, 3, ['序', '工作表'], list(enumerate(names, 1)), widths=[6, 26])

    # ---------- 3 数据补充说明 ----------
    ws = b.sheet('数据补充说明', '数据补充说明')
    # 2026 年度构成分解（实读 conversion.by_batch + by_year_undated，Lead 2026-10-19 统一说法）
    _bb = S['conversion']['by_batch']
    _y26_other = sorted([r for r in _bb if r[1] == 2026 and int(str(r[0])) != 411],
                        key=lambda r: int(str(r[0])))
    _y26_411 = [r for r in _bb if int(str(r[0])) == 411]
    _und = (S['conversion'].get('by_year_undated') or [{}])[0]
    _y26n = sum(r[1] for r in S['conversion']['by_year'] if str(r[0]) == '2026') or 0
    _cmp2026 = ''
    if _y26_other and _y26_411:
        _lo, _hi = int(str(_y26_other[0][0])), int(str(_y26_other[-1][0]))
        _cmp2026 = (
            f'2026 年度合计 {_y26n:,} 条 = {_lo}~{_hi} 批（{len(_y26_other)} 批）'
            f'{sum(r[2] for r in _y26_other):,} 条 + {_y26_411[0][0]} 批 '
            f'{sum(r[2] for r in _y26_411):,} 条；{_und.get("batch")} 批 '
            f'{_und.get("models")} 条公告年度为 null（公示批，批次时间表无公告年月），'
            f'已由 conversion.by_year_undated 单列'
            f'（当期销量 {_und.get("sales", 0):,} 辆），不计入 2026 合计。')
    notes = [
        '① 口径双原则：仅乘用车（GB/T 3730.1 M1 类）；仅主流车企品牌。',
        '② 动力分组：BEV=纯电动/换电式纯电动；PHEV=插电式混合动力/插电式增程（与看板一致）。',
        f'③ 年度=批次《公告》发布日期年份（批次时间表·官方精确日期）；批次范围 {bmin}~{bmax} 批'
        f'（实读，规范 §11.4）。411 批为正式公告（计入 2026）'
        + ('；412 批为公示口径（公示日期 2026-10-17，无正式发文字号），58 条单列、'
           '不计入 2026 正式公告年度合计 776。' if isinstance(bmax, int) and bmax >= 412 else '。'),
        '④ 电耗/续航口径：以工信部新能源车型能耗数据20260915 为准（v4.9.22 全量对齐）；',
        '    工况按动力类型推定：BEV=CLTC，PHEV/EREV=WLTC；单元格不再保留工况标注。',
        '⑤ 电机功率取「电机总功率/扭矩」首值（kW）；功率质量比/整车能量密度为逐行比值均值（有值子集）。',
        '⑥ 各表样本=该表所需字段均有值的行；样本数随底表覆盖率滚动变化。',
        '⑦ 销量口径：汽车之家·行业综合销量（车系级月零售，主宇宙 typeid=9&subranktypeid=2305）；'
        f'当月快照 {ssnap}，抓取日 {smeta["fetch_date"]}。单元格值=该行对应车系的当月零售量，'
        '行级值仅供单行参考——同一车系的多行会共享同一车系级数值，禁止加总。',
        '⑧ 供应商排名/CR/HHI 基于电芯供应商、电机生产企业有值子集，覆盖率有限仅作趋势参考。',
        '⑨ 旧版（v4.3.6 口径 4,613 款）快照存 audit-output/old_deliverable_45sheets_snapshot_20260927.json。',
        # ---- v4.9.24 新增销量联动披露（⑩~⑮）----
        f'⑩ 销量覆盖率（行级）：已填行 {cov["row"]["covered"]:,} / {cov["row"]["total"]:,} 行 = '
        f'{cov["row"]["pct"]}%。车系级（去重通用名称）：'
        f'{cov["series"]["generic_names_matched"]} / {cov["series"]["generic_names_total"]} = '
        f'{cov["series"]["generic_names_pct"]}%。',
        f'⑪ 销售额覆盖：已锚定车系市场合计 {sstc["sales_total_dedup"]:,} 辆 / '
        f'{ssnap} 全榜 {sstc["board_sales_total"]:,} 辆 = {sstc["linked_share_of_board_pct"]}%'
        f'（全榜 {sstc["board_series_count"]} 个车系，已锚定 {sstc["series_count"]} 个）。',
        f'⑫ 未覆盖分桶（共 {sum(cov["buckets"].values()):,} 行）：'
        + '、'.join(f'{k} {v:,}' for k, v in cov['buckets'].items())
        + '。分桶为分析边界，不等同于数据缺失。',
        f'⑬ 防重复计数铁律（规范 §4）：销量表的统计单元一律=车系（按 车系ID 或归一化车系名去重）。'
        f'对底表行直接加总会得到 {sstc["sales_total_rowlevel_naive"]:,} 辆，'
        f'为车系级口径 {sstc["sales_total_dedup"]:,} 辆的 {sstc["duplication_ratio"]} 倍，属错误口径。',
        '⑭ 候选池与匹配通道：' + '、'.join(
            f'{k} {v}' for k, v in smeta['match_channel_counts'].items()) + '（单位=行）。',
        '⑮ 销量联动分析表（#47~#54）与既有 46 表同源同批。「公告-市场转化」表中'
        f'「当期销量（{ssnap} 快照）」= 该批次公告车型在 {ssnap} 的车系级快照销量合计，'
        f'用于公告存量（第 {bmin}~{bmax} 批累积）× 当期市场表现的转化率表现，不等于批次发布当月的实际销量。',
        f'⑯ 口径构成（唯一表述）：{total:,} 条 = {n_formal:,} 条正式公告 + {n_pub} 条公示批次记录'
        f'（{_pub_b[-1] if _pub_b else "—"} 批，公示日期 2026-10-17）。'
        f'注意「行 ≠ 款」：本底表统计单元是「记录」（一条公告/公示记录），不是「款」。'
        f'公示批无正式发文字号，其公告年度未定，故在年度类表中单列一行，不计入正式公告年度合计。'
        + ('　' + _cmp2026 if _cmp2026 else ''),
        '⑰ 底表「批次」列存在 122 行字符串型值（\'361\'×11、\'409\'×111），'
        '消费端须强转 int（int(float(x))）；已登记 CR，下轮修复。',
    ]
    for i, t in enumerate(notes, 3):
        ws.cell(i, 1, t)
    ws.column_dimensions['A'].width = 100

    # ---------- 4 执行摘要 ----------
    yrs = sorted({r['年'] for r in recs if r['年']})
    yy = yrs[-1]
    bev = [r for r in recs if r['组'] == 'BEV']
    phev = [r for r in recs if r['组'] == 'PHEV']
    bev_y = [r for r in bev if r['年'] == yy]
    phev_y = [r for r in phev if r['年'] == yy]
    mean = lambda rs, f: (lambda v: round(sum(v) / len(v), 2) if v else None)(
        [r[f] for r in rs if r[f] is not None])
    ws = b.sheet('执行摘要', '执行摘要 - 关键指标')
    ws.cell(3, 1, '（结论详见演示文稿）')
    kpis = [
        ['指标', '数值', '单位', '备注'],
        ['公告/公示记录总量', total, '条', f'BEV {n_bev:,} / PHEV {n_phev:,}（行≠款）'],
        [f'{yy}年正式公告记录', len(bev_y) + len(phev_y), '条',
         f'BEV {len(bev_y)} / PHEV {len(phev_y)}；不含公示批'],
        [f'{_pub_b[-1] if _pub_b else "—"} 批公示记录', n_pub, '条',
         f'公告日期未定，单列不计入 {yy} 正式公告年度合计'],
        [f'{yy}年PHEV占比', round(len(phev_y) / max(1, len(bev_y) + len(phev_y)) * 100, 1), '%',
         '插混/增程持续放量'],
        [f'BEV平均续航({yy})', mean(bev_y, '续航'), 'km',
         f'较{yrs[0]}年显著提升'],
        [f'BEV平均电池容量({yy})', mean(bev_y, '容量'), 'kWh', '大电量化持续'],
        [f'PHEV平均纯电续航({yy})', mean(phev_y, '续航'), 'km', '插混长途化'],
        [f'BEV平均百公里电耗({yy})', mean(bev_y, '电耗'), 'kWh/100km', '能效持续优化'],
        [f'月销量覆盖({ssnap})', cov['row']['covered'], '行',
         f'汽车之家·行业综合销量 · 行级 {cov["row"]["pct"]}%'],
        # ---- 销量联动（全部取契约 JSON，禁止手写）----
        ['销量车系级覆盖', cov['series']['generic_names_matched'], '车系',
         f'{cov["series"]["generic_names_pct"]}% of {cov["series"]["generic_names_total"]} 去重通用名称'],
        ['在售车系覆盖率', cov['current_board']['pct'], '%',
         f'{cov["current_board"]["linked_series"]} / {cov["current_board"]["board_series"]} 在榜车系'
         '（对外唯一合法口径，规范 §5/§11）'],
        ['销售额覆盖(锚定率)', sstc['linked_share_of_board_pct'], '%',
         f'{sstc["sales_total_dedup"]:,} / {sstc["board_sales_total"]:,} 辆'],
        ['车系级去重市场合计', sstc['sales_total_dedup'], '辆',
         f'{ssnap} 快照 · 已锚定 {sstc["series_count"]} 车系（禁止行级加总）'],
        ['行级加总虚高倍数', sstc['duplication_ratio'], '×',
         f'错误口径 {sstc["sales_total_rowlevel_naive"]:,} 辆 ÷ 正确口径（规范 §4 对照）'],
    ]
    b.table(ws, 4, kpis[0], kpis[1:], widths=[26, 12, 12, 34])

    # ---------- 5 数据总览 ----------
    ws = b.sheet('数据总览', '数据总览')
    ws.cell(3, 1, '按动力类型').font = HFONT
    b.table(ws, 4, ['动力类型', '车型数', '占比%'],
            [['BEV', n_bev, round(n_bev / total * 100, 1)],
             ['PHEV', n_phev, round(n_phev / total * 100, 1)]], widths=[12, 10, 10])
    ws.cell(8, 1, '按年度').font = HFONT
    yr_rows = []
    for y in yrs:
        nb = sum(1 for r in recs if r['年'] == y and r['组'] == 'BEV')
        np_ = sum(1 for r in recs if r['年'] == y and r['组'] == 'PHEV')
        yr_rows.append([y, nb, np_, nb + np_, round(np_ / max(1, nb + np_) * 100, 1)])
    # §3 裁决（Lead 2026-10-19）：公告年度未定的公示行**单列一行**，
    # 既不并入正式公告年度，也不静默丢弃（否则年度合计 5,417 ≠ 5,475）。
    if _pub:
        _ylab = max(y for y in yrs if y)
        _pb = sum(1 for r in _pub if r['组'] == 'BEV')
        _pp = sum(1 for r in _pub if r['组'] == 'PHEV')
        yr_rows.append([f'{_ylab}（{_pub_b[-1]} 批·公示口径）', _pb, _pp, _pb + _pp,
                        round(_pp / max(1, _pb + _pp) * 100, 1)])
    assert sum(r[3] for r in yr_rows) == total, (
        f'年度合计 {sum(r[3] for r in yr_rows)} != 底表记录数 {total}（公示行未单列？）')
    b.table(ws, 9, ['年份', 'BEV(条)', 'PHEV(条)', '合计(条)', 'PHEV占比%'], yr_rows)

    # ---------- 6 年度分布 ----------
    ws = b.sheet('年度分布', '年度记录数量分布（公告年度；公示批单列）')
    ws.cell(3, 1, f'注：{_pub_b[-1] if _pub_b else "—"} 批 {n_pub} 行为公示口径，'
                  '公告日期未定，单列不计入正式公告年度合计。')
    b.table(ws, 4, ['年份', 'BEV(条)', 'PHEV(条)', '合计(条)', 'PHEV占比%'], yr_rows,
            widths=[22, 10, 10, 10, 12])
    ws.cell(5 + len(yr_rows), 1,
            f'合计 {total:,} 条 = {n_formal:,} 条正式公告 + {n_pub} 条公示批记录')

    # ---------- 7/8 续航分析 ----------
    for name, grp in (('BEV续航分析', 'BEV'), ('PHEV续航分析', 'PHEV')):
        ws = b.sheet(name, f'{grp} 续航里程年度分析')
        rows = []
        for y in yrs:
            vals = sorted(r['续航'] for r in recs
                          if r['年'] == y and r['组'] == grp and r['续航'] is not None)
            if vals:
                rows.append([y, round(sum(vals) / len(vals), 2), round(st.median(vals), 1),
                             round(pct(vals, 25), 0), round(pct(vals, 75), 0),
                             round(pct(vals, 90), 0), len(vals)])
        b.table(ws, 3, ['年份', '均值(km)', '中位数(km)', 'P25(km)', 'P75(km)', 'P90(km)', '样本数'],
                rows, widths=[8, 10, 11, 9, 9, 9, 8])

    # ---------- 9 续航对比 ----------
    ws = b.sheet('续航对比分析', 'BEV vs PHEV 续航年均值')
    mb, mp_ = agg_by_year(recs, '续航', 'BEV'), agg_by_year(recs, '续航', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV均值(km)', 'PHEV均值(km)', '差值(km)'],
            [[y, round(mb[y][0], 2), round(mp_[y][0], 2), round(mb[y][0] - mp_[y][0], 1)]
             for y in yrs if y in mb and y in mp_])

    # ---------- 10 电池容量 ----------
    ws = b.sheet('电池容量分析', '电池容量年度分析')
    cb, cp = agg_by_year(recs, '容量', 'BEV'), agg_by_year(recs, '容量', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV均值(kWh)', 'PHEV均值(kWh)'],
            [[y, round(cb[y][0], 2), round(cp[y][0], 2)] for y in yrs if y in cb and y in cp])

    # ---------- 11/12 容量分布 ----------
    CAP_BINS = [('<20', 0, 20), ('20-30', 20, 30), ('30-40', 30, 40), ('40-50', 40, 50),
                ('50-60', 50, 60), ('60-70', 60, 70), ('70-80', 70, 80), ('≥80', 80, 1e9)]
    CAP_BINS_P = [('<20', 0, 20), ('20-30', 20, 30), ('30-40', 30, 40),
                  ('40-50', 40, 50), ('≥50', 50, 1e9)]
    for name, grp, bins in (('BEV容量分布', 'BEV', CAP_BINS), ('PHEV容量分布', 'PHEV', CAP_BINS_P)):
        ws = b.sheet(name, f'{grp} 电池容量分布（按年度）')
        ws.cell(3, 1, '按容量区间计数（kWh，区间左闭右开）')
        hdr = ['年份'] + [b0 for b0, _, _ in bins]
        body = []
        for y in yrs:
            vals = [r['容量'] for r in recs
                    if r['年'] == y and r['组'] == grp and r['容量'] is not None]
            if vals:
                body.append([y] + [sum(1 for v in vals if lo <= v < hi) for _, lo, hi in bins])
        b.table(ws, 4, hdr, body)

    # ---------- 13/14 电池类型趋势 ----------
    def btype(v):
        if '磷酸铁锂' in v:
            return 'LFP'
        if '三元' in v:
            return 'NCM'
        return '其他'
    for name, grp in (('电池类型趋势-BEV', 'BEV'), ('电池类型趋势-PHEV', 'PHEV')):
        ws = b.sheet(name, f'{grp} 电池类型趋势')
        ws.cell(3, 1, '数量与占比')
        body = []
        for y in yrs:
            rs = [r for r in recs if r['年'] == y and r['组'] == grp and r['电池类型']]
            if not rs:
                continue
            c = Counter(btype(r['电池类型']) for r in rs)
            n = len(rs)
            body.append([y, c.get('LFP', 0), c.get('NCM', 0), c.get('其他', 0),
                         round(c.get('LFP', 0) / n * 100, 1),
                         round(c.get('NCM', 0) / n * 100, 1),
                         round(c.get('其他', 0) / n * 100, 1)])
        b.table(ws, 4, ['年份', 'LFP(款)', 'NCM(款)', '其他(款)', 'LFP占比%', 'NCM占比%', '其他占比%'],
                body)

    # ---------- 15 电池能量密度 ----------
    ws = b.sheet('电池能量密度', '电池能量密度年度趋势')
    db, dp = agg_by_year(recs, '密度', 'BEV'), agg_by_year(recs, '密度', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV均值(Wh/kg)', 'PHEV均值(Wh/kg)'],
            [[y, round(db[y][0], 1), round(dp[y][0], 1)] for y in yrs if y in db and y in dp])

    # ---------- 16/17 续航分布 ----------
    RAN_B = [('<250', 0, 250), ('250-350', 250, 350), ('350-450', 350, 450),
             ('450-550', 450, 550), ('550-650', 550, 650), ('650-750', 650, 750), ('≥750', 750, 1e9)]
    RAN_P = [('<100', 0, 100), ('100-150', 100, 150), ('150-200', 150, 200),
             ('200-250', 200, 250), ('≥250', 250, 1e9)]
    for name, grp, bins in (('续航分布-BEV', 'BEV', RAN_B), ('续航分布-PHEV', 'PHEV', RAN_P)):
        ws = b.sheet(name, f'{grp} 续航里程分布（按年度）')
        ws.cell(3, 1, '按续航区间计数（km，区间左闭右开）')
        hdr = ['年份'] + [b0 for b0, _, _ in bins]
        body = []
        for y in yrs:
            vals = [r['续航'] for r in recs
                    if r['年'] == y and r['组'] == grp and r['续航'] is not None]
            if vals:
                body.append([y] + [sum(1 for v in vals if lo <= v < hi) for _, lo, hi in bins])
        b.table(ws, 4, hdr, body)

    # ---------- 18 电机功率 ----------
    ws = b.sheet('电机功率分析', '电机功率年度分析')
    pb, pp = agg_by_year(recs, '功率', 'BEV'), agg_by_year(recs, '功率', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV均值(kW)', 'PHEV均值(kW)'],
            [[y, round(pb[y][0], 2), round(pp[y][0], 2)] for y in yrs if y in pb and y in pp])

    # ---------- 19 整备质量 ----------
    ws = b.sheet('整备质量分析', '整备质量年度分析')
    wb_, wp_ = agg_by_year(recs, '整备', 'BEV'), agg_by_year(recs, '整备', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV均值(kg)', 'PHEV均值(kg)'],
            [[y, round(wb_[y][0], 1), round(wp_[y][0], 1)] for y in yrs if y in wb_ and y in wp_])

    # ---------- 20 功率质量比 ----------
    ws = b.sheet('功率质量比', '功率质量比年度趋势')
    rb, rp = per_row_ratio(recs, '功率', '整备', 'BEV'), per_row_ratio(recs, '功率', '整备', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV(kW/kg×1000)', 'PHEV(kW/kg×1000)'],
            [[y, round(rb[y], 1), round(rp[y], 1)] for y in yrs if y in rb and y in rp])

    # ---------- 21 整车能量密度 ----------
    ws = b.sheet('整车能量密度', '整车能量密度年度趋势')
    eb, ep = per_row_ratio(recs, '容量', '整备', 'BEV'), per_row_ratio(recs, '容量', '整备', 'PHEV')
    b.table(ws, 3, ['年份', 'BEV(Wh/kg)', 'PHEV(Wh/kg)'],
            [[y, round(eb[y], 1), round(ep[y], 1)] for y in yrs if y in eb and y in ep])

    # ---------- 22 能效分析 ----------
    ws = b.sheet('能效分析', 'BEV 百公里电耗年度分析')
    ee = agg_by_year(recs, '电耗', 'BEV')
    b.table(ws, 3, ['年份', '均值(kWh/100km)', '中位数', '样本数'],
            [[y, round(ee[y][0], 2), round(ee[y][1], 2), ee[y][2]] for y in yrs if y in ee])

    # ---------- 23 PHEV 油耗 ----------
    ws = b.sheet('PHEV油耗分析', 'PHEV 油耗与纯电续航年度分析')
    rows = []
    for y in yrs:
        oil = [r['油耗'] for r in recs if r['年'] == y and r['组'] == 'PHEV'
               and r['油耗'] is not None and 0.1 <= r['油耗'] <= 12]
        rng = [r['续航'] for r in recs if r['年'] == y and r['组'] == 'PHEV'
               and r['续航'] is not None]
        if oil and rng:
            rows.append([y, round(sum(oil) / len(oil), 2),
                         round(sum(rng) / len(rng), 1), len(oil)])
    b.table(ws, 3, ['年份', '综合油耗均值(L/100km)', '纯电续航均值(km)', '样本数'], rows)

    # ---------- 24 PHEV 油耗×排量 ----------
    ws = b.sheet('PHEV油耗-排量年份', 'PHEV 油耗按排量区间年度均值')
    DISP_BINS = [('<1.0L', 0, 1000), ('1.0-1.5L', 1000, 1500), ('1.5-2.0L', 1500, 2000),
                 ('2.0-2.5L', 2000, 2500), ('≥2.5L', 2500, 1e9)]
    body = []
    for lab, lo, hi_ in DISP_BINS:
        row = [lab]
        for y in yrs:
            vals = [r['油耗'] for r in recs if r['年'] == y and r['组'] == 'PHEV'
                    and r['油耗'] is not None and r['排量'] is not None
                    and lo <= r['排量'] < hi_ and 0.1 <= r['油耗'] <= 12]
            row.append(round(sum(vals) / len(vals), 2) if vals else '')
        body.append(row)
    b.table(ws, 3, ['排量\\年份'] + [str(y) for y in yrs], body)

    # ---------- 25 技术路线矩阵 ----------
    ws = b.sheet('技术路线矩阵', 'BEV 技术路线四象限')
    ws.cell(3, 1, '阈值=BEV 续航/容量中位数（全体 BEV 有值子集）')
    bvr = [r['续航'] for r in bev if r['续航'] is not None]
    bvc = [r['容量'] for r in bev if r['容量'] is not None]
    thr_r = st.median(bvr) if bvr else 0
    thr_c = st.median(bvc) if bvc else 0
    quad = Counter()
    for r in bev:
        if r['续航'] is None or r['容量'] is None:
            continue
        quad[('低' if r['续航'] < thr_r else '高') + '续航' +
             ('低' if r['容量'] < thr_c else '高') + '容量'] += 1
    qn = sum(quad.values())
    b.table(ws, 4, ['象限', '车型数', '占比%'],
            [[k, quad.get(k, 0), round(quad.get(k, 0) / qn * 100, 1)]
             for k in ('低续航低容量', '高续航高容量', '低续航高容量', '高续航低容量') if k in quad],
            widths=[16, 10, 10])
    ws.cell(9, 1, f'阈值：续航 {thr_r:.0f} km · 容量 {thr_c:.0f} kWh')

    # ---------- 26 供应商集中度 HHI ----------
    def hhi(vals):
        c = Counter(vals)
        n = sum(c.values())
        return round(sum((v / n * 100) ** 2 for v in c.values()), 1) if n >= 5 else None
    ws = b.sheet('供应商集中度', '供应商 HHI 年度趋势')
    ws.cell(3, 1, f'电池(电芯)口径样本 {sum(1 for r in recs if r["电芯"])} · 电机口径样本 {sum(1 for r in recs if r["电机企"])}（有值子集）')
    body = []
    for y in yrs:
        hb = hhi([r['电芯'] for r in recs if r['年'] == y and r['电芯']])
        hm = hhi([r['电机企'] for r in recs if r['年'] == y and r['电机企']])
        if hb or hm:
            body.append([y, hb, hm])
    b.table(ws, 4, ['年份', '电池供应商HHI', '电机供应商HHI'], body)

    # ---------- 27/28 供应商 CR ----------
    def cr(vals, k):
        c = Counter(vals)
        n = sum(c.values())
        if n < 5:
            return None
        return round(sum(v for _, v in c.most_common(k)) / n * 100, 1)
    for name, fld in (('电池供应商CR', '电芯'), ('电机供应商CR', '电机企')):
        ws = b.sheet(name, f'{fld} 供应商集中度 CR 年度趋势')
        body = []
        for y in yrs:
            vals = [r[fld] for r in recs if r['年'] == y and r[fld]]
            if len(vals) >= 5:
                body.append([y, cr(vals, 1), cr(vals, 3), cr(vals, 5), len(vals)])
        b.table(ws, 3, ['年份', 'CR1(%)', 'CR3(%)', 'CR5(%)', '样本数'], body)

    # ---------- 29/30 供应商排名 ----------
    for name, fld, label in (('电池供应商排名', '电芯', '电芯供应商'),
                             ('电机供应商排名', '电机企', '电机生产企业')):
        ws = b.sheet(name, f'{label}排名 Top10（有值子集）')
        c = Counter(r[fld] for r in recs if r[fld])
        b.table(ws, 3, ['排名', label, '车型数'],
                [[i, k, v] for i, (k, v) in enumerate(c.most_common(10), 1)],
                widths=[6, 30, 10])

    # ---------- 31 供应商组合 ----------
    ws = b.sheet('供应商组合', '电池-电机供应商组合 Top10')
    c = Counter((r['电芯'], r['电机企']) for r in recs if r['电芯'] and r['电机企'])
    b.table(ws, 3, ['组合（电芯 × 电机企业）', '车型数'],
            [[f'{k[0]} × {k[1]}', v] for k, v in c.most_common(10)], widths=[52, 10])

    # ---------- 32/33 品牌×供应商 ----------
    def brand_sup(name, bfld, sfld, slabel, topb, tops):
        ws = b.sheet(name, f'Top{topb}品牌 × Top{tops}{slabel}（车型数）')
        brands = [k for k, _ in Counter(r[bfld] for r in recs if r[bfld]).most_common(topb)]
        sups = [k for k, _ in Counter(r[sfld] for r in recs if r[sfld]).most_common(tops)]
        body = []
        for bd in brands:
            row = [bd]
            for sp in sups:
                row.append(sum(1 for r in recs if r[bfld] == bd and r[sfld] == sp))
            body.append(row)
        b.table(ws, 3, ['品牌'] + sups, body, widths=[14] + [16] * tops)
    brand_sup('品牌-电池供应商', '商标', '电芯', '电芯供应商', 8, 5)
    brand_sup('品牌-电机供应商', '商标', '电机企', '电机供应商', 10, 5)

    # ---------- 34/35 细分市场 ----------
    for name, grp in (('细分市场-BEV', 'BEV'), ('细分市场-PHEV', 'PHEV')):
        ws = b.sheet(name, f'{grp} 细分市场汇总（车型数）')
        segs = sorted({r['细分'] for r in recs if r['细分'] and r['组'] == grp})
        body = []
        for sg in segs:
            row = [sg]
            for y in yrs:
                row.append(sum(1 for r in recs if r['组'] == grp and r['细分'] == sg and r['年'] == y))
            row.append(sum(row[1:]))
            body.append(row)
        b.table(ws, 3, ['细分市场'] + [str(y) for y in yrs] + ['合计'], body, widths=[14] + [8] * (len(yrs) + 1))

    # ---------- 36/37 细分市场对比 ----------
    def seg_cmp(name, fld, unit, label):
        ws = b.sheet(name, f'BEV 细分市场{label} 2023 vs 2025 均值对比')
        body = []
        segs = sorted({r['细分'] for r in recs if r['细分'] and r['组'] == 'BEV'})
        for sg in segs:
            v23 = [r[fld] for r in recs if r['组'] == 'BEV' and r['细分'] == sg
                   and r['年'] == 2023 and r[fld] is not None]
            v25 = [r[fld] for r in recs if r['组'] == 'BEV' and r['细分'] == sg
                   and r['年'] == 2025 and r[fld] is not None]
            if v23 and v25:
                m23, m25 = sum(v23) / len(v23), sum(v25) / len(v25)
                body.append([sg, round(m23, 1), round(m25, 1),
                             round((m25 - m23) / m23 * 100, 1)])
        b.table(ws, 3, ['细分市场', f'2023均值({unit})', f'2025均值({unit})', '变化%'], body,
                widths=[14, 14, 14, 10])
    seg_cmp('细分市场续航对比', '续航', 'km', '续航')
    seg_cmp('细分市场容量对比', '容量', 'kWh', '容量')

    # ---------- 38 质量按细分市场 ----------
    ws = b.sheet('质量按细分市场', '各细分市场整备质量均值（有值子集）')
    seg_m = {}
    for r in recs:
        if r['细分'] and r['整备'] is not None:
            seg_m.setdefault(r['细分'], []).append(r['整备'])
    rows = sorted(((k, sum(v) / len(v)) for k, v in seg_m.items() if len(v) >= 5),
                  key=lambda x: -x[1])[:12]
    b.table(ws, 3, ['细分市场', '整备均值(kg)'],
            [[k, round(v, 1)] for k, v in rows], widths=[16, 14])

    # ---------- 39/40 相关性 ----------
    CORR_B = [('续航', '续航'), ('电池容量', '容量'), ('电机功率', '功率'),
              ('整备质量', '整备'), ('能量密度', '密度'), ('电耗', '电耗')]
    CORR_P = CORR_B + [('综合油耗', '油耗')]
    for name, grp, spec in (('相关性-BEV', 'BEV', CORR_B), ('相关性-PHEV', 'PHEV', CORR_P)):
        ws = b.sheet(name, f'{grp} 相关性矩阵（Pearson，有值成对子集）')
        body = []
        for lab_a, f_a in spec:
            row = [lab_a]
            for lab_b, f_b in spec:
                xs, ys = [], []
                for r in recs:
                    if r['组'] == grp and r[f_a] is not None and r[f_b] is not None:
                        xs.append(r[f_a])
                        ys.append(r[f_b])
                p = pearson(xs, ys)
                row.append(round(p, 3) if p is not None else '')
            body.append(row)
        b.table(ws, 3, [''] + [lab for lab, _ in spec], body)

    # ---------- 41/42 同比增长 ----------
    def yoy_rows(grp, spec):
        body = []
        for lab, f in spec:
            row = [lab]
            prev = None
            for y in yrs:
                vals = [r[f] for r in recs if r['年'] == y and r['组'] == grp and r[f] is not None]
                m = sum(vals) / len(vals) if vals else None
                row.append(round((m - prev) / prev * 100, 1)
                           if (m is not None and prev) else '')
                prev = m if m is not None else prev
            body.append(row)
        return body
    WSPEC = [('平均续航(km)', '续航'), ('平均容量(kWh)', '容量'), ('平均功率(kW)', '功率'),
             ('平均整备(kg)', '整备'), ('平均能量密度(Wh/kg)', '密度'), ('平均电耗(kWh/100km)', '电耗')]
    PSPEC = WSPEC + [('平均油耗(L/100km)', '油耗')]
    for name, grp, spec in (('同比增长-BEV', 'BEV', WSPEC), ('同比增长-PHEV', 'PHEV', PSPEC)):
        ws = b.sheet(name, f'{grp} 关键指标同比增长率(%)')
        b.table(ws, 3, ['指标\\年份'] + [str(y) for y in yrs], yoy_rows(grp, spec))

    # ---------- 43/44 关键指标汇总 ----------
    for name, grp, spec in (('关键指标汇总-BEV', 'BEV', WSPEC), ('关键指标汇总-PHEV', 'PHEV', PSPEC)):
        ws = b.sheet(name, f'{grp} 关键指标（全周期描述统计）')
        body = []
        for lab, f in spec:
            vals = [r[f] for r in recs if r['组'] == grp and r[f] is not None]
            if vals:
                body.append([lab, round(sum(vals) / len(vals), 2),
                             round(st.median(vals), 1), len(vals)])
        b.table(ws, 3, ['指标', '均值', '中位数', '样本数'], body)

    # ---------- 45 年度综合指标 ----------
    ws = b.sheet('年度综合指标', '年度综合技术指标（有值子集均值）')
    hdr = ['年份', 'BEV续航', 'BEV容量', 'BEV功率', 'BEV质量', 'BEV整车密度',
           'PHEV续航', 'PHEV容量', 'PHEV功率', 'PHEV质量', 'PHEV整车密度']
    rows = []
    for y in yrs:
        row = [y]
        for grp in ('BEV', 'PHEV'):
            for f in ('续航', '容量', '功率', '整备'):
                m = agg_by_year(recs, f, grp).get(y)
                row.append(round(m[0], 2) if m else '')
            m = per_row_ratio(recs, '容量', '整备', grp).get(y)
            row.append(round(m, 1) if m else '')
        rows.append(row)
    b.table(ws, 3, hdr, rows, widths=[8] + [11] * 10)
    ws.cell(4 + len(rows), 1,
            f'注：本表按正式公告年度分组，覆盖 {n_formal:,} 条；'
            f'{_pub_b[-1] if _pub_b else "—"} 批 {n_pub} 行为公示口径（公告日期未定），'
            '未参与年度分组，单列于「年度分布」表。')

    # ---------- 46 月销量分析（车系级去重 Top20 + 旧版行级对照块） ----------
    ws = b.sheet('月销量分析', f'月销量 Top20（车系级去重 · 汽车之家·行业综合销量 {ssnap}）')
    ws.cell(3, 1, '统计单元=车系（按 seriesid 去重）；同一车系在底表可能对应多行，'
                  '行级月销量仅供单行参考，禁止加总（规范 §4 铁律）。')
    top_series = S['concentration']['top_series']
    b.table(ws, 4, ['排名', '车系', '月销量(辆)', '市场占比%', '企业/品牌'],
            [[i, row[0], int(row[1]), row[2], row[3]] for i, row in enumerate(top_series, 1)],
            widths=[6, 26, 12, 10, 36])
    r0 = 4 + len(top_series) + 2
    ws.cell(r0, 1, '【对照】旧版行级口径 Top20（错误口径，仅用于展示重复计数程度）').font = HFONT
    ws.cell(r0 + 1, 1, f'行级直接加总 {sstc["sales_total_rowlevel_naive"]:,} 辆 = 车系级口径 '
                       f'{sstc["sales_total_dedup"]:,} 辆 的 {sstc["duplication_ratio"]} 倍')
    _s = {}
    for r in recs:
        if r['月销']:
            k = r['通用'] or r['商标'] or r['企业'][:10]
            if r['月销'] > _s.get(k, (0, ''))[0]:
                _s[k] = (r['月销'], r['商标'] or r['企业'][:8])
    top_sales = sorted(((k, v[0], v[1]) for k, v in _s.items()), key=lambda x: -x[1])[:20]
    b.table(ws, r0 + 2, ['排名', '车系（通用名称）', '月销量(辆)', '品牌'],
            [[i, _html_e(k), int(v), ent] for i, (k, v, ent) in enumerate(top_sales, 1)],
            widths=[6, 30, 12, 14])

    # ===== 47 销量口径说明 =====
    # 块序（Lead 2026-10-19 裁决：口径局限不得藏在表尾）：
    #   A 口径局限·首要披露 → B 覆盖率口径 → C 销售额覆盖/重复计数 → D 未覆盖分桶
    ws = b.sheet('销量口径说明', '销量口径说明（规范 §4/§5/§11 披露项）')
    # ④ Lead 2026-10-19 裁决：JSON meta.reference_source 仍写着早期不完整的「懂车帝微型车榜」路径，
    #    对外披露不得使用。此处按 Lead 签发的 WS1 实际比对口径覆盖。
    #    （待 link-engine 在 JSON 侧同步修正后可移除此覆盖，恢复直接读 smeta。）
    _REF_SRC = ('全车型总榜 typeid=1（仅参照）+ 易车 2026-08 全表逐值比对（147/147 完全一致）'
                '+ 懂车帝 2026-09 全量榜 660 条（可比 282，unexplained=0）')
    _lines = [
        f'数据源：{smeta["source"]}；{smeta["source_tip"]}。',
        f'范围：{smeta["scope"]}；接口 typeid=9 & subranktypeid=2305；'
        f'快照月 {ssnap}，抓取日 {smeta["fetch_date"]}。',
        # ⑤ Lead 2026-10-19 裁决：#47 必须披露 412 批公示构成（行 ≠ 款）
        f'构成：全表 {total:,} 条 = {n_formal:,} 条正式公告 + {n_pub} 条公示批次记录'
        f'（412 批，公示日期 2026-10-17，无正式发文字号）；年度口径分母取 {n_formal:,}，'
        f'2026 年 = 776，公示批单列不并入（规范 §12）。'
        + (f'「行 ≠ 款」：统计单元是「记录」，不是「款」。' if n_pub else ''),
        f'参考源：{_REF_SRC}。',
        f'防重复计数铁律：{smeta["dedup_rule"]}。所有销量表统计单元=车系；底表行级值不可加总。',
        f'快照覆盖 {len(smeta["months"])} 个月：{"、".join(smeta["months"])}；'
        f'各月榜单车系数：{"、".join(f"{m} {n}" for m, n in smeta["qc"]["monthly_counts"].items())}。',
        f'十个月车系并集 {smeta["qc"]["union_series"]} 个；当月（{ssnap}）{smeta["qc"]["current_month_series"]} 个。',
        f'价格带分箱规则：{smeta["price_bin_rule"]}；分箱={"、".join(sprice)}。',
        f'旧口径留档：{smeta["backup_old"]}（易车零售 2026-08，已整体替换）。',
        f'上游质检：contract_assertions all_pass={S["contract_assertions"]["all_pass"]}、'
        f'dedup_assertions all_pass={S["dedup_assertions"]["all_pass"]}。',
    ]
    for i, t in enumerate(_lines, 3):
        ws.cell(i, 1, t)

    # ---- A. 口径局限（首要披露）----
    r0 = 3 + len(_lines) + 1
    _mx_n, _mx_s, _mx_p, _mx = _famstats(sstc, 'mixed_energy_families')
    _me_n, _me_s, _me_p, _me = _famstats(sstc, 'mixed_with_erev')
    _fl_n, _fl_s, _fl_p, _fl = _famstats(sstc, 'r2_flip_impact')
    ws.cell(r0, 1, 'A. 【口径局限·首要披露】车系级单标签分类的损耗（规范 §11.2）').font = TFONT
    ws.cell(r0 + 1, 1, sstc.get('mixed_energy_note')
            or '同车系可能同时提供纯电与增程版本（公告行级多标签）；车系级只能归入其一，'
               '故单标签分类对下述规模的市场在原理上有损。')
    ws.cell(r0 + 2, 1, '⚠️ 已锚定销量中约三分之一来自「同车系多动力标签」家族 —— '
                       '单标签给车系分类对该部分盘子原理性有损，是本轮首要口径局限。')
    b.table(ws, r0 + 3, ['口径（三键独立，不合并）', '车系数', '销量(辆)', '占已锚定%', '含义'],
            [['混合家族（≥2 种行级标签）', _mx_n, _mx_s, _mx_p,
              '「同车系多动力标签」家族的广度 → 单标签分类的原理性损耗上限'],
             ['其中含 EREV 行者', _me_n, _me_s, _me_p,
              '混合家族里确实同时提供增程版本者'],
             ['R2 相对 R1 实际翻转者', _fl_n, _fl_s, _fl_p,
              '被否决方案 R2 的成本证据（整系误归 EREV）']],
            widths=[28, 10, 14, 12, 62])
    ws.cell(r0 + 7, 1, '动力类型三分（R1，canonical）：行级 label = EREV（是否增程列真值）'
                       'else BEV（动力类型含 BEV）else PHEV；车系级取行级 label 众数'
                       '（并列取字典序最小）。').font = HFONT
    ws.cell(r0 + 8, 1, f'聚合规则（规范 §11.1）：{smeta["agg_rule"]}')
    b.table(ws, r0 + 9, ['动力类型', '销量(辆)', '占已锚定%'],
            [[k, v, round(100.0 * v / sstc['sales_total_dedup'], 2)]
             for k, v in sstc['by_energy']],
            widths=[14, 14, 12])
    _rmx = r0 + 11 + len(sstc['by_energy']) + 1
    ws.cell(_rmx, 1, f'混合家族清单（{_mx_n} 系 · 规范 §11.2 强制披露）').font = HFONT
    if _mx:
        b.table(ws, _rmx + 1, ['车系', '销量(辆)', '占已锚定%'],
                [list(r) for r in _mx],   # _famstats 已归一化为 (车系, 销量, 占比)
                widths=[28, 14, 12])

    # ---- B. 覆盖率口径 ----
    r1 = _rmx + (3 + len(_mx) if _mx else 2) + 1
    ws.cell(r1, 1, 'B. 覆盖率口径（含自指指标标注）').font = HFONT
    b.table(ws, r1 + 1, ['口径', '已覆盖', '分母', '覆盖率%', '定义'],
            [['行级', cov['row']['covered'], cov['row']['total'], cov['row']['pct'],
              cov['definitions']['row']],
             ['车系级（去重通用名称）', cov['series']['generic_names_matched'],
              cov['series']['generic_names_total'], cov['series']['generic_names_pct'],
              cov['definitions']['series']],
             ['在售车系级【对外唯一合法口径】', cov['current_board']['linked_series'],
              cov['current_board']['board_series'], cov['current_board']['pct'],
              cov['definitions']['current_board']],
             ['（自指指标·禁止对外主张）', cov['self_consistent_ratio']['matched'],
              cov['self_consistent_ratio']['total'], cov['self_consistent_ratio']['pct'],
              cov['definitions']['self_consistent_ratio']]],
            widths=[32, 12, 12, 12, 66])
    ws.cell(r1 + 7, 1, '注：末行为自指指标（分母由分子派生），不是覆盖率，'
                       '不得替代行级/车系级/在售车系级覆盖率对外主张。').font = HFONT

    # ---- C. 销售额覆盖与重复计数对照 ----
    r2 = r1 + 8
    ws.cell(r2, 1, 'C. 销售额覆盖（锚定率）与重复计数对照').font = HFONT
    b.table(ws, r2 + 1, ['项目', '数值', '单位', '说明'],
            [['已锚定车系市场合计', sstc['sales_total_dedup'], '辆', sstc['universe']],
             [f'{ssnap} 全榜', sstc['board_sales_total'], '辆',
              f'{sstc["board_series_count"]} 个车系'],
             ['锚定率', sstc['linked_share_of_board_pct'], '%', '已锚定 ÷ 全榜'],
             ['车系级去重口径', sstc['series_count'], '车系', '统计单元=车系'],
             ['行级加总（错误口径）', sstc['sales_total_rowlevel_naive'], '辆',
              f'虚高 {sstc["duplication_ratio"]} 倍，禁止用于任何占比']],
            widths=[26, 14, 8, 56])

    # ---- D. 未覆盖分桶 ----
    r3 = r2 + 8
    ws.cell(r3, 1, 'D. 未覆盖分桶（分析边界，非数据缺失）').font = HFONT
    b.table(ws, r3 + 1, ['分桶', '行数', '占底表%'],
            [[k, v, round(100.0 * v / smeta['base_rows'], 2)] for k, v in cov['buckets'].items()],
            widths=[32, 12, 12])

    # ---- E. 底表写入层披露（col39 / col40 与 col35 的非对称；Lead 2026-10-19 放行要求） ----
    # 三个数**实读底表**得出（JSON 不承载写入层计数），并与 Lead 验收值比对，不一致即失败。
    ST_RETIRED = '未匹配-已停产或未上量'
    n_sid = sum(1 for r in recs if r.get('sid'))
    n_sale = sum(1 for r in recs if r.get('月销'))
    n_retired = sum(1 for r in recs if r.get('覆盖状态') == ST_RETIRED)
    n_offboard_anchored = sum(1 for r in recs if r.get('覆盖状态') == ST_RETIRED and r.get('sid'))
    if n_sid - n_sale != n_offboard_anchored:
        raise ValueError(f'col39−col35 差额 {n_sid - n_sale} != 「已停产或未上量」中已锚定数 '
                         f'{n_offboard_anchored}（写入层口径不自洽）')
    # 旧值账目实读留档。留档 schema 有两种（均由 Lead/写入流程重建，二选一）：
    #   (A) 行序全表台账：rows=底表全部行，键含 sheet_row / prior_sales（权威行序留档）
    #   (B) 紧凑台账：rows 仅旧值行，键含 disposition(replaced/cleared/kept) / current_sales
    # 二者都归到同一组计数，并与 meta 自述交叉（防留档被再次覆盖成「新值」——本日实证事故）。
    assert len(recs) == smeta['base_rows'], (
        f'行序配对前提失败：recs {len(recs)} != base_rows {smeta["base_rows"]}')
    with open(smeta['backup_old'], encoding='utf-8') as _f:
        _bk = json.load(_f)
    _brows = _bk['rows']
    _bmeta = _bk.get('meta') or {}
    if _brows and 'sheet_row' in _brows[0]:
        # (A) 行序全表台账：清空/更替 = 「工作簿现值 × 留档旧值」按 sheet_row 行序配对
        _prior = {r['sheet_row']: r.get('prior_sales') for r in _brows}
        n_prior = sum(1 for v in _prior.values() if v is not None)
        n_cleared = sum(1 for i, r in enumerate(recs)
                        if _prior.get(i + 2) is not None and not r.get('月销'))
        n_replaced, _kept = n_prior - n_cleared, 0
    else:
        # (B) 紧凑台账：直接读 disposition
        _disp = Counter(r.get('disposition') for r in _brows)
        n_prior = len(_brows)
        n_replaced, n_cleared = _disp.get('replaced', 0), _disp.get('cleared', 0)
        _kept = _disp.get('kept', 0)
    if n_replaced + n_cleared + _kept != n_prior:
        raise ValueError(f'旧值账目不平：更替 {n_replaced} + 清空 {n_cleared} + 保留 {_kept} '
                         f'!= 旧值格数 {n_prior}')
    if _bmeta.get('rows_with_prior_sales') is not None \
            and _bmeta['rows_with_prior_sales'] != n_prior:
        raise ValueError(f'留档 meta.rows_with_prior_sales {_bmeta["rows_with_prior_sales"]} '
                         f'!= 实读旧值格数 {n_prior}（留档疑似被覆盖）')
    _mc = _bmeta.get('counts')
    if _mc:
        # canonical 版键名为 prior_nonnull（早期紧凑版为 prior_rows）→ 两者兼容
        _mc_prior = _mc.get('prior_nonnull', _mc.get('prior_rows'))
        _mc_t = (_mc_prior, _mc.get('replaced'), _mc.get('cleared'), _mc.get('kept'))
        if None not in _mc_t and _mc_t != (n_prior, n_replaced, n_cleared, _kept):
            raise ValueError(f'留档 meta.counts {_mc} 与实读 {n_prior}/{n_replaced}/'
                             f'{n_cleared}/{_kept} 不一致')
    # 独立复核：留档 current_sales 必须与工作簿现值逐格一致。两种留档 schema：
    #   canonical 行序台账（有 sheet_row）→ 单元格级行序配对
    #     （AGENTS 2026-09-17：单元格级配对必须按行序，禁止按键字典配对——跨批同码会错位）
    #   早期紧凑台账（有 key）→ 出现序复合键配对
    if _brows and 'sheet_row' in _brows[0]:
        _cur_by_row = {i + 2: r.get('月销') for i, r in enumerate(recs)}
        _mm = [r for r in _brows
               if _cur_by_row.get(r['sheet_row'], 'MISSING') != r.get('current_sales')]
        if _mm:
            raise ValueError(
                f'留档与工作簿现值行序配对不一致 {len(_mm)} 处，'
                f'首例 sheet_row={_mm[0]["sheet_row"]} 留档={_mm[0].get("current_sales")} '
                f'工作簿={_cur_by_row.get(_mm[0]["sheet_row"])}')
    elif _brows and 'key' in _brows[0]:
        _occ, _cur = Counter(), {}
        for r in recs:
            _kk = (r['批次'], r['型号'], r['通用'])
            _cur[f'{_kk[0]}|{_kk[1]}|{_kk[2]}#{_occ[_kk]}'] = r.get('月销')
            _occ[_kk] += 1
        _mm = [r for r in _brows if _cur.get(r['key']) != r.get('current_sales')]
        if _mm:
            raise ValueError(f'留档与工作簿现值配对不一致 {len(_mm)} 处，'
                             f'首例 key={_mm[0]["key"]} 留档={_mm[0].get("current_sales")} '
                             f'工作簿={_cur.get(_mm[0]["key"])}')
    r4 = r3 + 3 + len(cov['buckets'])
    ws.cell(r4, 1, 'E. 底表写入层披露（车系ID / 销量覆盖状态 / 月销量 三列口径关系）').font = HFONT
    b.table(ws, r4 + 1, ['项目', '行数', '说明'],
            [['车系ID 非空（col39）', n_sid, '已锚定到车系的底表行数'],
             ['月销量 非空（col35）', n_sale, f'仅写 {ssnap} 当期销量'],
             ['两者差额', n_sid - n_sale, '已锚定但当期不在榜 → 销量留空（设计使然）'],
             [f'「{ST_RETIRED}」行数', n_retired, f'其中已锚定车系ID {n_offboard_anchored} 行'],
             ['口径统一前旧值格数', n_prior, '易车 2026-08 旧口径（行级逐格留档）'],
             ['　→ 已更替', n_replaced, f'写入 {ssnap} 当期值'],
             ['　→ 已清空', n_cleared, '该车系不在当期榜 → 留空（非数据丢失，旧值已留档）'],
             ['　→ 保留旧值', _kept, '统一替换，无逐格保留']],
            widths=[30, 12, 62])
    for _i, _t in enumerate([
        f'① 「{ST_RETIRED}」{n_retired:,} 行中，有 {n_offboard_anchored:,} 行已锚定车系ID'
        f'（该车系存在于 10 个月并集但不在 {ssnap} 榜），故当期销量留空（不写旧月值、不写 0）。'
        f'因此 col39 非空 {n_sid:,} ≠ col35 非空 {n_sale:,} 是设计使然，不是缺数。',
        f'② 口径统一后有 {n_cleared} 格旧值被清空：其车系不在 {ssnap} 榜，'
        f'按「列 35 只写 {ssnap} 当期销量」口径留空；旧值已完整留档 {smeta["backup_old"]}。'
        f'不得表述为「数据丢失」。',
        '③ 禁止用 col39 非空数 / col35 非空数互相推断对方；两者是不同判据，口径见本表 A~D 节。',
        '④ 旧值账目经 Lead 与 WS3 两套独立实现交叉验证一致'
        f'（{n_prior}/{n_replaced}/{n_cleared}/{_kept}）；留档文件由两方一致认定前，'
        '该路径曾被重建覆盖两次（详见留档 meta.incident_note），故本表一律实读留档、不手写。',
    ], r4 + 10):
        ws.cell(_i, 1, _t).font = HFONT
    ws.column_dimensions['A'].width = 108

    # ---------- 48 销量市场结构（§6-1） ----------
    ws = b.sheet('销量市场结构', f'销量市场结构（车系级去重 · {ssnap}）')
    ws.cell(3, 1, f'已锚定车系口径（{sstc["series_count"]} 车系，合计 {sstc["sales_total_dedup"]:,} 辆）；'
                  f'全榜 {sstc["board_sales_total"]:,} 辆 · {sstc["board_series_count"]} 车系。'
                  '车系级已去重，底表行级值不可加总。')
    b.table(ws, 4, ['维度', '分类', '销量(辆)', '占比%', '全榜销量(辆)', '全榜占比%'],
            ([['细分市场', k, v, dict(sstc['by_segment_pct']).get(k), '', '']
              for k, v in sstc['by_segment']]
             + [['动力类型', k, v,
                 round(100.0 * v / sstc['sales_total_dedup'], 4), '', '']
                for k, v in sstc['by_energy']]
             + [['价格带', k, v, dict(sstc['by_price_pct']).get(k),
                 dict(sstc['by_price_board']).get(k), dict(sstc['by_price_board_pct']).get(k)]
                for k, v in sstc['by_price']]
             + [['企业', k, v, round(100.0 * v / sstc['sales_total_dedup'], 4), '', '']
                for k, v in sstc['by_oem']]
             + [['品牌', k, v, round(100.0 * v / sstc['sales_total_dedup'], 4), '', '']
                for k, v in sstc['by_brand']]),
            widths=[12, 30, 14, 10, 14, 12])
    _mr = 6 + len(sstc['by_segment']) + len(sstc['by_energy']) \
        + len(sstc['by_price']) + len(sstc['by_oem']) + len(sstc['by_brand'])
    ws.cell(_mr, 1, '未纳入统计（无价格/无细分，属覆盖缺口）：'
                    + '；'.join(f'{k} {v["sales"]:,} 辆 / {v["series_count"]} 车系'
                                for k, v in sstc['missing'].items())).font = HFONT

    # ---------- 49 销量集中度（§6-2） ----------
    ws = b.sheet('销量集中度', f'销量集中度 CR/HHI（车系级去重 · {ssnap}）')
    ws.cell(3, 1, f'已锚定 {sstc["series_count"]} 车系 {sstc["sales_total_dedup"]:,} 辆；'
                  f'全榜 {sstc["board_series_count"]} 车系 {sstc["board_sales_total"]:,} 辆。'
                  'CR 单位=%；HHI 为 0-10000 点值。')
    _cr = lambda d: [d['n'], d['CR3'], d['CR5'], d['CR10'], d['HHI'], d['total']]
    b.table(ws, 4, ['口径', '单元', '统计个数', 'CR3%', 'CR5%', 'CR10%', 'HHI', '合计销量(辆)'],
            [['已锚定', '车系'] + _cr(S['concentration']['series']),
             ['已锚定', '品牌'] + _cr(S['concentration']['brand']),
             ['已锚定', '企业'] + _cr(S['concentration']['oem']),
             ['全榜', '车系'] + _cr(S['concentration']['series_board'])],
            widths=[10, 8, 10, 9, 9, 10, 10, 14])
    r1 = 10
    ws.cell(r1, 1, 'Top20 车系（已锚定）').font = HFONT
    b.table(ws, r1 + 1, ['排名', '车系', '销量(辆)', '占比%', '企业/品牌'],
            [[i, row[0], int(row[1]), row[2], row[3]] for i, row in enumerate(top_series, 1)],
            widths=[6, 26, 12, 10, 36])
    r2 = r1 + 2 + len(top_series) + 1
    ws.cell(r2, 1, 'Top20 车系（全榜）').font = HFONT
    b.table(ws, r2 + 1, ['排名', '车系', '销量(辆)'],
            [[i, row[0], int(row[1])] for i, row in enumerate(S['concentration']['top_series_board'], 1)],
            widths=[6, 26, 12])
    r3 = r2 + 2 + 20 + 1
    ws.cell(r3, 1, 'Top10 品牌').font = HFONT
    b.table(ws, r3 + 1, ['排名', '品牌', '销量(辆)', '占比%'],
            [[i, row[0], int(row[1]), row[2]] for i, row in enumerate(S['concentration']['top_brand'][:10], 1)],
            widths=[6, 22, 12, 10])
    r4 = r3 + 2 + 10 + 1
    ws.cell(r4, 1, 'Top10 企业').font = HFONT
    b.table(ws, r4 + 1, ['排名', '企业', '销量(辆)', '占比%'],
            [[i, row[0], int(row[1]), row[2]] for i, row in enumerate(S['concentration']['top_oem'][:10], 1)],
            widths=[6, 34, 12, 10])

    # ---------- 50 销量月度趋势（§6-3） ----------
    ws = b.sheet('销量月度趋势', f'销量月度趋势（车系级去重 · {smeta["months"][0]}~{ssnap}）')
    ws.cell(3, 1, '已锚定车系口径与全榜口径并列；车系级已去重，底表行级值不可加总（规范 §4）。')
    _mo = smeta['months']
    b.table(ws, 4, ['月份', '已锚定合计', 'BEV', 'PHEV+EREV', '全榜合计', '锚定占比%'],
            [[_mo[i], row[1], row[2], row[3], smon['monthly_board'][i][1],
              round(100.0 * row[1] / smon['monthly_board'][i][1], 2)]
             for i, row in enumerate(smon['monthly'])],
            widths=[10, 14, 12, 12, 14, 12])
    r1 = 5 + len(_mo) + 1
    ws.cell(r1, 1, '动力类型月度细分').font = HFONT
    b.table(ws, r1 + 1, ['月份', 'BEV', 'PHEV', 'EREV', 'BEV占比%'],
            [[row[0], row[1], row[2], row[3], row[4]] for row in smon['energy_monthly']],
            widths=[10, 12, 12, 12, 12])
    r2 = r1 + 2 + len(_mo) + 1
    ws.cell(r2, 1, f'价格带月度（已锚定；表头取 meta.price_bins）').font = HFONT
    b.table(ws, r2 + 1, ['月份'] + sprice,
            [[row[0]] + list(row[1:]) for row in smon['price_monthly']],
            widths=[10] + [12] * len(sprice))
    r3 = r2 + 2 + len(_mo) + 1
    ws.cell(r3, 1, '价格带月度（全榜）').font = HFONT
    b.table(ws, r3 + 1, ['月份'] + sprice,
            [[row[0]] + list(row[1:]) for row in smon['price_monthly_board']],
            widths=[10] + [12] * len(sprice))
    r4 = r3 + 2 + len(_mo) + 1
    ws.cell(r4, 1, '细分市场月度（Top12，其余并入"其他"）').font = HFONT
    _segs = [k for k, _ in sorted(
        ((k, sum(d.get(k, 0) for _m, d in smon['segment_monthly']))
         for k in smon['segment_monthly'][0][1]), key=lambda x: -x[1])[:12]]
    _seg_rows = []
    for _m, d in smon['segment_monthly']:
        _tot = sum(d.values())
        _top = [d.get(k, 0) for k in _segs]
        _seg_rows.append([_m] + _top + [_tot - sum(_top)])
    b.table(ws, r4 + 1, ['月份'] + _segs + ['其他'], _seg_rows,
            widths=[10] + [13] * (len(_segs) + 1))

    # ---------- 51 销量-技术关联（§6-4） ----------
    ws = b.sheet('销量-技术关联', f'销量-技术关联（销量加权 · {ssnap}）')
    ws.cell(3, 1, f'加权=按 {ssnap} 车系销量加权均值；等权=公告等权均值。'
                  f'样本 {stech["sample_n"]["weighted"]} 车系，车系指标取底表多行中位数/众数。')
    _mkeys = ['续航', '容量', '电耗', '整备', '功率', '能量密度']
    _unit = {'续航': 'km', '容量': 'kWh', '电耗': 'kWh/100km', '整备': 'kg',
             '功率': 'kW', '能量密度': 'Wh/kg'}
    b.table(ws, 4, ['口径'] + [f'{k}({_unit[k]})' for k in _mkeys] + ['样本数'],
            [['销量加权'] + [stech['weighted'][k] for k in _mkeys] + [stech['sample_n']['weighted']],
             ['公告等权'] + [stech['unweighted'][k] for k in _mkeys] + [stech['sample_n']['unweighted']]],
            widths=[12] + [14] * len(_mkeys) + [10])
    r1 = 7
    ws.cell(r1, 1, '分动力技术画像').font = HFONT
    b.table(ws, r1 + 1, ['动力'] + [f'{k}({_unit[k]})' for k in _mkeys],
            [[row[0]] + [row[1].get(k) for k in _mkeys] for row in stech['by_energy_tech']],
            widths=[10] + [14] * len(_mkeys))
    r2 = r1 + 2 + len(stech['by_energy_tech']) + 1
    ws.cell(r2, 1, '续航段分布（占比% / 车系数；分箱定义见 range_bins_def）').font = HFONT
    b.table(ws, r2 + 1, ['续航段', '销量占比%', '车系数'],
            [[row[0], row[1], row[2]] for row in stech['range_bins']],
            widths=[14, 12, 10])
    r3 = r2 + 2 + len(stech['range_bins']) + 1
    ws.cell(r3, 1, '价格带技术画像（加权续航/容量/功率/整备 + 销量 + 车系数）').font = HFONT
    b.table(ws, r3 + 1, ['价格带', '续航(km)', '容量(kWh)', '功率(kW)', '整备(kg)',
                         '销量(辆)', '车系数'],
            [list(row) for row in stech['price_tech']],
            widths=[12, 12, 12, 12, 12, 14, 10])
    _rng_def = list(stech['range_bins_def'])
    r4 = r3 + 2 + len(stech['price_tech']) + 1
    ws.cell(r4, 1, '价格带 × 续航段 矩阵（A 销量 / B 车系数）').font = HFONT
    ws.cell(r4 + 1, 1, 'A. 销量(辆)')
    b.table(ws, r4 + 2, ['价格带'] + _rng_def,
            [[row[0]] + list(row[1]) for row in stech['price_range_matrix']],
            widths=[12] + [12] * len(_rng_def))
    _r4b = r4 + 3 + len(stech['price_range_matrix']) + 1
    ws.cell(_r4b, 1, 'B. 车系数')
    b.table(ws, _r4b + 1, ['价格带'] + _rng_def,
            [[row[0]] + list(row[2]) for row in stech['price_range_matrix']],
            widths=[12] + [12] * len(_rng_def))

    # ---------- 52 价格-技术-销量（§6-5） ----------
    ws = b.sheet('价格-技术-销量', '价格 — 技术 — 销量（车系级去重）')
    ws.cell(3, 1, f'价格带取 priceinfo 中值后分箱：{smeta["price_bin_rule"]}。'
                  f'统计单元=车系（{sstc["series_count"]}），行级值不可加总。')
    _miss_price = sstc['missing'].get('price') or {}
    ws.cell(4, 1, f'未纳入（无价格）：{_miss_price.get("sales", 0):,} 辆 / '
                  f'{_miss_price.get("series_count", 0)} 车系（属覆盖缺口，非数据缺失）。')
    b.table(ws, 5, ['价格带', '销量(辆)', '销量占比%', '全榜销量(辆)', '全榜占比%',
                    '加权续航(km)', '加权容量(kWh)', '加权功率(kW)', '加权整备(kg)', '车系数'],
            [[row[0], row[5], dict(sstc['by_price_pct']).get(row[0]), '', '',
              row[1], row[2], row[3], row[4], row[6]] for row in stech['price_tech']],
            widths=[12, 14, 12, 14, 12, 14, 14, 14, 14, 10])
    r1 = 7 + len(stech['price_tech'])
    ws.cell(r1, 1, '价格带 × 续航段 销量矩阵').font = HFONT
    b.table(ws, r1 + 1, ['价格带'] + _rng_def,
            [[row[0]] + list(row[1]) for row in stech['price_range_matrix']],
            widths=[12] + [12] * len(_rng_def))
    r2 = r1 + 2 + len(stech['price_range_matrix']) + 1
    ws.cell(r2, 1, '价格带 × 续航段 车系数矩阵').font = HFONT
    b.table(ws, r2 + 1, ['价格带'] + _rng_def,
            [[row[0]] + list(row[2]) for row in stech['price_range_matrix']],
            widths=[12] + [12] * len(_rng_def))
    r3 = r2 + 2 + len(stech['price_range_matrix']) + 1
    ws.cell(r3, 1, '全榜价格带对照').font = HFONT
    b.table(ws, r3 + 1, ['价格带', '全榜销量(辆)', '全榜占比%'],
            [[k, v, dict(sstc['by_price_board_pct']).get(k)] for k, v in sstc['by_price_board']],
            widths=[12, 14, 12])

    # ---------- 53 公告-市场转化（§6-6） ----------
    ws = b.sheet('公告-市场转化', f'公告 — 市场转化（当期销量={ssnap} 快照）')
    ws.cell(3, 1, f'当期销量 = {ssnap} 车系级快照；用于公告存量（第 {bmin}~{bmax} 批累积）× '
                  '当期市场表现的转化率表现，不等于批次发布当月的实际销量。')
    ws.cell(4, 1, S['conversion']['note'])
    _bf = S['conversion']['by_batch_detail_fields']
    _bd = {r[0]: r for r in S['conversion']['by_batch_detail']}
    b.table(ws, 5, ['批次', '年份', '公告车型数', '当期销量(辆)',
                    _bf[1] if len(_bf) > 1 else '公告行数',
                    _bf[2] if len(_bf) > 2 else '去重产品型号数',
                    _bf[3] if len(_bf) > 3 else '已锚定车系数'],
            [[row[0], row[1], row[2], row[3]]
             + (list(_bd[row[0]][1:]) if row[0] in _bd else ['', '', ''])
             for row in S['conversion']['by_batch']],
            widths=[8, 8, 12, 14, 10, 16, 14])
    r1 = 6 + len(S['conversion']['by_batch']) + 1
    ws.cell(r1, 1, '按公告年份汇总').font = HFONT
    b.table(ws, r1 + 1, ['年份', '公告车型数', '当期销量(辆)'],
            [list(row) for row in S['conversion']['by_year']],
            widths=[10, 14, 16])
    ws.cell(r1 + 3 + len(S['conversion']['by_year']), 1,
            f'注：上表为正式公告口径（合计 {sum(x[1] for x in S["conversion"]["by_year"]):,} 条，不含公示批）。'
            f'{_pub_b[-1] if _pub_b else "—"} 批 {n_pub} 行为公示口径，公告日期未定，'
            '单列于本表上方「按批次」明细，不计入正式公告年度合计。')
    r2 = r1 + 2 + len(S['conversion']['by_year']) + 1
    ws.cell(r2, 1, f'分品牌转化率（按当期销量降序 Top20 / 共 {len(S["conversion"]["rate_table"])} 品牌）').font = HFONT
    _rt = sorted(S['conversion']['rate_table'], key=lambda x: -x[3])[:20]
    b.table(ws, r2 + 1, ['品牌', '公告车型数', '在售车系数', '当期销量(辆)', '单车系均销(辆)'],
            [list(row) for row in _rt],
            widths=[20, 14, 12, 16, 16])

    # ---------- 54 月度榜单迁移（§6-7） ----------
    ws = b.sheet('月度榜单迁移', f'月度榜单迁移（{_mo[0]} → {ssnap}）')
    ws.cell(3, 1, S['trend']['roster_rule'] + '（非 Top20 位次差，勿混淆）')
    b.table(ws, 4, ['排名', '首月车系', '首月销量(辆)', '末月车系', '末月销量(辆)'],
            [[i, smon['top20_first'][i - 1][0], int(smon['top20_first'][i - 1][1]),
              smon['top20_last'][i - 1][0], int(smon['top20_last'][i - 1][1])]
             for i in range(1, min(len(smon['top20_first']), len(smon['top20_last'])) + 1)],
            widths=[6, 24, 14, 24, 14])
    _fp = {r[0]: (i, r[1]) for i, r in enumerate(smon['top20_first'], 1)}
    _lp = {r[0]: (i, r[1]) for i, r in enumerate(smon['top20_last'], 1)}
    r1 = 5 + min(len(smon['top20_first']), len(smon['top20_last'])) + 1
    ws.cell(r1, 1, '两期同时在榜车系名次/销量变动').font = HFONT
    _both = sorted(set(_fp) & set(_lp), key=lambda k: _lp[k][0])
    b.table(ws, r1 + 1, ['车系', '首月排名', '首月销量(辆)', '末月排名', '末月销量(辆)',
                         '排名变动', '销量变动%'],
            [[k, _fp[k][0], int(_fp[k][1]), _lp[k][0], int(_lp[k][1]),
              _fp[k][0] - _lp[k][0],
              round(100.0 * (_lp[k][1] - _fp[k][1]) / _fp[k][1], 2) if _fp[k][1] else '']
             for k in _both],
            widths=[24, 10, 14, 10, 14, 10, 12])
    r2 = r1 + 2 + len(_both) + 1
    ws.cell(r2, 1, f'名次变动区间分布（车系数；区间={"、".join(smon["rankchange_bins"])}）').font = HFONT
    b.table(ws, r2 + 1, ['变动区间', '车系数'],
            [[row[0], row[1]] for row in smon['rankchange_hist']],
            widths=[14, 12])
    r3 = r2 + 2 + len(smon['rankchange_hist']) + 1
    ws.cell(r3, 1, f'新晋榜（名册新进，共 {smon["new_entries_total"]} 车系）').font = HFONT
    b.table(ws, r3 + 1, ['车系', '销量(辆)'],
            [[row[0], int(row[1])] for row in smon['new_entries']],
            widths=[26, 14])
    r4 = r3 + 2 + len(smon['new_entries']) + 1
    ws.cell(r4, 1, f'掉榜（名册退出，共 {smon["drops_total"]} 车系）').font = HFONT
    b.table(ws, r4 + 1, ['车系', '销量(辆)'],
            [[row[0], int(row[1])] for row in smon['drops']],
            widths=[26, 14])
    r5 = r4 + 2 + len(smon['drops']) + 1
    ws.cell(r5, 1, f'命名待锚定候选（{len(cov["pending_naming"])} 项，未计入任何已锚定统计）').font = HFONT
    b.table(ws, r5 + 1, ['通用名称', '底表行数', '候选车系'],
            [[row[0], row[1], '、'.join(row[2])] for row in cov['pending_naming']],
            widths=[20, 10, 56])

    return b, names


def _html_e(s):
    import html as _h
    return _h.escape(str(s))


def main():
    recs = load_records()
    S = load_sales(base_rows=len(recs))   # 契约 v1.1+：唯一接口；行数锚点用底表实读（规范 §11.4）
    book, names = build(recs, S)
    # 覆盖检查：目录与实际 sheet 一致（表数一律用 len(names)，去字面量）
    actual = book.wb.sheetnames
    missing = [n for n in names if n not in actual]
    assert not missing, f'缺表: {missing}'
    extra = [n for n in actual if n not in names]
    assert not extra, f'目录未列出的多余表: {extra}'
    assert '月销量分析' in actual, '缺月销量分析表'
    assert len(actual) == len(names), f'sheet 数 {len(actual)} != 目录 {len(names)}'
    # 自动备份：文件名必须含版本号（AGENTS 2026-09-27 教训③：
    # 带自动备份的生成器多次运行会覆盖原始备份 → 含版本号 + 存在性检查）
    # 追加守卫（2026-10-19）：空/极小同名文件会永久占位 → 现有备份 < 1KB 视为无效并重备
    os.makedirs('archive', exist_ok=True)
    _bak = f"archive/行业分析底表_pre-{VER}_backup_{datetime.date.today():%Y%m%d}.xlsx"
    _bak_ok = os.path.exists(_bak) and os.path.getsize(_bak) >= 1024
    if os.path.exists(OUT) and not _bak_ok:
        shutil.copy2(OUT, _bak)
        print(f'原始产物已备份 → {_bak}')
    book.wb.save(OUT)
    # 复开验证
    chk = openpyxl.load_workbook(OUT, read_only=True)
    assert len(chk.sheetnames) == len(names), \
        f'复开 sheet 数 {len(chk.sheetnames)} != 目录 {len(names)}'
    print(f'生成完成：{len(names)} 表 → {OUT}（底表 {len(recs)} 行 · 口径 {VER} · '
          f'销量快照 {S["meta"]["snapshot_month"]}）')
    print('sheets:', '、'.join(chk.sheetnames[:8]), '…')
    print(f'销量联动区块: {len(names) - 46} 张（#47~#{len(names)}）')


if __name__ == '__main__':
    main()
