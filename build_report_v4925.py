# -*- coding: utf-8 -*-
"""build_report_v4925.py — 行业分析报告 docx 生成/扩展器（v4.9.36 口径，销量章扩展）

┌─ 本文件的两个角色（phase1 起新增角色②）────────────────────────────────────┐
│ ① 历史角色：从权威底表重算 T01~T59 基础章节（build_r1..build_r5）            │
│    ⚠️ 见 Update/销量分析报告大纲_20261019.md §0.2：本生成器只产出"表格层"，   │
│    叙事层②/样式层/封面+TOC 层由三个未入库的 beautify 提交完成——              │
│    因此**从零重跑 build_report_v4925.py 会丢失交付物现状**，               │
│    交付物更新一律走 fill_report_narrative_v4932.py 的增量注入路径。          │
│ ② 新增角色：销量章（第八章）数据契约 + 计算 + 表格生成 + 增量注入。            │
│    阶段 2 由此模块产表，由 fill_report_narrative_v4932.py 注入叙事与数字。    │
└──────────────────────────────────────────────────────────────────────────┘

数据源（口径锚点 = Update/销量口径规范_20261019.md，唯一权威）
  A-JSON : audit-output/sales_linkage_analytics_20261019.json   （task-2 产出）
  原始快照: audit-output/sales_autohome_series_monthly_20261019.json（61MB，仅校验用）
  底表    : NEV公告参数汇总表_合并版（341~410批）.xlsx      （只读，不写）

纪律（硬约束，任何违反视为缺陷）
  1. 禁止手写数字：一切数值经 NUM_SRC 注册表从 A-JSON 读取；本模块内除章节号/表号/
     批次号/月份标签外不得出现裸数字（见 guard_no_literals / G1 门禁）。
  2. 禁止 Table.cell() 写入：含 vMerge 的表格必须走 set_tc()（按 tr.tc_lst 定位）。
  3. 禁止全局数值映射：数字只在「(章节, 表ID, 行标签, 列头) → JSON路径」标签寻址下注入。
  4. 自动备份名必须含版本号（防多次运行覆盖原始备份）。
  5. 车系级口径：所有汇总以车系为统计单元；底表行级值不可加总（脚注强制）。

用法
  python build_report_v4925.py --out <docx路径> [--stages r1,r2,r3,r4,r5]   # 角色①（慎用）
  python build_report_v4925.py --sales-probe                                # 仅校验 A-JSON 契约
  python build_report_v4925.py --sales-draft --out <docx路径>                # 新章独立草稿
"""

import argparse
import datetime
import json
import os
import re
import statistics
import sys
from collections import defaultdict, Counter

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')

# ============================================================================
# 0. 版本与路径常量（阶段 2 由 Lead 确认递增号）
# ============================================================================
GENERATOR_VERSION = 'v4.9.34'          # Lead 裁决 D3：并发会话已发 [4.9.33]，本轮为其后一版
SNAPSHOT_ID = '20261019'
ANALYTICS_JSON = os.path.join(HERE, 'audit-output', 'sales_linkage_analytics_20261019.json')
AUTOHOME_JSON = os.path.join(HERE, 'audit-output', 'sales_autohome_series_monthly_20261019.json')
REPORT_DOCX = os.path.join(HERE, 'NEV公告车型行业分析报告.docx')
ARCHIVE_DIR = os.path.join(HERE, 'archive')

# 「销量写入前」底表快照的定位常量（见 _load_head_rows 说明）。
# 底表入库提交 d35891b 之后 HEAD 已是新态，因此必须按"写入前"的提交取。
# 双重校验：列数 == 39 且 col36（销量数据月份）== '2026-08'；不符则沿 git log 回溯。
SIZE_BASE_COMMIT = '474e897'
PREV_COLUMNS = 39
PREV_MONTH = '2026-08'
# 口径统一后的目标态（新终态 5,475 行 / 41 列 / 341~412 批）
NEW_COLUMNS = 41

# 样式常量 —— 与交付物现状完全一致（见大纲 §0.2 样式层）
STYLE_TABLE = 'Light Grid Accent 1'
STYLE_H1, STYLE_H2, STYLE_H3, STYLE_BODY = 'Heading 1', 'Heading 2', 'Heading 3', 'Normal'
FONT_CN = '微软雅黑'

# 车系级口径脚注（规范 §4 强制文本，逐表附带）
FOOT_SERIES = ('车系级口径：已按 车系ID 去重；底表行级销量值仅供单行参考，不可加总。'
               f'数据源：汽车之家·行业综合销量（车系级月度零售量，每月10日左右更新）；'
               f'快照月 {0}；快照ID {1}。').format('2026-09', SNAPSHOT_ID)
FOOT_PRICE = ('价格带分箱：由 priceinfo（如"26.35-36.90万"）取中值后分箱 '
              '≤10万 / 10-15 / 15-20 / 20-30 / 30-40 / 40万+。')

# ============================================================================
# 1. A-JSON 数据契约（大写常量 = 冻结接口，阶段 2 按此解析）
# ============================================================================
# 冻结契约 v1.0（Update/销量联动JSON契约_20261019.md）：路径逐字照抄，**不设别名**。
# 消费方 fail-fast：任一路径缺失 → 抛错并报告，不得用默认值、不得跳过。
SCHEMA_PATH = [
    # --- meta ---
    ('meta.snapshot_month',        ['meta.snapshot_month']),
    ('meta.months',                ['meta.months']),
    ('meta.fetch_date',            ['meta.fetch_date']),
    ('meta.source',                ['meta.source']),
    ('meta.source_tip',            ['meta.source_tip']),
    ('meta.source_url_template',   ['meta.source_url_template']),
    ('meta.reference_source',      ['meta.reference_source']),
    ('meta.scope',                 ['meta.scope']),
    ('meta.dedup_rule',            ['meta.dedup_rule']),
    ('meta.price_bins',            ['meta.price_bins']),
    ('meta.price_bin_rule',        ['meta.price_bin_rule']),
    ('meta.backup_old',            ['meta.backup_old']),
    ('meta.base_table_path',       ['meta.base_table_path']),
    ('meta.base_rows',             ['meta.base_rows']),
    ('meta.matched_rows',          ['meta.matched_rows']),
    # 规范 §11.4 禁硬编码：批次范围/批次数一律从 JSON 读，不写死 "341~410"/"70 批"
    ('meta.batch_range',           ['meta.batch_range']),
    ('meta.batch_count',           ['meta.batch_count']),
    ('meta.batch_range_note',      ['meta.batch_range_note']),
    ('meta.qc.union_series',       ['meta.qc.union_series']),
    # --- coverage（规范 §5：三口径 + 未覆盖分桶）---
    ('coverage.row.covered',       ['coverage.row.covered']),
    ('coverage.row.total',         ['coverage.row.total']),
    ('coverage.row.pct',           ['coverage.row.pct']),
    ('coverage.series.matched',    ['coverage.series.matched']),
    ('coverage.series.total',      ['coverage.series.total']),
    ('coverage.series.pct',        ['coverage.series.pct']),
    # 口径自纠（Lead m00348）：546/1299 = 42.03%，剔 12 个无身份通用名称
    ('coverage.series.generic_names_matched', ['coverage.series.generic_names_matched']),
    ('coverage.series.generic_names_total',   ['coverage.series.generic_names_total']),
    ('coverage.series.generic_names_pct',     ['coverage.series.generic_names_pct']),
    # Lead m00348 定性：分母由分子派生（348+33=381），属**同义反复**，不是覆盖率，
    # 已由 WS2 改名 self_consistent_ratio。仅登记路径以便门禁断言其存在，**绝不对外引用其数值**。
    ('coverage.self_consistent_ratio', ['coverage.self_consistent_ratio']),
    ('coverage.detail',            ['coverage.detail']),
    ('coverage.definitions',       ['coverage.definitions']),
    # 纲要 v1.0 §3.5：唯一合法的"在售覆盖"主张（WS2 已落地实际键名）
    ('coverage.current_board',              ['coverage.current_board']),
    ('coverage.current_board.board_series',   ['coverage.current_board.board_series']),
    ('coverage.current_board.linked_series',  ['coverage.current_board.linked_series']),
    ('coverage.current_board.pct',            ['coverage.current_board.pct']),
    ('coverage.current_board.linked_sales',   ['coverage.current_board.linked_sales']),
    ('coverage.current_board.board_sales',    ['coverage.current_board.board_sales']),
    ('coverage.current_board.sales_pct',      ['coverage.current_board.sales_pct']),
    ('coverage.status_counts',     ['coverage.status_counts']),
    ('coverage.buckets',           ['coverage.buckets']),
    ('coverage.pending_naming',    ['coverage.pending_naming']),
    # --- structure（规范 §6 表1）---
    ('structure.by_segment',       ['structure.by_segment']),
    ('structure.by_segment_pct',   ['structure.by_segment_pct']),
    ('structure.by_energy',        ['structure.by_energy']),
    ('structure.by_price',         ['structure.by_price']),
    ('structure.by_price_pct',     ['structure.by_price_pct']),
    ('structure.by_oem',           ['structure.by_oem']),
    ('structure.by_brand',         ['structure.by_brand']),
    ('structure.sales_total_dedup',        ['structure.sales_total_dedup']),
    ('structure.sales_total_rowlevel_naive', ['structure.sales_total_rowlevel_naive']),
    ('structure.board_sales_total',        ['structure.board_sales_total']),
    ('structure.board_series_count',       ['structure.board_series_count']),
    ('structure.linked_share_of_board_pct', ['structure.linked_share_of_board_pct']),
    ('structure.duplication_ratio',        ['structure.duplication_ratio']),
    ('structure.by_price_board',           ['structure.by_price_board']),
    ('structure.by_price_board_pct',       ['structure.by_price_board_pct']),
    # 规范 §11.2 强制披露：R1 众数口径下的混合动力家族清单（三数须区分，勿混用）
    ('structure.mixed_energy_families',    ['structure.mixed_energy_families']),
    ('structure.mixed_with_erev',          ['structure.mixed_with_erev']),
    ('structure.r2_flip_impact',           ['structure.r2_flip_impact']),
    # 纲要 v1.0 结论 4：在榜未关联分桶（WS2 已落地实际键名）
    ('structure.unlinked_board',   ['structure.unlinked_board']),
    # --- concentration（规范 §6 表2）---
    ('concentration.series',       ['concentration.series']),
    ('concentration.brand',        ['concentration.brand']),
    ('concentration.oem',          ['concentration.oem']),
    ('concentration.series_board', ['concentration.series_board']),
    ('concentration.top_series',   ['concentration.top_series']),
    ('concentration.top_series_board', ['concentration.top_series_board']),
    ('concentration.top_brand',    ['concentration.top_brand']),
    ('concentration.top_oem',      ['concentration.top_oem']),
    # --- trend（规范 §6 表3 + 表7）---
    ('trend.monthly',              ['trend.monthly']),
    ('trend.monthly_board',        ['trend.monthly_board']),
    ('trend.energy_monthly',       ['trend.energy_monthly']),
    ('trend.price_monthly',        ['trend.price_monthly']),
    ('trend.price_monthly_board',  ['trend.price_monthly_board']),
    ('trend.segment_monthly',      ['trend.segment_monthly']),
    ('trend.top20_first',          ['trend.top20_first']),
    ('trend.top20_last',           ['trend.top20_last']),
    ('trend.rankchange_hist',      ['trend.rankchange_hist']),
    ('trend.new_entries',          ['trend.new_entries']),
    ('trend.drops',                ['trend.drops']),
    # --- tech_sales（规范 §6 表4/表5）---
    ('tech_sales.weighted',        ['tech_sales.weighted']),
    ('tech_sales.unweighted',      ['tech_sales.unweighted']),
    ('tech_sales.sample_n',        ['tech_sales.sample_n']),
    ('tech_sales.range_bins',      ['tech_sales.range_bins']),
    ('tech_sales.range_bins_def',  ['tech_sales.range_bins_def']),
    ('tech_sales.price_tech',      ['tech_sales.price_tech']),
    ('tech_sales.price_range_matrix', ['tech_sales.price_range_matrix']),
    ('tech_sales.by_energy_tech',  ['tech_sales.by_energy_tech']),
    # --- conversion（规范 §6 表6）---
    ('conversion.by_batch',        ['conversion.by_batch']),
    ('conversion.by_year',         ['conversion.by_year']),
    ('conversion.brand_scatter',   ['conversion.brand_scatter']),
    ('conversion.rate_table',      ['conversion.rate_table']),
]

# 契约中必须齐备的**子键**（fail-fast 细分；缺失即报错，不降级）
SCHEMA_SUBKEYS = {
    'meta.match_channel_counts':      ['A_exact', 'B_brand', 'C_substr', 'D_suggest', 'hold'],
    'meta.qc.monthly_counts':         [],     # 至少 1 个月
    'meta.qc.cross_source':           ['yiche_202608', 'dongchedi_202609'],
    'meta.qc.cross_source.yiche_202608':  ['compared', 'within_1pct', 'unmatched',
                                           'over_10pct', 'mean_rel_diff'],
    'meta.qc.cross_source.dongchedi_202609': ['dcd_rows', 'matched', 'comparable_nev_count',
                                              'within_1pct', 'over_10pct_count',
                                              'over_10pct_unexplained_count'],
    'meta.dedup_selfcheck':           ['case', 'generic_names', 'resolved_seriesid',
                                       'collapsed_rows', 'distinct_series_in_top20'],
    'coverage.status_counts':         ['在售-已匹配', '在售-零销量', '未匹配-命名待锚定',
                                       '未匹配-已停产或未上量', '无车系身份', '未匹配-无车系ID'],
    'coverage.buckets':               ['无车系身份', '未匹配-已停产或未上量',
                                       '未匹配-命名待锚定', '未匹配-无车系ID'],
    'tech_sales.weighted':            ['续航', '容量', '电耗', '整备', '功率', '能量密度'],
    'tech_sales.unweighted':          ['续航', '容量', '电耗', '整备', '功率', '能量密度'],
    'tech_sales.sample_n':            ['weighted', 'unweighted'],
    'concentration.series':           ['CR3', 'CR5', 'CR10', 'HHI', 'n'],
    'concentration.brand':            ['CR3', 'CR5', 'CR10', 'HHI', 'n'],
    'concentration.oem':              ['CR3', 'CR5', 'CR10', 'HHI', 'n'],
    'concentration.series_board':     ['CR3', 'CR5', 'CR10', 'HHI', 'n'],
    'coverage.detail':                ['generic_names_all', 'generic_names_no_identity',
                                       'generic_names_denominator', 'generic_names_anchored',
                                       'generic_names_hold', 'generic_names_unmatched',
                                       'distinct_seriesid_anchored', 'pending_naming_total'],
    'coverage.definitions':           ['row', 'series', 'current_board',
                                       'self_consistent_ratio', 'buckets', 'suggest'],
    'coverage.current_board':         ['board_series', 'linked_series', 'pct',
                                       'linked_sales', 'board_sales', 'sales_pct'],
    # 结论 4 边界：在榜未关联分桶，成因必须按 cause_bucket 拆分披露（不得只说"未关联"）
    'structure.unlinked_board':       ['series_count', 'sales', 'sales_pct',
                                       'board_series', 'board_sales', 'cause_bucket'],
    'structure.unlinked_board.cause_bucket': ['命名未锚定', '公告库无痕迹'],
    # 混合动力家族三数须区分：74/32.36%（全部混合家族）≠ 16/6.35%（R2 翻转影响）
    #                              ≠ 23/7.35%（含 EREV 的混合家族）
    'structure.mixed_energy_families': ['series', 'sales', 'pct_of_linked', 'items'],
    'structure.r2_flip_impact':        ['series', 'sales', 'pct_of_linked', 'items'],
    'structure.mixed_with_erev':       ['series', 'sales', 'pct_of_linked', 'items'],
}

# ============================================================================
# 1b. 过渡基准值（纲要 v1.0 已独立复核，契约槽位由 WS2 补入中）
#     ⚠️ 这些值**不是**契约字段，仅用于"WS2 落地前"的早期集成；每次使用都打印并记账。
#     最终 docx（G9 门禁）要求 pending_used == 0，即必须改从 JSON 读；不得长期依赖。
#     来源：`Update/核心论断与建议纲要_20261019.md` §3.5 / 结论 4。
# ============================================================================
PENDING_SLOTS = {
    # ⚠️ WS2 已于 2026-10-19 16:56 版落地这两个槽位（键名以 JSON 为准）；
    #    此处保留同形兜底仅为"JSON 回退旧版"时不致崩溃。终稿 G9 要求 pending_used == 0。
    'coverage.current_board': {
        'board_series': 411, 'linked_series': 282, 'pct': 68.6131,
        'linked_sales': 797701, 'board_sales': 1092064, 'sales_pct': 73.0453,
        '_source': '纲要 v1.0 §3.5（WS2 补入契约前）',
    },
    'structure.unlinked_board': {
        'series_count': 129, 'sales': 294363, 'sales_pct': 26.9547,
        'board_series': 411, 'board_sales': 1092064,
        'cause_bucket': {
            '命名未锚定':   {'series': 88, 'sales': 223077},
            '公告库无痕迹': {'series': 41, 'sales': 71286},
        },
        '_source': '纲要 v1.0 结论 4 边界（WS2 补入契约前）',
    },
}


def apply_pending_slots(aj, allow_pending=True):
    """把契约尚未落地的槽位用纲要 v1.0 基准值补齐（仅限过渡期）。

    返回实际用到的槽位键列表；allow_pending=False 时缺失即抛错（最终交付路径用）。
    """
    used = []
    for path, val in PENDING_SLOTS.items():
        if dig(aj, path) is None:
            if not allow_pending:
                raise KeyError(f'契约槽位缺失且不允许过渡基准值: {path}（WS2 需补入）')
            node, *rest = path.split('.')
            cur = aj.setdefault(node, {})
            for r in rest[:-1]:
                cur = cur.setdefault(r, {})
            cur[rest[-1]] = {k: v for k, v in val.items() if not k.startswith('_')}
            used.append(path)
    if used:
        print(f'[PENDING] ⚠ 使用纲要 v1.0 过渡基准值 {len(used)} 处（WS2 落地后须改从 JSON 读）:')
        for p in used:
            print(f'          - {p}  ← {PENDING_SLOTS[p]["_source"]}')
    return used

# 未覆盖分桶的标准枚举（规范 §2.3 六取值；对应计数必须齐备）
COVERAGE_BUCKET_KEYS = [
    '在售-已匹配', '在售-零销量', '未匹配-命名待锚定',
    '未匹配-已停产或未上量', '无车系身份', '未匹配-无车系ID',
]


def dig(obj, dotted):
    """按 'a.b.c' 取嵌套值；任一层缺失返回 None（不抛）。"""
    cur = obj
    for k in dotted.split('.'):
        if isinstance(cur, dict):
            if k not in cur:
                return None
            cur = cur[k]
        elif isinstance(cur, list) and k.isdigit():
            i = int(k)
            if i >= len(cur):
                return None
            cur = cur[i]
        else:
            return None
    return cur


def load_analytics(path=ANALYTICS_JSON, allow_pending=True):
    """装载 A-JSON。allow_pending=True 时用纲要 v1.0 过渡基准值补齐 WS2 未落地的槽位。

    返回 (aj, pending_used)。最终交付路径（终稿）应传 allow_pending=False，
    并由 G9 门禁要求 pending_used == 0。
    """
    if not os.path.exists(path):
        print(f'[FATAL] A-JSON 不存在: {path}')
        print('        task-2 尚未产出。阶段 2 前必须就绪；不生成半成品。')
        sys.exit(2)
    with open(path, encoding='utf-8') as f:
        aj = json.load(f)
    used = apply_pending_slots(aj, allow_pending=allow_pending)
    return aj, used


def assert_schema(aj):
    """G0 前置门禁：冻结契约齐备才继续。缺项全部列出后 exit(2)（fail-fast，不降级）。"""
    missing, resolved, shape_bad, subbad = [], {}, [], []
    for logical, aliases in SCHEMA_PATH:
        v, hit = None, None
        for a in aliases:
            v = dig(aj, a)
            if v is not None:
                hit = a
                break
        if v is None:
            missing.append((logical, aliases))
        else:
            resolved[logical] = hit
            if logical.endswith('months') and not isinstance(v, list):
                shape_bad.append((logical, type(v).__name__, 'list'))
    # 子键齐备性（契约 fail-fast 细分）
    for logical, subkeys in SCHEMA_SUBKEYS.items():
        obj = dig(aj, logical)
        if obj is None:
            continue                       # 已由 SCHEMA_PATH 报过
        if subkeys:
            if not isinstance(obj, dict):
                subbad.append((logical, type(obj).__name__, 'object'))
                continue
            for k in subkeys:
                if k not in obj:
                    subbad.append((f'{logical}["{k}"]', 'absent', 'present'))
        else:                              # 空列表 = 至少 1 项
            if not isinstance(obj, dict) or not obj:
                subbad.append((logical, 'empty', 'non-empty object'))
    # 六类覆盖状态之和必须 == base_rows（契约硬约束）
    sc = dig(aj, 'coverage.status_counts')
    base_rows = dig(aj, 'meta.base_rows')
    sum_bad = []
    if isinstance(sc, dict) and isinstance(base_rows, (int, float)):
        got = sum(v for v in sc.values() if isinstance(v, (int, float)))
        if got != base_rows:
            sum_bad.append((got, base_rows))
    if missing or shape_bad or subbad or sum_bad:
        print('[G0] A-JSON 契约校验 FAIL —— 不生成文档')
        for logical, aliases in missing:
            print(f'  MISSING {logical}   (契约路径: {", ".join(aliases)})')
        for logical, got, want in shape_bad:
            print(f'  SHAPE   {logical}: got {got}, want {want}')
        for logical, got, want in subbad:
            print(f'  SUBKEY  {logical}: got {got}, want {want}')
        for got, want in sum_bad:
            print(f'  SUM     coverage.status_counts 合计 {got} != meta.base_rows {want}')
        print(f'  实到顶层键: {sorted(aj.keys())}')
        sys.exit(2)
    print(f'[G0] 契约校验 PASS —— {len(resolved)} 条路径 + '
          f'{len(SCHEMA_SUBKEYS)} 组子键全部解析，status_counts 合计 == base_rows')
    return resolved


# ============================================================================
# 2. NUM_SRC 数字注册表 —— 「模板变量 → (冻结契约 JSON 路径, 单位)」
#    模板里写 ${key}，由 render() 替换；未在此注册的键 → KeyError（fail-fast）。
#    路径一律逐字取自 `Update/销量联动JSON契约_20261019.md` v1.0，禁止别处另算。
#    状态标记：① 直接路径 = 契约已定义，阶段 2 只待 JSON；② 待 Lead 新增 = 需向生产方申请新增字段。
# ============================================================================
NUM_SRC = {
    # --- 摘要 ---
    'abs.snapshot_sales':     ('trend.monthly.9.1',                      '辆'),   # 2026-09 已锚定车系级月销
    'abs.board_total':        ('trend.monthly_board.9.1',                 '辆'),   # 2026-09 全榜月销（含未锚定车系）
    'abs.board_series':       ('structure.board_series_count',            '个'),
    'abs.naive_total':        ('structure.sales_total_rowlevel_naive',    '辆'),   # 反例口径
    'abs.dedup_ratio':        ('structure.duplication_ratio',             '倍'),
    'abs.series_cr10':        ('concentration.series_board.CR10',         '%'),   # 全榜 411 车系口径
    'abs.oem_cr10':           ('concentration.oem.CR10',                  '%'),
    'abs.w_km':               ('tech_sales.weighted.续航',                'km'),
    'abs.u_km':               ('tech_sales.unweighted.续航',              'km'),
    # 摘要里的"三个分母"声明（Lead m01044 要求进摘要而非附录）
    'abs.row_pct':            ('coverage.row.pct',                        '%'),
    'abs.row_covered':        ('coverage.row.covered',                    '行'),
    'abs.row_total':          ('coverage.row.total',                      '行'),
    'abs.board_pct':          ('coverage.current_board.pct',              '%'),
    'abs.board_linked':       ('coverage.current_board.linked_series',    '个'),
    'abs.board_series_n':     ('coverage.current_board.board_series',     '个'),
    'abs.board_sales_pct':    ('coverage.current_board.sales_pct',        '%'),
    # --- 第一章 数据与方法 ---
    'm1.legacy_cells':        ('CALC:git 历史（销量写入前）底表 col35 非空格数', '格'),
    'm1.head_rows':           ('CALC:git 历史（销量写入前）底表数据行数',        '行'),
    'm1.batch_lo':            ('meta.batch_range.0',                          '批'),
    'm1.batch_hi':            ('meta.batch_range.1',                          '批'),
    'm1.batch_count':         ('meta.batch_count',                            '批'),
    # 末批（公示口径）行数与正式公告/公示记录数拆分 —— Lead 裁决 D 条件④
    # 「5,475 条公告/公示记录（正式公告 5,417 条 + 末批公示 58 条）× 41 列」。
    # 全部现读底表，不手写。
    'm1.last_batch_rows':     ('CALC:底表中批次 == batch_range[1] 的行数（公示口径行数）', '条'),
    'm1.official_rows':       ('CALC:底表行数 − 末批公示行数（正式公告记录数）',           '条'),
    'm1.prev_month':          ('CALC:git 历史底表的销量月份（口径统一前的月份）', ''),
    'm1.prev_cols':           ('CALC:git 历史底表列数（口径统一前的表宽）',   '列'),
    # 附录 A 用的数据源与交叉校验条目（全部走契约路径，零手写）
    'm1.source':              ('meta.source',                                ''),
    'm1.source_tip':          ('meta.source_tip',                            ''),
    'm1.snapshot':            ('meta.snapshot_month',                        ''),
    'm1.fetch_date':          ('meta.fetch_date',                            ''),
    'm1.dcd_rows':            ('meta.qc.cross_source.dongchedi_202609.dcd_rows',                     '条'),
    'm1.dcd_comp':            ('meta.qc.cross_source.dongchedi_202609.comparable_nev_count',          '个'),
    'm1.dcd_within':          ('meta.qc.cross_source.dongchedi_202609.within_1pct',                   '个'),
    'm1.dcd_over':            ('meta.qc.cross_source.dongchedi_202609.over_10pct_count',              '项'),
    'm1.dcd_unexplained':     ('meta.qc.cross_source.dongchedi_202609.over_10pct_unexplained_count',  '项'),
    'm1.rows':                ('meta.base_rows',                             '条'),
    'm1.cols':                ('meta.base_table_columns',                    '列'),
    # 封面/脚注用的脚本常量（CFG 前缀，非契约字段：取自本脚本自身配置）
    'cfg.ver':                ('CFG:ver',                                    ''),
    'cfg.gen_date':           ('CFG:gen_date',                               ''),
    'm1.cross_compared':      ('meta.qc.cross_source.yiche_202608.compared',      '格'),
    'm1.cross_same':          ('meta.qc.cross_source.yiche_202608.within_1pct',   '格'),
    'm1.cross_unmatched':     ('meta.qc.cross_source.yiche_202608.unmatched',     '格'),
    'm1.cross_over10':        ('CALC:over_10pct 按名称去重后的条数',              '个'),
    'm1.row_pct':             ('coverage.row.pct',                        '%'),
    'm1.series_pct':          ('coverage.series.generic_names_pct',       '%'),   # 546/1299 = 42.03（D9/口径自纠后）
    'm1.board_pct':           ('coverage.current_board.pct',              '%'),   # 282/411 = 68.61 ← 对外主张"在售覆盖率"的唯一合法数字
    'm1.board_sales_pct':     ('coverage.current_board.sales_pct',        '%'),   # 73.05
    'm1.hold':                ('meta.match_channel_counts.hold',           '个'),
    # 命名待锚定的**三个不同口径**必须分开报，混用即失真（行级 116 / 通用名称 33 / 车系级 88）
    'm1.pending_rows':        ('coverage.status_counts.未匹配-命名待锚定',  '行'),
    'm1.pending_generic':     ('coverage.detail.generic_names_hold',        '个'),
    'm1.pending_series':      ('structure.unlinked_board.cause_bucket.命名未锚定.series', '个'),
    # Lead m00985 两处必披露（底表现读：col35/col39/col40 + 旧值留档行序配对）
    'm1.stopped_rows':        ('CALC:col40=="未匹配-已停产或未上量" 的行数',       '行'),
    'm1.stopped_with_id':     ('CALC:上述行中 col39 非空（已锚定车系ID）的行数',   '行'),
    'm1.col35_filled':        ('CALC:col35「月销量」非空格数',                     '格'),
    'm1.col39_filled':        ('CALC:col39「车系ID」非空格数（设计上大于 col35）', '格'),
    'm1.cleared_cells':       ('CALC:旧值被清空的格数（车系不在 2026-09 榜）',      '格'),
    'm1.replaced_cells':      ('CALC:旧值被更替的格数',                            '格'),
    'm1.kept_cells':          ('CALC:旧值保留原值的格数（应为 0）',                '格'),
    'm1.generic_all':         ('coverage.detail.generic_names_all',        '个'),
    'm1.generic_denom':       ('coverage.detail.generic_names_denominator','个'),
    'm1.generic_noident':     ('coverage.detail.generic_names_no_identity','个'),
    # --- 第二章 市场总量与结构 ---
    'm2.total':               ('structure.sales_total_dedup',             '辆'),
    'm2.board_total':         ('structure.board_sales_total',             '辆'),
    'm2.bev':                 ('structure.by_energy.0.1',                 '辆'),
    'm2.phev':                ('structure.by_energy.1.1',                 '辆'),
    'm2.erev':                ('structure.by_energy.2.1',                 '辆'),
    'm2.bev_pct':             ('CALC:BEV 销量 ÷ sales_total_dedup',       '%'),
    'm2.phev_pct':            ('CALC:PHEV 销量 ÷ sales_total_dedup',      '%'),
    'm2.erev_pct':            ('CALC:EREV 销量 ÷ sales_total_dedup',      '%'),
    'm2.top_segment':         ('structure.by_segment.0.0',                ''),
    'm2.top_segment_sales':   ('structure.by_segment.0.1',                '辆'),
    'm2.top_price_board':     ('CALC:by_price_board 中销量最大的价格带',    ''),
    'm2.naive_total':         ('structure.sales_total_rowlevel_naive',    '辆'),
    'm2.dup_ratio':           ('structure.duplication_ratio',             '倍'),
    # --- 第三章 技术-市场关联（加权 vs 等权，纲要结论 2） ---
    'm3.w_km':                ('tech_sales.weighted.续航',                'km'),
    'm3.u_km':                ('tech_sales.unweighted.续航',              'km'),
    'm3.d_km':                ('CALC:加权续航 − 等权续航',                'km'),
    'm3.w_kwh':               ('tech_sales.weighted.容量',                'kWh'),
    'm3.u_kwh':               ('tech_sales.unweighted.容量',              'kWh'),
    'm3.d_kwh':               ('CALC:加权容量 − 等权容量',                'kWh'),
    'm3.w_kwh100':            ('tech_sales.weighted.电耗',                'kWh/100km'),
    'm3.u_kwh100':            ('tech_sales.unweighted.电耗',              'kWh/100km'),
    'm3.d_kwh100':            ('CALC:加权电耗 − 等权电耗',                'kWh/100km'),
    'm3.w_kg':                ('tech_sales.weighted.整备',                'kg'),
    'm3.u_kg':                ('tech_sales.unweighted.整备',              'kg'),
    'm3.d_kg':                ('CALC:加权整备 − 等权整备',                'kg'),
    'm3.w_kw':                ('tech_sales.weighted.功率',                'kW'),
    'm3.u_kw':                ('tech_sales.unweighted.功率',              'kW'),
    'm3.d_kw':                ('CALC:加权功率 − 等权功率',                'kW'),
    'm3.w_whkg':              ('tech_sales.weighted.能量密度',            'Wh/kg'),
    'm3.u_whkg':              ('tech_sales.unweighted.能量密度',          'Wh/kg'),
    'm3.d_whkg':              ('CALC:加权能量密度 − 等权能量密度',        'Wh/kg'),
    'm3.n_w':                 ('tech_sales.sample_n.weighted',            '车系'),
    'm3.n_u':                 ('tech_sales.sample_n.unweighted',          '车系'),
    'm3.range_share_700':     ('CALC:续航 ≥700km 区间销量份额',           '%'),
    # --- 第三章补充：分动力画像（规范 §11.2 R1 众数口径，canonical） ---
    'm3.bev_km':              ('tech_sales.by_energy_tech.0.1.续航',      'km'),
    'm3.bev_kwh':             ('tech_sales.by_energy_tech.0.1.容量',      'kWh'),
    'm3.bev_kwh100':          ('tech_sales.by_energy_tech.0.1.电耗',      'kWh/100km'),
    'm3.bev_kg':              ('tech_sales.by_energy_tech.0.1.整备',      'kg'),
    'm3.bev_kw':              ('tech_sales.by_energy_tech.0.1.功率',      'kW'),
    'm3.bev_whkg':            ('tech_sales.by_energy_tech.0.1.能量密度',  'Wh/kg'),
    'm3.phev_km':             ('tech_sales.by_energy_tech.1.1.续航',      'km'),
    'm3.phev_kwh':            ('tech_sales.by_energy_tech.1.1.容量',      'kWh'),
    'm3.phev_kwh100':         ('tech_sales.by_energy_tech.1.1.电耗',      'kWh/100km'),
    'm3.phev_kg':             ('tech_sales.by_energy_tech.1.1.整备',      'kg'),
    'm3.phev_kw':             ('tech_sales.by_energy_tech.1.1.功率',      'kW'),
    'm3.phev_whkg':           ('tech_sales.by_energy_tech.1.1.能量密度',  'Wh/kg'),
    'm3.erev_km':             ('tech_sales.by_energy_tech.2.1.续航',      'km'),
    'm3.erev_kwh':            ('tech_sales.by_energy_tech.2.1.容量',      'kWh'),
    'm3.erev_kwh100':         ('tech_sales.by_energy_tech.2.1.电耗',      'kWh/100km'),
    'm3.erev_kg':             ('tech_sales.by_energy_tech.2.1.整备',      'kg'),
    'm3.erev_kw':             ('tech_sales.by_energy_tech.2.1.功率',      'kW'),
    'm3.erev_whkg':           ('tech_sales.by_energy_tech.2.1.能量密度',  'Wh/kg'),
    # 相对 BEV 的比率（R1 叙事的关键量化：45% / +28% / +36% / +62%）
    'm3.erev_km_ratio':       ('CALC:EREV 续航 ÷ BEV 续航 × 100',         '%'),
    'm3.erev_kw_ratio':       ('CALC:EREV 功率 ÷ BEV 功率 × 100',         '%'),
    'm3.erev_kg_ratio':       ('CALC:EREV 整备 ÷ BEV 整备 × 100',         '%'),
    'm3.erev_kwh100_ratio':   ('CALC:EREV 电耗 ÷ BEV 电耗 × 100',         '%'),
    # "高 X%" = 上述相对指数 − 100；语义与 _ratio 不同，不可互换（否则"为128%"被说成"高128%"）
    'm3.erev_kw_over':        ('CALC:(EREV 功率 ÷ BEV 功率 − 1) × 100',   '%'),
    'm3.erev_kg_over':        ('CALC:(EREV 整备 ÷ BEV 整备 − 1) × 100',   '%'),
    'm3.erev_kwh100_over':    ('CALC:(EREV 电耗 ÷ BEV 电耗 − 1) × 100',   '%'),
    # --- 混合动力家族披露（规范 §11.2 强制项） ---
    'm1.mixed_n':             ('CALC:混合动力家族数（structure.mixed_energy_families）', '个'),
    'm1.mixed_sales':         ('CALC:混合家族销量合计',                    '辆'),
    'm1.mixed_pct':           ('CALC:混合家族占已锚定销量比例',            '%'),
    'm1.mixed_erev_n':        ('CALC:含 EREV 行的混合家族数',              '个'),
    'm1.mixed_erev_sales':    ('CALC:含 EREV 行的混合家族销量合计',        '辆'),
    'm1.r2_flip_n':           ('CALC:被否决 R2 方案会翻转的车系数',       '个'),
    'm1.r2_flip_sales':       ('CALC:被否决 R2 方案会翻转车系的销量合计', '辆'),
    'm1.r2_flip_pct':         ('CALC:被否决 R2 方案翻转销量占已锚定比例', '%'),
    'm4.s_cr3':               ('concentration.series_board.CR3',          '%'),
    'm4.s_cr5':               ('concentration.series_board.CR5',          '%'),
    'm4.s_cr10':              ('concentration.series_board.CR10',         '%'),
    'm4.s_hhi':               ('concentration.series_board.HHI',          ''),
    'm4.s_n':                 ('concentration.series_board.n',            '个'),
    'm4.sa_cr10':             ('concentration.series.CR10',               '%'),   # 已锚定 282 车系口径
    'm4.sa_hhi':              ('concentration.series.HHI',                ''),
    'm4.sa_n':                ('concentration.series.n',                  '个'),
    'm4.b_cr3':               ('concentration.brand.CR3',                 '%'),
    'm4.b_cr5':               ('concentration.brand.CR5',                 '%'),
    'm4.b_cr10':              ('concentration.brand.CR10',                '%'),
    'm4.b_hhi':               ('concentration.brand.HHI',                 ''),
    'm4.b_n':                 ('concentration.brand.n',                   '个'),
    'm4.o_cr3':               ('concentration.oem.CR3',                   '%'),
    'm4.o_cr5':               ('concentration.oem.CR5',                   '%'),
    'm4.o_cr10':              ('concentration.oem.CR10',                  '%'),
    'm4.o_hhi':               ('concentration.oem.HHI',                   ''),
    'm4.o_n':                 ('concentration.oem.n',                     '家'),
    'm4.top_series':          ('concentration.top_series.0.0',            ''),
    'm4.top_series_sales':    ('concentration.top_series.0.1',            '辆'),
    'm4.top_series_pct':      ('concentration.top_series.0.2',            '%'),
    'm4.top_oem':             ('concentration.top_oem.0.0',               ''),
    'm4.top_oem_sales':       ('concentration.top_oem.0.1',               '辆'),
    'm4.top_oem_pct':         ('concentration.top_oem.0.2',               '%'),
    'm4.top_brand':           ('concentration.top_brand.0.0',             ''),
    'm4.top_brand_sales':     ('concentration.top_brand.0.1',             '辆'),
    'm4.top_brand_pct':       ('concentration.top_brand.0.2',             '%'),
    # --- 第五章 趋势与迁移 ---
    'm5.first':               ('trend.monthly_board.0.1',                 '辆'),
    'm5.last':                ('trend.monthly_board.9.1',                 '辆'),
    'm5.bev_first':           ('trend.energy_monthly.0.4',                '%'),
    'm5.bev_last':            ('trend.energy_monthly.9.4',                '%'),
    # 分动力份额与迁移（Lead m01005 锁定：+8.83pp，逐月 10/10 闭合、残差 0）
    'm5.bev_shift_pp':        ('CALC:BEV 份额 2026-09 − 2025-12（pp）',    'pp'),
    'm5.phev_shift_pp':       ('CALC:PHEV 份额 2026-09 − 2025-12（pp）',   'pp'),
    'm5.erev_shift_pp':       ('CALC:EREV 份额 2026-09 − 2025-12（pp）',   'pp'),
    'm5.phev_first':          ('CALC:2025-12 PHEV 份额（同分母）',         '%'),
    'm5.phev_last':           ('CALC:2026-09 PHEV 份额（同分母）',         '%'),
    'm5.erev_first':          ('CALC:2025-12 EREV 份额（同分母）',         '%'),
    'm5.erev_last':           ('CALC:2026-09 EREV 份额（同分母）',         '%'),
    'm5.bev_peak_month':      ('CALC:BEV 份额最高的月份',                  ''),
    'm5.bev_peak_pct':        ('CALC:BEV 份额峰值（%）',                   '%'),
    'm5.anchored_last':       ('trend.monthly.9.1',                       '辆'),
    'm5.price_2030_0303':     ('CALC:price_monthly 2026-03 的 20-30 万份额', '%'),    'm5.price_2030_0309':     ('CALC:price_monthly 2026-09 的 20-30 万份额', '%'),
    'm5.price_shift_pp':      ('CALC:20-30 万份额 2026-03→2026-09 变动(pp)', 'pp'),
    'm5.price_shift_3040_pp': ('CALC:30-40 万份额 2026-03→2026-09 变动(pp)', 'pp'),
    'm5.price_2030_2512':     ('CALC:2025-12（抢装月）20-30 万份额，仅作对照', '%'),
    'm5.entered_n':           ('CALC:len(trend.new_entries)',             '个'),
    'm5.dropped_n':           ('CALC:len(trend.drops)',                   '个'),
    'm5.union':               ('meta.qc.union_series',                    '个'),
    # --- 第六章 公告-市场转化（纲要结论 4） ---
    'm6.rows':                ('meta.base_rows',                          '行'),
    'm6.covered_rows':        ('coverage.row.covered',                    '行'),
    'm6.row_pct':             ('coverage.row.pct',                        '%'),
    'm6.board_pct':           ('coverage.current_board.pct',              '%'),
    'm6.board_sales_pct':     ('coverage.current_board.sales_pct',        '%'),
    'm6.batches':             ('CALC:len(conversion.by_batch)',           '个'),
    'm6.corr':                ('CALC:brand_scatter 皮尔逊相关系数',        ''),
    'm6.corr_r2':             ('CALC:上述相关系数的决定系数（r²，%）',     '%'),
    'm6.corr_n':              ('CALC:brand_scatter 样本量（品牌数）',      '家'),
    'm6.low_rate_brand':      ('CALC:rate_table 中公告车型数最多的品牌（大样本对照）', ''),
    # —— F3 效率离散度（Lead 终检订正：5.83 倍 ≠ 一个数量级，须给真实极差） ——
    'm6.eff_n':               ('CALC:可测算效率的品牌数（在售车系数>0）',        '家'),
    'm6.eff_median':          ('CALC:全部可测算品牌单车系均销中位数',           '辆'),
    'm6.eff_max_brand':       ('CALC:单车系均销最高的品牌',                      ''),
    'm6.eff_max':             ('CALC:上述品牌单车系均销',                        '辆'),
    'm6.eff_min_brand':       ('CALC:单车系均销最低的品牌（全样本）',            ''),
    'm6.eff_min':             ('CALC:上述品牌单车系均销（全样本，含极小样本）',   '辆'),
    'm6.eff_ratio':           ('CALC:全样本极差倍数（最高÷最低）',               '倍'),
    'm6.eff_ge10_n':          ('CALC:公告车型数≥10 的主流企业数',                '家'),
    'm6.eff_ge10_min_brand':  ('CALC:主流企业中单车系均销最低者',                ''),
    'm6.eff_ge10_min':        ('CALC:上述企业单车系均销',                        '辆'),
    'm6.eff_ge10_median':     ('CALC:主流企业单车系均销中位数',                  '辆'),
    'm6.eff_ge10_ratio':      ('CALC:主流企业极差倍数（最高÷最低）',             '倍'),
    'm6.eff_tesla_rows':      ('CALC:特斯拉公告车型数',                          '条'),
    'm6.eff_tesla_series':    ('CALC:特斯拉在售车系数',                          '个'),
    'm6.eff_tesla_sales':     ('CALC:特斯拉当期销量',                            '辆'),
    'm6.eff_tesla_rate':      ('CALC:特斯拉单车系均销',                          '辆'),
    'm6.eff_byd_rows':        ('CALC:比亚迪公告车型数',                          '条'),
    'm6.eff_byd_brand':       ('CALC:投放规模最大的对照品牌名',                  ''),
    'm6.eff_byd_series':      ('CALC:比亚迪在售车系数',                          '个'),
    'm6.eff_byd_sales':       ('CALC:比亚迪当期销量',                            '辆'),
    'm6.eff_byd_rate':        ('CALC:比亚迪单车系均销',                          '辆'),
    'm6.tesla_vs_byd':        ('CALC:特斯拉÷比亚迪 单车系均销倍数',              '倍'),
    'm6.low_rate_value':      ('CALC:该品牌单车系均销（辆/车系）',         '辆'),
    'm6.top_rate_brand':      ('CALC:rate_table 单车系均销最高的品牌',      ''),
    'm6.top_rate_value':      ('CALC:rate_table 单车系均销最高值',          '辆'),
    'm6.unlinked_series':     ('structure.unlinked_board.series_count',   '个'),
    'm6.unlinked_sales':      ('structure.unlinked_board.sales',          '辆'),
    'm6.unlinked_pct':        ('structure.unlinked_board.sales_pct',      '%'),
    'm6.naming_series':       ('structure.unlinked_board.cause_bucket.命名未锚定.series', '个'),
    'm6.naming_sales':        ('structure.unlinked_board.cause_bucket.命名未锚定.sales',  '辆'),
    'm6.notrace_series':      ('structure.unlinked_board.cause_bucket.公告库无痕迹.series', '个'),
    'm6.notrace_sales':       ('structure.unlinked_board.cause_bucket.公告库无痕迹.sales',  '辆'),
    'm6.pending_naming':      ('coverage.detail.pending_naming_total',    '个'),
    # --- 第七章 判断与建议（待《核心论断与建议纲要》v1.0 定稿后按 R1~R3 改写） ---
    'm7.opportunity':         ('TEXT:纲要 v1.0 结论 1~5 的合成判断',        ''),
    'm7.risk':                ('TEXT:纲要 v1.0 结论 4 的不确定性面',        ''),
    'm7.hold':                ('coverage.status_counts.未匹配-命名待锚定', '行'),
    'm7.metric':              ('TEXT:纲要 §2 观测指标',                    ''),
    'm7.r1':                  ('TEXT:纲要 §2 R1',                          ''),
    'm7.r2':                  ('TEXT:纲要 §2 R2',                          ''),
    'm7.r3':                  ('TEXT:纲要 §2 R3',                          ''),
    'm7.r4':                  ('TEXT:纲要 §2 R4',                          ''),
}

# 允许出现在正文里的**结构性字面量**（章节号/表图号/日期/批次范围/口径常量）。
# 用正则而非枚举：结构性字面量形态有限，而枚举必然漏项（每加一句都可能踩空）。
LITERAL_WHITELIST_PATTERNS = [
    r'^[一二三四五六七八九十]+$',        # 章节号"一、""第二章"
    r'^[1-9]$',                          # 单章号"第1章"
    r'^20\d\d$',                         # 年份"2026"
    r'^20\d\d-?\d\d$',                   # 月份"2026-09""202609"
    r'^20\d\d-?\d\d-?\d\d$',             # 日期"2026-10-19"
    r'^\d{3,4}$',                        # 批次号/底表列数"341""410""5392"
    r'^(8|A)-(0|[1-9]\d?)$',             # 表号"表 8-1""表 A-2"
    r'^\d+(\.\d+)+$',                    # 版本号"4.9.36"
    r'^(10|15|20|30|40)$',               # 价格带边界"10-15万"
]
LITERAL_WHITELIST = {
    '2026-09', '2026-10-19', '2025-12', '2026', '341', '410', '5392',
    '10', '15', '20', '30', '40', '3', '5', '6', '8', '9',
}


def is_whitelisted_literal(tok):
    """结构性字面量判定（枚举 + 形态正则）。"""
    if tok in LITERAL_WHITELIST or tok.replace(',', '') in LITERAL_WHITELIST:
        return True
    return any(re.match(p, tok) for p in LITERAL_WHITELIST_PATTERNS)

# 非契约路径的占位前缀及其含义 —— render() 遇到这些前缀不取数、只留可读占位符。
# 阶段 2 收口目标：'NOW' 直接删（改用真实来源）、'NEW' 向 Lead 申请由生产方新增字段、
# 'CALC'/'TEXT' 在脚本内实现（派生量或纲要定稿文案）。
PLACEHOLDER_PREFIXES = {
    'NOW':  '当前会话实测值（底表/原始快照），阶段 2 改为显式来源',
    'NEW':  '契约待新增字段 —— 须向 Lead 申请由生产方新增，不得改名既有字段',
    'CALC': '脚本内派生（由既有契约字段计算），阶段 2 实现',
    'TEXT': '《核心论断与建议纲要》定稿文案，非数值',
}


def _pearson(xs, ys):
    """皮尔逊相关系数。样本 <3 或零方差返回 None（不猜、不用 0 冒充）。"""
    n = len(xs)
    if n < 3:
        return None
    mx = sum(xs) / n
    my = sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    if sxx <= 0 or syy <= 0:
        return None
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    return sxy / (sxx ** 0.5 * syy ** 0.5)


_BASE_ROWS_CACHE = {}
_PRIOR_ROWS_CACHE = {}
_HEAD_ROWS_CACHE = {}


def _resolve_rel(p):
    """把 A-JSON 里的相对路径按仓库根解析（A-JSON 位于 audit-output/ 下）。"""
    if not p:
        return None
    if os.path.isabs(p):
        return p
    return os.path.normpath(
        os.path.join(os.path.dirname(os.path.abspath(ANALYTICS_JSON)), '..', p))


def _load_base_rows(aj):
    """现读底表（只读；read_only 模式，不写任何东西）。失败返回 None。"""
    path = _resolve_rel(dig(aj, 'meta.base_table_path'))
    if not path:
        return None
    if path in _BASE_ROWS_CACHE:
        return _BASE_ROWS_CACHE[path]
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:                      # noqa: BLE001 — 读不到即降级为占位
        print(f'[calc] 底表读取失败（报告相关披露将留占位）: {e}')
        return None
    try:
        ws = wb['NEV公告参数汇总']
        rows = list(ws.iter_rows(min_row=2, values_only=True))
    except Exception as e:                      # noqa: BLE001
        print(f'[calc] 底表 sheet 读取失败: {e}')
        wb.close()
        return None
    wb.close()
    _BASE_ROWS_CACHE[path] = rows
    return rows


def _load_prior_rows(aj):
    """现读旧值留档（meta.backup_old）。失败返回 None。"""
    path = _resolve_rel(dig(aj, 'meta.backup_old'))
    if not path:
        return None
    if path in _PRIOR_ROWS_CACHE:
        return _PRIOR_ROWS_CACHE[path]
    try:
        with open(path, encoding='utf-8') as f:
            bj = json.load(f)
        rows = bj.get('rows') if isinstance(bj, dict) else None
    except (OSError, ValueError) as e:
        print(f'[calc] 旧值留档读取失败（相关披露将留占位）: {e}')
        return None
    if not isinstance(rows, list):
        return None
    _PRIOR_ROWS_CACHE[path] = rows
    return rows


def _load_head_rows(aj):
    """从 **git 历史**取底表的"销量写入前"快照（只读）。

    为什么不用 meta.backup_old：该留档文件在 2026-10-19 19:32 被并发会话**覆盖**
    （prior_month 由 2026-08 变 2026-09、rows_with_prior_sales 由 806 变 2,185），
    原始的 806 格易车旧值证据已不在文件中。底表 xlsx 有 git 跟踪，因此改从 git 历史取。

    也不能直接取 HEAD：底表入库提交（如 d35891b "411批正式公告+412批公示入库"）
    之后，HEAD 版本已是 **41 列 / col36=2026-09** 的新态，不再是"写入前"。
    故按 SIZE_BASE_COMMIT 取基线；并**用两个特征自动校验**该版本确为写入前状态：
        col36（销量数据月份）== '2026-08'  且  列数 == 39
    校验不过就沿 `git log` 向前找最近的匹配提交（最多回溯若干条），
    找不到则返回 (None, None, None) → 相关披露留占位，绝不猜。

    返回 (rows, prev_col_count, prev_month)。
    """
    rel = dig(aj, 'meta.base_table_path')
    if not rel:
        return None, None, None
    if rel in _HEAD_ROWS_CACHE:
        return _HEAD_ROWS_CACHE[rel]
    import subprocess
    import tempfile

    def _git(*args):
        return subprocess.run(['git'] + list(args), stdout=subprocess.PIPE,
                              stderr=subprocess.DEVNULL, check=False)

    def _read(ref):
        tmp = os.path.join(tempfile.gettempdir(),
                           '_head_%d.xlsx' % (abs(hash(ref)) % 10 ** 8))
        try:
            with open(tmp, 'wb') as f:
                p = subprocess.run(['git', 'show', ref + ':' + rel],
                                   stdout=f, stderr=subprocess.DEVNULL, check=False)
            if p.returncode != 0 or os.path.getsize(tmp) == 0:
                return None
            import openpyxl
            wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
            ws = wb['NEV公告参数汇总']
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            wb.close()
            return rows or None
        except Exception:                       # noqa: BLE001
            return None
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    def _snapshot_state(rows):
        cols = len(rows[0]) if rows else 0
        month = None
        for r in rows or []:
            if len(r) > 36 and r[36]:
                month = r[36]
                break
        return cols, month

    # 候选提交序列：先 SIZE_BASE_COMMIT，再沿 git log 向前找
    refs = []
    if SIZE_BASE_COMMIT:
        refs.append(SIZE_BASE_COMMIT)
    log = _git('log', '--format=%H', '--', rel)
    if log.returncode == 0:
        for line in log.stdout.decode('utf-8', 'replace').splitlines():
            line = line.strip()
            if line and line not in refs:
                refs.append(line)

    for ref in refs[:12]:
        rows = _read(ref)
        if not rows:
            continue
        cols, month = _snapshot_state(rows)
        if cols == PREV_COLUMNS and str(month) == PREV_MONTH:
            out = (rows, cols, month)
            _HEAD_ROWS_CACHE[rel] = out
            return out
    print('[calc] git 历史中未找到"销量写入前"的底表快照'
          f'（期望 {PREV_COLUMNS} 列 / col36={PREV_MONTH}）——相关披露将留占位')
    return None, None, None


def _calc(key, aj):
    """派生量求值器（NUM_SRC 中以 'CALC:' 前缀登记）。

    只允许做**口径无损**的派生：同一契约字段内的加减、列表长度、极值、相关系数。
    返回 None 表示不可派生（render 会渲染为待办占位，绝不静默写 0）。
    """
    # ⚠️ `m1.legacy_cells` 在下方（与 m1.stopped_* 同组）统一从 **git 历史**派生。
    # 原先从 meta.backup_old 留档文件读的实现已删除：该文件 2026-10-19 19:32 被
    # 并发会话覆盖（prior_month 2026-08→2026-09、rows_with_prior_sales 806→2185），
    # 若继续用它会把 806 误报成 2185。

    if key in ('m5.bev_shift_pp', 'm5.phev_shift_pp', 'm5.erev_shift_pp'):
        # 分动力份额迁移：**同一分母**（已锚定合计 = BEV+PHEV+EREV）下的 pp 变动。
        # 三者合计恒为 0，是同一分母的镜像（Lead m01005 明确要求不要当三件事写）。
        rows = dig(aj, 'trend.energy_monthly') or []
        if len(rows) < 2:
            return None
        i = {'m5.bev_shift_pp': 1, 'm5.phev_shift_pp': 2, 'm5.erev_shift_pp': 3}[key]

        def share(r):
            tot = sum(x for x in r[1:4] if isinstance(x, (int, float)))
            return None if not tot else r[i] / tot * 100

        a, b = share(rows[0]), share(rows[-1])
        return None if a is None or b is None else b - a

    if key in ('m5.phev_first', 'm5.phev_last', 'm5.erev_first', 'm5.erev_last'):
        rows = dig(aj, 'trend.energy_monthly') or []
        if not rows:
            return None
        r = rows[0] if key.endswith('_first') else rows[-1]
        i = 2 if key.startswith('m5.phev') else 3
        tot = sum(x for x in r[1:4] if isinstance(x, (int, float)))
        return None if not tot else r[i] / tot * 100

    if key in ('m5.bev_peak_month', 'm5.bev_peak_pct'):
        rows = dig(aj, 'trend.energy_monthly') or []
        best, best_pct = None, None
        for r in rows:
            tot = sum(x for x in r[1:4] if isinstance(x, (int, float)))
            if not tot:
                continue
            p = r[1] / tot * 100
            if best_pct is None or p > best_pct:
                best, best_pct = r[0], p
        return (best if key.endswith('month') else best_pct)

    if key in ('m5.stopped_rows', 'm5.stopped_with_id', 'm5.col35_filled',
               'm5.col39_filled', 'm5.cleared_cells', 'm5.replaced_cells',
               'm5.kept_cells'):
        # 别名：统一走 m1.* 的底表现读实现（避免两套逻辑漂移）
        return _calc('m1.' + key[len('m5.'):], aj)

    if key in ('m1.last_batch_rows', 'm1.official_rows'):
        # 末批（公示口径）行数与正式公告记录数拆分。只依赖当前底表，不依赖写入前快照。
        base = _load_base_rows(aj)
        br = dig(aj, 'meta.batch_range') or []
        if base is None or len(br) != 2:
            return None

        def _bnum(v):
            try:
                return int(str(v).strip())
            except (TypeError, ValueError):
                return None

        last = sum(1 for r in base if len(r) > 0 and _bnum(r[0]) == _bnum(br[1]))
        return last if key == 'm1.last_batch_rows' else len(base) - last

    if key in ('m1.stopped_rows', 'm1.stopped_with_id', 'm1.col35_filled',
               'm1.col39_filled', 'm1.cleared_cells', 'm1.replaced_cells',
               'm1.kept_cells', 'm1.legacy_cells', 'm1.prev_month', 'm1.prev_cols',
               'm1.head_rows'):
        # 这些量**不在 A-JSON 里**，必须现读底表 + git HEAD 的"写入前"快照。
        # 行配对用 **(批次, 产品型号) 复合键**，不用行序 zip（底表在写入后被
        # 并发会话追加了 83 行，行序已不可靠），也不单用产品型号
        # （AGENTS 2026-09-17：跨批同码会错位）。键不唯一或配不到即返回 None。
        if key == 'm1.prev_month' or key == 'm1.prev_cols' or key == 'm1.head_rows':
            _h, cols, month = _load_head_rows(aj)
            if _h is None:
                return None
            if key == 'm1.prev_month':
                return month
            if key == 'm1.prev_cols':
                return cols
            return len(_h)
        base = _load_base_rows(aj)
        head, _cols, _month = _load_head_rows(aj)
        if base is None or head is None:
            return None
        if key == 'm1.col35_filled':
            return sum(1 for r in base if len(r) > 35 and r[35] not in (None, ''))
        if key == 'm1.col39_filled':
            return sum(1 for r in base if len(r) > 39 and r[39] not in (None, ''))
        if key == 'm1.stopped_rows':
            return sum(1 for r in base if len(r) > 40 and r[40] == '未匹配-已停产或未上量')
        if key == 'm1.stopped_with_id':
            return sum(1 for r in base if len(r) > 40 and r[40] == '未匹配-已停产或未上量'
                       and len(r) > 39 and r[39] not in (None, ''))

        def _key(r):
            return (str(r[0]).strip() if len(r) > 0 else '',
                    str(r[1]).strip() if len(r) > 1 else '')

        idx = {}
        for r in base:
            idx.setdefault(_key(r), []).append(r)
        total_head = sum(1 for r in head if len(r) > 35 and r[35] not in (None, ''))
        if key == 'm1.legacy_cells':
            return total_head
        replaced = cleared = kept = unmatched = 0
        for r in head:
            old = r[35] if len(r) > 35 else None
            if old in (None, ''):
                continue
            cand = idx.get(_key(r))
            if not cand or len(cand) != 1:
                unmatched += 1
                continue
            new = cand[0][35] if len(cand[0]) > 35 else None
            if new in (None, ''):
                cleared += 1
            elif abs(float(new) - float(old)) < 1e-9:
                kept += 1
            else:
                replaced += 1
        if unmatched:
            # 有配不上的行 → 宁可留占位，也不报一个偏小的数
            print(f'[calc] 旧值配对有 {unmatched} 行未唯一命中（披露将留占位）')
            return None
        return {'m1.cleared_cells': cleared, 'm1.replaced_cells': replaced,
                'm1.kept_cells': kept}[key]

    if key in ('m1.cross_over10_distinct', 'm1.cross_over10'):
        # over_10pct 在 JSON 中按"每格一条"重复列出（五菱宏光出现 4 次），
        # 对外主张"超 10% 偏差品牌数"必须按名称去重，否则会把 1 个说成 4 个。
        rows = dig(aj, 'meta.qc.cross_source.yiche_202608.over_10pct') or []
        return len({r[0] for r in rows if isinstance(r, (list, tuple)) and r})

    if key in ('m2.bev_pct', 'm2.phev_pct', 'm2.erev_pct'):
        rows = dig(aj, 'structure.by_energy') or []
        tot = dig(aj, 'structure.sales_total_dedup')
        want = {'m2.bev_pct': 'BEV', 'm2.phev_pct': 'PHEV', 'm2.erev_pct': 'EREV'}[key]
        if not tot:
            return None
        for r in rows:
            if r and r[0] == want:
                return r[1] / tot * 100
        return None

    if key.startswith('m3.d_'):
        metric = {'km': '续航', 'kwh': '容量', 'kwh100': '电耗',
                  'kg': '整备', 'kw': '功率', 'whkg': '能量密度'}[key[5:]]
        w = dig(aj, f'tech_sales.weighted.{metric}')
        u = dig(aj, f'tech_sales.unweighted.{metric}')
        return None if w is None or u is None else w - u

    if key in ('m5.entered_n', 'm5.dropped_n', 'm6.batches'):
        src = {'m5.entered_n': 'trend.new_entries', 'm5.dropped_n': 'trend.drops',
               'm6.batches': 'conversion.by_batch'}[key]
        rows = dig(aj, src)
        return len(rows) if isinstance(rows, list) else None

    if key in ('m5.price_shift_pp', 'm5.price_shift_3040_pp'):
        # ⚠️ 必须放在下方通用价格分支**之前**：'m5.price_shift_pp'.startswith('m5.price_')
        # 为真，否则会被当作价格带份额解析、正则不匹配后静默返回 None。
        band = '2030' if key == 'm5.price_shift_pp' else '3040'
        a = _calc(f'm5.price_pct_{band}_202603', aj)
        b = _calc(f'm5.price_pct_{band}_202609', aj)
        return None if a is None or b is None else b - a

    if key.startswith('m5.price_pct_') or key.startswith('m5.price_'):
        # 价格带份额（已锚定口径 price_monthly）。
        # NUM_SRC 注册形为 m5.price_<band>_<MMDD>（如 m5.price_2030_0303），
        # 内部形为 m5.price_pct_<band>_<YYYYMM>；用正则解析而非 split('_')
        # （后者会因为 m5.price 本身不含下划线而少一段，'0303' 被误当 band）。
        mm = re.search(r'_(10|1015|1520|2030|3040|40)_(\d{4}|\d{6})$', key)
        if not mm:
            return None
        band = {'10': '≤10万', '1015': '10-15万', '1520': '15-20万',
                '2030': '20-30万', '3040': '30-40万', '40': '40万+'}[mm.group(1)]
        tag = mm.group(2)
        if tag in ('0303', '0309', '2512'):
            mi = {'0303': 3, '0309': 9, '2512': 0}[tag]
        else:
            mi = {'202603': 3, '202609': 9, '202512': 0}[tag]
        rows = dig(aj, 'trend.price_monthly') or []
        if mi >= len(rows):
            return None
        row = rows[mi]
        vals = row[1:]
        tot = sum(v for v in vals if isinstance(v, (int, float)))
        if not tot:
            return None
        order = ['≤10万', '10-15万', '15-20万', '20-30万', '30-40万', '40万+']
        v = vals[order.index(band)]
        return v / tot * 100

    if key == 'm5.price_shift_pp':
        # 20-30 万 2026-03 → 2026-09 的份额变动（pp）——纲要结论 5 的主指标
        a = _calc('m5.price_pct_2030_202603', aj)
        b = _calc('m5.price_pct_2030_202609', aj)
        return None if a is None or b is None else b - a

    if key == 'm5.price_shift_3040_pp':
        a = _calc('m5.price_pct_3040_202603', aj)
        b = _calc('m5.price_pct_3040_202609', aj)
        return None if a is None or b is None else b - a

    if key == 'm5.price_pct_2030_202512':
        # 2025-12 基期（抢装月）—— 只作对照，不得作唯一基期（纲要 §4 第 9 条）
        band = {'≤10万': 0, '10-15万': 1, '15-20万': 2, '20-30万': 3, '30-40万': 4, '40万+': 5}['20-30万']
        rows = dig(aj, 'trend.price_monthly') or []
        row = rows[0]
        vals = row[1:]
        tot = sum(v for v in vals if isinstance(v, (int, float)))
        return None if not tot else vals[band] / tot * 100

    if key == 'm6.corr':
        # 公告车型数 × 当期销量：按品牌散点做皮尔逊相关（纲要结论 4 的关键证据）。
        # ⚠️ 实测 r ≈ 0.69（中等正相关），**不是纲要原文的"显著不相关"** ——
        #    报告按实测值写，并给出 r²（可解释的方差比例），
        #    再由"单车型产出差异达一个数量级"承接同一论点，不夸大也不改数。
        rows = dig(aj, 'conversion.brand_scatter') or []
        xs = [r[1] for r in rows if len(r) >= 3 and isinstance(r[1], (int, float))]
        ys = [r[2] for r in rows if len(r) >= 3 and isinstance(r[2], (int, float))]
        return _pearson(xs, ys) if len(xs) == len(ys) else None

    if key == 'm6.corr_r2':
        r = _calc('m6.corr', aj)
        # 保留 1 位：与 Lead 裁定的对外统一表述「可解释约 47.9%」逐字一致
        return None if r is None else round(r * r * 100, 1)

    if key == 'm6.corr_n':
        rows = dig(aj, 'conversion.brand_scatter') or []
        return len(rows) if rows else None

    if key.startswith('m6.eff_') or key in ('m6.tesla_vs_byd',):
        # 单车系效率离散度（Lead 终检 F3）：
        # 原稿用"特斯拉 32,418.5 vs 比亚迪 5,558.6 = 一个数量级"，实际仅 5.83 倍，
        # 属"论据不支撑结论"，会被心算推翻 → 改为按样本口径给出真实极差与中位数。
        # measurable = 在售车系数 > 0 的品牌（81 家）；
        # ge10 = 公告车型数 >= 10 的主流企业（60 家），避开 1~2 条公告的极端小样本。
        rows = [r for r in (dig(aj, 'conversion.rate_table') or [])
                if len(r) >= 5 and isinstance(r[2], (int, float)) and r[2] > 0
                and isinstance(r[4], (int, float))]
        if not rows:
            return None
        g10 = [r for r in rows if isinstance(r[1], (int, float)) and r[1] >= 10]
        by_brand = {r[0]: r for r in rows}
        med = lambda xs: statistics.median(xs) if xs else None
        if key == 'm6.eff_n':
            return len(rows)
        if key == 'm6.eff_median':
            return round(med([r[4] for r in rows]), 1)
        if key == 'm6.eff_max_brand':
            return max(rows, key=lambda r: r[4])[0]
        if key == 'm6.eff_max':
            return max(r[4] for r in rows)
        if key == 'm6.eff_min_brand':
            return min(rows, key=lambda r: r[4])[0]
        if key == 'm6.eff_min':
            return min(r[4] for r in rows)
        if key == 'm6.eff_ratio':
            lo = min(r[4] for r in rows)
            return round(max(r[4] for r in rows) / lo, 1) if lo else None
        if not g10:
            return None
        if key == 'm6.eff_ge10_n':
            return len(g10)
        if key == 'm6.eff_ge10_min_brand':
            return min(g10, key=lambda r: r[4])[0]
        if key == 'm6.eff_ge10_min':
            return min(r[4] for r in g10)
        if key == 'm6.eff_ge10_median':
            return round(med([r[4] for r in g10]), 1)
        if key == 'm6.eff_ge10_ratio':
            lo = min(r[4] for r in g10)
            return round(max(r[4] for r in g10) / lo, 1) if lo else None
        t, b = by_brand.get('特斯拉'), by_brand.get('比亚迪')
        if key == 'm6.tesla_vs_byd':
            return round(t[4] / b[4], 2) if t and b and b[4] else None
        if key.startswith('m6.eff_tesla_'):
            r = t
        elif key.startswith('m6.eff_byd_'):
            r = b
        else:
            return None
        if r is None:
            return None
        f = key.rsplit('_', 1)[1]
        return {'brand': r[0], 'rows': r[1], 'series': r[2],
                'sales': r[3], 'rate': r[4]}.get(f)

    if key in ('m6.low_rate_brand', 'm6.low_rate_value'):
        # 对照锚点：取 rate_table 中「公告车型数最多」的品牌（即投放规模最大的样本），
        # 与 top_rate（最高单车系产出）形成"数量" vs "效率"的同台对照。
        # 实测命中 比亚迪 478 条 / 5,558.6 辆，与特斯拉 32 条 / 32,418.5 辆对照。
        rows = dig(aj, 'conversion.rate_table') or []
        best = None
        for r in rows:
            if len(r) >= 5 and isinstance(r[1], (int, float)):
                if best is None or r[1] > best[1]:
                    best = r
        if best is None:
            return None
        return best[0] if key == 'm6.low_rate_brand' else best[4]

    if key == 'm6.top_rate_brand':
        rows = dig(aj, 'conversion.rate_table') or []
        best = None
        for r in rows:
            if len(r) >= 5 and isinstance(r[4], (int, float)):
                if best is None or r[4] > best[4]:
                    best = r
        return best[0] if best else None

    if key == 'm6.top_rate_value':
        rows = dig(aj, 'conversion.rate_table') or []
        best = None
        for r in rows:
            if len(r) >= 5 and isinstance(r[4], (int, float)):
                if best is None or r[4] > best[4]:
                    best = r
        return best[4] if best else None

    if key == 'm2.top_price_board':
        rows = dig(aj, 'structure.by_price_board') or []
        return max(rows, key=lambda r: r[1])[0] if rows else None

    if key.startswith('m3.erev_') and (key.endswith('_ratio') or key.endswith('_over')):
        # R1 口径下 EREV 相对 BEV 的比率（纲要仲裁 m00677）。
        # _ratio = EREV/BEV×100（"为 BEV 的 X%"）；_over = 该指数 − 100（"高 X%"）——
        # 两者语义不同，混用会把"为 128%"说成"高 128%"，是硬伤。
        base = key.rsplit('_', 1)[0]
        metric = {'m3.erev_km': '续航', 'm3.erev_kw': '功率',
                  'm3.erev_kg': '整备', 'm3.erev_kwh100': '电耗'}[base]
        rows = dig(aj, 'tech_sales.by_energy_tech') or []
        by = {r[0]: r[1] for r in rows if isinstance(r, list) and len(r) >= 2}
        bev, erev = by.get('BEV'), by.get('EREV')
        if not bev or not erev:
            return None
        b, e = bev.get(metric), erev.get(metric)
        if not b:
            return None
        idx = e / b * 100
        return idx if key.endswith('_ratio') else idx - 100

    if key in ('m1.mixed_n', 'm1.mixed_sales', 'm1.mixed_pct',
               'm1.r2_flip_n', 'm1.r2_flip_sales', 'm1.r2_flip_pct',
               'm1.mixed_erev_n', 'm1.mixed_erev_sales'):
        # JSON 中 mixed_energy_families / r2_flip_impact / mixed_with_erev 均为
        # **dict** {series, sales, pct_of_linked, items}，不是行列表。
        obj = {
            'm1.mixed_n': 'structure.mixed_energy_families',
            'm1.mixed_sales': 'structure.mixed_energy_families',
            'm1.mixed_pct': 'structure.mixed_energy_families',
            'm1.r2_flip_n': 'structure.r2_flip_impact',
            'm1.r2_flip_sales': 'structure.r2_flip_impact',
            'm1.r2_flip_pct': 'structure.r2_flip_impact',
            'm1.mixed_erev_n': 'structure.mixed_with_erev',
            'm1.mixed_erev_sales': 'structure.mixed_with_erev',
        }[key]
        node = dig(aj, obj)
        if not isinstance(node, dict):
            return None
        field = {'n': 'series', 'sales': 'sales', 'pct': 'pct_of_linked'}[key.rsplit('_', 1)[1]]
        return node.get(field)

    if key == 'm5.range_share_700':
        # 续航 ≥700km 区间的销量份额（纲要结论 2 的效率侧证据）
        defs = dig(aj, 'tech_sales.range_bins_def') or []
        rows = dig(aj, 'tech_sales.range_bins') or []
        if '≥700km' not in defs:
            return None
        i = defs.index('≥700km')
        return rows[i][1] if i < len(rows) else None

    return None


def resolve_num(key, aj, override=None):
    """注册表键 → 数值。未注册/未解析一律抛错（禁止静默 0）。

    override: 自测专用 —— {key: 合成值}，仅覆盖占位前缀(NOW/NEW/CALC/TEXT)与索引路径，
    绝不用于交付路径（交付时 override=None，缺失即 fail-fast）。
    """
    if key not in NUM_SRC:
        raise KeyError(f'NUM_SRC 未注册: {key}')
    path, unit = NUM_SRC[key]
    if override and key in override:
        return override[key], unit
    if path.startswith('CALC:'):
        return _calc(key, aj), unit
    if path.startswith('CFG:'):
        # 脚本自身配置常量（版本号 / 生成日期），非业务数据、非契约字段。
        # 生成日期用**运行日**，保证封面"生成"时间与实际一致。
        name = path.split(':', 1)[1]
        if name == 'ver':
            return GENERATOR_VERSION, unit
        if name == 'gen_date':
            return datetime.date.today().isoformat(), unit
        return None, unit
    return dig(aj, path), unit


def fmt_num(v, unit=''):
    """统一数字格式：千分位，整数不带小数，小数保留原精度。

    非数值（纲要定稿文案 TEXT 类）原样返回，不加单位。
    """
    if v is None:
        return '〔待填〕'
    if isinstance(v, str):
        return v
    if isinstance(v, float) and abs(v - round(v)) > 1e-9:
        s = f'{v:,.2f}'.rstrip('0').rstrip('.')
    else:
        s = f'{int(round(v)):,}'
    return f'{s}{unit}'


def render(tmpl, aj, strict=True, override=None):
    """把 ${key} 渲染为格式化数值。

    strict=True（默认，阶段 2 交付路径）：契约路径必须解析，缺失即抛错（消费方 fail-fast）；
    占位前缀（NOW/NEW/CALC/TEXT）渲染为可读占位符，并在 strict 下打印待办清单。
    override: 自测专用合成值（见 resolve_num）。
    """
    todos = []
    injected = []           # 由 NUM_SRC 注入的数值字面量 —— G1 必须豁免，否则自己的数据会被门禁拦下

    def rep(m):
        key = m.group(1)
        if key not in NUM_SRC:
            raise KeyError(f'NUM_SRC 未注册: {key}')
        path, unit = NUM_SRC[key]
        pre = path.split(':', 1)[0] if ':' in path else None
        if pre == 'CALC' and not (override and key in override):
            # 派生量：先尝试真实求值；求不出来（上游字段缺失）才降级为占位待办，
            # 绝不静默写 0 —— 这是 G2「禁止编造」在派生层的落实。
            val, unit = resolve_num(key, aj, override)
            if val is None:
                todos.append((key, 'CALC', path))
                return f'〔CALC:{key}〕'
            injected.append(fmt_num(val, ''))
            return fmt_num(val)
        if pre in PLACEHOLDER_PREFIXES and not (override and key in override):
            todos.append((key, pre, path))
            return f'〔{pre}:{key}〕'
        val, unit = resolve_num(key, aj, override)
        if val is None:
            if strict:
                raise KeyError(f'NUM_SRC[{key}] 契约路径未解析: {path}')
            todos.append((key, 'UNRESOLVED', path))
            return f'〔待填:{key}〕'
        if isinstance(val, (int, float)):
            injected.append(fmt_num(val, ''))
        elif isinstance(val, str):
            # 字符串型解析值（版本号/日期/数据源名）同样属于"脚本注入"，
            # 必须豁免 G1，否则 v4.9.34 / 2026-10-19 会被当成手写数字拦下。
            injected.append(val)
        # 单位由模板自带（"${x} 辆"）；NUM_SRC 的 unit 字段仅作表格表头/脚注元数据，
        # 不在此处追加，否则会渲染出"1,940辆 辆""13%%"这类重复单位。
        return fmt_num(val)

    out = re.sub(r'\$\{([A-Za-z0-9_.]+)\}', rep, tmpl)
    if todos:
        uniq = sorted({f'{p}·{k}' for k, p, _ in todos})
        print(f'[render] 占位待办 {len(uniq)} 项: {", ".join(uniq)}')
    # 占位符自身可能含数字（如 `〔CALC:m1.col39_filled〕` 里的 39、`NOW:x_806`），
    # 那是**未填写的标记**而不是手写业务数字，必须从 G1 豁免，否则门禁会误报。
    exempt = list(injected) + re.findall(r'〔[^〕]*〕', out)
    guard_no_literals(out, where=tmpl[:40], exempt=exempt)
    return out

def guard_no_literals(text, where='', exempt=()):
    """G1：模板渲染后仍出现的裸数字，必须全部在 LITERAL_WHITELIST 内。

    exempt: 由 NUM_SRC 注册表注入的数值字面量——这些正是授权的数据，必须豁免，
            否则门禁会把自己注入的真实数字当成"手写数字"拦下。
    数字 token 以「非数字/非小数点/非连字符开头」界定，避免把日期或批次范围
    （"2026-09"、"341~410"、"v4.9.36"）拆成碎片误报；日期/批次/章节号已在白名单显式登记。
    """
    ok = set(exempt)
    for tok in re.findall(r'(?<![\d.\-~])\d[\d,]*(?:\.\d+)?', text):
        if tok in ok or is_whitelisted_literal(tok):
            continue
        raise AssertionError(f'[G1] 正文出现未注册数字 "{tok}"（{where}）—— '
                             f'请改走 NUM_SRC 注册表')


# ============================================================================
# 3. 表格工具（严格沿用既有实现；新增 set_tc 绕过 vMerge 错位）
# ============================================================================
# markdown 加粗标记：叙事模板用它标重点，写入时须解析为真加粗 run。
# 若不解析，字面 `**` 会原样进入 docx（Lead 终检 F1 硬缺陷）。
MD_BOLD_RE = re.compile(r'\*\*(.+?)\*\*')


def split_bold(text):
    """把 `**粗体**` 解析为 [(片段, 是否加粗)]；无标记时返回单段。"""
    text = text or ''
    out, pos = [], 0
    for m in MD_BOLD_RE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out or [('', False)]


def _write_cell(cell, text, force_bold=None):
    """写单元格：`**` 解析为真加粗 run（不残留字面星号）。"""
    ps = cell.paragraphs
    for extra in ps[1:]:
        extra._element.getparent().remove(extra._element)
    set_rich_text(ps[0], str(text), force_bold=force_bold)
    return ps[0]


def add_table(doc, header, rows_data, style=STYLE_TABLE):
    t = doc.add_table(rows=1 + len(rows_data), cols=len(header))
    t.style = style
    rows_data = [list(rd[:len(header)]) + [''] * (len(header) - len(rd)) for rd in rows_data]
    for j, h in enumerate(header):
        _write_cell(t.rows[0].cells[j], h, force_bold=True)
    for i, rd in enumerate(rows_data, 1):
        for j, v in enumerate(rd):
            _write_cell(t.rows[i].cells[j], v)
    return t


def set_tc(t, ri, ci, txt):
    """绕过 vMerge 下 Table.cell() 的 _cells 索引错位（AGENTS 2026-09-29）。
    直接按 tr.tc_lst 定位 _Cell 后写入；`**` 同样解析为真加粗 run。"""
    from docx.table import _Cell
    tr = t._tbl.tr_lst[ri]
    tcs = tr.tc_lst
    tc = tcs[ci if ci < len(tcs) else -1]
    _write_cell(_Cell(tc, t), txt)


def set_rich_text(p, text, apply_font=True, force_bold=None):
    """把叙事文本写入段落，并将 `**...**` 渲染为**真加粗 run**。

    为什么必须做（Lead 终检 F1 硬缺陷）：叙事模板里用 `**` 标重点，
    直接 `paragraph.text = s` 会把星号**原样写进 docx**——领导层文档里
    出现字面 `**` 会显得"未完成"。故此处解析标记并拆成多个 run。
    """
    for r in list(p.runs):                       # 清空既有 run（保留 pPr/样式）
        r._element.getparent().remove(r._element)
    for seg, bold in split_bold(text):
        if seg == '':
            continue
        r = p.add_run(seg)
        r.font.bold = bold if force_bold is None else force_bold
        if apply_font:
            set_font(r)
    return p


def strip_md_bold(text):
    """剥离 markdown 加粗标记，只保留文字（供无法承载多 run 的位置降级使用）。"""
    return MD_BOLD_RE.sub(r'\1', text or '')


def assert_no_markdown(doc, where='docx'):
    """硬门禁 G11：全文（段落 + 全部表格单元格）不得残留字面 `**`。

    扫描范围是整篇文档而非仅新增分册——旧章、表头、单元格同样要查。
    """
    bad = []
    for i, p in enumerate(doc.paragraphs):
        if '**' in (p.text or ''):
            bad.append(('段落', i, p.text[:70]))
    for ti, t in enumerate(doc.tables):
        for ri, row in enumerate(t.rows):
            for ci, c in enumerate(row.cells):
                if '**' in (c.text or ''):
                    bad.append((f'表{ti}', f'行{ri}列{ci}', c.text[:70]))
    if bad:
        for b in bad[:20]:
            print(f'  [G11] 残留 markdown: {b[0]} {b[1]} → {b[2]!r}')
        raise AssertionError(
            f'[G11] 文档残留 markdown 加粗标记 `**` 共 {len(bad)} 处（{where}）')
    print(f'[G11] markdown 标记扫描 PASS —— 全文 0 处残留 `**`'
          f'（段落 {len(doc.paragraphs)} + 表 {len(doc.tables)} 全部单元格）')
    return True


def set_font(run, name=FONT_CN, size=10.5):
    """设置 run 的中西文字体。

    ⚠️ 修正两处长期潜伏的缺陷（此前 `insert_para_at` 从未被调用，故一直未暴露）：
    ① `qn('rFonts:eastAsia')` 的 `rFonts` 不是合法命名空间前缀，必抛 KeyError
       ——正确前缀是 `w`，属性名是 `eastAsia`；
    ② 直接访问 `run._element.rPr.rFonts` 在 rPr/rFonts 尚未建立时为 None，
       应走 get_or_add 语义，而不是假定其存在。
    """
    run.font.name = name
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), name)


def _tail(doc):
    """返回 (parent, 当前 body 最后一个子元素)；用于把新建元素搬到正文尾部再搬走。"""
    body = doc.element.body
    return body, list(body)[-1]


def insert_table_at(doc, anchor_el, header, rows_data, style=STYLE_TABLE):
    """在 anchor_el 之前插入一张表。返回 (table, table_element)。"""
    t = add_table(doc, header, rows_data, style=style)
    el = t._tbl
    anchor_el.addprevious(el)          # 提升到目标位置
    return t, el


def insert_para_at(doc, anchor_el, text, style=STYLE_BODY):
    """在 anchor_el 之前插入一个段落。返回 (paragraph, element)。

    实现：先在尾部 add_paragraph 拿到带正确样式的 p，再 addprevious 提升。
    """
    p = doc.add_paragraph(text, style=style)
    el = p._p
    anchor_el.addprevious(el)
    for r in p.runs:
        set_font(r)
    return p, el


def find_heading_el(doc, text_prefix):
    """按标题文本前缀找 H1 元素（用于定位插入锚点）。"""
    from docx.text.paragraph import Paragraph
    for child in doc.element.body.iterchildren():
        if child.tag.endswith('}p'):
            p = Paragraph(child, doc)
            st = p.style.name if p.style else ''
            if st == STYLE_H1 and p.text.strip().startswith(text_prefix):
                return child
    return None


# ============================================================================
# 4. 表格注册表 —— 表号 → 生成函数（数据源 = A-JSON）
#    冻结：新章占用 T60~T81（见大纲 §0.4/§3）。禁止在 T01~T59 之间插表。
# ============================================================================
TABLE_REGISTRY = [
    # (表ID, 大纲表号, 章节, 标题, 生成函数名)
    ('T60', '表 8-0',       '摘要',        '销量与市场表现速览卡',      't_summary_card'),
    ('T61', '表 8-1',       '8.1',        '数据源与快照信息',          't_source_card'),
    ('T62', '表 8-2',       '8.1',        '覆盖率三口径与未覆盖分桶',   't_coverage'),
    ('T63', '表 8-3',       '8.2',        '销量结构：动力类型',        't_struct_powertrain'),
    ('T64', '表 8-4',       '8.2',        '销量结构：细分市场 Top15',  't_struct_segment'),
    ('T65', '表 8-5',       '8.2',        '销量结构：价格带',          't_struct_price'),
    ('T66', '表 8-6',       '8.2',        '销量结构：企业 Top15',      't_struct_enterprise'),
    ('T67', '表 8-7',       '8.2',        '销量结构：品牌 Top15',      't_struct_brand'),
    ('T68', '表 8-8',       '8.3',        '销量加权 vs 车系均值（六指标）', 't_tech_weighted'),
    ('T69', '表 8-9',       '8.3',        '续航分箱 × 销量',           't_tech_rangebins'),
    ('T70', '表 8-10',      '8.3',        '价格带 × 续航段 × 销量矩阵', 't_price_matrix'),
    ('T71', '表 8-11',      '8.3',        '价格带技术均值',            't_price_tech'),
    ('T72', '表 8-12',      '8.4',        '三层级集中度 CR/HHI',       't_concentration'),
    ('T73', '表 8-13',      '8.4',        '企业份额梯队',              't_share_ladder'),
    ('T74', '表 8-14',      '8.5',        '月度总量与结构演进',        't_trend'),
    ('T75', '表 8-15',      '8.5',        '排位变动 rankchange 分布',  't_rankchange'),
    ('T76', '表 8-16',      '8.5',        '新晋 / 掉榜车系',           't_enter_exit'),
    ('T77', '表 8-17',      '8.6',        '批次公告-市场转化效率',     't_conv_batch'),
    ('T78', '表 8-18',      '8.6',        '转化率分档汇总',            't_conv_bucket'),
    ('T79', '表 8-19',      '8.7',        '判断-依据-建议对照',        't_judgement'),
    ('T80', '表 A-1',       '附录A',      '覆盖率三口径卡',            't_cov_card'),
    ('T81', '表 A-2',       '附录A',      '未覆盖分桶明细',            't_cov_buckets'),
]

# 章节函数骨架 —— 阶段 2 按大纲 §3 填充实现（此处仅占位，禁止写死数字）
def _todo(*_a, **_k):
    """占位实现：阶段 2 前调用即明确失败（不产出半成品）。"""
    raise NotImplementedError('phase2')


def is_stub(fn):
    """判定是否为未实现的占位函数（用函数对象身份判定，不依赖 docstring）。"""
    return getattr(fn, '__name__', '') == '_todo'


# ---------------------------------------------------------------------------
# 表格实现（阶段 2）——每张表返回 (header, rows_data)，全部数字取自 A-JSON。
# 硬约束：本区块**禁止出现写死的业务数字**；标题一律经 tbl_caption() 带口径与月份。
# ---------------------------------------------------------------------------
_DASH = '—'


def _n(v, nd=0):
    """整数千分位；None → 破折号。"""
    if v is None:
        return _DASH
    if isinstance(v, str):
        return v
    if nd == 0:
        return f'{int(round(float(v))):,}'
    return f'{float(v):,.{nd}f}'


def _pct(v, nd=2):
    return _DASH if v is None else f'{float(v):.{nd}f}%'


def _f(v, nd=2):
    return _DASH if v is None else f'{float(v):,.{nd}f}'


def tbl_caption(label, title, note='（2026-09 · 车系级去重口径）'):
    """表标题：必须含口径与月份（纲要 §4 第 6 条，G8 门禁校验）。"""
    return f'{label} {title}{note}'


def _top(pairs, k=None, pct_of=None):
    """[[键, 值], ...] → 前 k 行；pct_of 非空时补一列占比。"""
    rows = []
    for i, r in enumerate(pairs if k is None else pairs[:k]):
        key, val = r[0], r[1]
        if pct_of:
            rows.append([f'{i + 1}', key, _n(val), _pct(val / pct_of * 100)])
        else:
            rows.append([f'{i + 1}', key, _n(val)])
    return rows


def t_summary_card(aj):
    """T60 摘要速览卡。"""
    cb = dig(aj, 'coverage.current_board') or {}
    _ub = dig(aj, 'structure.unlinked_board') or {}
    ub_s, ub_v, ub_p = _ub.get('series_count'), _ub.get('sales'), _ub.get('sales_pct')
    rows = [
        ['数据源', dig(aj, 'meta.source') or _DASH, '车系级月度零售量'],
        ['快照月份 / 抓取日', f'{dig(aj, "meta.snapshot_month")} / {dig(aj, "meta.fetch_date")}',
         dig(aj, 'meta.source_tip') or _DASH],
        ['全榜新能源零售量', f'{_n(dig(aj, "structure.board_sales_total"))} 辆',
         f'全榜 {dig(aj, "structure.board_series_count")} 个在售车系'],
        ['已锚定车系 / 销量', f'{_n(cb.get("linked_series"))} 个 / {_n(cb.get("linked_sales"))} 辆',
         f'占全榜销量 {_pct(cb.get("sales_pct"))}'],
        ['行级覆盖率', _pct(dig(aj, 'coverage.row.pct')),
         f'{_n(dig(aj, "coverage.row.covered"))} / {_n(dig(aj, "coverage.row.total"))} 条公告记录'],
        ['车系 CR10（全榜）', _pct(dig(aj, 'concentration.series_board.CR10')),
         '产品层面高度分散'],
        ['企业 CR10', _pct(dig(aj, 'concentration.oem.CR10')),
         f'HHI {_f(dig(aj, "concentration.oem.HHI"), 0)}'],
        ['加权 vs 等权续航', f'{_f(dig(aj, "tech_sales.weighted.续航"), 1)} / '
                             f'{_f(dig(aj, "tech_sales.unweighted.续航"), 1)} km',
         '市场实际购买 vs 公告池等权'],
        ['2026-09 最大价格带', (dig(aj, 'structure.by_price_pct') or [['', 0]])[0][0],
         f'占已锚定 {_pct((dig(aj, "structure.by_price_pct") or [["", 0]])[0][1])}'],
        ['在榜未关联车系', f'{_n(ub_s)} 个 / {_n(ub_v)} 辆',
         f'占全榜销量 {_pct(ub_p)}（分析边界，非数据缺陷）'],
        ['公告存量→当期转化', _pct(dig(aj, 'coverage.row.pct')),
         f'{dig(aj, "meta.batch_count")} 个批次累积 '
         f'{_n(dig(aj, "meta.base_rows"))} 条公告/公示记录'],
    ]
    return ['指标', '数值', '口径说明'], rows


def t_source_card(aj):
    """T61 数据源与快照信息。"""
    ms = dig(aj, 'meta.months') or []
    rng = f'{ms[0]} ~ {ms[-1]}（{len(ms)} 个月）' if ms else _DASH
    br = dig(aj, 'meta.batch_range') or []
    rows = [
        ['快照月份', dig(aj, 'meta.snapshot_month') or _DASH],
        ['抓取日期', dig(aj, 'meta.fetch_date') or _DASH],
        ['数据源', dig(aj, 'meta.source') or _DASH],
        ['官方更新提示', dig(aj, 'meta.source_tip') or _DASH],
        ['口径范围', dig(aj, 'meta.scope') or _DASH],
        ['参照宇宙（仅校验，不入库）', dig(aj, 'meta.reference_source') or _DASH],
        ['快照区间', rng],
        ['公告批次范围', f'第 {br[0]}~{br[1]} 批（共 {dig(aj, "meta.batch_count")} 批）'
                         if len(br) == 2 else _DASH],
        ['公告批次口径说明', dig(aj, 'meta.batch_range_note') or _DASH],
        ['底表规模', f'{_n(dig(aj, "meta.base_rows"))} 行 × '
                     f'{dig(aj, "meta.base_table_columns")} 列'],
        ['价格带分箱规则', dig(aj, 'meta.price_bin_rule') or _DASH],
        ['去重规则', dig(aj, 'meta.dedup_rule') or _DASH],
    ]
    return ['项目', '内容'], rows


def t_coverage(aj):
    """T62 覆盖率三口径与未覆盖分桶。"""
    cb = dig(aj, 'coverage.current_board') or {}
    cs = dig(aj, 'coverage.detail') or {}
    rows = [
        ['行级', _n(dig(aj, 'coverage.row.covered')), _n(dig(aj, 'coverage.row.total')),
         _pct(dig(aj, 'coverage.row.pct')), '公告记录是否拿到销量'],
        ['当期在榜车系级', _n(cb.get('linked_series')), _n(cb.get('board_series')),
         _pct(cb.get('pct')), '对外主张"在售覆盖率"的合法口径'],
        ['销售额', _n(cb.get('linked_sales')), _n(cb.get('board_sales')),
         _pct(cb.get('sales_pct')), '已锚定车系贡献的全榜销量占比'],
        ['车系级（去重通用名称）', _n(cs.get('generic_names_anchored')),
         _n(cs.get('generic_names_denominator')), _pct(dig(aj, 'coverage.series.generic_names_pct')),
         '分母已剔除无车系身份的通用名称'],
    ]
    for k in ('无车系身份', '未匹配-已停产或未上量', '未匹配-命名待锚定', '未匹配-无车系ID'):
        v = dig(aj, f'coverage.buckets.{k}')
        rows.append(['【未覆盖分桶】' + k, _n(v), _n(dig(aj, 'meta.base_rows')),
                     _pct(v / dig(aj, 'meta.base_rows') * 100) if v is not None
                     and dig(aj, 'meta.base_rows') else _DASH, '四类分桶互斥且合计=底表行数'])
    return ['口径', '分子', '分母', '比例', '适用场景'], rows


def t_struct_powertrain(aj):
    """T63 销量结构：动力类型。"""
    tot = dig(aj, 'structure.sales_total_dedup')
    rows = [[k, _n(v), _pct(v / tot * 100) if tot else _DASH]
            for k, v in (dig(aj, 'structure.by_energy') or [])]
    return ['动力类型', '销量(辆)', '占比'], rows


def t_struct_segment(aj):
    """T64 销量结构：细分市场 Top15。"""
    return (['排名', '细分市场', '销量(辆)'],
            _top(dig(aj, 'structure.by_segment') or [], 15))


def t_struct_price(aj):
    """T65 销量结构：价格带（已锚定口径 + 全榜口径并列）。"""
    b = dict((x[0], x[1]) for x in (dig(aj, 'structure.by_price') or []))
    pb = dict((x[0], x[1]) for x in (dig(aj, 'structure.by_price_board') or []))
    tot = dig(aj, 'structure.sales_total_dedup') or 0
    totb = dig(aj, 'structure.board_sales_total') or 0
    rows = []
    for k, _v in (dig(aj, 'structure.by_price') or []):
        rows.append([k, _n(b.get(k)), _pct(b[k] / tot * 100) if tot else _DASH,
                     _n(pb.get(k)), _pct(pb[k] / totb * 100) if totb else _DASH])
    return ['价格带', '已锚定销量(辆)', '已锚定占比', '全榜销量(辆)', '全榜占比'], rows


def t_struct_enterprise(aj):
    """T66 销量结构：企业 Top15。"""
    return (['排名', '企业', '销量(辆)'],
            _top(dig(aj, 'structure.by_oem') or [], 15))


def t_struct_brand(aj):
    """T67 销量结构：品牌 Top15。"""
    return (['排名', '品牌', '销量(辆)'],
            _top(dig(aj, 'structure.by_brand') or [], 15))


_METRICS = (('续航', 'km', 1), ('容量', 'kWh', 1), ('电耗', 'kWh/100km', 2),
            ('整备', 'kg', 0), ('功率', 'kW', 1), ('能量密度', 'Wh/kg', 1))


def t_tech_weighted(aj):
    """T68 销量加权 vs 车系均值（六指标）。"""
    w = dig(aj, 'tech_sales.weighted') or {}
    u = dig(aj, 'tech_sales.unweighted') or {}
    rows = []
    for k, unit, nd in _METRICS:
        wv, uv = w.get(k), u.get(k)
        d = (wv - uv) if (wv is not None and uv is not None) else None
        rows.append([f'{k}（{unit}）', _f(wv, nd), _f(uv, nd), _f(d, nd),
                     _pct(d / uv * 100) if d is not None and uv else _DASH])
    sn = dig(aj, 'tech_sales.sample_n') or {}
    rows.append(['样本（车系数）', _n(sn.get('weighted')), _n(sn.get('unweighted')),
                 _DASH, _DASH])
    return ['技术指标', '销量加权', '公告等权', '差值', '相对差异'], rows


def t_tech_rangebins(aj):
    """T69 续航分箱 × 销量。"""
    rows = [[r[0], _pct(r[1]), _n(r[2])] for r in (dig(aj, 'tech_sales.range_bins') or [])]
    return ['续航段', '销量占比', '车系数'], rows


def t_price_matrix(aj):
    """T70 价格带 × 续航段 × 销量矩阵。"""
    cols = dig(aj, 'tech_sales.range_bins_def') or []
    header = ['价格带'] + [f'{c}\n销量(辆)' for c in cols] + ['合计(辆)']
    rows = []
    for r in (dig(aj, 'tech_sales.price_range_matrix') or []):
        sales = r[1] if len(r) > 1 and isinstance(r[1], list) else []
        rows.append([r[0]] + [_n(v) for v in sales] + [_n(sum(sales))])
    return header, rows


def t_price_tech(aj):
    """T71 价格带技术均值。"""
    rows = []
    for r in (dig(aj, 'tech_sales.price_tech') or []):
        r = list(r) + [None] * (7 - len(r))
        rows.append([r[0], _f(r[1], 1), _f(r[2], 1), _f(r[3], 1), _f(r[4], 0),
                     _n(r[5]), _n(r[6])])
    return ['价格带', '加权续航(km)', '加权容量(kWh)', '加权功率(kW)',
            '加权整备(kg)', '销量(辆)', '车系数'], rows


def t_concentration(aj):
    """T72 三层级集中度 CR/HHI。"""
    rows = []
    for label, path in (('车系（全榜）', 'concentration.series_board'),
                        ('车系（已锚定）', 'concentration.series'),
                        ('品牌', 'concentration.brand'),
                        ('企业', 'concentration.oem')):
        c = dig(aj, path) or {}
        rows.append([label, _n(c.get('n')), _pct(c.get('CR3')), _pct(c.get('CR5')),
                     _pct(c.get('CR10')), _f(c.get('HHI'), 0)])
    return ['层级', '主体数', 'CR3', 'CR5', 'CR10', 'HHI'], rows


def t_share_ladder(aj):
    """T73 企业份额梯队。"""
    tot = dig(aj, 'concentration.oem.total') or dig(aj, 'structure.sales_total_dedup') or 0
    return (['排名', '企业', '销量(辆)', '份额'],
            _top(dig(aj, 'concentration.top_oem') or [], 15, pct_of=tot))


def t_trend(aj):
    """T74 月度总量与结构演进。"""
    pcts = dict((r[0], r[4]) for r in (dig(aj, 'trend.energy_monthly') or [])
                if len(r) > 4)
    boards = dict((r[0], r[1]) for r in (dig(aj, 'trend.monthly_board') or []))
    rows = []
    for r in (dig(aj, 'trend.monthly') or []):
        rows.append([r[0], _n(boards.get(r[0])), _n(r[1]), _n(r[2]), _n(r[3]), _n(r[4]),
                     _pct(pcts.get(r[0]))])
    return ['月份', '全榜销量(辆)', '已锚定合计(辆)', 'BEV(辆)', 'PHEV(辆)', 'EREV(辆)',
            'BEV 份额'], rows


def t_rankchange(aj):
    """T75 排位变动 rankchange 分布。"""
    rows = [[r[0], _n(r[1])] for r in (dig(aj, 'trend.rankchange_hist') or [])]
    return ['排位变动区间', '车系数'], rows


def t_enter_exit(aj):
    """T76 新晋 / 掉榜车系。"""
    rows = []
    for label, path in (('新晋', 'trend.new_entries'), ('掉榜', 'trend.drops')):
        for r in (dig(aj, path) or []):
            rows.append([label, r[0], _n(r[1])])
    return ['类型', '车系', '2026-09 销量(辆)'], rows


def t_conv_batch(aj):
    """T77 批次公告-市场转化效率。"""
    rows = []
    for r in (dig(aj, 'conversion.by_batch') or []):
        r = list(r) + [None] * (4 - len(r))
        rows.append([_n(r[0]), _n(r[1]) if r[1] is not None else _DASH, _n(r[2]), _n(r[3])])
    return ['批次', '年份', '公告车型数', '当期销量(辆)'], rows


def t_conv_bucket(aj):
    """T78 转化率分档汇总（按年 + 单车系产出 Top10）。"""
    rows = []
    for r in (dig(aj, 'conversion.by_year') or []):
        rows.append(['按年', r[0], _n(r[1]), _n(r[2])])
    rt = sorted((dig(aj, 'conversion.rate_table') or []),
                key=lambda r: -(r[4] if len(r) > 4 and isinstance(r[4], (int, float)) else 0))
    for r in rt[:10]:
        rows.append(['单车系产出 Top', r[0], _n(r[1]), _n(r[4], 1)])
    return ['分档', '分组', '公告车型数', '当期销量 / 单车系均销(辆)'], rows


def t_judgement(aj):
    """T79 判断-依据-建议对照。"""
    def txt(k):
        v = resolve_num(k, aj)
        return v if isinstance(v, str) else _DASH
    rows = [
        ['R1', '结论 4', txt('m7.r1'), txt('m7.metric')],
        ['R2', '结论 2', txt('m7.r2'), '加权 / 等权六项技术画像差值、分动力技术画像'],
        ['R3', '结论 3 / 5', txt('m7.r3'), '价格带份额迁移矩阵、企业 CR10 与 HHI'],
        ['R4', '结论 1 / 4', txt('m7.r4'), '公告库最新批次与发布月、销量快照月与抓取日、生成时间戳'],
    ]
    return ['建议', '对应结论', '建议内容', '观测指标'], rows


def t_cov_card(aj):
    """T80 覆盖率三口径卡。"""
    return t_coverage(aj)


def t_cov_buckets(aj):
    """T81 未覆盖分桶明细（含在榜未关联）。"""
    ub = dig(aj, 'structure.unlinked_board') or {}
    cb = ub.get('cause_bucket') or {}
    rows = []
    for k in ('无车系身份', '未匹配-已停产或未上量', '未匹配-命名待锚定', '未匹配-无车系ID'):
        rows.append(['行级未覆盖', k, _n(dig(aj, f'coverage.buckets.{k}')), _DASH])
    rows.append(['行级未覆盖', '合计', _n(dig(aj, 'coverage.row.total')), _DASH])
    rows.append(['在榜未关联', '车系合计', _n(ub.get('series_count')), _n(ub.get('sales'))])
    for k in ('命名未锚定', '公告库无痕迹'):
        c = cb.get(k) or {}
        rows.append(['在榜未关联', k, _n(c.get('series')), _n(c.get('sales'))])
    return ['类别', '分桶', '车系数 / 行数', '销量(辆)'], rows


def count_pending_tables():
    """注册表里有几张表还没实现（阶段 2 进度指示）。"""
    impl = {f.__name__: f for f in globals().values()
            if callable(f) and f.__name__.startswith('t_')}
    return [tid for tid, _lab, _sec, _t, fn in TABLE_REGISTRY
            if fn not in impl or is_stub(impl[fn])]


# ============================================================================
# 4b. 叙事模板 —— 每章「结论框」正文。**模板内禁止出现业务数字**：
#     所有数字写作 ${key}，由 render() 从 A-JSON 取（G1 门禁会拒绝裸数字）。
#     模板按任务书章节树顺序排列；阶段 2 依《核心论断与建议纲要》v1.0 定稿措辞。
# ============================================================================
NARRATIVE = [
    # ⚠️ 章节键命名依赖 D2 未决：'ch7'/'ch8' 对应"新章作为第八章"的编号方案；
    #    若 Lead 裁定新章插入为独立章（原第七章顺延），只需改标题文案，键名与模板不变。
    # ⚠️ 批次范围/行数一律用 ${...} 注入（规范 §11.4：禁止硬编码 5,392 与 341~411）。
    ('abstract', '摘要 · 销量与市场表现（结论文本）',
     '本报告的一句话主张：中国新能源乘用车市场已进入"渗透率过半后的结构竞争期"——'
     '总量创新高但增长不再普惠，胜负手从"有没有电车"转向"在哪个价格带、'
     '用什么技术效率、以多高的单车型产出"竞争。'
     '2026-09 新能源乘用车零售 ${abs.board_total} 辆（全榜 ${abs.board_series} 个在售车系），'
     '其中已与公告库建立关联的 ${abs.snapshot_sales} 辆。'
     '五条结论：'
     '① 总量与渗透率——政策换挡造成 2026 年初深坑，9 个月修复并创新高；'
     '② 技术路线——市场用钱投票的方向是"更长续航、更轻、更省电"，而非"更大电池、更高能量密度"，'
     '加权续航 ${abs.w_km} km vs 公告等权 ${abs.u_km} km；'
     '③ 竞争格局——企业端高度集中而车系端极度分散（全榜车系 CR10 = ${abs.series_cr10}%，'
     '企业 CR10 = ${abs.oem_cr10}%），是"集团军 + 车海"；'
     '④ 公告→市场转化——投放数量与市场效率是两种不同能力；'
     '⑤ 价格带迁移——20–30 万持续扩容，30–40 万被两端分流。'
     '口径前置声明：全部销量为车系级去重口径；按底表行级直接加总会虚高到 ${abs.naive_total} 辆'
     '（${abs.dedup_ratio} 倍），这是口径错误而非市场规模。'
     '关于覆盖率，本报告对每一个覆盖率都同时给出**三个分母**：'
     '行级 ${abs.row_pct}%（${abs.row_covered}/${abs.row_total}）、'
     '当期在榜车系级 **${abs.board_pct}%**（${abs.board_linked}/${abs.board_series_n}）、'
     '销售额 ${abs.board_sales_pct}%（${abs.snapshot_sales}/${abs.board_total}）。'
     '同一份数据为何三个数相差如此之远、每个数适用于什么场景，'
     '本报告在「数据与方法」一章逐项定义；'
     '任何引用本报告覆盖率的外部文稿，请连同分母一并引用。'),

    ('ch1', '第一章 数据与方法 · 结论框',
     '数据源为汽车之家「行业综合销量」新能源榜车系级月度零售量，快照月 2026-09，'
     '官方提示"每月 10 日左右更新"。'
     '本期完成一次口径统一：底表原 ${m1.legacy_cells} 格月销量为易车 2026-08 口径，'
     '本期统一替换为汽车之家 2026-09，以确保全列单一月份、单一来源——'
     '口径一致性优先于"只补空"。'
     '换源经交叉校验证明不断层：与旧源逐格比对的可比 ${m1.cross_compared} 格中，'
     '${m1.cross_same} 格在 1% 内一致；超出 10% 的仅 ${m1.cross_over10} 个品牌，'
     '成因是燃油/纯电归并口径差异。'
     '三条口径铁律：一是车系级去重（底表行不等于车型，同车系跨批次多行共享同一销量值，'
     '行级不可加总）；二是未匹配一律留空而非填写零值；'
     '三是多候选一律挂起待锚定而非硬猜。'
     '因此覆盖率必须同时报告三个口径——行级 ${m1.row_pct}%、'
     '当期在榜车系级 ${m1.board_pct}%（销售额覆盖 ${m1.board_sales_pct}%）、'
     '车系级去重通用名称 ${m1.series_pct}%——'
     '三者分母不同、含义不同，只报最高者会误导读者。'
     '另须披露一处分类口径：同车系同时提供纯电与增程版本者，'
     '按公告行级标签的众数归入其一，涉及 ${m1.mixed_n} 个混合动力家族、'
     '${m1.mixed_sales} 辆、占已锚定销量 ${m1.mixed_pct}%；'
     '未采用"任一行为增程即整系归增程"的规则，因后者会被单行标记翻转整系归属'
     '（该被否决方案会实际翻转 ${m1.r2_flip_n} 个车系、${m1.r2_flip_sales} 辆，'
     '占已锚定 ${m1.r2_flip_pct}%，是其成本证据）。'
     '未覆盖部分按四类分桶披露；行级"命名待锚定"共 ${m1.pending_rows} 行，'
     '对应 ${m1.pending_generic} 个通用名称已主动挂起（车系级 ${m1.pending_series} 个），'
     '等待人工确认后再入库，不接受"取最相似"的猜测性匹配。'
     '此处须澄清两处易被误读为"缺数"的地方，均属设计使然而非数据缺陷：'
     '其一，"未匹配-已停产或未上量"的 ${m1.stopped_rows} 行中有 ${m1.stopped_with_id} 行'
     '**已锚定车系ID**（该车系在 10 个月并集内、但不在 2026-09 当期榜），'
     '按"只写当期销量"的规则当期销量留空——既不写旧月值、也不写零；'
     '因此「车系ID」列非空数（${m1.col39_filled}）大于「月销量」列非空数'
     '（${m1.col35_filled}）是设计使然，不是缺数。'
     '其二，口径统一完成后有 ${m1.cleared_cells} 格旧值被清空'
     '（其车系不在 2026-09 榜，按同一规则留空）。'
     '本次改动采用**双路径互证**，可逐格审计与还原：'
     '① 由写入前底表副本重建的逐格留档，'
     '含 ${m1.legacy_cells} 格的旧值/现值/处置（'
     '${m1.replaced_cells} 格更替、${m1.cleared_cells} 格清空、'
     '${m1.kept_cells} 格保留），副本 sha256 可核；'
     '② 版本库写入前提交的底表快照（'
     '${m1.head_rows} 行 / ${m1.prev_cols} 列 / 销量月份 ${m1.prev_month}），'
     '可按 (批次, 产品型号) 复合键逐格还原。'
     '两条独立路径结果一致，三者和等于改动前的 ${m1.legacy_cells} 格。'
     '该改动可完整回退：写入前整本备份「archive/v4.9.32_pre-sales_backup_20261019.xlsx」'
     '保留改动前 ${m1.head_rows} 行 × ${m1.prev_cols} 列的全量内容，'
     '配合版本库提交可逐格还原至改动前状态。'
     '即"口径可审计"与"改动可回退"是两个独立的信任承诺，本报告两者均成立。'
     '另需说明批次口径：本底表覆盖第 ${m1.batch_lo}~${m1.batch_hi} 批'
     '（共 ${m1.batch_count} 批），其中末批 ${m1.batch_hi} 批为**公示口径**'
     '（尚无正式发文字号），引用批次范围时须连同此说明一并引用。'),

    ('ch2', '第二章 市场总量与结构 · 结论框',
     '2026-09 新能源乘用车全榜零售 ${m2.board_total} 辆；'
     '已锚定车系级合计 ${m2.total} 辆，其中纯电 ${m2.bev} 辆（${m2.bev_pct}%）、'
     '插混 ${m2.phev} 辆（${m2.phev_pct}%）、增程 ${m2.erev} 辆（${m2.erev_pct}%）。'
     '结构上，销量集中于少数细分市场与价格带：最大细分市场为 ${m2.top_segment}'
     '（${m2.top_segment_sales} 辆），全榜最大价格带为 ${m2.top_price_board}。'
     '与公告侧结构对照可见供需错位的方向性信号：公告车型数占比高的细分市场，'
     '其销量占比未必同比例高——这正是"公告了但没卖动"的量化表达。'
     '本节所有销量数字均为车系级去重口径。'),

    ('ch3', '第三章 技术-市场关联 · 结论框',
     '把销量作为权重后，技术画像与公告池算术均值出现系统性偏差（各 ${m3.n_w}/${m3.n_u} 车系）：'
     '加权续航 ${m3.w_km} km vs 等权 ${m3.u_km} km（差 ${m3.d_km} km）；'
     '加权整备 ${m3.w_kg} kg vs 等权 ${m3.u_kg} kg（差 ${m3.d_kg} kg）；'
     '加权电耗 ${m3.w_kwh100} vs 等权 ${m3.u_kwh100} kWh/100km（差 ${m3.d_kwh100}）；'
     '而加权电池容量 ${m3.w_kwh} kWh 反而低于等权 ${m3.u_kwh} kWh（差 ${m3.d_kwh}），'
     '能量密度 ${m3.w_whkg} vs ${m3.u_whkg} Wh/kg 基本持平。'
     '含义：畅销车并非"电池更大、参数更堆料"，而是用更轻的车身、更少的电量跑出更长的续航——'
     '竞争实质是效率竞争；以公告池（供给面）均值代替市场（需求面）会得出相反的方向性误判。'
     '分动力画像（销量加权，按公告行级标签众数归类）：'
     'BEV ${m3.bev_km} km / ${m3.bev_kwh} kWh / ${m3.bev_kwh100} / ${m3.bev_kg} kg / ${m3.bev_kw} kW；'
     'PHEV ${m3.phev_km} / ${m3.phev_kwh} / ${m3.phev_kwh100} / ${m3.phev_kg} / ${m3.phev_kw}；'
     'EREV ${m3.erev_km} / ${m3.erev_kwh} / ${m3.erev_kwh100} / ${m3.erev_kg} / ${m3.erev_kw}。'
     '据此，增程不是"长纯电续航"路线，而是"大功率 + 重车身 + 中等电池"路线：'
     '其纯电续航仅为 BEV 的 ${m3.erev_km_ratio}%；'
     '而功率为 BEV 的 ${m3.erev_kw_ratio}%、整备质量为 ${m3.erev_kg_ratio}%、'
     '电耗为 ${m3.erev_kwh100_ratio}%（均以 BEV = 100 的相对指数，'
     '即功率高 ${m3.erev_kw_over}%、整备高 ${m3.erev_kg_over}%、电耗高 ${m3.erev_kwh100_over}%）。'
     '必须强调：以上是**关联而非因果**——价格与产品定位同时驱动技术参数与销量，'
     '不能推出"卖得好的车技术更好"。'),

    ('ch4', '第四章 集中度与竞争格局 · 结论框',
     '同一时点三个口径的结构性落差是本报告最重要的格局观察：'
     '企业（${m4.o_n} 家）CR3/CR5/CR10 = ${m4.o_cr3}/${m4.o_cr5}/${m4.o_cr10}%，HHI ${m4.o_hhi}；'
     '品牌（${m4.b_n} 个）CR3/CR5/CR10 = ${m4.b_cr3}/${m4.b_cr5}/${m4.b_cr10}%，HHI ${m4.b_hhi}；'
     '车系（全榜 ${m4.s_n} 个）CR3/CR5/CR10 = ${m4.s_cr3}/${m4.s_cr5}/${m4.s_cr10}%，HHI ${m4.s_hhi}。'
     '含义：前十家企业掌握近八成市场，但没有任何单一车系能形成垄断；'
     '头部集团份额靠"多品牌 × 多车系"矩阵覆盖——竞争在集团层面高度收敛，'
     '在产品层面高度发散。'
     '三个 CR 必须同时给出：CR5 在车系级衡量"爆款集中度"，在企业级衡量"集团控制力"，'
     '两者不可互相替代，单给一个会得出相反的格局判断。'
     '代表作：企业口径首位 ${m4.top_oem}（${m4.top_oem_sales} 辆，${m4.top_oem_pct}%），'
     '品牌口径首位 ${m4.top_brand}（${m4.top_brand_sales} 辆，${m4.top_brand_pct}%），'
     '单一车系最高 ${m4.top_series}（${m4.top_series_sales} 辆，${m4.top_series_pct}%）。'),

    ('ch5', '第五章 趋势与迁移 · 结论框',
     '10 个月窗口内总量与结构同步演进：全榜月销从 ${m5.first} 辆到 ${m5.last} 辆。'
     '动力结构（**分母为已锚定口径 ${m5.anchored_last} 辆，占全榜 ${m6.board_sales_pct}%，'
     '不是全榜三分**——全榜的 BEV/PHEV/EREV 三分本报告不做，'
     '因为只锚定了公告库覆盖到的车系）：'
     '纯电份额自 2025-12 的 ${m5.bev_first}% 升至 2026-09 的 ${m5.bev_last}%，'
     '**${m5.bev_shift_pp} pp**；同期插混份额 ${m5.phev_first}%→${m5.phev_last}%'
     '（${m5.phev_shift_pp} pp）、增程 ${m5.erev_first}%→${m5.erev_last}%'
     '（${m5.erev_shift_pp} pp），两者合计降幅恰好对冲纯电升幅'
     '（同一分母的镜像，不是三件独立的事）。'
     '其间 ${m5.bev_peak_month} 触及 ${m5.bev_peak_pct}% 的高点后小幅回落至 ${m5.bev_last}%，'
     '说明**纯电替代已进入高位平台期而非单向陡升**——'
     '这与"渗透率过半后进入结构竞争期"的主判断一致。'
     '价格带迁移以政策后首个正常月 2026-03 为基期：'
     '20–30 万区间份额由 ${m5.price_2030_0303}% 升至 ${m5.price_2030_0309}%'
     '（${m5.price_shift_pp} pp），是唯一显著扩容区间；'
     '30–40 万区间同期变动 ${m5.price_shift_3040_pp} pp，明显收缩。'
     '以 2025-12 为基期会得到同向但幅度更大的结果（该月为购置税调整前的抢装月，'
     '不宜作唯一基期；该口径下 20–30 万份额为 ${m5.price_2030_2512}%）。'
     '榜单名册层面（2026-09 vs 2026-08 新能源榜名册集合差，非 Top20 位次差）：'
     '${m5.entered_n} 个车系新进入、${m5.dropped_n} 个车系退出，'
     '排位变动分布显示位次尚未固化。'
     '边界声明：本窗口为 10 个月（跨年），不具备年度周期长度，**季节性不可辨识**；'
     '在榜车系并集 ${m5.union} 个远多于单月在榜数，说明榜单主体持续轮换。'),

    ('ch6', '第六章 公告-市场转化效率 · 结论框',
     '公告存量（${m6.batches} 个批次累积 ${m6.rows} 条公告车型记录）与当期市场表现的转化率观察：'
     '仅 ${m6.covered_rows} 条（${m6.row_pct}%）在 2026-09 有对应车系销量；'
     '按车系口径，已锚定车系占当月在榜 ${m6.board_pct}%、贡献全榜销量的 ${m6.board_sales_pct}%。'
     '公告车型投放数量与当期销量呈**中等正相关**'
     '（Pearson r = ${m6.corr}，可解释约 ${m6.corr_r2}% 的销量差异，n = ${m6.corr_n} 品牌）；'
     '但剩余一半以上的差异来自投放数量之外的因素：'
     '单车型产出效率的离散度极大——公告记录≥10 条的 ${m6.eff_ge10_n} 家主流企业中，'
     '最高 ${m6.eff_max_brand} ${m6.eff_max} 辆/车型、'
     '最低 ${m6.eff_ge10_min_brand} ${m6.eff_ge10_min} 辆/车型，'
     '极差约 ${m6.eff_ge10_ratio} 倍，中位数仅 ${m6.eff_ge10_median} 辆/车型'
     '（全样本 ${m6.eff_n} 家可测算企业中位数为 ${m6.eff_median} 辆/车型），'
     '效率是比规模更陡峭的分化维度。'
     '投放规模与市场表现并不成正比：${m6.eff_byd_brand} 以 ${m6.eff_byd_rows} 条公告、'
     '${m6.eff_byd_series} 个车系取得 ${m6.eff_byd_sales} 辆'
     '（${m6.eff_byd_rate} 辆/车型）；'
     '${m6.eff_max_brand} 仅 ${m6.eff_tesla_rows} 条公告、${m6.eff_tesla_series} 个车系'
     '即取得 ${m6.eff_tesla_sales} 辆（${m6.eff_tesla_rate} 辆/车型），'
     '单车型效率为前者的 ${m6.tesla_vs_byd} 倍。'
     '这说明"投放能力"与"单品爆款能力"是两种不同能力。'
     '口径声明：本节的"当期销量"是 2026-09 统一快照，不是各批次发布当月的销量；'
     '本报告不按月切片追溯每个批次的发布当期表现，以保持全案统一快照口径这一一致性锚点。'
     '这是自觉的口径取舍：为 ${m6.batches} 个批次分别取不同月份销量会打破快照一致性，'
     '所得"发布当期表现"与全案其他章节不可比。'
     '必须同页披露的边界（本结论最大可信度风险）：'
     '未覆盖分桶中"已停产或未上量"与"命名待锚定"难以区分；'
     '另有 ${m6.unlinked_series} 个在榜车系（${m6.unlinked_sales} 辆，占全榜 ${m6.unlinked_pct}%）'
     '尚未与底表建立关联，其中命名未锚定 ${m6.naming_series} 个车系（${m6.naming_sales} 辆）、'
     '公告库无痕迹 ${m6.notrace_series} 个车系（${m6.notrace_sales} 辆）。'
     '这属于分析边界而非数据质量缺陷：前者可通过车系名锚定表消解，后者反映真实的产品来源差异。'
     '故本结论仅在已锚定口径成立，不得表述为"行业整体"。'),

    ('ch7', '第七章 面向领导层的判断与建议 · 结论框',
     '综合前六章，四条建议各自对应一条结论，并各配一个可跟踪的观测指标，不写口号。'
     'R1（← 结论 4）：把"公告→市场"转化作为产品投放的常态化考核指标，'
     '并把车系名锚定表（公告通用名称 ↔ 媒体车系名）作为长期维护的数据资产立项；'
     '观测指标＝季度"公告→当季上量"转化率、单车型销量密度（辆/公告车型）、'
     '命名未锚定车系数（当前 ${m6.naming_series}）。'
     'R2（← 结论 2）：产品定义评审改用销量加权技术画像作基准，而非公告池等权均值；'
     '观测指标＝加权/等权六项技术画像差值、分动力技术画像。'
     'R3（← 结论 3/5）：产品投放向 20–30 万区间集中，审慎评估 30–40 万新增投放；'
     '观测指标＝价格带份额迁移矩阵、企业 CR10 与 HHI。'
     'R4（← 结论 1/4）：建立"公告 + 销量"月度同频刷新机制；'
     '观测指标＝公告库最新批次与发布月、销量快照月与抓取日、交付物生成时间戳。'
     '判断层面：三条结构性判断可支撑决策——其一，市场已从渗透率驱动的普遍增长，'
     '转入份额再分配阶段，增量不再自动惠及所有参与者；'
     '其二，企业端高集中与产品端高分散并存，说明规模优势来自产品矩阵的组织效率，'
     '而非单款车型的参数领先；其三，真正的效率缺口在"公告了但没卖动"，'
     '这一缺口可量化、可考核、可逐年改善。'
     '风险层面：需同时关注三项不确定性——一是政策换挡的后续扰动'
     '（购置税调整已造成 2026 年初单月渗透率回落 20 个百分点量级，后续节奏不可预判）；'
     '二是数据边界，当期在榜车系中仍有 ${m6.unlinked_series} 个未与公告库建立关联'
     '（占全榜销量 ${m6.unlinked_pct}%），本报告的转化类判断仅在此口径内成立；'
     '三是技术路线分类口径，混合动力家族占比不低，'
     '任何单一标签归并都会带来结构性误差，结论应作方向性参考而非精确测算。'),

    ('ch8', '第八章 结论 · 回扣',
     '回扣五条结论：总量与结构（政策换挡后的深坑与修复）、技术路线分化'
     '（市场选择效率而非堆料）、集中度（集团收敛而产品发散）、'
     '公告-市场转化（投放数量不等于市场效率）、价格带迁移（20–30 万扩容、30–40 万承压）。'
     '一句话总结：中国新能源乘用车已进入结构竞争期，增长不再普惠，'
     '胜负手在价格带选择、技术效率与单车型产出。'
     '局限性声明：本报告为 10 个月窗口、车系级去重口径、未覆盖部分语义正确；'
     '所有关联性观察均不构成因果结论；样本不足的分组不予解读。'),

    ('appendix', '附录 A 口径与覆盖率披露 · 摘要',
     '本附录把全案口径一次性摊开，供外部引用时连同分母一并摘录。'
     '① 数据源：${m1.source}，车系级月度零售量，官方提示"${m1.source_tip}"，'
     '快照月 ${m1.snapshot}，抓取日 ${m1.fetch_date}。'
     '② 口径延续性（换源不改口径）：${m1.cross_compared} 格可比对中 '
     '${m1.cross_same} 格在 1% 内一致，唯一超 10% 者为燃油/纯电归并差异。'
     '③ 第三方交叉校验：懂车帝 ${m1.dcd_rows} 条可比 NA 车系 ${m1.dcd_comp} 个，'
     '${m1.dcd_within} 个在 1% 内一致；超 10% 偏差 ${m1.dcd_over} 项已逐项定性为口径差，'
     '无法解释的数值错误 ${m1.dcd_unexplained} 项。'
     '④ 车系级去重铁律：底表行 ≠ 车型，任何汇总必须以车系为统计单元。'
     '⑤ 覆盖率三口径必须同时出现（行级 ${m1.row_pct}% / '
     '在榜车系级 ${m1.board_pct}% / 销售额 ${m1.board_sales_pct}%）；'
     '其中自指比率口径（分母由分子派生）不构成覆盖率，本报告不予采用。'
     '⑥ 未覆盖分桶与在榜未关联分桶逐项披露，'
     '成因已分桶，属分析边界而非数据质量缺陷。'
     '⑦ 历史快照处置采用双路径互证，可逐格还原。'
     '⑧ 性质边界：公告口径、零售量而非批发量/上险量、'
     '价格为厂商指导价区间中值而非成交价。'),
]

# 写作禁区（纲要 §4）——禁用词表。命中即在 G8 门禁 fail。
FORBIDDEN_CAUSAL = [
    '销量高.*技术', '销量第一.*技术', '卖得好.*所以', '领先.*说明',
    '证明.*技术.*第一', '必然.*领先',
]
FORBIDDEN_RHETORIC = [
    '遥遥领先', '断层第一', '全面碾压', '碾压', '吊打', '空前', '史上最强',
    '毫无疑问', '毋庸置疑', '必将', '唯一正确',
]
FORBIDDEN_UNSOURCED = [
    '据悉', '业内人士认为', '有关人士', '内部消息', '据说',
]
FORBIDDEN_PROXY = [               # 公告车型数 ≠ 在售车型数
    r'在售车型数\s*[=＝]\s*公告',
    r'公告车型数\s*(即|就是)\s*在售',
]
# 禁区 7（纲要 v1.0 §4 / Lead m00348）：91.34% 自指覆盖率一律禁用 ——
# 该比率分母 381 = 分子 348 + 33，属同义反复，不构成任何信息。
FORBIDDEN_SELFREF = [
    r'91\.3\d*\s*%', r'91\.34', r'94\.30',
]
# 禁区 8（纲要 v1.0 §4）：129 个在榜未关联车系属**分析边界**，不得表述为数据缺失/质量问题。
FORBIDDEN_MISLABEL = [
    r'(129|1\d\d)\s*个?(在榜)?车系.{0,12}(数据缺失|质量缺陷|数据错误|数据问题)',
    r'(数据缺失|质量缺陷|数据错误|数据问题).{0,12}(129|1\d\d)\s*个?(在榜)?车系',
]
# 禁区 9（纲要 v1.0 §4）：价格结构迁移不得以 2025-12 作唯一基期（抢装月）。
# 判定方式：正文出现"2025-12"且同时出现"基期"却未提及 2026-03 → 违规（见 gate_base_period）。
# R1 仲裁附带（Lead m00677）：EREV 续航 = BEV 的 71% 是 R2 错误口径的产物，禁止出现。
FORBIDDEN_STALE_R2 = [
    r'增程.{0,10}续航.{0,10}(7[01]|7\d\.\d)\s*%',
    r'(7[01]|7\d\.\d)\s*%.{0,10}(BEV|纯电).{0,6}续航',
]


def gate_forbidden_text(text, where=''):
    """G8 门禁：写作禁区（纲要 v1.0 §4 九条）的文本类。返回违规清单。"""
    bad = []
    for pat in FORBIDDEN_CAUSAL:
        if re.search(pat, text):
            bad.append(('因果臆断', pat))
    for w in FORBIDDEN_RHETORIC:
        if w in text:
            bad.append(('不可量化修辞', w))
    for w in FORBIDDEN_UNSOURCED:
        if w in text:
            bad.append(('无源表述', w))
    for pat in FORBIDDEN_PROXY:
        if re.search(pat, text):
            bad.append(('公告数当在售数', pat))
    for pat in FORBIDDEN_SELFREF:
        if re.search(pat, text):
            bad.append(('自指覆盖率 91.34%/94.30% 已禁用', pat))
    for pat in FORBIDDEN_MISLABEL:
        if re.search(pat, text):
            bad.append(('把分析边界误标为数据质量缺陷', pat))
    for pat in FORBIDDEN_STALE_R2:
        if re.search(pat, text):
            bad.append(('R2 作废口径残留（EREV 续航≈71% BEV）', pat))
    if bad:
        print(f'[G8] 写作禁区命中 @ {where}:')
        for kind, hit in bad:
            print(f'     {kind}: {hit!r}')
    return bad


def gate_base_period(text):
    """禁区 9：价格结构迁移不得以 2025-12 作唯一基期。

    规则：若正文提及价格带迁移且出现"2025-12"，则必须同时出现"2026-03"，
    或显式注明该月为抢装月/仅作对照。返回违规描述或 None。
    """
    if '2025-12' not in text:
        return None
    mentions_migration = ('基期' in text) or ('迁移' in text) or ('pp' in text)
    if not mentions_migration:
        return None
    if '2026-03' in text:
        return None
    if ('抢装' in text) or ('仅作对照' in text) or ('不宜作唯一基期' in text):
        return None
    return '提及 2025-12 基期但既未给 2026-03 基期、也未注明该月为抢装月/仅作对照'


def gate_table_caption(caption):
    """G8 附加：图表标题必须含口径与月份（纲要 §4 第 6 条）。"""
    ok = ('2026-09' in caption) and ('车系级' in caption)
    return ok


# ---------------------------------------------------------------------------
# G10 残留旧数字扫描（Lead 裁决 D1 条件③）
# 增量注入会有"上一版遗留数字"，命中数必须为 0。若某处**必须**出现历史数字
# （如"原 5,392 行"的历史叙述），必须列入 LEGACY_WHITELIST 并给出理由，
# 门禁输出逐条列出 —— 不允许静默放过。
# ---------------------------------------------------------------------------
LEGACY_FORBIDDEN = [
    '5,392',        # 旧底表行数（现 5,475）
    '40.52',        # 旧行级覆盖率（现 40.18）
    '42.03',        # 自纠前的车系级去重覆盖率（现 41.99）
    '41.6%',        # 已作废的口径错配覆盖率
    '91.34',        # 自指比率，纲要 §4 明令禁用
    '11.69',        # 旧虚高倍数（现 11.75）
    '2,185',        # 旧 col35 非空数（现 2,200）
    '5,392 款',     # 旧规模表述（且"款"字当用"条"）
    'Top20 新晋',   # 口径错误：新晋/掉榜是名册集合差，非 Top20 位次差
]

# 白名单：值 → 允许出现的理由（必须具体，门禁会打印）
LEGACY_WHITELIST = {
    '5,392': '第一章 §口径统一 披露中"写入前底表快照 5,392 行 / 39 列"的历史叙述——'
             '这是改动前状态的正确记录，非残留错误数字',
}


def gate_legacy_numbers(texts, where='', whitelist=None):
    """扫描残留旧数字。返回 (命中清单, 白名单命中清单)。

    命中 = 既在 LEGACY_FORBIDDEN 里、又不在 whitelist 里。
    """
    wl = LEGACY_WHITELIST if whitelist is None else whitelist
    blob = '\n'.join(t for t in texts if isinstance(t, str))
    hits, allowed = [], []
    for tok in LEGACY_FORBIDDEN:
        if tok in blob:
            if tok in wl:
                allowed.append((tok, wl[tok]))
            else:
                hits.append(tok)
    if hits:
        print(f'[G10] 残留旧数字命中 @ {where}: {hits}')
    if allowed:
        print(f'[G10] 白名单放行 @ {where}:')
        for tok, why in allowed:
            print(f'       · {tok} —— 理由：{why}')
    return hits, allowed


def narrate(chapter_key, aj, check=True, override=None):
    """按 chapter_key 渲染结论框文本；check=True 时跑 G8 门禁（九条禁区）。

    G1 零字面量检查在 render() 内完成（必须带 injected 豁免，否则注册表注入的
    真实数字会被重复检查误判为手写数字）。
    """
    for k, _title, tmpl in NARRATIVE:
        if k == chapter_key:
            txt = render(tmpl, aj, override=override)
            if check:
                bad = gate_forbidden_text(txt, where=f'NARRATIVE[{k}]')
                if bad:
                    raise AssertionError(f'NARRATIVE[{k}] 命中写作禁区：{bad}')
                bp = gate_base_period(txt)
                if bp:
                    raise AssertionError(f'NARRATIVE[{k}] 禁区 9（基期）：{bp}')
            return txt
    raise KeyError(f'未注册的章节叙事键: {chapter_key}')


def selftest_overrides(aj):
    """自测专用：把占位前缀(NOW/NEW/CALC/TEXT)映射为合成值。

    只处理无法从合成夹具直接解析的键；其余键仍走契约路径（可验证路径正确性）。
    这些合成值**只存在于自测内存**，永不进入交付物。
    """
    yz = aj['meta']['qc']['cross_source']['yiche_202608']
    st = aj['structure']
    tsw, tsu = aj['tech_sales']['weighted'], aj['tech_sales']['unweighted']
    ov = {
        # NOW：底表实测值（阶段 2 改为显式来源）
        'm1.legacy_cells': 806,
        # 底表/git 历史派生（真实路径下由 _load_base_rows/_load_head_rows 现读；
        # 自测夹具没有真实底表文件，故显式给合成值——只验证模板与门禁，不碰真数据）
        'm1.prev_month': PREV_MONTH,
        'm1.prev_cols': PREV_COLUMNS,
        'm1.head_rows': 5392,
        'm1.stopped_rows': 2749,
        'm1.stopped_with_id': 418,
        'm1.col35_filled': 2200,
        'm1.col39_filled': 2618,
        'm1.cleared_cells': 49,
        'm1.replaced_cells': 757,
        'm1.kept_cells': 0,
        'm1.last_batch_rows': 58,
        'm1.official_rows': 5417,
        # CALC：相关系数族（自测夹具的 brand_scatter 样本不足以产生稳定解，
        # 故显式给合成值；真实路径由 _calc 从 conversion.brand_scatter 现算）
        'm6.corr': 0.69,
        'm6.corr_r2': 47.9,
        'm6.corr_n': 188,
        'm6.eff_n': 81,
        'm6.eff_median': 766.2,
        'm6.eff_max_brand': '特斯拉',
        'm6.eff_max': 32418.5,
        'm6.eff_min_brand': '甲',
        'm6.eff_min': 2.0,
        'm6.eff_ratio': 16209.2,
        'm6.eff_ge10_n': 60,
        'm6.eff_ge10_min_brand': '开沃',
        'm6.eff_ge10_min': 30.0,
        'm6.eff_ge10_median': 1127.6,
        'm6.eff_ge10_ratio': 1080.6,
        'm6.eff_tesla_rows': 32,
        'm6.eff_tesla_series': 2,
        'm6.eff_tesla_sales': 64837,
        'm6.eff_tesla_rate': 32418.5,
        'm6.eff_byd_brand': '比亚迪',
        'm6.eff_byd_rows': 478,
        'm6.eff_byd_series': 30,
        'm6.eff_byd_sales': 166758,
        'm6.eff_byd_rate': 5558.6,
        'm6.tesla_vs_byd': 5.83,
        'm6.low_rate_brand': '比亚迪',
        'm6.low_rate_value': 5558.6,
        # CALC：脚本内派生（自测里显式算一遍，验证 _calc 的口径与模板一致）
        'm1.cross_over10': len({r[0] for r in yz['over_10pct'] if r}),
        'm2.bev_pct': st['by_energy'][0][1] / st['sales_total_dedup'] * 100,
        'm2.phev_pct': st['by_energy'][1][1] / st['sales_total_dedup'] * 100,
        'm2.erev_pct': st['by_energy'][2][1] / st['sales_total_dedup'] * 100,
        'm2.top_price_board': max(st['by_price_board'], key=lambda r: r[1])[0],
        'm3.d_km': tsw['续航'] - tsu['续航'],
        'm3.d_kwh': tsw['容量'] - tsu['容量'],
        'm3.d_kwh100': tsw['电耗'] - tsu['电耗'],
        'm3.d_kg': tsw['整备'] - tsu['整备'],
        'm3.d_kw': tsw['功率'] - tsu['功率'],
        'm3.d_whkg': tsw['能量密度'] - tsu['能量密度'],
        'm3.range_share_700': aj['tech_sales']['range_bins'][-1][1],
        'm5.price_shift_pp': 3.54,
        'm5.price_shift_3040_pp': -2.09,
        'm5.price_2030_2512': 19.42,
        'm5.entered_n': len(aj['trend']['new_entries']),
        'm5.dropped_n': len(aj['trend']['drops']),
        'm6.batches': len(aj['conversion']['by_batch']),
        'm6.corr': -0.2,
        'm6.top_rate_brand': '特斯拉',
        'm6.top_rate_value': 32418.5,
        # NEW：契约待新增字段（须向 Lead 申请）
        'm5.price_2030_0303': 19.30,
        'm5.price_2030_0309': 22.84,
        # TEXT：纲要 v1.0 定稿文案（t_judgement / T79 使用；模板与建议口径一致）
        'm7.opportunity': '市场已从渗透率驱动的普遍增长转入份额再分配；'
                          '企业端高集中与产品端高分散并存，规模优势来自产品矩阵的组织效率；'
                          '真正的效率缺口在"公告了但没卖动"，可量化、可考核、可逐年改善。',
        'm7.risk': '政策换挡的后续扰动（购置税调整曾致单月渗透率回落二十个百分点量级）；'
                   '数据边界（部分在榜车系未与公告库关联，转化类判断仅在此口径内成立）；'
                   '技术路线分类口径（混合动力家族以单一标签归并会带来结构性误差）。',
        'm7.metric': '季度"公告→当季上量"转化率、单车型销量密度、命名未锚定车系数',
        'm7.r1': '把"公告→市场"转化作为产品投放的常态化考核指标，'
                 '并把车系名锚定表（公告通用名称 ↔ 媒体车系名）作为长期维护的数据资产立项。',
        'm7.r2': '产品定义评审改用销量加权技术画像作基准，而非公告池等权均值；'
                 '对增程路线需评估其"大电池 + 大功率 + 重车身"的成本与法规风险。',
        'm7.r3': '产品投放向 20–30 万区间集中，审慎评估 30–40 万新增投放；'
                 '≤10 万代步车市场容量仍在，可作走量与积分贡献位。',
        'm7.r4': '建立"公告 + 销量"月度同频刷新机制，'
                 '在看板固化"数据口径 + 最新更新日期"，'
                 '把夜间自动化管线延伸为月度闭环。',
    }
    return ov


# ============================================================================
# 5. 自动备份（硬约束③：备份名必须含版本号 + 时间戳，且存在性检查）
# ============================================================================
def backup_report(tag='pre-sales'):
    import datetime
    import shutil
    if not os.path.exists(REPORT_DOCX):
        return None
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    bak = os.path.join(ARCHIVE_DIR, f'报告_{tag}_{GENERATOR_VERSION}_backup_{stamp}.docx')
    if os.path.exists(bak):                       # 存在性检查：绝不覆盖既有备份
        print(f'[backup] 已存在，跳过: {bak}')
        return bak
    shutil.copy2(REPORT_DOCX, bak)
    assert os.path.getsize(bak) == os.path.getsize(REPORT_DOCX), 'backup 字节数不一致'
    print(f'[backup] {os.path.getsize(bak):,} B → {bak}')
    return bak


# ============================================================================
# 6. 阶段 1 工具：A-JSON 契约探针（现在即可运行，不依赖 task-2 完成后的内容）
# ============================================================================
def sales_probe():
    """只读探针：报告 A-JSON 是否就绪 + 契约齐备性 + 与原始快照的交叉核对。

    现在（阶段 1）预期输出：A-JSON MISSING → 退出码 2（这是正确行为，不是缺陷）。
    """
    print(f'=== build_report {GENERATOR_VERSION} · A-JSON 契约探针 ===')
    for p, label in ((ANALYTICS_JSON, 'A-JSON(分析底表)'), (AUTOHOME_JSON, '原始快照')):
        if os.path.exists(p):
            print(f'{label:16s} {os.path.getsize(p):>12,} B  {p}')
        else:
            print(f'{label:16s} {"MISSING":>12s}  {p}')
    if os.path.exists(AUTOHOME_JSON):
        raw = json.load(open(AUTOHOME_JSON, encoding='utf-8'))
        meta = raw.get('meta', {})
        per_month = meta.get('per_month', {}) or {}
        print(f'  原始快照: {len(meta.get("months", []))} 个月 {meta.get("months")}')
        print(f'  月榜车系数: ' + ' '.join(f'{k}={v.get("count")}' for k, v in per_month.items()))
        print(f'  union_series={meta.get("series_union_count")}  captured_at={meta.get("captured_at")}')
        tip = next((v.get('tip') for v in per_month.values() if v.get('tip')), None)
        print(f'  官方口径 tip: {tip!r}')
        if os.path.exists(REPORT_DOCX):
            print(f'  交付物现状: {os.path.getsize(REPORT_DOCX):,} B  {REPORT_DOCX}')
    if not os.path.exists(ANALYTICS_JSON):
        print('[probe] A-JSON 未就绪 → 阶段 2 阻塞在 task-2。契约定义见 '
              'Update/销量分析报告大纲_20261019.md §2。')
        return 2
    aj = load_analytics()
    assert_schema(aj)
    return 0


def sales_draft(out_path):
    """新章独立草稿（阶段 2 用；表生成函数未实现时明确失败，不产出半成品）。"""
    aj = load_analytics()
    assert_schema(aj)
    doc = Document()
    todo = count_pending_tables()
    print(f'[draft] 注册表 {len(TABLE_REGISTRY)} 张表；待实现 {len(todo)} 张 → 阶段 2 任务')
    doc.add_heading('销量分析与公告-市场转化（草稿）', level=1)
    doc.save(out_path)
    print(f'saved: {out_path}')


# ============================================================================
# 6b. 自测：用**契约形状的合成夹具**（全合成值，不入库、不写交付物）
#     验证 G0 契约门禁、G1 零裸数字、G8 写作禁区、渲染管线与 stub 检测。
#     运行：python build_report_v4925.py --sales-selftest
# ============================================================================
SELFTEST_SCALE = 0.6          # 集成长度因子（把 10 个月压成紧凑夹具）


def make_selftest_analytics():
    """构造契约形状的合成 A-JSON（**全部为合成值**，仅用于门禁自测）。"""
    months = ['2025-12', '2026-01', '2026-02', '2026-03', '2026-04',
              '2026-05', '2026-06', '2026-07', '2026-08', '2026-09']
    n_batches, base_rows = 70, 5392
    s = int(base_rows * SELFTEST_SCALE)          # 过大而无法直接累加的合成规模
    d = int(s * 0.5)                             # 去重后规模
    n = int(d * 1.3)
    status_counts = {
        '在售-已匹配': int(base_rows * 0.11), '在售-零销量': int(base_rows * 0.02),
        '未匹配-命名待锚定': 17, '未匹配-已停产或未上量': int(base_rows * 0.74),
        '无车系身份': 44, '未匹配-无车系ID': 3,
    }
    status_counts['在售-零销量'] += base_rows - sum(status_counts.values())
    cov_row = {'covered': int(base_rows * 0.13), 'total': base_rows, 'pct': 0.0}
    cov_row['pct'] = round(cov_row['covered'] / cov_row['total'] * 100, 1)
    cov_series = {'matched': int(n * 0.13), 'total': n, 'pct': 0.0}
    cov_series['pct'] = round(cov_series['matched'] / cov_series['total'] * 100, 1)
    cov_onsale = {'matched': int(n * 0.72), 'total': int(n * 0.75), 'pct': 0.0}
    cov_onsale['pct'] = round(cov_onsale['matched'] / cov_onsale['total'] * 100, 1)
    buckets = {'无车系身份': status_counts['无车系身份'],
               '未匹配-已停产或未上量': status_counts['未匹配-已停产或未上量'],
               '未匹配-命名待锚定': status_counts['未匹配-命名待锚定'],
               '未匹配-无车系ID': status_counts['未匹配-无车系ID']}
    monthly, energy_monthly, price_monthly, segment_monthly = [], [], [], []
    for i, m in enumerate(months):
        tot = int(d * (0.75 + 0.05 * i))
        bev = int(tot * (0.55 + 0.01 * i))
        erev = int(tot * 0.12)
        monthly.append([m, tot, bev, tot - bev])
        energy_monthly.append([m, bev, tot - bev - erev, erev,
                               round(bev / tot * 100, 1)])
        price_monthly.append([m, int(tot * 0.18), int(tot * 0.24), int(tot * 0.22),
                              int(tot * 0.19), int(tot * 0.10), int(tot * 0.07)])
        segment_monthly.append([m, {'紧凑型SUV': int(tot * 0.3), '中型轿车': int(tot * 0.2),
                                    '小型车': int(tot * 0.15)}])
    pair = lambda keys, v: [[k, v] for k in keys]
    tech_keys = ['续航', '容量', '电耗', '整备', '功率', '能量密度']
    weighted = dict(zip(tech_keys, [601.5, 76.4, 13.2, 1901.0, 208.3, 158.7]))
    unweighted = dict(zip(tech_keys, [489.7, 63.9, 13.7, 1823.9, 195.7, 152.7]))
    b = [f'{340 + i}批' for i in range(1, n_batches + 1)]
    bands = ['≤10万', '10-15万', '15-20万', '20-30万', '30-40万', '40万+']
    rbands = ['<300km', '300-400km', '400-500km', '500-600km', '600-700km', '≥700km']
    board_total = int(d * 1.37)                 # 全榜（含未锚定车系）> 已锚定合计
    cov_series.update({'generic_names_matched': int(n * 0.21), 'generic_names_total': n,
                       'generic_names_pct': round(int(n * 0.21) / n * 100, 1)})
    cov_onsale.update({'generic_names_matched': int(n * 0.21),
                       'generic_names_hold': 17,
                       'generic_names_pct': round(int(n * 0.21) / int(n * 0.22) * 100, 1)})
    return {
        'meta': {
            'snapshot_month': '2026-09', 'months': months, 'fetch_date': '2026-10-19',
            'source': '汽车之家 · 行业综合销量（车系级月度零售量）',
            'source_tip': '* 数据源于行业综合销量，每月10日左右更新',
            'source_url_template': 'https://example.invalid/getList?pageindex={page}',
            'reference_source': '懂车帝微型车榜', 'scope': '新能源榜（typeid=9&subranktypeid=2305）',
            'dedup_rule': '车系级去重：同车系跨批次多行共享同一值，行级不可加总',
            'price_bins': bands,
            'price_bin_rule': 'priceinfo 取中值后分箱',
            'backup_old': 'audit-output/sales_prior_yiche_202608_backup_20261019.json',
            'base_table_path': 'NEV公告参数汇总表_合并版（341~412批）.xlsx',
            'base_rows': base_rows, 'matched_rows': status_counts['在售-已匹配'],
            # 批次范围/批次数：与真实 A-JSON 同形（禁硬编码，报告一律从这三个键读）
            'batch_range': [341, 411], 'batch_count': n_batches,
            'batch_range_note': '合成夹具：末批 412 为公示口径（未正式公告）',
            'match_channel_counts': {'A_exact': 512, 'B_brand': 41, 'C_substr': 7,
                                     'D_suggest': 2, 'hold': 17},
            'qc': {'monthly_counts': {m: 340 + 7 * i for i, m in enumerate(months)},
                   'union_series': 507,
                   'cross_source': {
                       'yiche_202608': {'compared': 483, 'within_1pct': 479,
                                        'within_5pct': 479, 'unmatched': 74,
                                        'mean_rel_diff': 0.0009, 'over_10pct': []},
                       # 与真实 A-JSON 同形（WS1 全量口径，非旧 compared/mismatch）
                       'dongchedi_202609': {'dcd_rows': 660, 'matched': 323,
                                            'comparable_nev_count': 282, 'within_1pct': 278,
                                            'over_10pct_count': 18,
                                            'over_10pct_unexplained_count': 0}}},
            'dedup_selfcheck': {'case': '星愿', 'generic_names': ['星愿、星愿UP', '星愿STAR、星愿',
                                                                 '星愿、星愿Ultra', '星愿UP、星愿',
                                                                 '星愿、派派、龙猫'],
                                'resolved_seriesid': '7806', 'collapsed_rows': 16,
                                'distinct_series_in_top20': 15},
        },
        'coverage': {
            'row': cov_row, 'series': cov_series, 'on_sale': cov_onsale,
            'status_counts': status_counts, 'buckets': buckets,
            'pending_naming': [['待锚定样本A', 5, ['候选一', '候选二']],
                               ['待锚定样本B', 3, ['候选三']]],
            'definitions': {'row': '行级覆盖定义（合成夹具）', 'series': '车系级定义（合成夹具）',
                            'current_board': '当期在榜车系级定义（合成夹具）',
                            'self_consistent_ratio': '同义反复比率定义（合成夹具）',
                            'buckets': '分桶定义（合成夹具）',
                            'suggest': 'suggest 通道未启用（合成夹具）'},
            'detail': {'generic_names_all': n + 12, 'generic_names_no_identity': 12,
                       'generic_names_denominator': n,
                       'generic_names_anchored': int(n * 0.21),
                       'generic_names_hold': 17,
                       'generic_names_unmatched': n - int(n * 0.21) - 17,
                       'distinct_seriesid_anchored': int(n * 0.13),
                       'pending_naming_total': 17, 'pending_naming_shown': 2},
            # 键名与真实 A-JSON 一致（board_series/linked_series/linked_sales/board_sales）
            'current_board': {'board_series': int(n * 0.75),
                              'linked_series': int(n * 0.72),
                              'pct': round(int(n * 0.72) / max(1, int(n * 0.75)) * 100, 4),
                              'linked_sales': d, 'board_sales': board_total,
                              'sales_pct': round(d / board_total * 100, 4),
                              'definition': '合成夹具'},
            'self_consistent_ratio': {'matched': int(n * 0.70), 'total': int(n * 0.72),
                                      'pct': 0.0, 'definition': '合成夹具，禁用'},
        },
        'structure': {
            'by_segment': pair(['紧凑型SUV', '中型轿车', '小型车'], int(d * 0.3)),
            'by_segment_pct': pair(['紧凑型SUV', '中型轿车', '小型车'], 30.0),
            'by_energy': [['BEV', int(d * 0.66)], ['PHEV', int(d * 0.29)], ['EREV', int(d * 0.05)]],
            'by_price': pair(bands, int(d * 0.2)),
            'by_price_pct': pair(bands, 6.2),
            'by_oem': pair(['企业甲', '企业乙', '企业丙'], int(d * 0.15)),
            'by_brand': pair(['品牌甲', '品牌乙', '品牌丙'], int(d * 0.12)),
            'sales_total_dedup': d, 'sales_total_rowlevel_naive': s,
            'board_sales_total': board_total, 'board_series_count': int(n * 0.75),
            'linked_share_of_board_pct': round(d / board_total * 100, 2),
            'duplication_ratio': round(s / d, 3),
            'by_price_board': pair(bands, int(board_total * 0.2)),
            'by_price_board_pct': pair(bands, 16.6),
            # 三套混合家族清单均为 dict{series,sales,pct_of_linked,items}，与真实 A-JSON 同形
            'mixed_energy_families': {'series': 5, 'sales': int(d * 0.03),
                                      'pct_of_linked': 3.0,
                                      'items': [{'车系': '混合家族示例A', '销量': int(d * 0.02),
                                                 '占比': 2.0},
                                                {'车系': '混合家族示例B', '销量': int(d * 0.01),
                                                 '占比': 1.0}]},
            'r2_flip_impact': {'series': 2, 'sales': int(d * 0.006), 'pct_of_linked': 0.6,
                               'items': [{'车系': '翻转示例', '销量': int(d * 0.006), '占比': 0.6}]},
            'mixed_with_erev': {'series': 3, 'sales': int(d * 0.008), 'pct_of_linked': 0.8,
                                'items': []},
            # 键名与真实 A-JSON 一致（series_count/sales/sales_pct/cause_bucket）
            'unlinked_board': {
                'series_count': int(n * 0.03),
                'sales': board_total - d,
                'sales_pct': round((board_total - d) / board_total * 100, 4),
                'board_series': int(n * 0.75), 'board_sales': board_total,
                'cause_bucket': {
                    '命名未锚定':   {'series': int(n * 0.02),
                                     'sales': int((board_total - d) * 0.75)},
                    '公告库无痕迹': {'series': int(n * 0.01),
                                     'sales': int((board_total - d) * 0.25)},
                    'unlinked_series': int(n * 0.03),
                    'unlinked_sales': board_total - d,
                },
                'top': [['未关联示例', 100, '10.00-15.00万']],
            },
        },
        'concentration': {
            'series': {'CR3': 21.4, 'CR5': 32.1, 'CR10': 51.2, 'HHI': 184.0, 'n': n},
            'brand': {'CR3': 27.3, 'CR5': 40.5, 'CR10': 62.8, 'HHI': 232.0, 'n': 210},
            'oem': {'CR3': 30.1, 'CR5': 44.7, 'CR10': 66.9, 'HHI': 268.0, 'n': 96},
            'series_board': {'CR3': 11.2, 'CR5': 15.3, 'CR10': 22.8, 'HHI': 112.7,
                             'n': int(n * 0.75)},
            'top_series': [['车系示例', int(d * 0.05), 5.0, '企业甲']] * 20,
            'top_series_board': [['车系示例', int(d * 0.05), 4.0, '企业甲']] * 20,
            'top_brand': [['品牌示例', int(d * 0.06), 6.0]] * 20,
            'top_oem': [['企业示例', int(d * 0.08), 8.0]] * 20,
        },
        'trend': {
            'monthly': monthly, 'energy_monthly': energy_monthly,
            'monthly_board': [[m, int(tot * 1.37)] for m, tot in
                              [(r[0], r[1]) for r in monthly]],
            'price_monthly': price_monthly,
            'price_monthly_board': [[r[0]] + [int(v * 1.37) for v in r[1:]]
                                    for r in price_monthly],
            'segment_monthly': segment_monthly,
            'top20_first': [['车系示例', int(d * 0.05)]] * 20,
            'top20_last': [['车系示例', int(d * 0.05)]] * 20,
            'rankchange_hist': [['上升1-3位', 61], ['持平', 120], ['下降1-3位', 48]],
            'new_entries': [['新晋示例', int(d * 0.01)]], 'drops': [['掉榜示例', int(d * 0.01)]],
        },
        'tech_sales': {
            'weighted': weighted, 'unweighted': unweighted,
            'sample_n': {'weighted': int(n * 0.13), 'unweighted': int(n * 0.12)},
            'range_bins': [[rbands[i], 16.6, int(n * 0.2)] for i in range(len(rbands))],
            'range_bins_def': rbands,
            'price_tech': [[p, 601.5, 76.4, 208.3, 1901.0, int(d * 0.2), 42] for p in bands],
            'price_range_matrix': [[p] + [round(601.5 + 3 * i) for i in range(len(rbands))]
                                   for p in bands],
            'by_energy_tech': [['BEV', weighted], ['PHEV', unweighted], ['EREV', weighted]],
        },
        'conversion': {
            'by_batch': [[b[i], 2021 + i % 6, 40 + i % 30, int(d * 0.004)] for i in range(n_batches)],
            'by_year': [[2021 + i, 900 + 20 * i, int(d * 0.04)] for i in range(6)],
            'brand_scatter': [['品牌甲', 120, int(d * 0.06)], ['品牌乙', 80, int(d * 0.05)]],
            'rate_table': [['品牌甲', 120, 21, int(d * 0.06), 4100], ['品牌乙', 80, 11, int(d * 0.05), 3600]],
        },
    }


def sales_selftest():
    """阶段 1 门禁自测：合成夹具 → G0/G1/G8/stub 检测的通过 + 负向用例。"""
    print(f'=== build_report {GENERATOR_VERSION} · 契约门禁自测（合成夹具，不入库） ===')
    fail = []

    # 1) 正向：合成夹具必须通过 G0
    aj = make_selftest_analytics()
    try:
        assert_schema(aj)
        print('[selftest] G0 契约门禁（合成夹具，正向）: PASS')
    except SystemExit:
        fail.append('G0 正向用例未通过（合成夹具不满足自身门禁——夹具或门禁有 bug）')

    # 2) 负向：status_counts 合计 != base_rows 必须被抓
    bad = make_selftest_analytics()
    bad['coverage']['status_counts']['在售-已匹配'] += 1
    import io
    import contextlib
    buf = io.StringIO()
    caught = False
    with contextlib.redirect_stdout(buf):
        try:
            assert_schema(bad)
        except SystemExit:
            caught = True
    print(f'[selftest] G0 负向（status_counts 合计错误）: {"PASS" if caught else "FAIL"}')
    if not caught:
        fail.append('G0 负向用例未拦截 status_counts 合计错误')
    elif 'SUM' not in buf.getvalue():
        fail.append('G0 负向用例未给出 SUM 诊断行')

    # 3) 负向：删一个契约路径必须被抓
    bad2 = make_selftest_analytics()
    del bad2['concentration']['oem']['HHI']
    caught2 = False
    with contextlib.redirect_stdout(io.StringIO()):
        try:
            assert_schema(bad2)
        except SystemExit:
            caught2 = True
    print(f'[selftest] G0 负向（缺 concentration.oem.HHI）: {"PASS" if caught2 else "FAIL"}')
    if not caught2:
        fail.append('G0 负向用例未拦截缺失子键')

    # 4) 渲染 + G1：全部章节叙事必须可渲染且零裸数字
    try:
        ov = selftest_overrides(aj)
        unregistered = sorted(set(re.findall(r'\$\{([A-Za-z0-9_.]+)\}',
                                             ''.join(t for _k, _ti, t in NARRATIVE))) - set(NUM_SRC))
        assert not unregistered, f'模板引用了未注册的 NUM_SRC 键: {unregistered}'
        for key, _title, _tmpl in NARRATIVE:
            txt = narrate(key, aj, override=ov)
            assert txt and len(txt) > 40, f'{key} 渲染结果过短'
        print(f'[selftest] G1 渲染（{len(NARRATIVE)} 章，零裸数字）: PASS')
        print(f'[selftest] NUM_SRC 注册 {len(NUM_SRC)} 键，模板引用全部已注册: PASS')
    except Exception as e:                                     # noqa: BLE001
        fail.append(f'G1/G8 渲染失败: {e}')
        print(f'[selftest] G1 渲染: FAIL — {e}')

    # 5) G8：禁用词必须被拦
    hits = gate_forbidden_text('该车系销量高说明技术领先，遥遥领先，据悉明年更强。', where='selftest')
    print(f'[selftest] G8 写作禁区拦截: {"PASS" if len(hits) >= 3 else "FAIL"}（命中 {len(hits)} 类）')
    if len(hits) < 3:
        fail.append('G8 未拦住全部禁用词类别')

    # 6) 表标题口径门禁
    ok1 = gate_table_caption('销量结构：动力类型（2026-09 车系级月销 · 新能源榜）')
    ok2 = gate_table_caption('销量结构：动力类型')
    print(f'[selftest] G8 表标题口径: {"PASS" if ok1 and not ok2 else "FAIL"}')
    if not (ok1 and not ok2):
        fail.append('G8 表标题口径门禁判断错误')

    # 7) stub 检测
    todo = count_pending_tables()
    print(f'[selftest] 表注册表 {len(TABLE_REGISTRY)} 张；待实现 {len(todo)} 张'
          f'（阶段 2 进度项，非失败）')

    if fail:
        print('[selftest] FAIL:')
        for f_ in fail:
            print(f'   - {f_}')
        return 2
    print('[selftest] ALL PASS —— 骨架可在 task-2 数据到位前完成门禁验证')
    return 0


# ============================================================================
# 7. 历史角色：底表重算（build_r1..r5，保持原样，勿动）
# ============================================================================
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


def build_r1(doc, rows):
    """执行摘要 + 第一章（T00 概览 / T03 年度批次 / T05 BEV占比 / T06-T07 细分市场）"""
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
    T03, T05 = [], []
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
    for ptname in ('BEV', 'PHEV/EREV'):
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
    for pt, lbl in ptmap:
        st = per_year_full(rows, pt, '纯电续航里程(km)')
        doc.add_heading(f'3.{1 if pt=="BEV" else 2} {lbl}续航里程趋势', level=2)
        add_table(doc, ['年份', '均值', '中位数', 'P25', '样本'],
                  [[y, fmt(s['mean']), fmt(s['median'], 0), fmt(s['p25'], 0), s['n']] for y, s in st.items()])
    stb, stp = per_year_full(rows, 'BEV', '纯电续航里程(km)'), per_year_full(rows, 'PHEV/EREV', '纯电续航里程(km)')
    doc.add_heading('3.3 BEV vs PHEV续航对比', level=2)
    add_table(doc, ['年份', 'BEV均值', 'PHEV均值', '差值'],
              [[y, fmt(stb[y]['mean']), fmt(stp[y]['mean']), fmt(stb[y]['mean'] - stp[y]['mean'])] for y in sorted(stb) if y in stp])
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

    def bt_by_year(pt):
        by = defaultdict(Counter)
        tot = Counter()
        for r in rows:
            if r['PT'] == pt and r['year']:
                bt = str(r.get('电池类型') or '').strip().replace('电池', '')
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
    st_ed = per_year_full(rows, 'BEV', '电池能量密度(Wh/kg)')
    doc.add_heading('3.10 BEV电池能量密度趋势', level=2)
    add_table(doc, ['年份', '均值', '中位数', '样本数'],
              [[y, fmt(s['mean'], 2), fmt(s['median'], 1), s['n']] for y, s in st_ed.items()])

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
    for pt, lbl in (('BEV', 'BEV'), ('PHEV/EREV', 'PHEV')):
        rng = [fnum(r.get('纯电续航里程(km)')) for r in rows if r['PT'] == pt]
        cap = [fnum(r.get('电池容量(kWh)')) for r in rows if r['PT'] == pt]
        rng = [v for v in rng if v]
        cap = [v for v in cap if v]
        if not rng or not cap:
            continue
        import statistics
        med_r = statistics.median(rng)
        med_c = statistics.median(cap)
        q = defaultdict(list)
        for r in rows:
            if r['PT'] != pt:
                continue
            rv, cv = fnum(r.get('纯电续航里程(km)')), fnum(r.get('电池容量(kWh)'))
            if rv is None or cv is None:
                continue
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
    doc.add_heading('第五章 细分市场深度分析', level=1)
    for pt, lbl in (('BEV', 'BEV'), ('PHEV/EREV', 'PHEV')):
        seg = defaultdict(list)
        for r in rows:
            if r['PT'] == pt:
                seg[str(r.get('细分市场') or '未知')].append(r)
        doc.add_heading(f'6.{1 if pt=="BEV" else 2} {lbl}细分市场统计', level=2)
        add_table(doc, ['细分市场', '样本数', '续航均值(km)', '容量均值(kWh)', '整备均值(kg)'],
                  [[k, len(v),
                    fmt(sum(x for x in (fnum(r.get('纯电续航里程(km)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('纯电续航里程(km)'))))),
                    fmt(sum(x for x in (fnum(r.get('电池容量(kWh)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('电池容量(kWh)'))))),
                    fmt(sum(x for x in (fnum(r.get('整备质量(kg)')) for r in v) if x) / max(1, sum(1 for r in v if fnum(r.get('整备质量(kg)')))))]
                   for k, v in sorted(seg.items(), key=lambda kv: -len(kv[1]))])
    seg_b = defaultdict(list)
    for r in rows:
        if r['PT'] == 'BEV' and r['year'] in (2023, 2025):
            seg_b[(str(r.get('细分市场') or '未知'), r['year'])].append(r)
    segs = sorted({k[0] for k in seg_b})
    doc.add_heading('6.3 BEV细分市场续航对比 - 2023年 vs 2025年', level=2)
    rows63 = []
    for sgm in segs:
        v23 = [x for x in (fnum(r.get('纯电续航里程(km)')) for r in seg_b.get((sgm, 2023), [])) if x]
        v25 = [x for x in (fnum(r.get('纯电续航里程(km)')) for r in seg_b.get((sgm, 2025), [])) if x]
        if v23 or v25:
            rows63.append([sgm, fmt(sum(v23)/len(v23)) if v23 else '-', fmt(sum(v25)/len(v25)) if v25 else '-', len(v23), len(v25)])
    add_table(doc, ['细分市场', '2023续航均值', '2025续航均值', '2023样本', '2025样本'], rows63)
    doc.add_heading('6.4 BEV细分市场容量对比 - 2023年 vs 2025年', level=2)
    rows64 = []
    for sgm in segs:
        v23 = [x for x in (fnum(r.get('电池容量(kWh)')) for r in seg_b.get((sgm, 2023), [])) if x]
        v25 = [x for x in (fnum(r.get('电池容量(kWh)')) for r in seg_b.get((sgm, 2025), [])) if x]
        if v23 or v25:
            rows64.append([sgm, fmt(sum(v23)/len(v23)) if v23 else '-', fmt(sum(v25)/len(v25)) if v25 else '-', len(v23), len(v25)])
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
        doc.add_heading(f'7.{num+4} {lbl}年度综合指标', level=2)
        add_table(doc, ['年份', '续航均值', '容量均值', '总功率均值', '样本数'],
                  [[y, fmt(st_all[y]['mean']), fmt(sc[y]['mean']), fmt(sp[y]['mean']), st_all[y]['n']] for y in years])
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


# ============================================================================
# 7b. 增量注入 —— 销量分册（摘要 + 八章 + 附录 A）
#
# 为什么是"追加分册"而不是插入原七章之间：
#   Phase-1 审计已确证（大纲 §0.2）本生成器无法从零重放交付物现状——
#   叙事层②/样式层/封面+TOC 层由三个未入库的 beautify 提交完成。
#   因此新内容一律**追加**，绝不打散既有 59 表与 121 段的位置（避免表号漂移）。
#
# D2（Lead 裁决：新章编号我定，但目录与正文必须同源生成、不得手写）：
#   下面 SALES_SECTIONS 是**唯一编号来源**；标题文案由它生成，
#   目录由 Word 的 TOC 域按 Heading 样式自动生成 → 二者同源，不存在脱节。
# ============================================================================
SALES_PART_TITLE = '销量与市场表现联动分析（2026-09 快照 · 车系级去重口径）'

SALES_SECTIONS = [
    # (叙事键, 章标题, 归属表号段或 None)
    ('abstract', '摘要 · 销量与市场表现',        '摘要'),
    ('ch1',      '第一章 数据与方法',            '8.1'),
    ('ch2',      '第二章 市场总量与结构',        '8.2'),
    ('ch3',      '第三章 技术-市场关联',         '8.3'),
    ('ch4',      '第四章 集中度与竞争格局',      '8.4'),
    ('ch5',      '第五章 趋势与迁移',            '8.5'),
    ('ch6',      '第六章 公告-市场转化效率',     '8.6'),
    ('ch7',      '第七章 面向领导层的判断与建议', '8.7'),
    ('ch8',      '第八章 结论',                  None),
    ('appendix', '附录 A 口径与覆盖率披露',      '附录A'),
]

# 封面第三行：规模表述按 Lead 裁决 D 条件④ —— "条"而非"款"，并含 5,417+58 拆分。
COVER_LINE3 = ('权威底表 ${cfg.ver} 口径 · ${m1.rows} 条公告/公示记录'
               '（正式公告 ${m1.official_rows} 条 + 第 ${m1.batch_hi} 批公示 '
               '${m1.last_batch_rows} 条） × ${m1.cols} 列 · '
               '工信部公告第 ${m1.batch_lo}~${m1.batch_hi} 批 · 生成 ${cfg.gen_date}')


def set_para_text(p, text, bold=None):
    """覆盖段落文本：保留段落样式与首个 run 的字体，清空其余 run。

    不新建段落（避免破坏封面排版与 TOC 域）。
    """
    runs = p.runs
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    if bold is not None:
        runs[0].font.bold = bold
    for r in runs[1:]:
        r.text = ''


def fix_cover(doc, aj):
    """封面规模/版本/批次修正（P1 列数、P2 批次范围、版本滞后）。

    匹配必须**只命中封面那一行**：早期实现用 '权威底表' 做关键词，
    会连带命中执行摘要里的"以下概览数据由权威底表全量重算（v4.9.25 口径）。"
    —— 那行会被整行替换成封面文案，属于静默损坏。故改为锚定封面独有串
    '工信部公告第'（正文不会出现该串），并限定在前 5 段内。
    """
    done = []
    for p in doc.paragraphs[:5]:
        t = p.text or ''
        if '工信部公告第' not in t:
            continue
        new = render(COVER_LINE3, aj)
        set_para_text(p, new)
        done.append(new)
    if len(done) != 1:
        print(f'[cover] ⚠ 封面规模行命中 {len(done)} 处（应为 1 处）——请人工核对')
    for d in done:
        print(f'[cover] 已更新: {d}')
    return done


def _write_narrative(doc, text, style=STYLE_BODY):
    """插入一段叙事正文，`**重点**` 渲成真加粗 run（而非字面星号）。"""
    p = doc.add_paragraph(style=style)
    set_rich_text(p, text)
    return p


def append_sales_part(doc, aj):
    """把销量分册追加到文末（Heading 样式 → 自动进入目录）。"""
    legacy_texts = []          # 供 G10 残留旧数字扫描
    n_tables = 0
    doc.add_paragraph(SALES_PART_TITLE, style=STYLE_H1)

    for key, title, sec in SALES_SECTIONS:
        doc.add_paragraph(title, style=STYLE_H2)
        txt = narrate(key, aj)          # render() 内已跑 G1 并豁免脚本注入值
        hits = gate_forbidden_text(txt, where=f'分册/{key}')
        if hits:
            raise AssertionError(f'G8 写作禁区命中 @ {key}: {hits}')
        legacy_texts.append(txt)
        # 叙事正文：`**` 须渲染为真加粗 run（G11 在 save 前兜底断言）
        _write_narrative(doc, txt)

        if not sec:
            continue
        for tid, label, tsec, ttitle, fn in TABLE_REGISTRY:
            if tsec != sec:
                continue
            f = globals().get(fn)
            if f is None or is_stub(f):
                raise AssertionError(f'{tid} 未实现（{fn}）——拒绝产出半成品报告')
            cap = tbl_caption(label, ttitle)
            if not gate_table_caption(cap):
                raise AssertionError(f'G8 表标题缺口径/月份: {cap}')
            p = doc.add_paragraph(style=STYLE_BODY)
            r = p.add_run(cap)
            r.font.bold = True
            legacy_texts.append(cap)
            header, rows = f(aj)
            add_table(doc, header, rows)
            foot = doc.add_paragraph(FOOT_SERIES, style=STYLE_BODY)
            for run in foot.runs:
                run.font.size = Pt(8)
            n_tables += 1

    # 全分册统一脚注段
    doc.add_paragraph(FOOT_PRICE, style=STYLE_BODY)

    # G10 残留旧数字扫描（Lead 裁决 D1 条件③）
    hits, _allowed = gate_legacy_numbers(legacy_texts, where='销量分册')
    if hits:
        raise AssertionError(f'G10 残留旧数字未清：{hits}'
                             f'（必须清掉，或列入 LEGACY_WHITELIST 并写明理由）')
    print(f'[G10] 残留旧数字扫描 PASS —— 0 命中（已扫 {len(legacy_texts)} 段）')
    return n_tables


def _norm_pt(v):
    """动力类型归一：底表该列是双形态（BEV/PHEV 代码 + '纯电动轿车' 等中文全称）。"""
    s = str(v or '').strip()
    if s == 'BEV' or '纯电' in s:
        return 'BEV'
    if not s:
        return ''
    return 'PHEV/EREV'


def fix_legacy_summary(doc, aj):
    """旧章（销量分册之前）的存量数字同步。

    背景：底表由 5,392 行变 5,475 行（并发会话入库 411/412 批），
    旧章的"总记录数"及其**同表构成项**同时失效。只改总数不改构成项
    会造成表格自相矛盾，故此处从当前底表现读重算。

    纪律（AGENTS 2026-09-29 教训）：**不做文档级全局数值映射**——
    替换范围严格限定在"销量分册之前"的正文，且只针对已确证失效的
    少数 token；每一处改动都打印留痕，不静默放过。
    """
    base = _load_base_rows(aj)
    if base is None:
        print('[legacy] ⚠ 底表读取失败 —— 旧章存量数字未同步（将留旧值，需人工核对）')
        return []
    br = dig(aj, 'meta.batch_range') or [None, None]
    total = len(base)
    n_bev = sum(1 for r in base if len(r) > 9 and _norm_pt(r[9]) == 'BEV')
    n_phev = sum(1 for r in base if len(r) > 9 and _norm_pt(r[9]) == 'PHEV/EREV')
    # 品牌/企业去重口径：按「产品商标」/「企业名称」原值 strip 后去重，排除空与 None。
    # ⚠️ 必须显式判 `is not None`：`str(None)` == 'None' 且为真值，会把一个空商标
    #    当成名为 "None" 的品牌计入（实测使品牌数由 201 虚增到 202）。
    brands = len({str(r[2]).strip() for r in base
                  if len(r) > 2 and r[2] is not None and str(r[2]).strip()})
    oems = len({str(r[3]).strip() for r in base
                if len(r) > 3 and r[3] is not None and str(r[3]).strip()})
    # 空值行数 + 「牌」字归一后的品牌数：均为"口径说明"栏的必需内容，
    # 必须现算而非手写（摘要头条 KPI 的口径要能被追问时立刻答上来）。
    # ⚠️ 归一确实会合并若干商标（实测 12 组同名，如「大众」/「大众牌」），
    #    故 201 与 190 是两个都正确的数——差别在口径，须显式并列披露。
    n_no_brand = sum(1 for r in base
                     if len(r) <= 2 or r[2] is None or not str(r[2]).strip())
    n_no_oem = sum(1 for r in base
                   if len(r) <= 3 or r[3] is None or not str(r[3]).strip())
    brand_norm = len({re.sub(r'牌$', '', str(r[2]).strip()) for r in base
                      if len(r) > 2 and r[2] is not None and str(r[2]).strip()})

    # 分册起始位置：其后属于本轮新增内容，不参与替换
    part_at = None
    for i, p in enumerate(doc.paragraphs):
        if (p.text or '').strip() == SALES_PART_TITLE:
            part_at = i
            break
    if part_at is None:
        part_at = len(doc.paragraphs)

    # 1) 正文散文里的失效数字（只动分册之前）
    prose_rules = [
        ('5,392 款', f'{total:,} 条'),
        ('5,392款', f'{total:,}条'),
        ('5,392 条', f'{total:,} 条'),
        ('5,392条', f'{total:,}条'),
        ('5,392', f'{total:,}'),
        ('341~411', f'{br[0]}~{br[1]}'),
        ('341-411', f'{br[0]}-{br[1]}'),
    ]
    changes = []
    for i, p in enumerate(doc.paragraphs[:part_at]):
        t = p.text or ''
        if not t:
            continue
        new = t
        for old, rep in prose_rules:
            if old in new:
                new = new.replace(old, rep)
        if new != t:
            set_para_text(p, new)
            changes.append(('P', i, t[:60], new[:60]))

    # 2) 概览表的构成项（首列为"总记录数"的 3 列表）
    cell_map = {
        '总记录数': (f'{total:,}', f'第{br[0]}-{br[1]}批公告/公示记录（乘用车纯净口径）'),
        'BEV车型数': (f'{n_bev:,}', f'纯电动车型，占比{n_bev / total * 100:.1f}%'),
        'PHEV/EREV车型数': (f'{n_phev:,}', f'插电混动/增程车型，占比{n_phev / total * 100:.1f}%'),
        '品牌数': (f'{brands:,}',
                   f'产品商标原值去重，排除空值（{n_no_brand} 行无商标）；'
                   f'未做「牌」字归一，归一后为 {brand_norm}'),
        '企业数': (f'{oems:,}',
                   f'企业名称原值去重，排除空值（{n_no_oem} 行无企业名称）'),
    }
    for ti, t in enumerate(doc.tables):
        try:
            first = t.rows[0].cells[0].text
        except (IndexError, AttributeError):
            continue
        if first != '维度' or len(t.columns) < 3:
            continue
        for ri, row in enumerate(t.rows):
            key = row.cells[0].text.strip()
            if key in cell_map:
                val, note = cell_map[key]
                if row.cells[1].text.strip() != val or row.cells[2].text.strip() != note:
                    changes.append(('T', ti, ri, key,
                                    row.cells[1].text.strip(), val))
                    set_tc(t, ri, 1, val)
                    set_tc(t, ri, 2, note)

    for c in changes:
        print(f'[legacy] {c[0]}{c[1:] if c[0] == "T" else ""} → {c[-2:]}')
    print(f'[legacy] 旧章存量数字同步 {len(changes)} 处'
          f'（底表 {total:,} 行 / BEV {n_bev:,} / PHEV+EREV {n_phev:,} / '
          f'品牌 {brands:,} / 企业 {oems:,}）')
    return changes


def remove_sales_part(doc):
    """移除已存在的销量分册（使生成可**重复运行**而不叠加）。

    为什么必须做：`generate_report` 是"在现值 docx 上追加"的增量注入。
    若第二次运行不清除，会把整册再追加一遍（实测：段落 187→253、
    表格 81→103，即 22 张表被注入两次）——生成器必须幂等。

    实现：按 body 子元素顺序定位分册标题段落，删除其后的所有元素，
    但保留结尾的 `w:sectPr`（节属性，删除会导致文档损坏）。
    """
    body = doc.element.body
    kids = list(body)
    start = None
    for i, el in enumerate(kids):
        if el.tag == qn('w:p'):
            txt = ''.join(t.text or '' for t in el.iter(qn('w:t')))
            if txt.strip() == SALES_PART_TITLE:
                start = i
                break
    if start is None:
        return 0
    removed = 0
    for el in kids[start:]:
        if el.tag == qn('w:sectPr'):
            continue
        body.remove(el)
        removed += 1
    print(f'[clean] 已移除上一次注入的销量分册：{removed} 个元素（保证重复运行不叠加）')
    return removed


def generate_report(out_path=None, do_backup=True):
    """阶段 2 主入口：在现值 docx 上增量注入销量分册，产出终稿。"""
    aj, pending = load_analytics()
    assert_schema(aj)
    if pending:
        print(f'[G9] ⚠ 仍在使用过渡基准值 {pending} —— 终稿要求 pending_used == 0，'
              f'请等 WS2 落地后重跑。')

    out = out_path or REPORT_DOCX
    if do_backup and os.path.exists(out):
        backup_report('pre-sales')

    doc = Document(out)
    before_p = len(doc.paragraphs)
    before_t = len(doc.tables)
    fix_cover(doc, aj)
    fix_legacy_summary(doc, aj)
    remove_sales_part(doc)          # 幂等：重复运行不叠加分册
    n_tables = append_sales_part(doc, aj)

    # G10 全文残留扫描（Lead 裁决 D1 条件③）—— 覆盖正文 + 表格单元格，
    # 不只看新增分册：旧章里的失效数字同样会被用户读到。
    blob = [p.text or '' for p in doc.paragraphs]
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                blob.append(c.text or '')
    hits, _allowed = gate_legacy_numbers(blob, where='全文')
    if hits:
        raise AssertionError(f'G10 全文残留旧数字：{hits}（已中止，未写出文件）')

    # G11：markdown 加粗标记必须已渲染为真加粗 run，不得残留字面 `**`
    #（领导层文档出现星号属硬缺陷；同样在写出文件前拦截）
    assert_no_markdown(doc, where=out)

    doc.save(out)

    size = os.path.getsize(out)
    print(f'[OK] 已生成: {out}')
    print(f'     大小 {size:,} B；段落 {before_p} → {len(doc.paragraphs)}；'
          f'表格 {before_t} → {len(doc.tables)}（新增 {n_tables} 张）')
    return out


# ============================================================================
# 8. CLI
# ============================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out')
    ap.add_argument('--stages', default='r1')
    ap.add_argument('--sales-probe', action='store_true',
                    help='仅校验 A-JSON 契约（阶段 1 可运行）')
    ap.add_argument('--sales-draft', action='store_true',
                    help='生成销量新章独立草稿（阶段 2）')
    ap.add_argument('--sales-selftest', action='store_true',
                    help='用合成夹具自测 G0/G1/G8 门禁与渲染管线（不读真实数据、不写交付物）')
    ap.add_argument('--backup', action='store_true',
                    help='仅对现报告做含版本号的自动备份')
    ap.add_argument('--sales-generate', action='store_true',
                    help='【阶段2 主入口】在现值 docx 上增量注入销量分册，生成终稿')
    ap.add_argument('--no-backup', action='store_true',
                    help='配合 --sales-generate：跳过自动备份（不建议）')
    args = ap.parse_args()

    if args.sales_probe:
        return sales_probe()
    if args.sales_selftest:
        return sales_selftest()
    if args.backup:
        backup_report()
        return 0
    if args.sales_generate:
        generate_report(args.out, do_backup=not args.no_backup)
        return 0
    if args.sales_draft:
        if not args.out:
            ap.error('--sales-draft 需要 --out')
        return sales_draft(args.out)

    # 历史角色：底表重算（会丢失叙事/样式/封面层，见模块 docstring ①）
    if not args.out:
        ap.error('--out is required')
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
    print(f'saved: {args.out} (stages={args.stages}, rows={len(rows)}, gen={GENERATOR_VERSION})')
    return 0


if __name__ == '__main__':
    sys.exit(main())
