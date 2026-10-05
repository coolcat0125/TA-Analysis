#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NEV公告数据看板生成器 v3（开箱即用 · 本地离线运行）
====================================================
读取公告参数汇总Excel → 生成自包含HTML交互式看板（内嵌ECharts+SheetJS，双击即可打开）

v3.7 更新：
  1. 新增【分布格局】分析页（配置驱动，后续增减维度只改 DIST_DEFS 数组）：
     EV Range vs Length / Energy vs EV Range / Energy Consumption vs EV Range /
     Energy Density vs EV Range / Curb Mass vs Length（散点 + BEV/PHEV 独立趋势线）
  2. 全部图表纵轴/横轴标注名称与单位
  3. 看板展示层移除免征购置税相关信息（底层数据表保留该维度）
  4. 底层表新增「车长(mm)」占位列；Length 视角暂以公告轴距(mm)呈现，
     车长数据补齐后仅需将 DIST_DEFS 中 xKey 由 ab 换为 lg

v3.5 更新：
  1. 新增【企业与品牌】分析页：头部企业批次演进 / 企业×细分市场矩阵 /
     各批次新进入企业跟踪 / 品牌 Top 15
  2. 车型明细查询表：当前筛选内搜索（企业/商标/型号）、列排序、翻页浏览
  3. 记录字段扩展：产品型号 / 产品商标 / 是否减免购置税（导出同步）

v3 更新：
  1. 全局 PHEV 展示名统一为 PHEV/EREV（含图表、KPI、导出）
  2. 全局"拟合线"更名为"趋势线"；多参数图每参数独立趋势线
  3. 细分市场"/"占位值视为未填报，不显示、不计入统计
  4. 电池类型归一化（磷酸铁锂电池/蓄电池/LFP等合并口径），
     构成趋势按已有数据百分比占比核算
  5. 散点图 BEV / PHEV/EREV 分列趋势线

v2 功能：
  1. 细分市场筛选改为多选下拉菜单
  2. 统计口径：缺失/未填报不计入统计与百分比分母
  3. 散点图趋势线（点击图例开/关）
  4. 折线图缺失点以趋势线衔接（同色淡色虚线，不抢焦点）
  5. 【电池与续航】→【电池系统】，新增百公里电耗趋势
  6. 【数据质量】移至页面底部
  7. 数据管理页：导出当前筛选Excel/CSV + 导入最新数据表页面内刷新
  8. 白天/黑夜模式切换（CSS变量+ECharts联动）
  9. 导入新Excel后全局数据实时刷新

用法：
  python3 generate_dashboard.py                          # 使用默认文件名（同目录）
  python3 generate_dashboard.py <输入.xlsx> <输出.html>  # 指定输入输出
  python3 generate_dashboard.py --release                # 发布版（隐藏数据质量区块，仅保留版权页脚）

v3.7.1 更新：
  1. 分布格局五图标题中文化（英文名保留为副标题注记）
  2. 细分市场下拉按 Car/SUV/MPV 分组、字母级别→数字级别→未分级排序（Python/JS 双侧同键）
  3. 全部图表卡片新增「一句话核心要点」行（.ci），随全局筛选实时联动

v3.7.2 更新：
  1. 所有散点图悬浮窗补上车型【通用名称】（记录新增 gn 字段，页内导入同步）
  2. 车型明细查询表新增【通用名称】列；导出字段同步
  3. 坐标轴标签完整显示：网格自适应留白（containLabel），取消类目标签截断

v3.7.3 更新：
  1. 分布格局要点改为分燃料类型（BEV / PHEV/EREV）的均值与极值，随全局筛选联动；移除相关系数描述
  2. 移除分布格局卡片的两条口径注释说明
  3. 底表补充车长(mm)：以轴距为基准按车身形式系数估算（fill_length_v373.py，浅紫底纹+台账标记，
     填报率约65%）；完成率不高，看板 Length 分析仍以轴距为基准

依赖：openpyxl（pip install openpyxl）、同目录下 echarts.min.js 与 xlsx.full.min.js
"""
import openpyxl
import json
import os
import re
import sys
import datetime

# ============================================================
# 配置
# ============================================================
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_IN = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
DEFAULT_OUT = os.path.join(HERE, 'NEV公告数据看板.html')
DEFAULT_RELEASE_OUT = os.path.join(HERE, 'NEV公告数据看板_发布版.html')
COPYRIGHT = 'Copyright (c) 2026 David YEAH'
ECHARTS_PATH = os.path.join(HERE, 'echarts.min.js')
ECHARTSGL_PATH = os.path.join(HERE, 'echarts-gl.min.js')   # 三维展示（scatter3D/grid3D）
XLSXLIB_PATH = os.path.join(HERE, 'xlsx.full.min.js')
SHEET_NAME = 'NEV公告参数汇总'

FIELDS = ['b', 't', 's', 'e', 'w', 'r', 'c', 'bt', 'ed', 'ec',
          'fp', 'tp', 'ms', 'tq', 'fo', 'dv', 'ep', 'es', 'src',
          'm', 'bd', 'tx', 'ab', 'lg', 'gn', 'rt', 'rp', 'drive']


def clean_num(v):
    """数值清洗：'不适用(BEV)'、空值、异常文本 → None；多配置'285/302'取首值"""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if '不适用' in s or s in ('-', '—', '/', 'None', 'nan'):
        return None
    if '/' in s:
        s = s.split('/')[0].strip()  # 多配置/区间值取首值（如 285/302 → 285）
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def clean_str(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s or '不适用' in s:
        return None
    return s


def parse_torque(v):
    """解析'功率/扭矩'字段中的扭矩值"""
    if not v:
        return None
    s = str(v).strip()
    if '/' not in s:
        return None
    part = s.split('/')[-1].strip()
    try:
        return float(part)
    except (ValueError, TypeError):
        return None


def parse_pt(v):
    """v4.4.0：解析 'P/T' 文本（如 '230/472'、'130'）→ (功率, 扭矩)，缺失段为 None"""
    if v is None:
        return None, None
    s = str(v).strip()
    if not s or '不适用' in s:
        return None, None
    parts = s.split('/')
    try:
        p = float(parts[0]) if parts[0].strip() else None
    except (ValueError, TypeError):
        p = None
    try:
        t = float(parts[1]) if len(parts) > 1 and parts[1].strip() else None
    except (ValueError, TypeError):
        t = None
    return p, t


def clean_seg(v):
    """细分市场清洗：'/' 等占位符视为未填报，不计入统计与分类展示"""
    s = clean_str(v)
    if s is None or s.strip('/') == '' or s in ('-', '—', 'N/A'):
        return None
    return s


# ---------------------------------------------------------------------------
# 细分市场规范化（v4.7.1）
#   【用户口径】车身形式以**公告分类**为准（产品名称：轿车/运动型乘用车/
#   多用途乘用车/越野乘用车），具体级别以**尺寸**为准（《车型级别定义》轴距阈值），
#   最终整合为「<公告车身>-<尺寸级别>」（如 SUV-C），并确保**每一行都落到具体级别**。
#
#   落位优先级：
#     1. 已是具体级别且车身与公告一致 → 保持
#     2. 车身 = 公告分类（产品类型/产品名称；「多用途乘用车」借同车型既有标签或型号编码区分 SUV/MPV）
#     3. 级别 = 轴距按阈值分级；轴距缺失时依次用：同车型众数轴距 → 型号编码推定车身＋车长
#     4. Lux / Sports 按《车型级别定义》保留各自序列
# ---------------------------------------------------------------------------
SEG_BODIES = ('Car', 'SUV', 'MPV')
SEG_LUX_PREFIX = 'Lux '

# 轴距阈值（与《车型级别定义》完全一致）
SEG_WB_CAR = (2650, 2740, 2850, 3000)      # A <2650 ≤B <2740 ≤C <2850 ≤D <3000 ≤E
SEG_WB_SUV = (2680, 2850, 3000)            # B <2680 ≤C <2850 ≤D <3000 ≤E
SEG_WB_MPV = (3000,)                       # C <3000 ≤D

# 产品型号编码首段数字 → 车身形式（由已规范行实测统计：2→SUV 95%、7→Car 92%、6→SUV 86%）
# 注：6 为「客车」类编码，既含 SUV 也含 MPV，故仅作辅助，需再借长度/同车型信息
SEG_CODE_BODY = {'2': 'SUV', '7': 'Car'}

# 各车身形式车长合理区间（兜底校验用）
SEG_LEN_RANGE = {'Car': (2810, 5850), 'SUV': (2830, 5860), 'MPV': (3640, 5970)}


def seg_wheelbase_level(body, mm):
    """轴距阈值分级（与《车型级别定义》一致）。"""
    if mm is None or body not in SEG_BODIES:
        return None
    if body == 'Car':
        a, b, c, d = SEG_WB_CAR
        return ('Car-A' if mm < a else 'Car-B' if mm < b else
                'Car-C' if mm < c else 'Car-D' if mm < d else 'Car-E')
    if body == 'SUV':
        b, c, d = SEG_WB_SUV
        return 'SUV-B' if mm < b else 'SUV-C' if mm < c else 'SUV-D' if mm < d else 'SUV-E'
    if body == 'MPV':
        return 'MPV-C' if mm < SEG_WB_MPV[0] else 'MPV-D'
    return None


def seg_body_from_announcement(product_type, product_name):
    """【公告分类】车身形式。『多用途乘用车』按公告口径同时涵盖 SUV 与 MPV，返回 MPV/SUV 待细化。"""
    t = f'{product_type or ""} {product_name or ""}'
    if '越野' in t:
        return 'SUV'
    if '运动型' in t:
        return 'SUV'
    if '轿车' in t:
        return 'Car'
    if '多用途乘用车' in t:
        return 'MPV/SUV'
    if '客车' in t:
        return 'MPV'
    return None


def seg_body_from_code(model_code):
    """产品型号编码首段数字 → 车身形式（辅助判据）。"""
    m = re.search(r'\d', str(model_code or ''))
    return SEG_CODE_BODY.get(m.group(0)) if m else None


def seg_body_from_label(label):
    if not label:
        return None
    b = label[len(SEG_LUX_PREFIX):].strip() if label.startswith(SEG_LUX_PREFIX) else label
    for cand in SEG_BODIES:
        if b == cand or b.startswith(cand + '-'):
            return cand
    return None


def seg_is_specific(label):
    """是否为「具体」细分市场（可枚举、可展示、可统计）。"""
    if not label:
        return False
    if label == 'Sports' or label.startswith(SEG_LUX_PREFIX):
        return True
    if '-' not in label:
        return False
    body, lv = label.split('-', 1)
    if body not in SEG_BODIES:
        return False
    return (len(lv) == 1 and lv.isalpha()) or lv.isdigit()


def build_seg_ref_from_records(records):
    """建立参照表：通用名称 / 产品型号基码 → 各批次轴距·车长·既有级别。

    两级键：通用名称优先（同车型直接可比）；通用名称无尺寸时退化到产品型号基码
    （同基码 = 同一车型同一尺寸，可跨批次补全）。
    """
    ref = {}
    for rec in records:
        g = str(rec[GN_IDX] or '').strip().split(',')[0].strip()
        keys = []
        if g:
            keys.append('n:' + g)
        cb = seg_code_base(rec[MODEL_IDX])
        if cb:
            keys.append('c:' + cb)
        s = rec[SEG_IDX]
        for k in keys:
            e = ref.setdefault(k, {'wb': {}, 'ln': {}, 'seg': set()})
            if s and seg_is_specific(s) and not s.startswith(SEG_LUX_PREFIX):
                e['seg'].add(s)
            for key, ix in (('wb', AB_IDX), ('ln', LG_IDX)):
                v = rec[ix]
                if isinstance(v, (int, float)):
                    e[key][int(v)] = e[key].get(int(v), 0) + 1
    return ref


def seg_code_base(model_code):
    """产品型号基码：去掉末尾变体配置码，得到「制造商码+车型序号+动力码」。"""
    s = str(model_code or '').strip()
    if not s:
        return None
    m = re.match(r'^([A-Z0-9]+?)([A-Z]\d?[A-Z]?)$', s)
    return m.group(1) if m else s


SEG_IDX = 2          # 记录内 细分市场 位
PTYPE_IDX = 1        # 记录内 动力类型（非公告产品类型）
GN_IDX = 24          # 记录内 通用名称
AB_IDX = 22          # 记录内 轴距(mm)
LG_IDX = 23          # 记录内 车长(mm)
ANNO_BODY_IDX = 28   # 记录内 公告车身形式（新增，供核验与悬浮展示）
MODEL_IDX = 19       # 记录内 产品型号


def _modal(counter):
    return max(counter.items(), key=lambda kv: kv[1])[0] if counter else None


def resolve_segment(rec, ptype_raw, pname_raw, ref):
    """把一行规范为具体级别。返回 (细分市场, 依据, 公告车身)。"""
    raw = str(rec[SEG_IDX] or '').strip()
    g = str(rec[GN_IDX] or '').strip().split(',')[0].strip()
    en = ref.get('n:' + g) or {}
    cb = seg_code_base(rec[MODEL_IDX])
    ec = ref.get('c:' + cb) if cb else None
    ann = seg_body_from_announcement(ptype_raw.get(id(rec)), pname_raw.get(id(rec)))

    # 1) Lux / Sports：按《车型级别定义》保留各自序列
    if raw.startswith(SEG_LUX_PREFIX) and seg_body_from_label(raw):
        return raw, 'lux', seg_body_from_label(raw)
    if raw == 'Sports' or ann == 'Sports':
        return 'Sports', 'sports', 'Sports'

    # 2) 车身形式：公告优先；『多用途乘用车』借既有标签 → 型号编码
    body = ann if ann in SEG_BODIES else None
    if body is None:
        segs = set()
        for e in (en, ec or {}):
            segs |= {s.split('-')[0] for s in e.get('seg', set()) if '-' in s and not s.startswith(SEG_LUX_PREFIX)}
        segs = {s for s in segs if s in SEG_BODIES}
        if segs:
            body = max(segs, key=lambda s: sum(1 for e in (en, ec or {}) for x in e.get('seg', set()) if x.startswith(s + '-')))
        else:
            body = seg_body_from_code(rec[MODEL_IDX])
    if body is None and raw and '-' in raw:
        b0 = raw.split('-')[0]
        if b0 in SEG_BODIES:
            body = b0

    # 3) 级别：轴距（本行 → 同车型众数 → 同型号基码众数）
    wb = rec[AB_IDX]
    src = '轴距分级'
    if not isinstance(wb, (int, float)):
        mw = _modal(en.get('wb', {}))
        if mw:
            wb, src = mw, '同车型众数轴距'
        else:
            mw = _modal((ec or {}).get('wb', {}))
            if mw:
                wb, src = mw, '同型号基码轴距'
    if body and isinstance(wb, (int, float)):
        s = seg_wheelbase_level(body, wb)
        if s:
            return s, src, body

    # 4) 车长兜底：同类车身长度换算近似轴距（标注依据，不冒充实测）
    ln = rec[LG_IDX]
    if body and not isinstance(ln, (int, float)):
        ln = _modal(en.get('ln', {})) or _modal((ec or {}).get('ln', {}))
    if body and isinstance(ln, (int, float)):
        approx = {'Car': ln - 2050, 'SUV': ln - 1900, 'MPV': ln - 1900}.get(body)
        s = seg_wheelbase_level(body, approx)
        if s:
            return s, '车长推定', body

    # 5) 无尺寸：显式标注未分级（车身形式可判，级别不可判）——不臆造级别
    if body:
        return f'{body}-未分级', '无尺寸(未分级)', body
    return None, 'unresolved', None


def normalize_segments(records, ptype_raw, pname_raw):
    """就地规范化全部记录的细分市场，确保每行落到具体级别。

    返回 (统计字典, 未落位样例, 公告车身写入回调数据)
    """
    ref = build_seg_ref_from_records(records)
    stat = {}
    unresolved = []
    for rec in records:
        s, how, body = resolve_segment(rec, ptype_raw, pname_raw, ref)
        rec[SEG_IDX] = s
        stat[how] = stat.get(how, 0) + 1
        if s is None:
            unresolved.append((rec[0], rec[GN_IDX]))
    return stat, unresolved


def norm_bt(v):
    """电池类型归一化：杂称（磷酸铁锂电池/蓄电池/LFP/储能单体种类…）合并为标准口径；
    '未提供'等视为缺失。核算依据 = 已有数据中的百分比占比。"""
    s = clean_str(v)
    if s is None or s in ('未提供', '未提供具体数据', '无', '/'):
        return None
    t = s.upper()
    if '钠' in s:
        return '钠离子'
    if '钛酸锂' in s:
        return '钛酸锂'
    if '锰酸锂' in s and '三元' not in s:
        return '锰酸锂'
    if '磷酸铁锂' in s or 'LFP' in t:
        return '磷酸铁锂'
    if '三元' in s or '镍钴锰' in s or 'NCM' in t:
        return '三元锂'
    if '锂' in s:
        return '其他锂离子'
    return '其他'


def norm_tax(v):
    """是否减免购置税归一化：仅接受 是/否，其余视为缺失。"""
    s = clean_str(v)
    if s in ('是', '否'):
        return s
    return None


# 物理合理范围：超出视为录入异常，置空（避免离群值污染统计）
# (字段名, BEV范围, PHEV/EREV范围)
NUM_RANGES = {
    'w':  ((300, 4500), (500, 4500)),    # 整备质量 kg
    'r':  ((50, 1500), (20, 600)),       # 纯电续航 km
    'c':  ((2, 250), (1, 120)),          # 电池容量 kWh
    'ed': ((50, 400), (30, 400)),        # 能量密度 Wh/kg
    'ec': ((3, 40), (3, 40)),            # 百公里电耗
    'fp': ((5, 1500), (5, 800)),         # 前电机功率 kW（v4.4.0，原峰值口径）
    'tp': ((5, 1500), (5, 800)),         # 电机总功率 kW（前+后）
    'ft': ((20, 30000), (20, 30000)),    # 前电机扭矩 Nm
    'rt': ((20, 30000), (20, 30000)),    # 后电机扭矩 Nm
    'tq': ((20, 30000), (20, 30000)),    # 扭矩 Nm（兼容保留）
    'fo': (None, (0.1, 20)),             # 综合油耗 L/100km（BEV不适用）
    'dv': (None, (300, 6000)),           # 排量 mL（BEV不适用）
    'ep': (None, (10, 600)),             # 发动机功率 kW（BEV不适用）
    'ab': ((1800, 4200), (1800, 4200)),  # 轴距 mm（乘用车合理区间）
    'lg': ((2600, 6600), (2600, 6600)),  # 车长 mm（预留列，补数后启用）
}


def sanitize(field, val, ptype):
    """按物理合理范围过滤异常值"""
    if val is None:
        return None
    rng_pair = NUM_RANGES.get(field)
    if not rng_pair:
        return val
    rng = rng_pair[0] if ptype == 'BEV' else rng_pair[1]
    if rng is None:
        return None  # 该动力类型不适用
    if rng[0] <= val <= rng[1]:
        return val
    return None


def load_records(path):
    """读取Excel并转为紧凑记录数组（v4.4.0：按表头名索引，电机三列 P/T 文本）"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    records = []
    skipped = 0
    cleaned = 0
    ptype_raw = {}          # id(rec) → 公告产品类型（细分市场规范化用）
    pname_raw = {}          # id(rec) → 公告产品名称（车身形式判据）
    it = ws.iter_rows(min_row=1, values_only=True)
    hdr = [str(c).strip() if c is not None else '' for c in next(it)]
    hi = {h: i for i, h in enumerate(hdr)}

    def cell(row, name):
        i = hi.get(name)
        return row[i] if i is not None and i < len(row) else None
    for row in it:
        if not row or row[0] is None:
            continue
        try:
            batch = int(str(row[0]).strip())
        except (ValueError, TypeError):
            skipped += 1
            continue
        ptype = clean_str(row[9]) or '未知'
        fp, ft = parse_pt(cell(row, '前电机功率/扭矩'))
        rp, rt = parse_pt(cell(row, '后电机功率/扭矩'))
        tp, tq = parse_pt(cell(row, '电机总功率/扭矩'))
        raw = {
            'w':  clean_num(cell(row, '整备质量(kg)')),
            'ab': clean_num(cell(row, '轴距(mm)')),
            'r':  clean_num(cell(row, '纯电续航里程(km)')),
            'c':  clean_num(cell(row, '电池容量(kWh)')),
            'ed': clean_num(cell(row, '电池能量密度(Wh/kg)')),
            'ec': clean_num(cell(row, '百公里电耗(kWh/100km)')),
            'fp': fp,                              # 前电机功率kW
            'tp': tp,                              # 电机总功率kW（前+后）
            'ft': ft, 'rt': rt,                    # 前/后电机扭矩 Nm
            'fo': clean_num(cell(row, '综合油耗(L/100km)')),
            'dv': clean_num(cell(row, '发动机排量(mL)')),
            'ep': clean_num(cell(row, '发动机功率(kW)')),
        }
        # 兼容：旧口径扭矩列缺失时退回总功率/扭矩第二段
        if raw['ft'] is None and tq is not None and not cell(row, '后电机功率/扭矩'):
            raw['ft'] = tq
        val = {}
        for k, v in raw.items():
            s = sanitize(k, v, ptype)
            if v is not None and s is None:
                cleaned += 1
            val[k] = s
        drive = 1 if (val['fp'] is not None and rp is not None) else (0 if val['fp'] is not None else None)
        # 系统扭矩：四驱=前+后，两驱=前（后置空时沿用总功率/扭矩第二段兜底）
        tq_sys = val['ft']
        if drive == 1 and val['rt'] is not None and tq_sys is not None:
            tq_sys = val['ft'] + val['rt']
        rec = [
            batch, ptype,
            clean_seg(row[8]),          # s 细分市场（'/'占位符视为未填报）
            clean_str(row[3]),           # e 企业名称
            val['w'], val['r'], val['c'],
            norm_bt(cell(row, '电池类型')),   # bt 电池类型（归一化口径）
            val['ed'], val['ec'], val['fp'], val['tp'],
            clean_str(cell(row, '电机生产企业')),  # ms 电机生产企业
            tq_sys, val['fo'], val['dv'], val['ep'],
            clean_str(cell(row, '发动机生产企业')),  # es 发动机生产企业
            clean_str(cell(row, '数据来源')),   # src 数据来源
            clean_str(row[1]),           # m 产品型号
            clean_str(row[2]),           # bd 产品商标
            norm_tax(row[10]),           # tx 是否减免购置税
            val['ab'],                   # ab 轴距mm（分布格局 Length 视角现用口径）
            clean_num(cell(row, '车长(mm)')),  # lg 车长mm（占位列）
            clean_str(cell(row, '通用名称')),   # gn 通用名称（散点悬浮/明细表展示）
            val['rt'], rp, drive,        # rt 后电机扭矩 / rp 后电机功率 / drive 两驱0·四驱1
            str(cell(row, '产品名称') or ''),   # pn 公告产品名称（公告分类口径，供核验展示）
            clean_num(cell(row, '月销量(辆)')),  # sn 月销量(参考项,易车车系零售口径)
        ]
        ptype_raw[id(rec)] = str(cell(row, '产品类型') or '')
        pname_raw[id(rec)] = str(cell(row, '产品名称') or '')
        records.append(rec)
    wb.close()

    # v4.7.1：细分市场规范化（公告车身 + 尺寸级别 → 具体级别，确保每行落位）
    seg_stat, seg_unresolved = normalize_segments(records, ptype_raw, pname_raw)
    # 回填「公告车身形式 | 公告产品名称」到记录 pn 位（供悬浮窗核验公告口径）
    for rec in records:
        raw_pn = pname_raw.get(id(rec), '')
        ann_body = seg_body_from_announcement(ptype_raw.get(id(rec)), raw_pn) or ''
        rec.append(f'{ann_body}|{raw_pn}')
    return records, skipped, cleaned, seg_stat, seg_unresolved


def seg_key(s):
    """细分市场排序键：Car/SUV/MPV 分组 → 字母级别(A→E) → 数字级别(2→5) → 未分级 → 其他"""
    t = str(s).strip()
    if t == 'Sports':
        return (3, 999, t)
    rest = t[4:] if t.startswith('Lux ') else t
    if rest == 'Car' or rest.startswith('Car-'):
        body = 0
    elif rest == 'SUV' or rest.startswith('SUV-'):
        body = 1
    elif rest == 'MPV' or rest.startswith('MPV-'):
        body = 2
    else:
        body = 9
    dash = rest.find('-')
    lvl = '' if dash < 0 else rest[dash+1:]
    if not lvl:
        lr = 999
    elif lvl.isalpha() and len(lvl) == 1:
        lr = ord(lvl.upper()) - 65
    elif lvl.isdigit():
        lr = 100 + int(lvl)
    else:
        lr = 500
    return (body, lr, t)


def build_meta(records, src_name):
    batches = sorted({r[0] for r in records})
    segs = sorted({r[2] for r in records if r[2]}, key=seg_key)
    types = sorted({r[1] for r in records})
    return {
        'total': len(records),
        'batchMin': batches[0] if batches else 0,
        'batchMax': batches[-1] if batches else 0,
        'batchCount': len(batches),
        'batches': batches,
        'segments': segs,
        'types': types,
        'generatedAt': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'),
        'sourceFile': os.path.basename(src_name),
        'fields': FIELDS,
    }


HTML_HEAD = r'''<!DOCTYPE html>
<!-- NEV公告数据看板 __REL_NAME__ · __COPYRIGHT__ · MIT License -->
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NEV公告数据看板__TITLE_SUFFIX__</title>
<style>
:root{
  --bg:#14181F; --bg2:#171C25; --card:#1C222D; --card2:#212936;
  --border:#2A3242; --border2:#39445A;
  --text:#ECE8E1; --dim:#8B94A7; --faint:#5C6577;
  --clay:#D97757; --steel:#6B9BD1; --sage:#7FA97F; --gold:#D4A94E;
  --plum:#A97FA9; --teal:#5BAFAF; --rose:#C97B7B;
  --grid:#202839;
  --ttbg:rgba(24,29,39,.96); --ttbd:#39445A;
  --serif:Georgia,'Noto Serif SC','Songti SC','SimSun',serif;
  --sans:-apple-system,'PingFang SC','Microsoft YaHei','Segoe UI',sans-serif;
  --mono:'SF Mono','Cascadia Code',Consolas,'Courier New',monospace;
  --shadow:0 8px 32px rgba(0,0,0,.35);
}
/* ---------- 日间主题 ---------- */
body.theme-light{
  --bg:#F4F2ED; --bg2:#EDEAE3; --card:#FFFFFF; --card2:#F7F5F0;
  --border:#E0DCD2; --border2:#C9C3B6;
  --text:#2B3A4A; --dim:#5F6B7C; --faint:#98A1B0;
  --grid:#E8E4DA;
  --ttbg:rgba(255,255,255,.98); --ttbd:#D5D0C4;
  --shadow:0 6px 24px rgba(80,70,55,.12);
}
body.theme-light .kpi-v{color:#1E2A38}
body.theme-light ::-webkit-scrollbar-thumb{background:#C9C3B6}
body.theme-light ::-webkit-scrollbar-track{background:#EDEAE3}
*{margin:0;padding:0;box-sizing:border-box}
html{scroll-behavior:smooth}
body{background:var(--bg);color:var(--text);font-family:var(--sans);font-size:14px;line-height:1.6;min-height:100vh;transition:background .3s,color .3s}
body::before{content:'';position:fixed;inset:0;z-index:0;pointer-events:none;
  background:radial-gradient(ellipse 90% 55% at 50% -12%,rgba(217,119,87,.09),transparent 60%),
             linear-gradient(rgba(128,128,128,.03) 1px,transparent 1px),
             linear-gradient(90deg,rgba(128,128,128,.03) 1px,transparent 1px);
  background-size:100% 100%,52px 52px,52px 52px}
.wrap{position:relative;z-index:1;max-width:1460px;margin:0 auto;padding:0 28px 60px}

/* ---------- Header ---------- */
header{padding:44px 0 26px;border-bottom:1px solid var(--border);position:relative}
header::after{content:'';position:absolute;left:0;bottom:-1px;width:120px;height:2px;background:var(--clay)}
.h-top{display:flex;justify-content:space-between;align-items:flex-end;gap:24px;flex-wrap:wrap}
.h-title{font-family:var(--serif);font-size:34px;font-weight:600;letter-spacing:.5px}
.h-title em{font-style:normal;color:var(--clay)}
.h-sub{color:var(--dim);font-size:14px;margin-top:6px;letter-spacing:.3px}
.h-right{display:flex;flex-direction:column;align-items:flex-end;gap:10px}
.h-tools{display:flex;gap:8px}
.tool-btn{background:var(--card2);border:1px solid var(--border2);color:var(--dim);border-radius:8px;
  padding:6px 14px;font-size:12.5px;cursor:pointer;transition:all .18s;font-family:var(--sans);white-space:nowrap}
.tool-btn:hover{color:var(--clay);border-color:var(--clay)}
.tool-btn.icon{font-size:15px;padding:6px 10px}
.h-meta{text-align:right;color:var(--dim);font-size:12.5px;line-height:1.9}
.h-meta b{color:var(--text);font-family:var(--mono);font-weight:600}
.h-meta .tag{display:inline-block;background:rgba(217,119,87,.12);border:1px solid rgba(217,119,87,.35);
  color:var(--clay);padding:2px 10px;border-radius:20px;font-size:11.5px;margin-left:8px;letter-spacing:1px}

/* ---------- Filter bar ---------- */
.filterbar{position:sticky;top:0;z-index:50;background:var(--bg);border:1px solid var(--border);border-radius:12px;
  padding:14px 18px;margin:22px 0;display:flex;align-items:center;gap:16px;flex-wrap:wrap;box-shadow:var(--shadow)}
body::before{z-index:0}
.f-group{display:flex;align-items:center;gap:10px}
.f-label{color:var(--faint);font-size:12px;letter-spacing:1px;white-space:nowrap}
select{background:var(--card2);color:var(--text);border:1px solid var(--border2);border-radius:7px;
  padding:6px 10px;font-family:var(--mono);font-size:13px;cursor:pointer;outline:none;transition:border-color .2s}
select:hover,select:focus{border-color:var(--clay)}
.btn-group{display:flex;gap:0;border:1px solid var(--border2);border-radius:7px;overflow:hidden}
.btn-group button{background:transparent;color:var(--dim);border:none;padding:6px 16px;cursor:pointer;
  font-size:13px;font-family:var(--sans);transition:all .18s}
.btn-group button:hover{color:var(--text);background:var(--card2)}
.btn-group button.on{background:var(--clay);color:#fff;font-weight:600}
.quick{display:flex;gap:6px}
.quick button{background:var(--card2);color:var(--dim);border:1px solid var(--border2);border-radius:6px;
  padding:4px 10px;font-size:12px;cursor:pointer;transition:all .18s}
.quick button:hover{color:var(--clay);border-color:var(--clay)}
.f-count{margin-left:auto;color:var(--dim);font-size:13px;white-space:nowrap}
.f-count b{color:var(--clay);font-family:var(--mono);font-size:16px}
.f-reset{background:transparent;border:1px solid var(--border2);color:var(--dim);border-radius:7px;
  padding:6px 14px;cursor:pointer;font-size:12.5px;transition:all .18s}
.f-reset:hover{color:var(--rose);border-color:var(--rose)}

/* ---------- 多选下拉 ---------- */
.mselect{position:relative}
.ms-btn{background:var(--card2);border:1px solid var(--border2);color:var(--text);border-radius:7px;
  padding:6px 12px;font-size:13px;cursor:pointer;min-width:170px;text-align:left;transition:border-color .2s;
  font-family:var(--sans)}
.ms-btn:hover,.ms-btn.open{border-color:var(--clay)}
.ms-btn .caret{float:right;color:var(--faint);margin-left:8px;transition:transform .2s}
.ms-btn.open .caret{transform:rotate(180deg)}
.ms-btn .n{color:var(--clay);font-family:var(--mono)}
.ms-panel{display:none;position:absolute;top:calc(100% + 6px);left:0;z-index:99;background:var(--card);
  border:1px solid var(--border2);border-radius:10px;box-shadow:var(--shadow);min-width:230px;max-width:280px;padding:8px}
.ms-panel.open{display:block}
.ms-actions{display:flex;justify-content:space-between;padding:4px 8px 8px;border-bottom:1px solid var(--border);margin-bottom:4px}
.ms-actions a{color:var(--steel);font-size:12px;cursor:pointer;text-decoration:none}
.ms-actions a:hover{text-decoration:underline}
.ms-list{max-height:260px;overflow-y:auto;padding:2px}
.ms-item{display:flex;align-items:center;gap:8px;padding:5px 8px;border-radius:6px;cursor:pointer;font-size:12.5px;transition:background .15s}
.ms-item:hover{background:var(--card2)}
.ms-item input{accent-color:var(--clay);cursor:pointer}
.ms-item .c{margin-left:auto;color:var(--faint);font-family:var(--mono);font-size:11px}

/* ---------- KPI ---------- */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:26px}
.kpi{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:18px 20px;
  position:relative;overflow:hidden;transition:transform .2s,border-color .2s;opacity:0;animation:rise .5s forwards}
.kpi:nth-child(1){animation-delay:.05s}.kpi:nth-child(2){animation-delay:.1s}
.kpi:nth-child(3){animation-delay:.15s}.kpi:nth-child(4){animation-delay:.2s}
.kpi:nth-child(5){animation-delay:.25s}.kpi:nth-child(6){animation-delay:.3s}
.kpi:hover{transform:translateY(-3px);border-color:var(--border2)}
.kpi::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--ac,var(--clay))}
.kpi-v{font-family:var(--mono);font-size:30px;font-weight:600;color:var(--text);letter-spacing:-.5px}
.kpi-v small{font-size:14px;color:var(--dim);font-weight:400;margin-left:3px}
.kpi-l{color:var(--dim);font-size:12.5px;margin-top:4px;letter-spacing:.5px}
.kpi-d{color:var(--faint);font-size:11.5px;margin-top:2px;font-family:var(--mono)}
@keyframes rise{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}

/* ---------- Tabs ---------- */
.tabs{display:flex;gap:2px;border-bottom:1px solid var(--border);margin-bottom:24px;flex-wrap:nowrap;overflow-x:auto}
.tab{background:transparent;border:none;color:var(--dim);font-size:16.5px;padding:12px 22px;cursor:pointer;
  position:relative;transition:color .18s;font-family:var(--sans);letter-spacing:.5px;white-space:nowrap;flex:0 0 auto}
.tab:hover{color:var(--text)}
.tab.on{color:var(--clay);font-weight:600}
.tab.on::after{content:'';position:absolute;left:14px;right:14px;bottom:-1px;height:2px;background:var(--clay)}

/* ---------- Panels & cards ---------- */
.panel{display:none}
.panel.on{display:block;animation:fadeUp .35s}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
.grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px}
.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:18px 20px 12px;
  transition:border-color .2s;min-width:0}
.card:hover{border-color:var(--border2)}
.c-s6{grid-column:span 6}.c-s12{grid-column:span 12}.c-s4{grid-column:span 4}.c-s8{grid-column:span 8}
.c-s5{grid-column:span 5}.c-s7{grid-column:span 7}.c-s3{grid-column:span 3}
.c-h{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px;gap:10px}
.c-t{font-size:15px;font-weight:600;letter-spacing:.4px}
.c-t::before{content:'▎';color:var(--clay);margin-right:4px}
.c-s{color:var(--faint);font-size:11.5px;white-space:nowrap}
.ci{color:var(--dim);font-size:12px;line-height:1.7;margin:2px 0 4px}
.ci em{font-style:normal;color:var(--clay);font-family:var(--mono);font-size:11.5px}
.chart{width:100%;height:340px}
.chart.tall{height:400px}
.chart.xl{height:640px}
.chart.short{height:280px}
/* 自定义分布控制条 */
.cd-bar{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:12px 0 6px;
  padding:10px 12px;background:var(--card2);border:1px solid var(--border);border-radius:9px}
/* 控制条分两行：第一行坐标轴与视图，第二行风格/配色/球径/聚合与操作按钮 */
.cd-line{display:flex;flex-wrap:wrap;align-items:center;gap:8px;width:100%}
.cd-line2{margin-top:9px;padding-top:9px;border-top:1px solid var(--border)}
.cd-bar select{min-width:150px;max-width:230px}
.cd-bar .f-label{margin-left:4px}
.cd-bar .f-label:first-child{margin-left:0}
.cd-zscale{display:flex;align-items:center;gap:8px;padding-left:10px;border-left:1px solid var(--border)}
.cd-zscale.hide{display:none}
.cd-zscale input[type=range]{width:120px;accent-color:var(--clay);cursor:pointer}
.cd-zval{color:var(--clay);font-family:var(--mono);font-size:11.5px;min-width:42px}
.cd-btn{background:var(--card);color:var(--dim);border:1px solid var(--border2);border-radius:6px;
  padding:6px 12px;font-size:12.5px;cursor:pointer;font-family:var(--sans);transition:all .15s}
.cd-btn:hover{color:var(--clay);border-color:var(--clay)}
.cd-btn.off{opacity:.5;border-style:dashed}
.cd-btn.dis{opacity:.45;cursor:not-allowed}
.cd-btn.on{color:var(--clay);border-color:var(--clay);opacity:1}
.cd-bar .hide{display:none!important}
/* 全屏态：卡片纵向弹性布局，图表吃满剩余空间 */
.card.cd-fs,.card:fullscreen{display:flex;flex-direction:column;gap:8px;padding:12px 16px;overflow:auto;background:var(--card)}
.card.cd-fs .c-h,.card:fullscreen .c-h,.card.cd-fs .ci,.card:fullscreen .ci,
.card.cd-fs .cd-bar,.card:fullscreen .cd-bar{flex:0 0 auto}
.card.cd-fs .chart.xl,.card:fullscreen .chart.xl{flex:1 1 auto;height:auto;min-height:420px}
.cd-hint{color:var(--faint);font-size:11.5px;margin-left:auto}
@media(max-width:1080px){.c-s6,.c-s4,.c-s8,.c-s5,.c-s7,.c-s3{grid-column:span 12}}

/* ---------- Insight strip ---------- */
.insight{background:linear-gradient(135deg,rgba(217,119,87,.08),rgba(107,155,209,.06));
  border:1px solid rgba(217,119,87,.25);border-radius:12px;padding:16px 22px;margin-bottom:22px;
  display:flex;gap:14px;align-items:flex-start}
.insight .ic{font-family:var(--serif);font-size:22px;color:var(--clay);line-height:1.3}
.insight .tx{color:var(--dim);font-size:13px;line-height:1.8}
.insight .tx b{color:var(--text)}
.insight .tx em{font-style:normal;color:var(--clay);font-family:var(--mono)}

/* ---------- 数据质量（底部） ---------- */
.quality-sec{margin-top:34px}
.q-head{display:flex;align-items:center;gap:14px;background:var(--card);border:1px solid var(--border);
  border-radius:12px;padding:16px 22px;cursor:pointer;transition:border-color .2s;user-select:none}
.q-head:hover{border-color:var(--clay)}
.q-head .q-ic{width:38px;height:38px;border-radius:10px;background:rgba(217,119,87,.12);display:flex;
  align-items:center;justify-content:center;font-size:18px;color:var(--clay)}
.q-head .q-t{font-size:15.5px;font-weight:600}
.q-head .q-d{color:var(--faint);font-size:12px;margin-top:2px}
.q-head .q-arrow{margin-left:auto;color:var(--faint);font-size:13px;transition:transform .25s}
.q-head.open .q-arrow{transform:rotate(180deg)}
.q-body{display:none;margin-top:14px}
.q-body.open{display:block;animation:fadeUp .3s}

/* ---------- 数据管理页 ---------- */
.dm-grid{display:grid;grid-template-columns:repeat(12,1fr);gap:16px}
.dm-note{color:var(--dim);font-size:12.5px;line-height:1.9;margin:8px 0 14px}
.dm-note b{color:var(--text)}
.dm-note code{background:var(--card2);border:1px solid var(--border);border-radius:5px;padding:1px 8px;
  font-family:var(--mono);font-size:11.5px;color:var(--sage)}
.dropzone{border:2px dashed var(--border2);border-radius:12px;padding:34px 20px;text-align:center;
  cursor:pointer;transition:all .2s;color:var(--dim);font-size:13.5px}
.dropzone:hover,.dropzone.over{border-color:var(--clay);background:rgba(217,119,87,.05);color:var(--text)}
.dropzone .dz-ic{font-size:30px;display:block;margin-bottom:8px;color:var(--clay)}
.dropzone .dz-hint{font-size:11.5px;color:var(--faint);margin-top:6px}
.exp-btn{background:var(--clay);color:#fff;border:none;border-radius:8px;padding:10px 22px;font-size:14px;
  cursor:pointer;font-family:var(--sans);font-weight:600;transition:all .18s;margin-right:10px}
.exp-btn:hover{filter:brightness(1.08);transform:translateY(-1px)}
.exp-btn.alt{background:var(--card2);color:var(--text);border:1px solid var(--border2);font-weight:400}
.exp-btn.alt:hover{border-color:var(--clay);color:var(--clay)}
.fld-list{display:flex;flex-wrap:wrap;gap:6px;margin:10px 0 16px}
.fld-chip{display:flex;align-items:center;gap:6px;background:var(--card2);border:1px solid var(--border);
  border-radius:6px;padding:4px 10px;font-size:12px;cursor:pointer;transition:all .15s;user-select:none}
.fld-chip:hover{border-color:var(--clay)}
.fld-chip input{accent-color:var(--clay);cursor:pointer}
.fld-chip.off{opacity:.45}
.dm-stat{display:flex;gap:26px;flex-wrap:wrap;margin-top:6px}
.dm-stat .s{text-align:left}
.dm-stat .v{font-family:var(--mono);font-size:24px;font-weight:600;color:var(--clay)}
.dm-stat .l{color:var(--faint);font-size:11.5px;margin-top:2px}

/* ---------- 明细查询表 ---------- */
.tbl-bar{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:12px}
.tbl-bar input{flex:1;min-width:220px;background:var(--card2);color:var(--text);border:1px solid var(--border2);
  border-radius:8px;padding:8px 14px;font-size:13px;font-family:var(--sans);outline:none;transition:border-color .2s}
.tbl-bar input:focus{border-color:var(--clay)}
#tblInfo{color:var(--dim);font-size:12.5px;white-space:nowrap}
#tblInfo b{color:var(--clay);font-family:var(--mono)}
.pg-btn{background:var(--card2);border:1px solid var(--border2);color:var(--dim);border-radius:7px;
  padding:6px 14px;font-size:12.5px;cursor:pointer;transition:all .18s;font-family:var(--sans)}
.pg-btn:hover:not(:disabled){color:var(--clay);border-color:var(--clay)}
.pg-btn:disabled{opacity:.35;cursor:not-allowed}
.tbl-wrap{overflow:auto;max-height:520px;border:1px solid var(--border);border-radius:10px}
#tblDetail{width:100%;border-collapse:collapse;font-size:12.5px;white-space:nowrap}
#tblDetail th{position:sticky;top:0;z-index:2;background:var(--card2);color:var(--dim);font-weight:600;
  padding:9px 12px;text-align:left;cursor:pointer;border-bottom:1px solid var(--border2);user-select:none;white-space:nowrap}
#tblDetail th:hover{color:var(--clay)}
#tblDetail th.on{color:var(--clay)}
#tblDetail td{padding:7px 12px;border-bottom:1px solid var(--border);color:var(--text);max-width:260px;overflow:hidden;text-overflow:ellipsis}
#tblDetail tbody tr:hover td{background:var(--card2)}
#tblDetail td.na{color:var(--faint)}

/* ---------- Footer ---------- */
footer{margin-top:44px;border-top:1px solid var(--border);padding-top:20px;color:var(--faint);
  font-size:12.5px;line-height:2;display:flex;justify-content:space-between;gap:20px;flex-wrap:wrap}
footer b{color:var(--dim)}
footer code{background:var(--card2);border:1px solid var(--border);border-radius:5px;padding:1px 8px;
  font-family:var(--mono);font-size:12px;color:var(--sage)}
#toTop{position:fixed;right:26px;bottom:26px;z-index:60;width:40px;height:40px;border-radius:50%;
  background:var(--card);border:1px solid var(--border2);color:var(--dim);font-size:16px;cursor:pointer;
  box-shadow:var(--shadow);display:none;transition:color .2s}
#toTop:hover{color:var(--clay)}
#toTop.show{display:block}

/* ---------- 批次断档图表：默认收起，可开关 ---------- */
.card.gap-hidden{display:none}
body.show-gaps .card.gap-hidden{display:block}
.card.gap-hidden-note{display:none;text-align:center;color:var(--faint);font-size:12.5px;padding:16px}
body.show-gaps .card.gap-hidden-note{display:none}
#btnGap.on{color:var(--clay);border-color:var(--clay)}

/* ---------- 趋势线公式 Toast ---------- */
#formulaToast{position:fixed;left:50%;bottom:38px;transform:translateX(-50%) translateY(18px);
  background:rgba(24,29,39,.96);color:#ECE8E1;border:1px solid var(--border2);border-radius:10px;
  padding:10px 20px;font-family:var(--mono);font-size:13px;z-index:120;opacity:0;pointer-events:none;
  transition:.25s;box-shadow:var(--shadow);max-width:82vw;text-align:center;line-height:1.6}
#formulaToast.show{opacity:1;transform:translateX(-50%) translateY(0)}

::-webkit-scrollbar{width:9px;height:9px}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:5px}
::-webkit-scrollbar-track{background:var(--bg2)}
</style>
</head>
<body>
<div class="wrap">

<header>
  <div class="h-top">
    <div>
      <div class="h-title">NEV公告数据<em>看板</em></div>
      <div class="h-sub">新能源汽车公告车型参数多维透视 · 批次 / 动力类型 / 驱动系统特征</div>
    </div>
    <div class="h-right">
      <div class="h-tools">
        <button class="tool-btn" id="btnImport" title="导入最新数据表">⟳ 导入数据</button>
        <button class="tool-btn icon" id="btnTheme" title="切换白天/黑夜模式">☀</button>
      </div>
      <div class="h-meta">
        <div>数据源：<b id="hSrc">__SRC_FILE__</b></div>
        <div>批次：<b id="hBRange">__BMIN__ ~ __BMAX__</b>（<b id="hBCnt">__BCNT__</b> 批）__REL_TAG__</div>
        <div>生成日期：<b id="hGen">__GEN_AT__</b></div>
      </div>
    </div>
  </div>
</header>

<div class="filterbar">
  <div class="f-group">
    <span class="f-label">批次范围</span>
    <select id="bFrom"></select>
    <span class="f-label" style="color:var(--faint)">→</span>
    <select id="bTo"></select>
  </div>
  <div class="quick">
    <button data-q="all">全部批次</button>
    <button data-q="last10">近10批</button>
    <button data-q="last20">近20批</button>
    <button data-q="latest">最新一批</button>
  </div>
  <div class="f-group">
    <span class="f-label">动力类型</span>
    <div class="btn-group" id="typeGroup">
      <button data-t="ALL" class="on">全部</button>
      <button data-t="BEV">BEV</button>
      <button data-t="PHEV">PHEV/EREV</button>
    </div>
  </div>
  <div class="f-group">
    <span class="f-label">月销量(参考)</span>
    <select id="salesF">
      <option value="ALL">全部</option>
      <option value="s5w">≥5万辆</option>
      <option value="s1w">1万~5万</option>
      <option value="s3k">3千~1万</option>
      <option value="lt3k">&lt;3千</option>
      <option value="na">未填报</option>
    </select>
  </div>
  <button class="tool-btn" id="btnGap" title="显示/隐藏存在批次断档（源数据缺失）的图表">◇ 缺口图表 <span id="gapBtnLabel">0</span></button>
  <div class="f-group">
    <span class="f-label">细分市场</span>
    <div class="mselect" id="segSelect">
      <button class="ms-btn" id="segBtn">全部细分市场<span class="caret">▾</span></button>
      <div class="ms-panel" id="segPanel">
        <div class="ms-actions"><a id="segSelAll">全选</a><a id="segClrAll">清空</a></div>
        <div class="ms-list" id="segList"></div>
      </div>
    </div>
  </div>
  <div class="f-count" id="fCount">已选 <b>0</b> / 0 条</div>
  <button class="f-reset" id="fReset">重置筛选</button>
</div>

<div class="kpis" id="kpis"></div>

<nav class="tabs" id="tabs">
  <button class="tab on" data-p="p1">总览</button>
  <button class="tab" data-p="p2">电机系统</button>
  <button class="tab" data-p="p3">发动机系统</button>
  <button class="tab" data-p="p4">电池系统</button>
  <button class="tab" data-p="pDist">分布格局</button>
  <button class="tab" data-p="pCustom">自定义分布</button>
  <button class="tab" data-p="p5">企业与品牌</button>
  <button class="tab" data-p="p6">数据管理</button>
</nav>

<!-- ============ Panel 1 总览 ============ -->
<section class="panel on" id="p1">
  <div class="insight"><div class="ic">§</div><div class="tx" id="insight1Tx"></div></div>
  <div class="grid">
    <div class="card c-s8"><div class="c-h"><div class="c-t">批次公告车型数量趋势</div><div class="c-s">BEV / PHEV/EREV 堆积柱状 · 悬停查看占比</div></div><div class="chart" id="chBatch"></div></div>
    <div class="card c-s4"><div class="c-h"><div class="c-t">动力类型结构</div><div class="c-s">当前筛选范围</div></div><div class="chart" id="chType"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">细分市场分布</div><div class="c-s">已填报口径 · 数量与占比</div></div><div class="chart" id="chSeg"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">公告企业 Top 12</div><div class="c-s">已填报口径 · 按公告车型数</div></div><div class="chart" id="chEnt"></div></div>
    <div class="card c-s12"><div class="c-h"><div class="c-t">驱动系统核心参数 · 批次演进</div><div class="c-s">按批次均值 · 淡色虚线为缺失批次趋势线（每参数独立）</div></div><div class="chart tall" id="chEvo"></div></div>
  </div>
</section>

<!-- ============ Panel 2 电机系统 ============ -->
<section class="panel" id="p2">
  <div class="insight"><div class="ic">§</div><div class="tx" id="insight2Tx"></div></div>
  <div class="grid">
    <div class="card c-s6"><div class="c-h"><div class="c-t">电机总功率分布</div><div class="c-s">BEV / PHEV/EREV 分组直方 · 已填报口径</div></div><div class="chart" id="chPwDist"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">电机供应商 Top 15</div><div class="c-s">配套数量与占比</div></div><div class="chart" id="chMsTop"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">功率 × 扭矩 散点</div><div class="c-s">点击图例"趋势线"可开关</div></div><div class="chart" id="chPtq"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">电机功率批次趋势</div><div class="c-s">均值 · 淡色虚线为趋势线</div></div><div class="chart" id="chPwTrend"></div></div>
    <div class="card c-s12"><div class="c-h"><div class="c-t">功率质量比分布</div><div class="c-s">电机总功率 / 整备质量（kW/t）· 反映驱动系统推重水平</div></div><div class="chart" id="chPwr"></div></div>
  </div>
</section>

<!-- ============ Panel 3 发动机系统 ============ -->
<section class="panel" id="p3">
  <div class="insight"><div class="ic">§</div><div class="tx" id="insight3Tx"></div></div>
  <div class="grid">
    <div class="card c-s6"><div class="c-h"><div class="c-t">发动机排量分布</div><div class="c-s">PHEV/EREV 车型 · 已填报口径</div></div><div class="chart" id="chDv"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">排量 × 功率 散点</div><div class="c-s">点击图例"趋势线"可开关</div></div><div class="chart" id="chDvEp"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">发动机供应商 Top 12</div><div class="c-s">配套数量与占比</div></div><div class="chart" id="chEsTop"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">综合油耗分布</div><div class="c-s">PHEV/EREV 亏电油耗 L/100km</div></div><div class="chart" id="chFo"></div></div>
    <div class="card c-s12"><div class="c-h"><div class="c-t">发动机参数批次趋势</div><div class="c-s">均值 · 缺失批次以趋势线衔接（淡色虚线）</div></div><div class="chart" id="chDvTrend"></div></div>
  </div>
</section>

<!-- ============ Panel 4 电池系统 ============ -->
<section class="panel" id="p4">
  <div class="insight"><div class="ic">§</div><div class="tx" id="insight4Tx"></div></div>
  <div class="grid">
    <div class="card c-s6"><div class="c-h"><div class="c-t">BEV 续航分布</div><div class="c-s">纯电续航里程 km · 已填报口径</div></div><div class="chart" id="chRgB"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">PHEV/EREV 纯电续航分布</div><div class="c-s">纯电续航里程 km</div></div><div class="chart" id="chRgP"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">容量 × 续航 散点</div><div class="c-s">点击图例"趋势线"可开关</div></div><div class="chart" id="chCr"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">电池类型构成趋势</div><div class="c-s">按批次 100% 堆积 · 已填报口径占比</div></div><div class="chart" id="chBtTrend"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">电池能量密度趋势</div><div class="c-s">Wh/kg 均值（BEV）· 趋势线衔接</div></div><div class="chart" id="chEd"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">百公里电耗趋势</div><div class="c-s">kWh/100km 均值 · 淡色虚线为趋势线</div></div><div class="chart" id="chEcTrend"></div></div>
    <div class="card c-s12"><div class="c-h"><div class="c-t">百公里电耗分布</div><div class="c-s">当前筛选范围 · 已填报口径</div></div><div class="chart" id="chEc"></div></div>
  </div>
</section>

<!-- ============ Panel D 分布格局 ============ -->
<section class="panel" id="pDist">
  <div class="insight"><div class="ic">§</div><div class="tx" id="insightDTx"></div></div>
  <div class="grid" id="distGrid"></div>
</section>

<!-- ============ Panel Custom 自定义分布 ============ -->
<section class="panel" id="pCustom">
  <div class="insight"><div class="ic">◎</div><div class="tx" id="insightCustomTx"></div></div>
  <div class="grid">
    <div class="card c-s12">
      <div class="c-h">
        <div class="c-t">自定义分布</div>
      </div>
      <div class="ci" id="ci_chCustom"></div>
      <div class="cd-bar">
        <div class="cd-line">
          <span class="f-label">横坐标 X</span>
          <select id="cdX"></select>
          <span class="f-label">纵坐标 Y</span>
          <select id="cdY"></select>
          <span class="f-label">高度坐标 Z</span>
          <select id="cdZ" title="Z 轴字段：三维=高度/位置，二维=泡泡直径（不选则二维等径散点）"></select>
          <span class="f-label">视图</span>
          <select id="cdView" title="展示模式：二维泡泡 · 三维柱状（方案A 海拔）· 三维泡泡（方案B 梯度）· 梯度曲面（方案C 地形）· 二维热力（三维类需先选定 Z 字段）">
            <option value="d2">二维散点/泡泡</option>
            <option value="bar">三维柱状 · 海拔（A）</option>
            <option value="bub">三维泡泡 · 梯度（B）</option>
            <option value="surf">梯度曲面 · 地形（C）</option>
            <option value="heat">二维热力</option>
          </select>
        </div>
        <div class="cd-line cd-line2">
          <span id="cdAggWrap" class="hide"><span class="f-label">聚合</span>
            <select id="cdAgg" title="曲面高度聚合方式：均值 / 中位数 / 样本计数">
              <option value="mean">均值</option>
              <option value="median">中位数</option>
              <option value="count">样本计数</option>
            </select>
          </span>
          <span id="cdStyleWrap">
            <span class="f-label">风格</span>
            <select id="cdStyle" title="三维/热力视觉风格（8 款按需切换）"></select>
          </span>
          <span id="cdColorWrap">
            <span class="f-label">配色</span>
            <select id="cdColorMode" title="配色模式：动力类型双色系（BEV 暖橙 / PHEV-EREV 冷蓝）或 Z 梯度单色系">
              <option value="dual">动力类型双色系</option>
              <option value="z">Z 梯度单色系</option>
            </select>
          </span>
          <span id="cdSizeWrap" class="hide"><span class="f-label">球径</span>
            <select id="cdSize" title="泡泡直径字段（不选则等大；仅方案B 三维泡泡生效）"></select>
          </span>
          <span class="cd-zscale hide" id="cdZScaleWrap">
            <span class="f-label">泡泡比例</span>
            <input type="range" id="cdZScale" min="20" max="260" value="100" step="5">
            <span class="cd-zval" id="cdZVal">100%</span>
          </span>
          <span id="cd3dOnly" class="hide">
            <button class="cd-btn" id="cdViewReset" title="恢复默认视角（等轴测，等同双击图表）">视角复位</button>
            <button class="cd-btn" id="cdViewTop" title="切换为俯视视角（正对 X-Y 平面，透视距离已拉远）">俯视 X-Y</button>
          </span>
          <span id="cdAxisGrp">
            <span class="f-label">反向</span>
            <button class="cd-btn" id="cdInvX" title="X 轴排序方向反转">X↓</button>
            <button class="cd-btn" id="cdInvY" title="Y 轴排序方向反转">Y↓</button>
            <button class="cd-btn" id="cdInvZ" title="Z 梯度方向反转（三维/热力）">Z↓</button>
          </span>
          <button class="cd-btn" id="cdSwap" title="交换 X / Y">⇄ 交换</button>
          <span id="cd2dOnly">
            <button class="cd-btn" id="cdLevels" title="在轴距轴上标注 A/B/C 等级区段">等级区段 ✓</button>
            <button class="cd-btn" id="cdAuto" title="随缩放自动调节散点/泡泡显示大小">自动调节 ✓</button>
            <button class="cd-btn" id="cdZoomReset" title="恢复完整视野（等同双击图表）">复位视窗</button>
          </span>
          <button class="cd-btn" id="cdFull" title="全屏展示（Esc 退出）">⛶ 全屏</button>
          <button class="cd-btn" id="cdReset" title="恢复默认维度与视图">重置</button>
        </div>
      </div>
      <div class="chart xl" id="chCustom"></div>
    </div>
  </div>
</section>

<!-- ============ Panel 5 企业与品牌 ============ -->
<section class="panel" id="p5">
  <div class="insight"><div class="ic">§</div><div class="tx" id="insightEntTx"></div></div>
  <div class="grid">
    <div class="card c-s12"><div class="c-h"><div class="c-t">Top 8 企业公告量 · 批次演进</div><div class="c-s">堆叠面积图 · 跟踪头部企业公告格局变化</div></div><div class="chart tall" id="chEntTrend"></div></div>
    <div class="card c-s7"><div class="c-h"><div class="c-t">企业 × 细分市场矩阵</div><div class="c-s">Top 10 企业 × 已填报细分市场 · 公告数量热力</div></div><div class="chart tall" id="chEntSeg"></div></div>
    <div class="card c-s5"><div class="c-h"><div class="c-t">各批次新进入企业数</div><div class="c-s">首次出现公告的企业数与累计在营企业 · 跟踪新玩家入场节奏</div></div><div class="chart tall" id="chNewEnt"></div></div>
    <div class="card c-s6"><div class="c-h"><div class="c-t">品牌 Top 15</div><div class="c-s">产品商标口径 · 按公告车型数</div></div><div class="chart" id="chBrandTop"></div></div>
  </div>
</section>

<!-- ============ Panel 6 数据管理 ============ -->
<section class="panel" id="p6">
  <div class="insight"><div class="ic">⇅</div><div class="tx">导出当前筛选数据供进一步细化分析；导入最新公告数据表后，<b>全局数据与图表将实时刷新</b>（无需重新运行生成脚本）。</div></div>
  <div class="dm-grid">
    <div class="card c-s7">
      <div class="c-h"><div class="c-t">导出当前筛选数据</div><div class="c-s">Excel / CSV</div></div>
      <div class="dm-stat" id="dmStat"></div>
      <p class="dm-note" style="margin-top:14px">将导出<b>当前筛选条件</b>命中的全部记录（含所选批次、动力类型、细分市场）。可勾选需要保留的字段；缺失值以空单元格呈现，不参与统计。</p>
      <div class="fld-list" id="fldList"></div>
      <div>
        <button class="exp-btn" id="btnXlsx">导出 Excel (.xlsx)</button>
        <button class="exp-btn alt" id="btnCsv">导出 CSV</button>
      </div>
    </div>
    <div class="card c-s5">
      <div class="c-h"><div class="c-t">导入最新数据表</div><div class="c-s">页面内实时刷新</div></div>
      <p class="dm-note">支持 <code>.xlsx / .xls</code>，需包含 <code>NEV公告参数汇总</code> 工作表（或首表），字段结构同原始汇总表。导入后自动完成清洗、物理范围校验与全图表刷新。</p>
      <div class="dropzone" id="dropzone">
        <span class="dz-ic">⬆</span>
        点击选择或拖拽 Excel 文件到此处
        <div class="dz-hint">导入成功后页头数据源信息与全部筛选选项同步更新</div>
      </div>
      <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none">
      <p class="dm-note" id="importStatus" style="margin-top:12px;color:var(--faint)"></p>
    </div>
    <div class="card c-s12">
      <div class="c-h"><div class="c-t">车型明细查询</div><div class="c-s">当前筛选范围内检索 · 点击列头排序 · 缺失值以 – 呈现</div></div>
      <div class="tbl-bar">
        <input id="tblSearch" placeholder="搜索企业 / 商标 / 产品型号…" autocomplete="off">
        <span id="tblInfo"></span>
        <button class="pg-btn" id="tblPrev">← 上一页</button>
        <button class="pg-btn" id="tblNext">下一页 →</button>
      </div>
      <div class="tbl-wrap">
        <table id="tblDetail"><thead><tr></tr></thead><tbody></tbody></table>
      </div>
    </div>
  </div>
</section>

<!-- ============ 数据质量（页面底部） ============ -->
<section class="quality-sec" id="qualitySec">
  <div class="q-head open" id="qHead">
    <div class="q-ic">◈</div>
    <div>
      <div class="q-t">数据质量检查</div>
      <div class="q-d" id="qDesc">字段填报完整度 · 数据来源构成 · 批次×字段热力 · 点击折叠/展开</div>
    </div>
    <span class="q-arrow">▾</span>
  </div>
  <div class="q-body open" id="qBody">
    <div class="insight" style="margin-top:14px"><div class="ic">§</div><div class="tx" id="insight5Tx"></div></div>
    <div class="grid">
      <div class="card c-s7"><div class="c-h"><div class="c-t">字段填报完整度</div><div class="c-s">非空率 %（当前筛选范围）</div></div><div class="chart tall" id="chFill"></div></div>
      <div class="card c-s5"><div class="c-h"><div class="c-t">数据来源构成</div><div class="c-s">合并来源批次</div></div><div class="chart" id="chSrc"></div></div>
      <div class="card c-s12"><div class="c-h"><div class="c-t">批次 × 字段 填报热力</div><div class="c-s">绿=完整 红=缺失</div></div><div class="chart tall" id="chHeat"></div></div>
    </div>
  </div>
</section>

<footer>
  <div style="text-align:center;width:100%">
    <span style="font-size:12px;color:var(--dim)">__COPYRIGHT__ · 保留所有权利</span>
  </div>
</footer>

</div>
<button id="toTop" title="返回顶部">↑</button>

<script>__ECHARTS_LIB__</script>
<script>__ECHARTSGL_LIB__</script>
<script>__XLSX_LIB__</script>
<script>
const META = __META_JSON__;
const RAW_INIT = __DATA_JSON__;
</script>
<script>
'use strict';
/* ==================== 索引与全局 ==================== */
const I = {b:0,t:1,s:2,e:3,w:4,r:5,c:6,bt:7,ed:8,ec:9,fp:10,tp:11,ms:12,tq:13,fo:14,dv:15,ep:16,es:17,src:18,
  m:19,bd:20,tx:21,ab:22,lg:23,gn:24,rt:25,rp:26,drive:27,pn:28,sn:29};
const TX_IDX = I.tx;   // 免征购置税：数据层保留，展示层全量隐藏（v3.7口径）
const COL_NAMES = ['批次','动力类型','细分市场','企业名称','整备质量(kg)','纯电续航(km)','电池容量(kWh)','电池类型',
  '能量密度(Wh/kg)','百公里电耗(kWh/100km)','前电机功率(kW)','电机总功率(kW)','电机生产企业','系统扭矩(Nm)',
  '综合油耗(L/100km)','发动机排量(mL)','发动机功率(kW)','发动机生产企业','数据来源',
  '产品型号','产品商标','是否减免购置税','轴距(mm)','车长(mm)','通用名称','后电机扭矩(Nm)','后电机功率(kW)','驱动形式','公告产品名称','月销量(辆)'];
let RAW = RAW_INIT;
let META_CUR = META;

/* ==================== 主题 ==================== */
const THEMES = {
  dark:  {axis:'#8B94A7',axisLine:'#2A3242',split:'#202839',lbl:'#B9C0CE',
          ttBg:'rgba(24,29,39,.96)',ttBd:'#39445A',ttTx:'#ECE8E1',pieBd:'#1C222D',
          fitOp:.38, scatterA:.55},
  light: {axis:'#5F6B7C',axisLine:'#D5D0C4',split:'#E8E4DA',lbl:'#3D4A5C',
          ttBg:'rgba(255,255,255,.98)',ttBd:'#C9C3B6',ttTx:'#2B3A4A',pieBd:'#FFFFFF',
          fitOp:.45, scatterA:.5}
};
let THEME = 'dark';
const TC = ()=>THEMES[THEME];
let TT = {};
function refreshTT(){ TT = {backgroundColor:TC().ttBg,borderColor:TC().ttBd,borderWidth:1,
  textStyle:{color:TC().ttTx,fontSize:12},
  extraCssText:'box-shadow:0 8px 24px rgba(0,0,0,.25);border-radius:8px;'}; }
refreshTT();

function toggleTheme(){
  THEME = THEME==='dark'?'light':'dark';
  document.body.classList.toggle('theme-light', THEME==='light');
  document.getElementById('btnTheme').textContent = THEME==='dark'?'☀':'☾';
  refreshTT();
  updateAll();
}

const PAL = ['#D97757','#6B9BD1','#7FA97F','#D4A94E','#A97FA9','#5BAFAF','#C97B7B','#8B94A7','#B08968','#6A8FB5','#8FBF8F','#C4B454','#B98BC9','#7BC9C9'];
const AXS = (extra)=>Object.assign({
  axisLabel:{color:TC().axis,fontSize:11},
  axisLine:{lineStyle:{color:TC().axisLine}},
  axisTick:{show:false},
  splitLine:{lineStyle:{color:TC().split}}
},extra||{});
const LG = (extra)=>Object.assign({textStyle:{color:TC().axis,fontSize:12},itemWidth:14,itemHeight:9,top:0},extra||{});
const GRID = (extra)=>Object.assign({left:56,right:40,top:42,bottom:56,containLabel:true},extra||{});
const fmt = n => n==null?'-':n.toLocaleString('zh-CN');
const pct = (a,b)=> b>0 ? (a/b*100) : 0;
const p1 = (a,b)=> pct(a,b).toFixed(1);

/* ==================== 卡片要点行（随全局筛选联动） ==================== */
/* initCardInsights 为每个图表卡片自动插入 .ci 要素行；图表函数内调用 setInsight 写入一句话核心要点 */
const CHART_INSIGHT_IDS = ['chBatch','chType','chSeg','chEnt','chEvo',
  'chPwDist','chMsTop','chPtq','chPwTrend','chPwr',
  'chDv','chDvEp','chEsTop','chFo','chDvTrend',
  'chRgB','chRgP','chCr','chBtTrend','chEd','chEcTrend','chEc',
  'dRvsLen','dCvsRg','dEcVsRg','dEdVsRg','dMassLen','chCustom',
  'chEntTrend','chEntSeg','chNewEnt','chBrandTop',
  'chFill','chSrc','chHeat'];
function initCardInsights(){
  for(const id of CHART_INSIGHT_IDS){
    const el=document.getElementById(id); if(!el) continue;
    const card=el.closest('.card'); if(!card) continue;
    if(document.getElementById('ci_'+id)) continue;
    const d=document.createElement('div'); d.className='ci'; d.id='ci_'+id;
    const h=card.querySelector('.c-h');
    if(h) h.insertAdjacentElement('afterend', d); else card.appendChild(d);
  }
}
function setInsight(id, html){
  const el=document.getElementById('ci_'+id);
  if(el) el.innerHTML = html;
}
/* v3.7.3：分布格局要点为分燃料类型均值与极值（原先的 r 值描述已按需求移除） */

/* ==================== 细分市场排序（CAR/SUV/MPV 分组，字母级别→数字级别→未分级） ==================== */
function segKey(s){
  const t=String(s).trim();
  let body=9, rest=t;
  if(t==='Sports') return [3,999,t];
  if(t.startsWith('Lux ')){ rest=t.slice(4); }
  const bb = rest==='Car'||rest.startsWith('Car-') ? 0
           : rest==='SUV'||rest.startsWith('SUV-') ? 1
           : rest==='MPV'||rest.startsWith('MPV-') ? 2 : 9;
  body = bb;
  const dash = rest.indexOf('-');
  const lvl = dash<0 ? '' : rest.slice(dash+1);
  let lr;
  if(!lvl) lr = 999;
  else if(/^[A-Za-z]$/.test(lvl)) lr = lvl.toUpperCase().charCodeAt(0)-65;      // A=0…
  else if(/^\d+$/.test(lvl)) lr = 100+parseInt(lvl,10);                          // 2,3,4,5…
  else lr = 500;
  return [body,lr,t];
}

/* ==================== 状态 ==================== */
/* segs：已勾选（纳入统计）的细分市场集合；segAll=true 表示全选（不过滤） */
const state = {bFrom:META.batchMin, bTo:META.batchMax, type:'ALL', segs:new Set(), segAll:true, sales:'ALL'};

/* ==================== 数据工具（缺失不计入统计） ==================== */
function filtered(){
  return RAW.filter(r=>{
    if(r[I.b] < state.bFrom || r[I.b] > state.bTo) return false;
    if(state.type!=='ALL' && r[I.t]!==state.type) return false;
    if(!state.segAll){
      const s = r[I.s];
      if(s==null || !state.segs.has(s)) return false;   // 仅有勾选的细分市场纳入统计
    }
    if(state.sales!=='ALL'){
      const v = r[I.sn];
      if(state.sales==='na'){ if(v!=null && v!=='') return false; }
      else {
        if(v==null || v==='') return false;
        if(state.sales==='s5w' && v<50000) return false;
        if(state.sales==='s1w' && (v<10000 || v>=50000)) return false;
        if(state.sales==='s3k' && (v<3000 || v>=10000)) return false;
        if(state.sales==='lt3k' && v>=3000) return false;
      }
    }
    return true;
  });
}
/* 分组计数：跳过未填报（不计入统计与百分比） */
function groupCount(rows,idx,topN){
  const m = new Map();
  for(const r of rows){
    const v = r[idx];
    if(v==null||v==='') continue;   // 缺失不计入
    m.set(v,(m.get(v)||0)+1);
  }
  let arr = [...m.entries()].sort((a,b)=>b[1]-a[1]);
  if(topN) arr = arr.slice(0,topN);
  return arr;
}
function avg(rows,idx){
  let s=0,n=0;
  for(const r of rows){const v=r[idx];if(v!=null&&isFinite(v)){s+=v;n++;}}
  return n?s/n:null;
}
function buckets(rows,idx,edges,labels){
  const cnt = labels.map(()=>0);
  for(const r of rows){
    const v=r[idx];
    if(v==null) continue;   // 缺失不计入
    let k=labels.length-1;
    for(let i=0;i<edges.length;i++){ if(v<edges[i]){k=i;break;} }
    cnt[k]++;
  }
  return cnt;
}
function byBatch(rows){
  const bm = new Map();
  for(const r of rows){
    if(!bm.has(r[I.b])) bm.set(r[I.b],[]);
    bm.get(r[I.b]).push(r);
  }
  return META_CUR.batches.filter(b=>b>=state.bFrom&&b<=state.bTo).map(b=>[b, bm.get(b)||[]]);
}

/* ==================== 趋势线 ==================== */
/* 折线缺失点：线性插值填充；边界缺失：最近值延伸 */
function linInterp(vals){
  const n = vals.length;
  const filled = vals.slice();
  const fitFlag = vals.map(()=>false);
  let has = false;
  for(const v of vals){ if(v!=null){has=true;break;} }
  if(!has) return {filled,fitFlag};
  for(let i=0;i<n;i++){
    if(vals[i]!=null) continue;
    let prev=null,next=null;
    for(let j=i-1;j>=0;j--){if(vals[j]!=null){prev=j;break;}}
    for(let j=i+1;j<n;j++){if(vals[j]!=null){next=j;break;}}
    if(prev!=null&&next!=null){
      const t=(i-prev)/(next-prev);
      filled[i]=+(vals[prev]+(vals[next]-vals[prev])*t).toFixed(2);
    } else if(prev!=null){ filled[i]=vals[prev]; }
    else if(next!=null){ filled[i]=vals[next]; }
    fitFlag[i]=true;
  }
  return {filled,fitFlag};
}
/* 生成 [主系列(实线,connectNulls:false), 趋势线系列(淡色虚线,仅缺失点,connectNulls:true)] */
function mkFitPair(name, vals, color, extra){
  const {filled,fitFlag} = linInterp(vals);
  const fitOnly = filled.map((v,i)=>fitFlag[i]?v:null);
  const base = Object.assign({name,type:'line',data:vals,smooth:true,symbol:'circle',symbolSize:5,
    connectNulls:false,itemStyle:{color},lineStyle:{width:2.5}},extra||{});
  const fit = Object.assign({name:name+'趋势线',type:'line',data:fitOnly,smooth:true,symbol:'none',
    connectNulls:true,itemStyle:{color,opacity:0},lineStyle:{color,width:1.5,type:'dashed',opacity:TC().fitOp},
    silent:false,tooltip:{formatter:p=>`${p.seriesName}<br/>${p.name}：<b>${p.value}</b>（趋势线）`}},extra||{});
  return [base,fit];
}
/* 散点趋势线（最小二乘），独立系列加入图例可开关；name 参数用于分组/分参数命名 */
function trendSeries(pts, color, name){
  if(!pts||pts.length<3) return null;
  const n=pts.length;
  const sx=pts.reduce((a,p)=>a+p[0],0), sy=pts.reduce((a,p)=>a+p[1],0);
  const sxx=pts.reduce((a,p)=>a+p[0]*p[0],0), sxy=pts.reduce((a,p)=>a+p[0]*p[1],0);
  const den=n*sxx-sx*sx;
  if(!den) return null;
  const slope=(n*sxy-sx*sy)/den, icept=(sy-slope*sx)/n;
  const xs=pts.map(p=>p[0]);
  const x1=Math.min(...xs), x2=Math.max(...xs);
  const formula=`${name||'趋势线'} · 线性趋势：y = ${slope.toFixed(3)}x + ${icept.toFixed(1)}`;
  return {name:name||'趋势线',type:'line',data:[[x1,+(slope*x1+icept).toFixed(1)],[x2,+(slope*x2+icept).toFixed(1)]],
    showSymbol:false,smooth:false,lineStyle:{color,width:2,type:'dashed',opacity:.65},
    itemStyle:{color},tooltip:{formatter:()=>formula},__formula:formula,z:9};
}

/* ==================== 图表实例池 ==================== */
const CHARTS = {};
function chart(id){
  const el = document.getElementById(id);
  if(!el) return null;
  if(!CHARTS[id]) CHARTS[id] = echarts.init(el,null,{renderer:'canvas'});
  if(el.offsetWidth>0 && el.offsetHeight>0) CHARTS[id].resize();
  return CHARTS[id];
}
window.addEventListener('resize',()=>{Object.values(CHARTS).forEach(c=>{try{c.resize();}catch(e){}});});

/* ==================== 批次断档检测（源数据缺失致图表不连续） ==================== */
const GAP_CANDIDATES = ['chEvo','chPwTrend','chDvTrend','chEd','chEcTrend'];
const GAP_IDS = new Set();
function resetGap(){
  GAP_IDS.clear();
  GAP_CANDIDATES.forEach(id=>{ const el=document.getElementById(id); if(el&&el.closest('.card')) el.closest('.card').classList.remove('gap-hidden'); });
}
/* 判定：批次轴上任意相邻两个已出现批次间隔 ≥3（即整段批次源数据缺失），视为断档 */
function batchGapped(labels){
  const nums=[...new Set(labels.map(Number))].sort((a,b)=>a-b);
  for(let i=1;i<nums.length;i++){ if(nums[i]-nums[i-1]>=3) return true; }
  return false;
}
function markGap(id, labels){
  if(!batchGapped(labels)) return;
  GAP_IDS.add(id);
  const el=document.getElementById(id); if(el&&el.closest('.card')) el.closest('.card').classList.add('gap-hidden');
}
function updateGapBtn(){
  const lab=document.getElementById('gapBtnLabel'); if(lab) lab.textContent=`缺口图表 (${GAP_IDS.size})`;
  const gb=document.getElementById('btnGap'); if(gb) gb.classList.toggle('on', document.body.classList.contains('show-gaps'));
}
function initGapBtn(){
  const gb=document.getElementById('btnGap'); if(!gb) return;
  gb.onclick=()=>{ document.body.classList.toggle('show-gaps'); updateGapBtn(); };
}

/* ==================== 趋势线公式 Toast ==================== */
const TREND_REGS = {};   // chartId -> {seriesName: 公式文本}
let _toastEl=null, _toastT;
function showToast(msg){
  if(!_toastEl){ _toastEl=document.createElement('div'); _toastEl.id='formulaToast'; document.body.appendChild(_toastEl); }
  _toastEl.innerHTML=msg; _toastEl.classList.add('show');
  clearTimeout(_toastT); _toastT=setTimeout(()=>_toastEl.classList.remove('show'),4200);
}
/* 为散点图绑定「点击趋势线显示公式」（幂等，仅绑一次） */
function bindTrendClick(chartId){
  const inst=chart(chartId); if(!inst||inst.__tc) return; inst.__tc=true;
  inst.on('click', p=>{
    const m=TREND_REGS[chartId]; if(!m) return;
    const f=m[p.seriesName]; if(f) showToast(f);
  });
}
/* 散点 tooltip 车型名：产品商标 + 通用名称 */
const brandName = p => { const b=p.data.b, g=p.data.n; return [b,g].filter(x=>x&&x!=='-').join(' · ') || '–'; };

/* ==================== KPI ==================== */
function animNum(el,to,dec){
  const from = parseFloat(el.dataset.v||0)||0;
  const t0 = performance.now(), dur=600;
  function step(t){
    const p = Math.min(1,(t-t0)/dur), e = 1-Math.pow(1-p,3);
    const v = from+(to-from)*e;
    el.textContent = dec? v.toFixed(dec) : Math.round(v).toLocaleString('zh-CN');
    if(p<1) requestAnimationFrame(step); else el.dataset.v = to;
  }
  requestAnimationFrame(step);
}
const KPI_DEFS = [
  {k:'车型总数',u:'款',ac:'#D97757'},
  {k:'覆盖批次',u:'批',ac:'#6B9BD1'},
  {k:'公告企业',u:'家',ac:'#7FA97F'},
  {k:'BEV 占比',u:'%',ac:'#D4A94E',dec:1},
  {k:'平均纯电续航',u:'km',ac:'#A97FA9',dec:0},
  {k:'平均电机总功率',u:'kW',ac:'#5BAFAF',dec:0},
];
function initKPIs(){
  document.getElementById('kpis').innerHTML = KPI_DEFS.map(d=>
    `<div class="kpi" style="--ac:${d.ac}">
       <div class="kpi-v"><span id="kv_${d.k}">0</span><small>${d.u}</small></div>
       <div class="kpi-l">${d.k}</div>
       <div class="kpi-d" id="kd_${d.k}"></div>
     </div>`).join('');
}
function updateKPIs(rows){
  const n = rows.length;
  const bev = rows.filter(r=>r[I.t]==='BEV').length;
  const batchSet = new Set(rows.map(r=>r[I.b]));
  const entSet = new Set(rows.map(r=>r[I.e]).filter(Boolean));
  animNum(document.getElementById('kv_车型总数'), n);
  animNum(document.getElementById('kv_覆盖批次'), batchSet.size);
  animNum(document.getElementById('kv_公告企业'), entSet.size);
  animNum(document.getElementById('kv_BEV 占比'), n?bev/n*100:0, 1);
  animNum(document.getElementById('kv_平均纯电续航'), avg(rows,I.r)||0, 0);
  animNum(document.getElementById('kv_平均电机总功率'), avg(rows,I.tp)||0, 0);
  document.getElementById('kd_车型总数').textContent = `当前筛选 ${((n/META_CUR.total)*100).toFixed(1)}%`;
  document.getElementById('kd_覆盖批次').textContent = `范围 ${state.bFrom}~${state.bTo}`;
  document.getElementById('kd_公告企业').textContent = `BEV ${bev} / PHEV/EREV ${n-bev}`;
  document.getElementById('kd_BEV 占比').textContent = `PHEV/EREV ${(n? (n-bev)/n*100:0).toFixed(1)}%`;
  const rB = avg(rows.filter(r=>r[I.t]==='BEV'),I.r), rP = avg(rows.filter(r=>r[I.t]==='PHEV'),I.r);
  document.getElementById('kd_平均纯电续航').textContent = (rB||rP)?`BEV ${(rB||0).toFixed(0)} / PHEV/EREV ${(rP||0).toFixed(0)}`:'暂无数据';
  const pB = avg(rows.filter(r=>r[I.t]==='BEV'),I.tp), pP = avg(rows.filter(r=>r[I.t]==='PHEV'),I.tp);
  document.getElementById('kd_平均电机总功率').textContent = (pB||pP)?`BEV ${(pB||0).toFixed(0)} / PHEV/EREV ${(pP||0).toFixed(0)}`:'暂无数据';
}

/* ==================== Panel 1 ==================== */
function chBatch(rows){
  const bs = byBatch(rows);
  const labels = bs.map(x=>x[0]+'');
  const bev = bs.map(x=>x[1].filter(r=>r[I.t]==='BEV').length);
  const phev = bs.map(x=>x[1].filter(r=>r[I.t]==='PHEV').length);
  const totB = bs.map(x=>x[1].length);
  const pkI = totB.indexOf(Math.max(...totB));
  setInsight('chBatch', `区间 <em>${labels.length}</em> 批共 <em>${fmt(rows.length)}</em> 款 · 峰值第 <em>${labels[pkI]}</em> 批（<em>${fmt(totB[pkI])}</em> 款）`);
  chart('chBatch').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter(ps){
        const i=ps[0].dataIndex, t=bev[i]+phev[i];
        return `<b>第 ${labels[i]} 批</b><br>`+
          `<span style="color:#D97757">■</span> BEV：${bev[i]} 款（${t?pct(bev[i],t).toFixed(1):0}%）<br>`+
          `<span style="color:#6B9BD1">■</span> PHEV/EREV：${phev[i]} 款（${t?pct(phev[i],t).toFixed(1):0}%）<br>`+
          `合计：<b>${t}</b> 款`;
      }},TT),
    legend:LG({data:['BEV','PHEV/EREV']}),
    grid:GRID({bottom:64}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10,interval:0})})),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series:[
      {name:'BEV',type:'bar',stack:'t',data:bev,itemStyle:{color:'#D97757'},barMaxWidth:22},
      {name:'PHEV/EREV',type:'bar',stack:'t',data:phev,itemStyle:{color:'#6B9BD1'},barMaxWidth:22}
    ]
  },true);
}
function chType(rows){
  const bev = rows.filter(r=>r[I.t]==='BEV').length, phev = rows.filter(r=>r[I.t]==='PHEV').length;
  const data = [];
  if(bev>0) data.push({value:bev,name:'BEV',itemStyle:{color:'#D97757'}});
  if(phev>0) data.push({value:phev,name:'PHEV/EREV',itemStyle:{color:'#6B9BD1'}});
  setInsight('chType', `BEV <em>${fmt(bev)}</em> 款（<em>${p1(bev,rows.length)}%</em>）· PHEV/EREV <em>${fmt(phev)}</em> 款（<em>${p1(phev,rows.length)}%</em>）`);
  chart('chType').setOption({
    tooltip:Object.assign({trigger:'item',formatter:p=>`${p.name}：<b>${p.value}</b> 款（${p.percent}%）`},TT),
    series:[{type:'pie',radius:['48%','72%'],center:['50%','54%'],
      label:{color:TC().axis,formatter:'{b}\n{d}%',fontSize:12},
      labelLine:{lineStyle:{color:TC().axisLine}},
      itemStyle:{borderColor:TC().pieBd,borderWidth:2},data}]
  },true);
}
/* 横向条形：分母=已填报总数（缺失不计入百分比）；带双轴名称标注 */
function hbar(id, entries, color, unit, catName){
  const total = entries.reduce((a,b)=>a+b[1],0);   // 已填报口径
  const names = entries.map(e=>e[0]).reverse();
  const vals = entries.map(e=>e[1]).reverse();
  chart(id).setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>`${ps[0].name}<br>数量：<b>${ps[0].value}</b> ${unit||'款'}<br>占比（已填报口径）：<b>${total?pct(ps[0].value,total).toFixed(1):0}%</b>`},TT),
    grid:Object.assign({},GRID({left:10,right:70,top:24,bottom:8}),{containLabel:true}),
    xAxis:Object.assign({type:'value',name:`${catName?'车型数':'数量'}（${unit||'款'}）`,nameTextStyle:{color:TC().axis}},AXS({splitLine:{show:true}})),
    yAxis:Object.assign({type:'category',data:names,name:catName||undefined,nameTextStyle:{color:TC().axis}},AXS({axisLabel:{color:TC().lbl,fontSize:11.5}})),
    series:[{type:'bar',data:vals,itemStyle:{color:color,borderRadius:[0,4,4,0]},barMaxWidth:16,
      label:{show:true,position:'right',color:TC().axis,fontSize:11,fontFamily:'Consolas,monospace',
        formatter:p=>`${p.value} · ${total?pct(p.value,total).toFixed(1):0}%`}}]
  },true);
}
function chSeg(rows){
  const es = groupCount(rows,I.s,10);
  const filled = rows.filter(r=>r[I.s]!=null).length;
  setInsight('chSeg', es.length?`填报率 <em>${p1(filled,rows.length)}%</em> · 首位 <em>${es[0][0]}</em>（<em>${fmt(es[0][1])}</em> 款，占填报 <em>${p1(es[0][1],filled)}%</em>）`:'当前筛选无细分市场填报');
  hbar('chSeg', es, '#7FA97F', '款', '细分市场');
}
function chEnt(rows){
  const es = groupCount(rows,I.e,12);
  setInsight('chEnt', es.length?`首位 <em>${es[0][0]}</em>（<em>${fmt(es[0][1])}</em> 款，<em>${p1(es[0][1],rows.length)}%</em>）· 入榜 12 强合计 <em>${p1(es.reduce((a,b)=>a+b[1],0),rows.length)}%</em>`:'暂无企业数据');
  hbar('chEnt', es, '#D4A94E', '款', '企业名称');
}
function chEvo(rows){
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  markGap('chEvo', labels);
  const mk = idx => bs.map(x=>{const v=avg(x[1],idx);return v?+v.toFixed(1):null;});
  const defs = [
    ['电机总功率kW', mk(I.tp), '#D97757', {}],
    ['纯电续航km',   mk(I.r),  '#6B9BD1', {}],
    ['电池容量kWh',  mk(I.c),  '#7FA97F', {}],
    ['整备质量百kg', bs.map(x=>{const v=avg(x[1],I.w);return v?+(v/100).toFixed(1):null;}), '#D4A94E', {}],
    ['能量密度Wh/kg',mk(I.ed), '#A97FA9', {}]
  ];
  const rV = mk(I.r).filter(v=>v!=null), cV = mk(I.c).filter(v=>v!=null);
  setInsight('chEvo', (rV.length>1)?`纯电续航均值首末 <em>${rV[0]}</em>→<em>${rV[rV.length-1]}</em> km · 电池容量 <em>${cV.length?cV[0]+'→'+cV[cV.length-1]:'-'}</em> kWh`:'区间样本不足');
  const series = [];
  const legendData = defs.map(d=>d[0]).concat(defs.map(d=>d[0]+'趋势线'));
  defs.forEach(d=>{ series.push(...mkFitPair(d[0],d[1],d[2],d[3])); });
  chart('chEvo').setOption({
    tooltip:Object.assign({trigger:'axis'},TT),
    legend:LG({data:legendData,selected:(()=>{const s={};defs.forEach(d=>s[d[0]+'趋势线']=false);return s;})()}),
    grid:GRID({bottom:64,top:56}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value',name:'均值（单位见图例，混合口径）'},AXS()),
    series
  },true);
}
function insight1(rows){
  const bev = rows.filter(r=>r[I.t]==='BEV').length, phev = rows.filter(r=>r[I.t]==='PHEV').length;
  const bs = byBatch(rows);
  const last = bs[bs.length-1], first = bs[0];
  document.getElementById('insight1Tx').innerHTML =
    `当前筛选 <em>${rows.length}</em> 款车型：BEV <em>${bev}</em> 款（<em>${rows.length?pct(bev,rows.length).toFixed(1):0}%</em>）、PHEV/EREV <em>${phev}</em> 款（<em>${rows.length?pct(phev,rows.length).toFixed(1):0}%</em>）。`+
    (last?`最新批次（第 <em>${last[0]}</em> 批）共 <em>${last[1].length}</em> 款。`:'')+
    `平均纯电续航 <em>${(avg(rows,I.r)||0).toFixed(0)}</em> km，平均电机总功率 <em>${(avg(rows,I.tp)||0).toFixed(0)}</em> kW。所有占比均按<b>已填报口径</b>计算，缺失数据不计入分母。`;
}

/* ==================== Panel 2 电机系统 ==================== */
function chPwDist(rows){
  const edges=[100,150,200,250,300,400], labels=['<100','100-150','150-200','200-250','250-300','300-400','≥400'];
  const b1 = buckets(rows.filter(r=>r[I.t]==='BEV'),I.tp,edges,labels);
  const b2 = buckets(rows.filter(r=>r[I.t]==='PHEV'),I.tp,edges,labels);
  const t1=b1.reduce((a,b)=>a+b,0), t2=b2.reduce((a,b)=>a+b,0);
  let pkN=-1, pkL=labels[0];
  labels.forEach((L,i)=>{ const c=(b1[i]||0)+(b2[i]||0); if(c>pkN){pkN=c;pkL=L;} });
  setInsight('chPwDist', pkN>0?`主流功率段 <em>${pkL} kW</em>（合计 <em>${fmt(pkN)}</em> 款）· 样本 BEV ${fmt(t1)} / PHEV/EREV ${fmt(t2)}`:'暂无功率填报');
  chart('chPwDist').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>{let h=`<b>${ps[0].name} kW</b>`;ps.forEach(p=>{
        const tot=p.seriesName==='BEV'?t1:t2;
        h+=`<br><span style="color:${p.color}">■</span> ${p.seriesName}：<b>${p.value}</b> 款（${tot?pct(p.value,tot).toFixed(1):0}%）`;});return h;}},TT),
    legend:LG({data:['BEV','PHEV/EREV']}),
    grid:GRID(),
    xAxis:Object.assign({type:'category',data:labels},AXS({name:'kW',nameTextStyle:{color:TC().axis}})),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series:[
      {name:'BEV',type:'bar',data:b1,itemStyle:{color:'#D97757'},barMaxWidth:26},
      {name:'PHEV/EREV',type:'bar',data:b2,itemStyle:{color:'#6B9BD1'},barMaxWidth:26}
    ]
  },true);
}
function chMsTop(rows){
  const es = groupCount(rows,I.ms,15);
  const filled = es.reduce((a,b)=>a+b[1],0);
  const cr3 = es.slice(0,3).reduce((a,b)=>a+b[1],0);
  setInsight('chMsTop', filled?`已填报 <em>${fmt(filled)}</em> 款 · 首位 ${es[0][0]}（<em>${p1(es[0][1],filled)}%</em>）· CR3 <em>${p1(cr3,filled)}%</em>`:'暂无电机供应商填报');
  hbar('chMsTop', es, '#D97757', '款', '电机供应商');
}
function chPtq(rows){
  // v4.4.0：按驱动形式拆分散点（两驱=前电机功率/扭矩；四驱=系统总功率/总扭矩），各组独立趋势线
  const grp={
    two:[],       // 两驱（单电机）
    four:[]       // 四驱（双电机，总=前+后）
  };
  for(const r of rows){
    if(r[I.drive]==null||!r[I.tq]) continue;
    if(r[I.drive]===0 && r[I.fp]&&r[I.tq]){
      grp.two.push({value:[r[I.fp],r[I.tq]], n:r[I.gn], b:r[I.bd]});
    }else if(r[I.drive]===1 && r[I.tp]){
      grp.four.push({value:[r[I.tp],r[I.tq]], n:r[I.gn], b:r[I.bd]});
    }
  }
  const trT = trendSeries(grp.two.map(p=>p.value),'#D97757','两驱趋势线');
  const trF = trendSeries(grp.four.map(p=>p.value),'#7C5CBF','四驱趋势线');
  const trends = [trT,trF].filter(Boolean);
  TREND_REGS['chPtq'] = Object.fromEntries(trends.map(t=>[t.name,t.__formula]));
  const nAll = grp.two.length+grp.four.length;
  setInsight('chPtq', nAll?`散点样本 <em>${fmt(nAll)}</em> 组（两驱 <em>${fmt(grp.two.length)}</em> / 四驱 <em>${fmt(grp.four.length)}</em>）· 两驱取前电机功率×扭矩，四驱取系统总功率×总扭矩（总=前+后）· 虚线为分组线性趋势`:'暂无功率×扭矩匹配样本');
  const fin = [
    {name:'两驱（单电机）',type:'scatter',data:grp.two,symbolSize:6,itemStyle:{color:`rgba(217,119,87,${TC().scatterA})`},large:true,largeThreshold:800},
    {name:'四驱（双电机）',type:'scatter',data:grp.four,symbolSize:7,symbol:'diamond',itemStyle:{color:`rgba(124,92,191,${TC().scatterA})`},large:true,largeThreshold:800}
  ];
  trends.forEach(t=>fin.push(t));
  chart('chPtq').setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{
        const t=trends.find(x=>x.name===p.seriesName);
        if(t) return t.tooltip.formatter();
        const four=p.seriesName.indexOf('四驱')===0;
        return `${p.seriesName} · 品牌车型：<b>${brandName(p)}</b><br>${four?'系统总功率':'前电机功率'} <b>${p.data.value[0]}</b> kW<br>${four?'系统总扭矩':'前电机扭矩'} <b>${p.data.value[1]}</b> Nm`;
      }},TT),
    legend:LG({data:['两驱（单电机）','四驱（双电机）',...trends.map(t=>t.name)]}),
    grid:GRID(),
    xAxis:Object.assign({type:'value',name:'功率 kW',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true},AXS()),
    yAxis:Object.assign({type:'value',name:'扭矩 Nm',nameTextStyle:{color:TC().axis},scale:true},AXS()),
    series:fin
  },true);
}
function chPwTrend(rows){
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  markGap('chPwTrend', labels);
  const defs = [
    ['BEV平均总功率', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='BEV'),I.tp);return v?+v.toFixed(0):null;}), '#D97757'],
    ['PHEV/EREV平均总功率', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='PHEV'),I.tp);return v?+v.toFixed(0):null;}), '#6B9BD1'],
    ['BEV平均前电机功率', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='BEV'),I.fp);return v?+v.toFixed(0):null;}), '#D4A94E', {lineStyle:{width:1.5,type:'dashed'},symbol:'none'}]
  ];
  const b0 = defs[0][1].filter(v=>v!=null);
  setInsight('chPwTrend', b0.length>1?`BEV总功率批次均值首末 <em>${b0[0]}</em>→<em>${b0[b0.length-1]}</em> kW（${b0[b0.length-1]>=b0[0]?'上行':'下行'}）`:'区间样本不足');
  const series=[]; const legendData=[];
  defs.forEach(d=>{ series.push(...mkFitPair(d[0],d[1],d[2],d[3])); legendData.push(d[0],d[0]+'趋势线'); });
  const sel={}; defs.forEach(d=>sel[d[0]+'趋势线']=false);
  chart('chPwTrend').setOption({
    tooltip:Object.assign({trigger:'axis'},TT),
    legend:LG({data:legendData,selected:sel}),
    grid:GRID({bottom:64,top:56}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value',name:'kW',nameTextStyle:{color:TC().axis}},AXS()),
    series
  },true);
}
function chPwr(rows){
  const edges=[40,60,80,100,120,150,200], labels=['<40','40-60','60-80','80-100','100-120','120-150','150-200','≥200'];
  const cnt = labels.map(()=>0);
  for(const r of rows){
    if(r[I.tp]==null||r[I.w]==null||r[I.w]<=0) continue;
    const v = r[I.tp]/(r[I.w]/1000);
    let k = labels.length-1;
    for(let i=0;i<edges.length;i++){ if(v<edges[i]){k=i;break;} }
    cnt[k]++;
  }
  const total = cnt.reduce((a,b)=>a+b,0);
  const pkI = cnt.indexOf(Math.max(...cnt));
  setInsight('chPwr', total?`主流功率质量比 <em>${labels[pkI]} kW/t</em>（<em>${p1(cnt[pkI],total)}%</em>）· 样本 ${fmt(total)} 款`:'暂无功率×质量样本');
  chart('chPwr').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>`${ps[0].name} kW/t：<b>${ps[0].value}</b> 款（${total?pct(ps[0].value,total).toFixed(1):0}%）`},TT),
    grid:GRID(),
    xAxis:Object.assign({type:'category',data:labels,name:'kW/t',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis}},AXS()),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series:[{type:'bar',data:cnt,itemStyle:{color:'#5BAFAF'},barMaxWidth:34,
      label:{show:true,position:'top',color:TC().axis,fontSize:10,fontFamily:'Consolas',formatter:p=>p.value>0?(total?pct(p.value,total).toFixed(1):0)+'%':''}}]
  },true);
}
function insight2(rows){
  const ms = groupCount(rows,I.ms);
  const msTotal = ms.reduce((a,b)=>a+b[1],0);
  const top3 = ms.slice(0,3).reduce((a,b)=>a+b[1],0);
  let pwrSum=0,pwrN=0,pwrCnt=0;
  for(const r of rows){
    if(r[I.tp]!=null&&r[I.w]!=null&&r[I.w]>0){
      const v=r[I.tp]/(r[I.w]/1000);
      pwrSum+=v;pwrN++;
      if(v>=80&&v<=150)pwrCnt++;
    }
  }
  document.getElementById('insight2Tx').innerHTML =
    `已填报电机供应商的车型 <em>${msTotal}</em> 款（供应商 <em>${ms.length}</em> 家），Top3（${ms.slice(0,3).map(e=>e[0]).join('、')}）合计 <em>${top3}</em> 款，CR3 = <em>${msTotal?pct(top3,msTotal).toFixed(1):0}%</em>（已填报口径）。`+
    `电机总功率均值 <em>${(avg(rows,I.tp)||0).toFixed(0)}</em> kW（样本 ${(rows.filter(r=>r[I.tp]!=null)).length} 款）；`+
    (pwrN?`功率质量比均值 <em>${(pwrSum/pwrN).toFixed(0)}</em> kW/t（样本 <em>${pwrN}</em> 款，均同时具备功率与整备质量数据）。`:``);
}

/* ==================== Panel 3 发动机系统 ==================== */
function chDv(rows){
  const phev = rows.filter(r=>r[I.dv]!=null);
  const edges=[1000,1500,2000,2500], labels=['<1.0L','1.0-1.5L','1.5-2.0L','2.0-2.5L','≥2.5L'];
  const cnt = labels.map(()=>0);
  for(const r of phev){
    const v=r[I.dv];
    let k=labels.length-1;
    for(let i=0;i<edges.length;i++){ if(v<edges[i]){k=i;break;} }
    cnt[k]++;
  }
  const total=cnt.reduce((a,b)=>a+b,0);
  const pkI = cnt.indexOf(Math.max(...cnt));
  setInsight('chDv', total?`主流排量段 <em>${labels[pkI]}</em>（<em>${p1(cnt[pkI],total)}%</em>）· 样本 ${fmt(total)} 款（PHEV/EREV）`:'暂无排量填报');
  chart('chDv').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>`${ps[0].name}：<b>${ps[0].value}</b> 款（${total?pct(ps[0].value,total).toFixed(1):0}%）`},TT),
    grid:GRID(),
    xAxis:Object.assign({type:'category',data:labels,name:'发动机排量（L）',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis}},AXS()),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series:[{type:'bar',data:cnt,itemStyle:{color:'#A97FA9'},barMaxWidth:44,
      label:{show:true,position:'top',color:TC().lbl,fontSize:11,fontFamily:'Consolas',
        formatter:p=>`${p.value}${total&&p.value>0?'\n'+pct(p.value,total).toFixed(1)+'%':''}`}}]
  },true);
}
function chDvEp(rows){
  const pts = rows.filter(r=>r[I.dv]!=null&&r[I.ep]!=null).map(r=>({value:[+(r[I.dv]/1000).toFixed(2),r[I.ep]], n:r[I.gn], b:r[I.bd]}));
  const tr = trendSeries(pts.map(p=>p.value),'#D4A94E','PHEV/EREV趋势线');
  TREND_REGS['chDvEp'] = tr?{[tr.name]:tr.__formula}:{};
  setInsight('chDvEp', pts.length?`样本 <em>${fmt(pts.length)}</em> 组 · 平均排量 <em>${(pts.reduce((a,p)=>a+p.value[0],0)/pts.length).toFixed(2)} L</em> · 平均功率 <em>${(pts.reduce((a,p)=>a+p.value[1],0)/pts.length).toFixed(0)} kW</em>`:'暂无排量×功率样本');
  const series=[{type:'scatter',name:'PHEV/EREV',data:pts,symbolSize:7,itemStyle:{color:`rgba(169,127,169,${TC().scatterA})`}}];
  if(tr) series.push(tr);
  chart('chDvEp').setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>p.seriesName==='PHEV/EREV趋势线'?tr.tooltip.formatter():`PHEV/EREV · 品牌车型：<b>${brandName(p)}</b><br>排量 <b>${p.data.value[0]}</b> L · 功率 <b>${p.data.value[1]}</b> kW`},TT),
    legend:LG({data:tr?['PHEV/EREV','PHEV/EREV趋势线']:['PHEV/EREV']}),
    grid:GRID(),
    xAxis:Object.assign({type:'value',name:'排量 L',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true},AXS()),
    yAxis:Object.assign({type:'value',name:'功率 kW',nameTextStyle:{color:TC().axis},scale:true},AXS()),
    series
  },true);
}
function chEsTop(rows){
  const es = groupCount(rows,I.es,12);
  const filled = es.reduce((a,b)=>a+b[1],0);
  const cr3 = es.slice(0,3).reduce((a,b)=>a+b[1],0);
  setInsight('chEsTop', filled?`已填报 <em>${fmt(filled)}</em> 款 · 首位 ${es[0][0]}（<em>${p1(es[0][1],filled)}%</em>）· CR3 <em>${p1(cr3,filled)}%</em>`:'暂无发动机供应商填报');
  hbar('chEsTop', es, '#C97B7B', '款', '发动机供应商');
}
function chFo(rows){
  const src = rows.filter(r=>r[I.fo]!=null);
  const edges=[1,2,3,4,5,6], labels=['<1','1-2','2-3','3-4','4-5','5-6','≥6'];
  const cnt = labels.map(()=>0);
  for(const r of src){
    const v=r[I.fo]; let k=labels.length-1;
    for(let i=0;i<edges.length;i++){ if(v<edges[i]){k=i;break;} }
    cnt[k]++;
  }
  const total=cnt.reduce((a,b)=>a+b,0);
  const meanFo = src.length? (src.reduce((a,r)=>a+r[I.fo],0)/src.length).toFixed(2):'-';
  setInsight('chFo', total?`亏电油耗样本 <em>${fmt(total)}</em> 款 · 均值 <em>${meanFo} L/100km</em>`:'暂无油耗填报');
  chart('chFo').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>`${ps[0].name} L/100km：<b>${ps[0].value}</b> 款（${total?pct(ps[0].value,total).toFixed(1):0}%）`},TT),
    grid:GRID(),
    xAxis:Object.assign({type:'category',data:labels,name:'L/100km',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis}},AXS()),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series:[{type:'bar',data:cnt,itemStyle:{color:'#B08968'},barMaxWidth:34,
      label:{show:true,position:'top',color:TC().axis,fontSize:10,fontFamily:'Consolas',formatter:p=>p.value>0?(total?pct(p.value,total).toFixed(1):0)+'%':''}}]
  },true);
}
function chDvTrend(rows){
  const bs = byBatch(rows.filter(r=>r[I.dv]!=null||r[I.ep]!=null));
  const labels = bs.map(x=>String(x[0]));
  markGap('chDvTrend', labels);
  const defs = [
    ['平均排量 L(×10)', bs.map(x=>{const v=avg(x[1],I.dv);return v?+(v/100).toFixed(2):null;}), '#A97FA9'],
    ['平均功率 kW', bs.map(x=>{const v=avg(x[1],I.ep);return v?+v.toFixed(0):null;}), '#C97B7B']
  ];
  const dv0 = defs[0][1].filter(v=>v!=null), ep0 = defs[1][1].filter(v=>v!=null);
  setInsight('chDvTrend', dv0.length>1?`排量均值首末 <em>${dv0[0]}</em>→<em>${dv0[dv0.length-1]}</em> L · 功率均值 <em>${ep0.length?ep0[0]+'→'+ep0[ep0.length-1]:'-'}</em> kW`:'区间样本不足');
  const series=[]; const legendData=[]; const sel={};
  defs.forEach(d=>{ series.push(...mkFitPair(d[0],d[1],d[2])); legendData.push(d[0],d[0]+'趋势线'); sel[d[0]+'趋势线']=false; });
  chart('chDvTrend').setOption({
    tooltip:Object.assign({trigger:'axis'},TT),
    legend:LG({data:legendData,selected:sel}),
    grid:GRID({bottom:64,top:56}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value'},AXS()),
    series
  },true);
}
function insight3(rows){
  const pe = rows.filter(r=>r[I.dv]!=null);
  const es = groupCount(rows,I.es);
  const esTotal = es.reduce((a,b)=>a+b[1],0);
  const top3 = es.slice(0,3).reduce((a,b)=>a+b[1],0);
  document.getElementById('insight3Tx').innerHTML =
    `已填报排量的 PHEV/EREV 车型 <em>${pe.length}</em> 款，平均排量 <em>${((avg(pe,I.dv)||0)/1000).toFixed(2)}</em> L，平均功率 <em>${(avg(pe,I.ep)||0).toFixed(0)}</em> kW（缺失记录已剔除，不计入均值）。`+
    (esTotal?`已填报发动机供应商车型 <em>${esTotal}</em> 款（供应商 <em>${es.length}</em> 家），CR3 = <em>${pct(top3,esTotal).toFixed(1)}%</em>。`:'')+
    `平均综合油耗 <em>${(avg(pe,I.fo)||0).toFixed(2)}</em> L/100km。`;
}

/* ==================== Panel 4 电池系统 ==================== */
function hist(id, rows, idx, edges, labels, color, name){
  const cnt = buckets(rows,idx,edges,labels);
  const total = cnt.reduce((a,b)=>a+b,0);
  chart(id).setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>`${ps[0].name} ${name}：<b>${ps[0].value}</b> 款（${total?pct(ps[0].value,total).toFixed(1):0}%）`},TT),
    grid:GRID(),
    xAxis:Object.assign({type:'category',data:labels},AXS({name:name,nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis}})),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series:[{type:'bar',data:cnt,itemStyle:{color:color},barMaxWidth:30,
      label:{show:true,position:'top',color:TC().axis,fontSize:10,fontFamily:'Consolas',formatter:p=>p.value>0?(total?pct(p.value,total).toFixed(1):0)+'%':''}}]
  },true);
}
function chRgB(rows){
  const sub = rows.filter(r=>r[I.t]==='BEV');
  const RG_L=['<300','300-400','400-500','500-600','600-700','700-800','≥800'];
  const cnt = buckets(sub, I.r, [300,400,500,600,700,800], RG_L);
  const tot = cnt.reduce((a,b)=>a+b,0);
  const pk = cnt.indexOf(Math.max(...cnt));
  setInsight('chRgB', tot?`BEV主流续航段 <em>${RG_L[pk]} km</em>（<em>${p1(cnt[pk],tot)}%</em>）· 样本 ${fmt(tot)} 款`:'暂无BEV续航填报');
  hist('chRgB', sub, I.r, [300,400,500,600,700,800], RG_L, '#D97757','纯电续航里程(km)');
}
function chRgP(rows){
  const sub = rows.filter(r=>r[I.t]==='PHEV');
  const RG_L=['<75','75-100','100-125','125-150','150-200','≥200'];
  const cnt = buckets(sub, I.r, [75,100,125,150,200], RG_L);
  const tot = cnt.reduce((a,b)=>a+b,0);
  const pk = cnt.indexOf(Math.max(...cnt));
  setInsight('chRgP', tot?`PHEV/EREV主流纯电续航段 <em>${RG_L[pk]} km</em>（<em>${p1(cnt[pk],tot)}%</em>）· 样本 ${fmt(tot)} 款`:'暂无PHEV/EREV续航填报');
  hist('chRgP', sub, I.r, [75,100,125,150,200], RG_L, '#6B9BD1','纯电续航里程(km)');
}
function chCr(rows){
  const bevP=[],phevP=[];
  for(const r of rows){
    if(r[I.c]!=null&&r[I.r]!=null){ (r[I.t]==='BEV'?bevP:phevP).push({value:[r[I.c],r[I.r]], n:r[I.gn], b:r[I.bd]}); }
  }
  // BEV 与 PHEV/EREV 分别提供趋势线（散点对象取 value 数组）
  const allPts = bevP.concat(phevP).map(p=>p.value);
  const allN = allPts.length;
  const meanC = allN? (allPts.reduce((a,p)=>a+p[0],0)/allN).toFixed(1):'-';
  const meanR = allN? (allPts.reduce((a,p)=>a+p[1],0)/allN).toFixed(0):'-';
  setInsight('chCr', allN?`散点样本 <em>${fmt(allN)}</em> 组 · 容量均值 <em>${meanC} kWh</em> · 续航均值 <em>${meanR} km</em> · 容量越大续航越长（分组趋势线）`:'暂无容量×续航样本');
  const trB = trendSeries(bevP.map(p=>p.value),'#D97757','BEV趋势线');
  const trP = trendSeries(phevP.map(p=>p.value),'#6B9BD1','PHEV/EREV趋势线');
  const trends = [trB,trP].filter(Boolean);
  TREND_REGS['chCr'] = Object.fromEntries(trends.map(t=>[t.name,t.__formula]));
  const series = [
    {name:'BEV',type:'scatter',data:bevP,symbolSize:6,itemStyle:{color:`rgba(217,119,87,${TC().scatterA})`},large:true,largeThreshold:800},
    {name:'PHEV/EREV',type:'scatter',data:phevP,symbolSize:6,itemStyle:{color:`rgba(107,155,209,${TC().scatterA})`},large:true,largeThreshold:800}
  ];
  const lg=['BEV','PHEV/EREV'];
  trends.forEach(t=>{series.push(t);lg.push(t.name);});
  chart('chCr').setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{
        const t=trends.find(x=>x.name===p.seriesName);
        if(t) return t.tooltip.formatter();
        return `${p.seriesName} · 品牌车型：<b>${brandName(p)}</b><br>容量 <b>${p.data.value[0]}</b> kWh · 续航 <b>${p.data.value[1]}</b> km`;
      }},TT),
    legend:LG({data:lg}),
    grid:GRID(),
    xAxis:Object.assign({type:'value',name:'容量 kWh',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true},AXS()),
    yAxis:Object.assign({type:'value',name:'续航 km',nameTextStyle:{color:TC().axis},scale:true},AXS()),
    series
  },true);
}
function chBtTrend(rows){
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  // 核算依据：已有数据（归一化后）中的百分比占比，按实际出现的类型动态取列
  const types = groupCount(rows,I.bt,6).map(e=>e[0]);
  const series = types.map((t,i)=>({
    name:t,type:'line',stack:'x',areaStyle:{opacity:.35},symbol:'none',smooth:false,
    itemStyle:{color:PAL[i]},lineStyle:{width:1},
    data:bs.map(x=>{
      const n = x[1].filter(r=>r[I.bt]===t).length;
      const tot = x[1].filter(r=>r[I.bt]!=null).length;   // 已填报口径
      return tot?+(n/tot*100).toFixed(1):null;
    })
  }));
  const lastB = bs[bs.length-1];
  let ltPct = null;
  if(lastB && types.length){
    const fN = lastB[1].filter(r=>r[I.bt]!=null).length;
    ltPct = fN? +(lastB[1].filter(r=>r[I.bt]===types[0]).length/fN*100).toFixed(1):null;
  }
  setInsight('chBtTrend', types.length?`最新批次占比首位 <em>${types[0]}</em>（<em>${ltPct==null?'-':ltPct+'%'}</em>，已填报口径）`:'暂无电池类型填报');
  chart('chBtTrend').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'cross'},
      formatter:ps=>{let h=`<b>第 ${ps[0].axisValue} 批</b>`;ps.forEach(p=>{if(p.value!=null)h+=`<br><span style="color:${p.color}">■</span> ${p.seriesName}：<b>${p.value}%</b>`});return h;}},TT),
    legend:LG({data:types}),
    grid:GRID({bottom:64,top:56}),
    xAxis:Object.assign({type:'category',boundaryGap:false,data:labels},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value',max:100,axisLabel:{formatter:'{value}%'}},AXS()),
    series
  },true);
}
function chEd(rows){
  const bs = byBatch(rows.filter(r=>r[I.t]==='BEV'));
  const labels = bs.map(x=>String(x[0]));
  markGap('chEd', labels);
  const vals = bs.map(x=>{const v=avg(x[1],I.ed);return v?+v.toFixed(0):null;});
  const edV = vals.filter(v=>v!=null);
  setInsight('chEd', edV.length>1?`BEV能量密度批次均值首末 <em>${edV[0]}</em>→<em>${edV[edV.length-1]}</em> Wh/kg（${edV[edV.length-1]>=edV[0]?'上行':'下行'}）`:'区间样本不足');
  const [main,fit] = mkFitPair('能量密度',vals,'#7FA97F',{areaStyle:{color:{type:'linear',x:0,y:0,x2:0,y2:1,
    colorStops:[{offset:0,color:'rgba(127,169,127,.25)'},{offset:1,color:'rgba(127,169,127,0)'}]}}});
  chart('chEd').setOption({
    tooltip:Object.assign({trigger:'axis',valueFormatter:v=>v+' Wh/kg'},TT),
    legend:LG({data:['能量密度','能量密度趋势线'],selected:{'能量密度趋势线':false}}),
    grid:GRID({bottom:64,top:56}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value',name:'Wh/kg',nameTextStyle:{color:TC().axis},scale:true},AXS()),
    series:[main,fit]
  },true);
}
function chEcTrend(rows){
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  markGap('chEcTrend', labels);
  const defs = [
    ['BEV平均电耗', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='BEV'),I.ec);return v?+v.toFixed(1):null;}), '#D97757'],
    ['PHEV/EREV平均电耗', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='PHEV'),I.ec);return v?+v.toFixed(1):null;}), '#6B9BD1']
  ];
  const ec0 = defs[0][1].filter(v=>v!=null);
  setInsight('chEcTrend', ec0.length>1?`BEV电耗批次均值首末 <em>${ec0[0]}</em>→<em>${ec0[ec0.length-1]}</em> kWh/100km（${ec0[ec0.length-1]<=ec0[0]?'持续优化':'回升'}）`:'区间样本不足');
  const series=[]; const legendData=[]; const sel={};
  defs.forEach(d=>{ series.push(...mkFitPair(d[0],d[1],d[2])); legendData.push(d[0],d[0]+'趋势线'); sel[d[0]+'趋势线']=false; });
  chart('chEcTrend').setOption({
    tooltip:Object.assign({trigger:'axis',valueFormatter:v=>v+' kWh/100km'},TT),
    legend:LG({data:legendData,selected:sel}),
    grid:GRID({bottom:64,top:56}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value',name:'kWh/100km',nameTextStyle:{color:TC().axis},scale:true},AXS()),
    series
  },true);
}
function chEc(rows){
  const sub = rows.filter(r=>r[I.ec]!=null);
  const EC_L=['<10','10-12','12-14','14-16','16-18','≥18'];
  const cnt = buckets(sub, I.ec, [10,12,14,16,18], EC_L);
  const tot = cnt.reduce((a,b)=>a+b,0);
  const pk = cnt.indexOf(Math.max(...cnt));
  setInsight('chEc', tot?`主流电耗段 <em>${EC_L[pk]} kWh/100km</em>（<em>${p1(cnt[pk],tot)}%</em>）· 样本 ${fmt(tot)} 款`:'暂无电耗填报');
  hist('chEc', sub, I.ec, [10,12,14,16,18], EC_L, '#6A8FB5','百公里电耗(kWh/100km)');
}
function insight4(rows){
  const b = rows.filter(r=>r[I.t]==='BEV'), p = rows.filter(r=>r[I.t]==='PHEV');
  const btFilled = rows.filter(r=>r[I.bt]!=null);
  document.getElementById('insight4Tx').innerHTML =
    `BEV 平均续航 <em>${(avg(b,I.r)||0).toFixed(0)}</em> km / 平均容量 <em>${(avg(b,I.c)||0).toFixed(1)}</em> kWh（样本 ${b.filter(r=>r[I.r]!=null).length} / ${b.filter(r=>r[I.c]!=null).length} 款）；`+
    `PHEV/EREV 纯电平均续航 <em>${(avg(p,I.r)||0).toFixed(0)}</em> km / 平均容量 <em>${(avg(p,I.c)||0).toFixed(1)}</em> kWh。`+
    `已填报电池类型 <em>${btFilled.length}</em> 款中：磷酸铁锂 <em>${btFilled.length?pct(btFilled.filter(r=>r[I.bt]==='磷酸铁锂').length,btFilled.length).toFixed(1):0}%</em>，`+
    `三元锂 <em>${btFilled.length?pct(btFilled.filter(r=>r[I.bt]==='三元锂').length,btFilled.length).toFixed(1):0}%</em>。`;
}

/* ==================== Panel D 分布格局（配置驱动，增减维度改 DIST_DEFS） ==================== */
const spl = n => { const m=String(n).match(/^(.*?)\((.*)\)$/); return m?[m[1],m[2]]:[String(n),'']; };
const DIST_DEFS = [
  {id:'dRvsLen',  t:'纯电续航 × 尺寸',                 sub:'EV Range vs Length · 散点 · BEV/PHEV 分列趋势线', yK:I.r, xK:I.ab, yN:'纯电续航里程(km)', xN:'轴距(mm)'},
  {id:'dCvsRg',   t:'电池容量 × 纯电续航',             sub:'Energy vs EV Range · 容量-续航匹配关系散点',      yK:I.c, xK:I.r, yN:'电池容量(kWh)',    xN:'纯电续航里程(km)'},
  {id:'dEcVsRg',  t:'百公里电耗 × 纯电续航',           sub:'Energy Consumption vs EV Range · 能效水平散点',   yK:I.ec,xK:I.r, yN:'百公里电耗(kWh/100km)', xN:'纯电续航里程(km)'},
  {id:'dEdVsRg',  t:'能量密度 × 纯电续航',             sub:'Energy Density vs EV Range · 技术水平散点',       yK:I.ed,xK:I.r, yN:'电池能量密度(Wh/kg)', xN:'纯电续航里程(km)'},
  {id:'dMassLen', t:'整备质量 × 尺寸',                 sub:'Curb Mass vs Length · 重量-尺寸分布散点',         yK:I.w, xK:I.ab, yN:'整备质量(kg)',     xN:'轴距(mm)'}
];
function initDist(){
  const g=document.getElementById('distGrid'); if(!g)return;
  g.innerHTML = DIST_DEFS.map(d=>
    `<div class="card c-s6"><div class="c-h"><div class="c-t">${d.t}</div><div class="c-s">${d.sub}</div></div>`+
    `<div class="chart" id="${d.id}"></div></div>`).join('');
}
function renderDists(rows){
  for(const d of DIST_DEFS){
    if(!document.getElementById(d.id)) continue;
    const bevP=[], phevP=[];
    for(const r of rows){
      const x=r[d.xK], y=r[d.yK];
      if(x!=null&&y!=null){ (r[I.t]==='BEV'?bevP:phevP).push({value:[x,y], n:r[I.gn], b:r[I.bd]}); }
    }
    const trB=trendSeries(bevP.map(p=>p.value),'#D97757','BEV趋势线');
    const trP=trendSeries(phevP.map(p=>p.value),'#6B9BD1','PHEV/EREV趋势线');
    const trends=[trB,trP].filter(Boolean);
    TREND_REGS[d.id]=Object.fromEntries(trends.map(t=>[t.name,t.__formula]));
    const series=[
      {name:'BEV',type:'scatter',data:bevP,symbolSize:6,itemStyle:{color:`rgba(217,119,87,${TC().scatterA})`},large:true,largeThreshold:800},
      {name:'PHEV/EREV',type:'scatter',data:phevP,symbolSize:6,itemStyle:{color:`rgba(107,155,209,${TC().scatterA})`},large:true,largeThreshold:800}
    ];
    trends.forEach(t=>series.push(t));
    /* 要点：分燃料类型的均值与极值（随全局筛选联动） */
    const fv = v => v>=100 ? Math.round(v).toLocaleString('zh-CN') : (+v.toFixed(1));
    const stat = ps=>{
      if(!ps.length) return null;
      let xs=0,ys=0,xmin=Infinity,xmax=-Infinity,ymin=Infinity,ymax=-Infinity;
      for(const p of ps){const x=p.value[0],y=p.value[1];xs+=x;ys+=y;
        if(x<xmin)xmin=x;if(x>xmax)xmax=x;if(y<ymin)ymin=y;if(y>ymax)ymax=y;}
      return {xm:xs/ps.length,xmin,xmax,ym:ys/ps.length,ymin,ymax};
    };
    const [xn,xu]=spl(d.xN), [yn,yu]=spl(d.yN);
    const part=(nm,s)=>s?`${nm}：${xn}均值 <em>${fv(s.xm)}</em>（极值 ${fv(s.xmin)}~${fv(s.xmax)}）${xu} · ${yn}均值 <em>${fv(s.ym)}</em>（极值 ${fv(s.ymin)}~${fv(s.ymax)}）${yu}`:'';
    const sB=stat(bevP), sP=stat(phevP);
    const segs=[['BEV',sB],['PHEV/EREV',sP]].filter(x=>x[1]).map(x=>part(x[0],x[1])).join(' ｜ ');
    setInsight(d.id, segs?`${segs}（样本 <em>${fmt(bevP.length+phevP.length)}</em> 组）`:'当前筛选无有效样本');
    chart(d.id).setOption({
      tooltip:Object.assign({trigger:'item',
        formatter:p=>{
          const t=trends.find(x=>x.name===p.seriesName);
          if(t) return t.tooltip.formatter();
          const [xn,xu]=spl(d.xN), [yn,yu]=spl(d.yN);
          return `${p.seriesName} · 品牌车型：<b>${brandName(p)}</b><br>${xn}：<b>${p.data.value[0]} ${xu}</b><br>${yn}：<b>${p.data.value[1]} ${yu}</b>`;
        }},TT),
      legend:LG({data:['BEV','PHEV/EREV',...trends.map(t=>t.name)]}),
      grid:GRID(),
      xAxis:Object.assign({type:'value',name:d.xN,nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true},AXS()),
      yAxis:Object.assign({type:'value',name:d.yN,nameTextStyle:{color:TC().axis},scale:true},AXS()),
      series
    },true);
    bindTrendClick(d.id);
  }
}
function insightD(rows){
  const el=document.getElementById('insightDTx'); if(!el)return;
  const nAb=rows.filter(r=>r[I.ab]!=null&&r[I.ab]!=null&&r[I.r]!=null).length;
  const nCR=rows.filter(r=>r[I.c]!=null&&r[I.r]!=null).length;
  const nEc=rows.filter(r=>r[I.ec]!=null&&r[I.r]!=null).length;
  const nEd=rows.filter(r=>r[I.ed]!=null&&r[I.r]!=null).length;
  document.getElementById('insightDTx').innerHTML =
    `分布格局从五个角度呈现车型参数的<u>两两关系</u>：续航-尺寸、容量-续航、电耗-续航、密度-续航、质量-尺寸。`+
    `当前筛选内有效样本：续航×轴距 <em>${nAb}</em>、容量×续航 <em>${nCR}</em>、电耗×续航 <em>${nEc}</em>、密度×续航 <em>${nEd}</em> 款（缺失不计入）。`+
    `右侧【自定义分布】可自由选择 X / Y / Z 维度与视图模式（二维泡泡 / 三维柱状 / 二维热力），风格 8 款按需切换。`;
}

/* ==================== 自定义分布（X/Y · 2D 泡泡 / 3D 柱状 / 2D 热力 · 风格化） ==================== */
/* 字段目录：分布格局已用字段优先置顶，其余按语义顺序排列。
   仅纳入「可量化」字段（数值型），标签类字段不作为 Z 轴备选。 */
const CD_FIELDS = (()=>{
  const prio = [];
  const seen = new Set();
  // 1) 分布格局已用字段优先
  for(const d of DIST_DEFS){
    for(const [k,n] of [[d.xK,d.xN],[d.yK,d.yN]]){
      if(!seen.has(k)){ seen.add(k); prio.push({k,n,num:true}); }
    }
  }
  // 2) 其余可量化字段，按语义顺序
  const rest = [
    [I.w,  '整备质量(kg)',            true],
    [I.r,  '纯电续航里程(km)',         true],
    [I.c,  '电池容量(kWh)',           true],
    [I.ed, '电池能量密度(Wh/kg)',      true],
    [I.ec, '百公里电耗(kWh/100km)',    true],
    [I.fp, '前电机功率(kW)',           true],
    [I.tp, '电机总功率(kW)',           true],
    [I.tq, '系统扭矩(Nm)',            true],
    [I.fo, '综合油耗(L/100km)',        true],
    [I.dv, '发动机排量(mL)',           true],
    [I.ep, '发动机功率(kW)',           true],
    [I.ab, '轴距(mm)',                true],
    [I.lg, '车长(mm)',                true],
    [I.b,  '批次',                    true]
  ];
  for(const [k,n,num] of rest){
    if(!seen.has(k)){ seen.add(k); prio.push({k,n,num}); }
  }
  return prio;
})();
const CD_LABEL = k => (CD_FIELDS.find(f=>f.k===k)||{n:'—'}).n;
const CD_DEFAULT = {x:I.ab, y:I.r, z:null};
/* 视觉风格库（出处：ECharts GL 官方 bar3D 示例 bevel 倒角 / Plotly Viridis·Magma·Turbo 色标 / GitHub Contributions 绿阶）
   bg=场景底色 ramps=BEV/PHEV 双色系浅→深梯度 bevel=[size,smoothness]=柱体倒角 split/axisLine/label=网格与文字 */
const CD_STYLES = {
  A: { name:'学术极简', bg:'', barSize:[1.5,1.5], shading:'color',
    bevRamp:['#F6D8CB','#EFB39A','#E29067','#D97757','#A94F30'],
    phevRamp:['#D4E2F4','#AECBEC','#8AB2DF','#6B9BD1','#45699C'],
    bevBase:'rgba(217,119,87,.92)', phevBase:'rgba(107,155,209,.92)',
    border:null, bw:0, opacity:.95, ambient:.55, main:1.4,
    split:TC().split, axisLine:TC().axisLine, label:TC().axis, bevel:null },
  B: { name:'深空大屏', bg:'radial-gradient(1100px 620px at 62% 38%, #122A45 0%, #060B15 72%)', barSize:[1.8,1.8], shading:'lambert',
    bevRamp:['#FFE4CC','#FFB37E','#FF8A3D','#FF6A2B','#C2410C'],
    phevRamp:['#D6ECFF','#8FC2FF','#4D9AFF','#2E7CFF','#1D4ED8'],
    bevBase:'rgba(255,138,61,.95)', phevBase:'rgba(77,154,255,.95)',
    border:f=>f==='bev'?'rgba(255,179,126,.8)':'rgba(126,184,255,.8)', bw:1, opacity:.96, ambient:.75, main:1.5,
    split:'rgba(86,156,255,.22)', axisLine:'rgba(120,170,235,.6)', label:'#A8C6F0', bevel:null },
  C: { name:'柔和白模', bg:'#F4EFE7', barSize:[1.6,1.6], shading:'lambert',
    bevRamp:['#F7DBC6','#EFAF8C','#E18F63','#C96A4A','#9E4E32'],
    phevRamp:['#DEE9F6','#B3CBE8','#87ACD6','#5C85B8','#3E6390'],
    bevBase:'rgba(201,106,74,.95)', phevBase:'rgba(92,133,184,.95)',
    border:null, bw:0, opacity:.97, ambient:.95, main:.8,
    split:'#DFD7C9', axisLine:'#C9BFB0', label:'#5A5248', bevel:null },
  D: { name:'玻璃塔楼', bg:'radial-gradient(1100px 620px at 55% 40%, #0E1526 0%, #05080F 75%)', barSize:[2,2], shading:'color',
    bevRamp:['rgba(240,150,100,.3)','rgba(235,130,80,.46)','rgba(225,110,60,.62)','rgba(217,90,40,.8)','rgba(170,60,25,.92)'],
    phevRamp:['rgba(120,170,230,.3)','rgba(100,150,215,.46)','rgba(80,130,200,.62)','rgba(60,110,185,.8)','rgba(35,80,150,.92)'],
    bevBase:'rgba(240,150,100,.5)', phevBase:'rgba(110,155,215,.5)',
    border:f=>f==='bev'?'rgba(255,190,140,.8)':'rgba(150,195,245,.8)', bw:1.1, opacity:1, ambient:.6, main:1.2,
    split:'rgba(140,165,195,.16)', axisLine:'rgba(150,175,205,.5)', label:'#93A4BD', bevel:null },
  E: { name:'Viridis/Magma 科研渐变', bg:'#101216', barSize:[1.4,1.4], shading:'color',
    bevRamp:['#1D1147','#51127C','#B63679','#E65164','#FB8861'],
    phevRamp:['#414487','#2A788E','#22A884','#7AD151'],
    bevBase:'rgba(182,54,121,.92)', phevBase:'rgba(42,120,142,.92)',
    border:null, bw:0, opacity:.95, ambient:.8, main:.7,
    split:'#23262E', axisLine:'#3A3F4B', label:'#B9C0CE', bevel:[.2,2] },
  F: { name:'Turbo 地形热力', bg:'#0B1220', barSize:[1.9,1.9], shading:'lambert',
    bevRamp:['#F5C93A','#F5872E','#E14B1F','#A31507'],
    phevRamp:['#30123B','#2050C8','#1FA7D6','#2AE0B8'],
    bevBase:'rgba(241,110,45,.94)', phevBase:'rgba(32,110,190,.94)',
    border:null, bw:0, opacity:.96, ambient:.55, main:1.4,
    split:'rgba(90,140,200,.18)', axisLine:'rgba(140,180,230,.5)', label:'#C7D3E4', bevel:[.3,3] },
  G: { name:'单色深空', bg:'radial-gradient(1100px 620px at 55% 38%, #0A1526 0%, #04070D 75%)', barSize:[1.5,1.5], shading:'color',
    bevRamp:['#0E3A52','#155E80','#1F8FB4','#35C4E8','#8FE8FF'],
    phevRamp:['#1B2A5E','#28407E','#3A5CA8','#5B82D4','#8FB2F0'],
    bevBase:'rgba(31,143,180,.92)', phevBase:'rgba(58,92,168,.92)',
    border:null, bw:0, opacity:.92, ambient:.65, main:1.1,
    split:'rgba(70,110,170,.14)', axisLine:'rgba(120,160,210,.4)', label:'#8FA6C4', bevel:[.2,2] },
  H: { name:'GitHub 天际线', bg:'#0D1117', barSize:[2.6,2.6], shading:'lambert',
    bevRamp:['#0E4429','#006D32','#26A641','#39D353'],
    phevRamp:['#134E4A','#0F766E','#14B8A6','#5EEAD4'],
    bevBase:'rgba(38,166,65,.94)', phevBase:'rgba(20,118,110,.94)',
    border:null, bw:0, opacity:.96, ambient:.7, main:.9,
    split:'#21262D', axisLine:'#30363D', label:'#8B949E', bevel:[.35,4] }
};
let CD_STATE = {x:CD_DEFAULT.x, y:CD_DEFAULT.y, z:CD_DEFAULT.z, zScale:100, levels:true, autoSize:true,
  view:'d2', style:'A', invX:false, invY:false, invZ:false, insightFull:false, agg:'mean',
  colorMode:'dual', sizeField:null};

/* 轴距等级区段定义（与《车型级别定义》阈值一致；用于图上辅助示意） */
const CD_LEVELS = {
  'Car': {color:'#D97757', name:'轿车',  bands:[['Car-A',null,2650],['Car-B',2650,2740],['Car-C',2740,2850],['Car-D',2850,3000],['Car-E',3000,null]]},
  'SUV': {color:'#6B9BD1', name:'SUV',   bands:[['SUV-B',null,2680],['SUV-C',2680,2850],['SUV-D',2850,3000],['SUV-E',3000,null]]},
  'MPV': {color:'#7FA97F', name:'MPV',   bands:[['MPV-C',null,3000],['MPV-D',3000,null]]}
};
/* 生成等级区段 markLine（横轴或纵轴为轴距时生效） */
function cdLevelMarks(xN, yN){
  const onX = xN.indexOf('轴距') === 0, onY = yN.indexOf('轴距') === 0;
  if(!onX && !onY) return null;
  const data=[];
  for(const key of ['Car','SUV','MPV']){
    const L=CD_LEVELS[key];
    for(const [nm,lo] of L.bands.map(b=>[b[0],b[1]])){
      if(lo==null) continue;
      data.push(onX
        ? {xAxis:lo, label:{formatter:nm, position:'insideEndTop', color:L.color, fontSize:10},
           lineStyle:{color:L.color,type:'dashed',width:1,opacity:.6}}
        : {yAxis:lo, label:{formatter:nm, position:'insideEndTop', color:L.color, fontSize:10},
           lineStyle:{color:L.color,type:'dashed',width:1,opacity:.6}});
    }
  }
  return data.length ? {silent:true, symbol:'none', animation:false, data} : null;
}

function cdFillSelect(el, val, allowNone){
  el.innerHTML = (allowNone?`<option value="">（不使用 · 散点图）</option>`:'') +
    CD_FIELDS.map(f=>`<option value="${f.k}">${f.n}</option>`).join('');
  el.value = val==null?'':String(val);
}
/* 说明折叠：核心一行 + 展开开关；detail 仅展开态渲染 */
function cdToggleInsight(){ CD_STATE.insightFull=!CD_STATE.insightFull; renderCustom(filtered()); }
function cdInsightWrap(core, detail){
  const tg=`<a href="javascript:void(0)" onclick="cdToggleInsight()" style="color:#D97757;opacity:.85;font-size:11px;margin-left:8px;cursor:pointer;white-space:nowrap">${CD_STATE.insightFull?'▴ 收起说明':'▾ 展开说明'}</a>`;
  return core + tg + (CD_STATE.insightFull ? `<br>${detail}` : '');
}
function cdInitUI(){
  const sx=document.getElementById('cdX'), sy=document.getElementById('cdY'),
        sz=document.getElementById('cdZ'), ss=document.getElementById('cdZScale'),
        cv=document.getElementById('cdView'), cs=document.getElementById('cdStyle');
  if(!sx) return;
  cdFillSelect(sx, CD_STATE.x, false);
  cdFillSelect(sy, CD_STATE.y, false);
  cdFillSelect(sz, CD_STATE.z==null?'':CD_STATE.z, true);
  cs.innerHTML = Object.keys(CD_STYLES).map(k=>`<option value="${k}">${CD_STYLES[k].name}</option>`).join('');
  cs.value = CD_STATE.style;
  sx.onchange=()=>{ CD_STATE.x=+sx.value; renderCustom(filtered()); };
  sy.onchange=()=>{ CD_STATE.y=+sy.value; renderCustom(filtered()); };
  sz.onchange=()=>{ CD_STATE.z = sz.value===''?null:+sz.value;
    if(CD_STATE.z==null && CD_STATE.view!=='d2') CD_STATE.view='d2';   // 清空 Z → 回二维
    cdModeUI(); renderCustom(filtered()); };
  cv.onchange=()=>{ let v=cv.value;
    if(v!=='d2' && CD_STATE.z==null){
      /* 未选 Z 直接切三维/热力：自动补默认 Z=电池容量，保持可用 */
      CD_STATE.z=I.c; sz.value=String(I.c);
      showToast('已自动选定 Z=电池容量(kWh)，可在「高度坐标 Z」更换字段');
    }
    CD_STATE.view=v; cdModeUI(); renderCustom(filtered()); };
  cs.onchange=()=>{ CD_STATE.style=cs.value;
    /* 风格切换必须 clear+全量重建：各风格梯度色带长度不同（4/5 色），
       merge 会残留旧风格色阶条目→插值出彩虹乱色（实证：切到学术极简显示红黄紫乱色，多次切换后自愈） */
    const ch=chart('chCustom');
    if(ch){ const cam=cdCaptureCam(); if(cam) CD_CAM=cam; ch.clear(); ch.__cd3dFirst=true; }
    renderCustom(filtered()); };
  ss.oninput =()=>{ CD_STATE.zScale=+ss.value;
    document.getElementById('cdZVal').textContent = ss.value+'%'; renderCustom(filtered()); };
  document.getElementById('cdSwap').onclick=()=>{
    const t=CD_STATE.x; CD_STATE.x=CD_STATE.y; CD_STATE.y=t;
    sx.value=String(CD_STATE.x); sy.value=String(CD_STATE.y); renderCustom(filtered());
  };
  const lvBtn=document.getElementById('cdLevels');
  const syncLv=()=>{ lvBtn.textContent = '等级区段 ' + (CD_STATE.levels?'✓':'✗');
                     lvBtn.classList.toggle('off', !CD_STATE.levels); };
  lvBtn.onclick=()=>{ CD_STATE.levels=!CD_STATE.levels; syncLv(); renderCustom(filtered()); };
  syncLv();
  /* 自动调节（随视窗缩放调整散点/泡泡大小） */
  const auBtn=document.getElementById('cdAuto');
  const syncAu=()=>{ auBtn.textContent = '自动调节 ' + (CD_STATE.autoSize?'✓':'✗');
                     auBtn.classList.toggle('off', !CD_STATE.autoSize); };
  auBtn.onclick=()=>{ CD_STATE.autoSize=!CD_STATE.autoSize; syncAu(); renderCustom(filtered()); };
  syncAu();
  /* 坐标排序方向（X/Y=轴反向 · Z=梯度方向反转） */
  const invBtns=[['cdInvX','invX'],['cdInvY','invY'],['cdInvZ','invZ']];
  for(const [id,key] of invBtns){
    document.getElementById(id).onclick=()=>{ CD_STATE[key]=!CD_STATE[key]; cdAxisUI(); renderCustom(filtered()); };
  }
  /* 曲面聚合方式（均值/中位数/样本计数）：系列数据整体重算，clear 后重建 */
  const ag=document.getElementById('cdAgg');
  if(ag){
    ag.value=CD_STATE.agg;
    ag.onchange=()=>{ CD_STATE.agg=ag.value; cdModeUI(); cdRebuild3D(); };
  }
  /* 配色模式（动力类型双色系 / Z 梯度单色系）：视觉映射改变，clear 后重建 */
  const cm=document.getElementById('cdColorMode');
  if(cm){
    cm.value=CD_STATE.colorMode;
    cm.onchange=()=>{ CD_STATE.colorMode=cm.value; cdRebuild3D(); };
  }
  /* 球径字段（仅方案B 三维泡泡生效）：逐点直径改变，clear 后重建 */
  const sf=document.getElementById('cdSize');
  if(sf){
    cdFillSelect(sf, CD_STATE.sizeField==null?'':CD_STATE.sizeField, true);
    sf.onchange=()=>{ CD_STATE.sizeField = sf.value===''?null:+sf.value; cdRebuild3D(); };
  }
  /* 视角复位 / 俯视（三维） */
  document.getElementById('cdViewReset').onclick=()=>cdSetView(CD_VIEW);
  document.getElementById('cdViewTop').onclick=()=>cdSetView({alpha:88, beta:0, distance:640});
  /* 全屏展示 */
  const card=document.getElementById('chCustom').closest('.card');
  document.getElementById('cdFull').onclick=()=>{
    if(!document.fullscreenElement){ if(card.requestFullscreen) card.requestFullscreen().catch(()=>showToast('当前环境不允许全屏')); }
    else document.exitFullscreen();
  };
  document.addEventListener('fullscreenchange', ()=>{
    const fs=!!document.fullscreenElement;
    card.classList.toggle('cd-fs', fs);
    document.getElementById('cdFull').textContent = fs?'✕ 退出全屏':'⛶ 全屏';
    /* 全屏布局重排分多帧完成，GL 画布需多次校正才能铺满 */
    const rs=()=>{ const c=chart('chCustom'); if(c) c.resize(); };
    rs(); setTimeout(rs, 90); setTimeout(rs, 260);
  });
  cdModeUI();
  /* 监听 dataZoom：更新视窗状态；开启自动调节时按新视窗重算散点大小（2D 专用） */
  const ch=chart('chCustom');
  if(ch && !ch.__cdZoomBound){
    ch.__cdZoomBound = true;
    ch.on('dataZoom', ()=>{
      try{
        const op = ch.getOption();
        const dz = op.dataZoom || [];
        for(const d of dz){
          if(d.xAxisIndex!=null){ CD_ZOOM.xStart=d.start; CD_ZOOM.xEnd=d.end; }
          if(d.yAxisIndex!=null){ CD_ZOOM.yStart=d.start; CD_ZOOM.yEnd=d.end; }
        }
      }catch(e){}
      if(CD_STATE.autoSize) scheduleCdRender();
      else updateCdZoomText();
    });
    /* 双击图表：三维=视角复位 / 二维=复位视窗 */
    const el=document.getElementById('chCustom');
    if(el) el.addEventListener('dblclick', ()=>cdResetZoom());
  }
}
/* 三维内部重建（曲面聚合方式切换）：捕获相机后 clear，再全量重建 */
function cdRebuild3D(){
  if(!is3dKind(CD_STATE.view)){ renderCustom(filtered()); return; }
  const ch=chart('chCustom');
  if(ch){ const cam=cdCaptureCam(); if(cam) CD_CAM=cam; ch.clear(); ch.__cd3dFirst=true; }
  renderCustom(filtered());
}
/* 视图模式显隐联动：风格下拉仅 3D/热力；3D 按钮组/2D 按钮组/泡泡比例按需出现；
   聚合下拉仅曲面（方案C）显示 */
function cdModeUI(){
  const m=CD_STATE.view, noZ=CD_STATE.z==null, d3=is3dKind(m);
  const cv=document.getElementById('cdView');
  if(cv) cv.value=m;   // 三维/热力不再禁用：未选 Z 时选择后自动补默认 Z 字段
  document.getElementById('cdStyleWrap').classList.toggle('hide', m==='d2');
  document.getElementById('cd3dOnly').classList.toggle('hide', !d3);
  document.getElementById('cd2dOnly').classList.toggle('hide', m!=='d2');
  document.getElementById('cdZScaleWrap').classList.toggle('hide', m!=='d2' || noZ);
  const zi=document.getElementById('cdInvZ');
  if(zi) zi.classList.toggle('hide', m==='d2');
  const aw=document.getElementById('cdAggWrap');
  if(aw) aw.classList.toggle('hide', m!=='surf');
  /* 球径字段仅方案B（三维泡泡）可用 */
  const sw=document.getElementById('cdSizeWrap');
  if(sw) sw.classList.toggle('hide', m!=='bub');
  /* 配色模式对所有三维样式生效（曲面同样支持双/单色系） */
  const cw=document.getElementById('cdColorWrap');
  if(cw) cw.classList.toggle('hide', !d3);
  cdAxisUI();
}
function cdAxisUI(){
  for(const [id,key] of [['cdInvX','invX'],['cdInvY','invY'],['cdInvZ','invZ']]){
    const b=document.getElementById(id); if(!b) continue;
    b.textContent = key.slice(-1).toUpperCase() + (CD_STATE[key]?'↑':'↓');
    b.classList.toggle('on', CD_STATE[key]);
  }
}
/* 缩放中的重绘节流：一帧内只重算一次，避免拖动滑块时卡顿 */
let __cdRaf=null;
function scheduleCdRender(){
  if(__cdRaf) return;
  __cdRaf = requestAnimationFrame(()=>{ __cdRaf=null; renderCustom(filtered()); });
}
function cdResetZoom(silent){
  if(is3dKind(CD_STATE.view)){ cdSetView(CD_VIEW); return; }   // 三维模式下双击 = 视角复位
  if(CD_STATE.view==='heat') return;
  CD_ZOOM={xStart:0,xEnd:100,yStart:0,yEnd:100};
  const ch=chart('chCustom');
  if(ch){
    try{ ch.dispatchAction({type:'dataZoom', start:0, end:100}); }catch(e){}
    try{ ch.dispatchAction({type:'dataZoom', yAxisIndex:0, start:0, end:100}); }catch(e){}
  }
  if(!silent) renderCustom(filtered());
}
function updateCdZoomText(){
  const t=document.getElementById('ci_chCustom');
  if(t && t.innerHTML.indexOf('视窗')<0){
    t.innerHTML += `<br>视窗：X <em>${(CD_ZOOM.xEnd-CD_ZOOM.xStart).toFixed(0)}%</em> · Y <em>${(CD_ZOOM.yEnd-CD_ZOOM.yStart).toFixed(0)}%</em>（拖动滑块或框选可放大；双击图表复位）`;
  }
}
function cdBubbleSize(z, zmin, zmax, scale){
  /* 泡泡直径：Z 在 [zmin,zmax] 内的位置 → 直径区间；直径 ∝ √t（面积≈正比数值）。
     scale 为「自动调节」系数（视窗放大时增大），1 = 基准。 */
  const LO=12, HI=54;
  const span=(zmax-zmin);
  const t = span>0 ? (z-zmin)/span : 0.5;
  const r = Math.sqrt(Math.max(0,Math.min(1,t)));
  return (LO+(HI-LO)*r) * (CD_STATE.zScale/100) * (scale||1);
}

/* ==================== 三维柱状（风格化 · 梯度区间选取 · 轴反向） ==================== */
/* 视角预设与三维箱体尺寸；「视角复位」回等轴测，「俯视 X-Y」正对平面（等同原二维布局） */
/* 三维类视图（方案A 柱状 / 方案B 泡泡 / 方案C 曲面） */
function is3dKind(k){ return k==='bar'||k==='bub'||k==='surf'; }
const CD_VIEW = {alpha:32, beta:-48, distance:340};
const CD_BOX = {w:170, d:170, h:120};
let CD_CAM = null;   // 用户最近一次设置的视角（重建时恢复；轴反向重建前由 cdCaptureCam 刷新）
/* 从 gl 控制器捕获实时相机（拖动旋转写入控制器而非 CD_CAM） */
function cdCaptureCam(){
  try{
    const ch=chart('chCustom');
    const g3=(ch._componentsViews||[]).find(x=>x && x.type==='grid3D');
    const c=g3 && g3._control;
    if(c) return {alpha:c.getAlpha(), beta:c.getBeta(), distance:(c.getDistance?c.getDistance():null)||CD_VIEW.distance};
  }catch(e){}
  return null;
}
/* 视角切换/复位：gl 的 ViewControl 有 _notFirst 守卫——仅首次渲染应用 option 的
   viewControl，之后相机归控制器所有、merge 更新不再生效。因此先 merge 写入模型，
   再调控制器的 setFromViewControlModel 强制应用（公开方法、无重建）。 */
function cdSetView(v){
  CD_CAM = {alpha:v.alpha, beta:v.beta, distance:v.distance||CD_VIEW.distance};
  const ch=chart('chCustom');
  if(!ch || !is3dKind(ch.__cdKind)) return;
  try{
    ch.setOption({grid3D:{viewControl:Object.assign({target:[0,0,0]}, CD_CAM)}}, false);
    const g3=(ch._componentsViews||[]).find(x=>x && x.type==='grid3D');
    const ctrl=g3 && g3._control;
    if(ctrl && ctrl.setFromViewControlModel){
      const m=ch.getModel().getComponent('grid3D', 0).getModel('viewControl');
      if(ctrl.stopAllAnimation) ctrl.stopAllAnimation();
      ctrl.setFromViewControlModel(m);
    }
  }catch(e){}
}
function cdExtent(ps,k){
  let lo=Infinity,hi=-Infinity;
  for(const p of ps){ const v=p.value[k]; if(isFinite(v)){ if(v<lo)lo=v; if(v>hi)hi=v; } }
  return isFinite(lo)?[lo,hi]:[0,1];
}
function cdPadded(e){
  /* 轴两端留 6% 边距避免点贴箱壁；极差为零（单值）时按 10% 绝对幅度展开防退化坐标轴 */
  if(e[1]-e[0] > 1e-9){ const m=(e[1]-e[0])*0.06; return [e[0]-m, e[1]+m]; }
  const m=Math.max(1, Math.abs(e[0])*0.1); return [e[0]-m, e[1]+m];
}
/* 三维值轴通用构造（风格化配色 + 反向镜像标签 + padding 端点隐藏；bar/bub/surf 共用） */
function cdAX3(S,name,ext,z0,inv){ return {type:'value',name,
    min:z0?0:ext[0], max:ext[1], nameGap:18,
    nameTextStyle:{color:S.label,fontSize:12},
    axisLabel:{color:S.label,fontSize:10,textStyle:{color:S.label,fontSize:10},
      /* 抑制 gl 坐标轴 padding 引出的浮点尾数（如 7.8209999…）；千位分隔；隐藏两端 padding 标签；
         反向时标签显示镜像值（刻度位置不变、数值反读） */
      formatter:v=>{ if(z0&&v===0) return '0'; if(v===ext[0]||v===ext[1]) return '';
        let d=v; if(inv) d=(ext[0]+ext[1])-v;
        const r=Math.round(d*10)/10;return Math.abs(r)>=1000?Math.round(r).toLocaleString('zh-CN'):String(r);}},
    axisLine:{lineStyle:{color:S.axisLine}},
    splitLine:{lineStyle:{color:S.split}}}; }
/* 三维渲染（方案A bar3D 海拔柱 / 方案B scatter3D 梯度泡泡）：返回 true=成功；false=WebGL 失败（调用方回退二维）。
   首次渲染 notMerge 全量构建（含 viewControl 默认视角）+ resize 校正 GL 画布；
   后续合并更新——不下发 viewControl，用户旋转/缩放后的相机状态得以保留。 */
function renderCustom3D(ch, kind, bev, phev, m, avg, rng, fv){
  if(kind==='surf') return renderSurface3D(ch, bev, phev, m, avg, rng, fv);
  const S=CD_STYLES[CD_STATE.style] || CD_STYLES.A;
  const xN=m.xN, yN=m.yN, zN=m.zN;
  const el=document.getElementById('chCustom');
  if(el) el.style.background=S.bg;
  /* Z 梯度方向反转 = 色带反转（轴系不动，颜色浅深互换） */
  const brB=CD_STATE.invZ?[...S.bevRamp].reverse():S.bevRamp;
  const brP=CD_STATE.invZ?[...S.phevRamp].reverse():S.phevRamp;
  /* 配色模式：dual=两系列各自色带（BEV 暖橙 / PHEV-EREV 冷蓝）；z=单色带统一着色 */
  const dual = CD_STATE.colorMode!=='z';
  const ramp = CD_STATE.invZ?[...S.bevRamp].reverse():S.bevRamp;
  const all=bev.concat(phev);
  const mkBar=(nm,ps,base,fam)=>({name:nm,type:'bar3D',data:ps,barSize:S.barSize,
    shading:S.shading,
    itemStyle:Object.assign({color:base,opacity:S.opacity},
      S.border?{borderColor:S.border(fam),borderWidth:S.bw}:{},
      S.bevel?{bevelSize:S.bevel[0],bevelSmoothness:S.bevel[1]}:{}),
    emphasis:{itemStyle:{opacity:1}}});
  /* 球径字段：选定字段时按该字段归一化到 3~20 px（直径∝√数值）；未选则等大 8 */
  const szK = (kind==='bub') ? CD_STATE.sizeField : null;
  let szLo=0, szHi=1, szN='';
  if(szK!=null){
    szN=CD_LABEL(szK);
    for(const p of all){ const v=p.sv; if(v!=null&&isFinite(v)){ if(v<szLo)szLo=v; if(v>szHi)szHi=v; } }
    for(const p of all){
      const v=p.sv;
      p.symbolSize = (v!=null&&isFinite(v)&&szHi>szLo)
        ? +(3+17*Math.max(0,Math.min(1,(v-szLo)/(szHi-szLo)))).toFixed(2) : 8;
    }
  }
  const mkBub=(nm,ps,base)=>({name:nm,type:'scatter3D',data:ps,
    symbolSize: szK!=null ? undefined : 8,
    shading:S.shading,
    itemStyle:{color:base,opacity:Math.min(1,S.opacity+.05)},
    emphasis:{itemStyle:{opacity:1}}});
  const exX=cdPadded(cdExtent(all,0)), exY=cdPadded(cdExtent(all,1));
  const invX=CD_STATE.invX, invY=CD_STATE.invY;
  const statB=[avg(bev,0),avg(bev,1),avg(bev,2)], statP=[avg(phev,0),avg(phev,1),avg(phev,2)];
  /* 轴反向=数据空间镜像：gl 轴不支持 inverse 且 min>max 会被归一化回升序（实证选项变画面不变），
     故保持轴几何不变，柱体按 (min+max-值) 镜像换位、刻度标签显示镜像值；ox/oy 保留真实值供悬浮 */
  if(invX||invY){
    const mx=exX[0]+exX[1], my=exY[0]+exY[1];
    for(const p of all){
      if(invX){ p.ox=p.value[0]; p.value[0]=+(mx-p.value[0]).toFixed(4); }
      if(invY){ p.oy=p.value[1]; p.value[1]=+(my-p.value[1]).toFixed(4); }
    }
  }
  const total=all.length;
  const kindName = kind==='bar' ? '三维柱状图（方案A · 海拔）' : '三维泡泡图（方案B · 梯度）';
  const dimTxt = kind==='bar' ? `× 柱高 Z <em>${zN}</em>` : `× Z 位置 <em>${zN}</em> · 颜色=Z 梯度`;
  let core=`<b>${kindName}</b>：X <em>${xN}</em> × Y <em>${yN}</em> ${dimTxt}${szK!=null?` × 球径 <em>${szN}</em>`:''} · 风格 <em>${S.name}</em> · 配色 <em>${dual?'动力类型双色系':'Z 梯度单色系'}</em> · 有效样本 <em>${fmt(total)}</em> 组（BEV <em>${fmt(bev.length)}</em> / PHEV-EREV <em>${fmt(phev.length)}</em>，缺失不计入）`;
  const colorTxt = dual
    ? 'BEV 暖橙 / PHEV-EREV 冷蓝两套色系，各配一条梯度刻度'
    : '单色带统一着色，两系列共用一条梯度刻度';
  let detail= kind==='bar'
    ? `颜色深度=海拔：<em>${zN}</em> 取值范围 <em>${fv(m.zmin)} ~ ${fv(m.zmax)}</em>，同色系内 Z 越大颜色越深${CD_STATE.invZ?'（梯度已反向）':''}；${colorTxt}；<b>拖动右侧色带两端手柄选取数值区间</b>，区间外柱体自动置灰，即高亮目标区域`
    : `颜色=Z 梯度：<em>${zN}</em> 取值范围 <em>${fv(m.zmin)} ~ ${fv(m.zmax)}</em>，同色系内 Z 越大颜色越深${CD_STATE.invZ?'（梯度已反向）':''}；${colorTxt}；<b>拖动右侧色带两端手柄选取数值区间</b>，区间外泡泡自动置灰；切换「方案A」可看逐车型柱高读数`;
  if(szK!=null) detail += `<br>球径字段 <em>${szN}</em> 取值 <em>${fv(szLo)} ~ ${fv(szHi)}</em>，直径按 √数值归一化映射到 3~20 px（面积近似正比于数值）。`;
  const part3=(nm,st)=>st.every(v=>v!=null)?`${nm}：<em>${xN}</em> 均值 <em>${fv(st[0])}</em> · <em>${yN}</em> 均值 <em>${fv(st[1])}</em> · <em>${zN}</em> 均值 <em>${fv(st[2])}</em>`:'';
  const seg=[part3('BEV',statB),part3('PHEV/EREV',statP)].filter(Boolean).join(' ｜ ');
  if(seg) detail += `<br>${seg}`;
  detail += `<br>视角操作：<b>左键拖动=旋转</b> · <b>滚轮=缩放</b> · <b>右键拖动=平移</b> · 双击图表或「视角复位」=复位（等轴测）·「俯视 X-Y」=正对平面`;
  const html=cdInsightWrap(core,detail);
  setInsight('chCustom', html);
  const panelTx=document.getElementById('insightCustomTx');
  if(panelTx) panelTx.innerHTML = html;
  const g3={boxWidth:CD_BOX.w, boxDepth:CD_BOX.d, boxHeight:CD_BOX.h,
    light:{main:{intensity:S.main,shadow:false},ambient:{intensity:S.ambient}},
    axisLine:{lineStyle:{color:S.axisLine}},
    splitLine:{lineStyle:{color:S.split}},
    axisPointer:{show:false}};
  /* 首次构建（含视角切换后的重建）下发 viewControl；数据刷新走 merge 不下发，保留用户相机 */
  if(ch.__cd3dFirst) g3.viewControl=Object.assign({}, CD_CAM||CD_VIEW,
    {target:[0,0,0],minDistance:60,maxDistance:1000,rotateSensitivity:1,zoomSensitivity:1,panSensitivity:1,autoRotate:false});
  const opt={
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{ const d=p.data, v=d.value;
        const ox=d.ox!=null?d.ox:v[0], oy=d.oy!=null?d.oy:v[1];
        return `${p.seriesName} · 品牌车型：<b>${brandName(p)}</b><br>`+
          `${xN}：<b>${fv(ox)} ${xuOf(xN)}</b><br>${yN}：<b>${fv(oy)} ${xuOf(yN)}</b>`+
          `<br>${zN}：<b>${fv(v[2])} ${xuOf(zN)}</b>${kind==='bar'?'（柱高）':'（Z 位置）'}`+
          (szK!=null?`<br>${szN}：<b>${fv(d.sv)} ${xuOf(szN)}</b>（球径）`:''); }},TT),
    legend:LG({data:['BEV','PHEV/EREV']}),
    /* 梯度刻度=区间选择器：calculable 手柄拖动选取区间，区间外柱体置灰（outOfRange）
       双色系=两系列各自色带；单色系=共用一条色带 */
    visualMap: dual ? [
      {type:'continuous',seriesIndex:0,dimension:2,min:m.zmin,max:m.zmax,calculable:true,show:true,
       orient:'vertical',right:8,top:'middle',itemWidth:13,itemHeight:110,hoverLink:true,
       text:[`BEV ${fv(m.zmax)}`,fv(m.zmin)],textStyle:{color:S.label,fontSize:10},
       inRange:{color:brB}, outOfRange:{color:'rgba(132,140,152,.12)'}},
      {type:'continuous',seriesIndex:1,dimension:2,min:m.zmin,max:m.zmax,calculable:true,show:true,
       orient:'vertical',right:66,top:'middle',itemWidth:13,itemHeight:110,hoverLink:true,
       text:[`PHEV/EREV ${fv(m.zmax)}`,fv(m.zmin)],textStyle:{color:S.label,fontSize:10},
       inRange:{color:brP}, outOfRange:{color:'rgba(132,140,152,.12)'}}
    ] : [
      /* 单色系：不绑定 seriesIndex，两系列共用一条 Z 梯度色带 */
      {type:'continuous',dimension:2,min:m.zmin,max:m.zmax,calculable:true,show:true,
       orient:'vertical',right:8,top:'middle',itemWidth:13,itemHeight:110,hoverLink:true,
       text:[`Z ${fv(m.zmax)}`,fv(m.zmin)],textStyle:{color:S.label,fontSize:10},
       inRange:{color:ramp}, outOfRange:{color:'rgba(132,140,152,.12)'}}
    ],
    grid3D:g3,
    xAxis3D:cdAX3(S,xN,exX,false,invX),
    yAxis3D:cdAX3(S,yN,exY,false,invY),
    zAxis3D:cdAX3(S,zN,cdPadded([m.zmin,m.zmax]),true,false),
    series: kind==='bar'
      ? [mkBar('BEV',bev,S.bevBase,'bev'), mkBar('PHEV/EREV',phev,S.phevBase,'phev')]
      : [mkBub('BEV',bev,S.bevBase), mkBub('PHEV/EREV',phev,S.phevBase)]
  };
  try{
    ch.setOption(opt, !!ch.__cd3dFirst);
    if(ch.__cd3dFirst) ch.resize();   // 面板首建时 GL 层画布可能沿用隐藏期尺寸，强制对齐
    ch.__cd3dFirst = false;
    TREND_REGS['chCustom']={};
    return true;
  }catch(e){ return false; }
}
/* ===== 方案C：梯度曲面图（surface3D，X-Y 网格聚合，曲面高=Z 聚合值，颜色=高度梯度） ===== */
/* 与 renderHeat 同源的分箱思路，但输出规则网格供 surface3D 成面；
   空格相应位置透明镂空（与泡泡图一致，只留网格线，不遮挡山峰），即数据覆盖盲区；聚合方式由 CD_STATE.agg 控制 */
function renderSurface3D(ch, bev, phev, m, avg, rng, fv){
  const S=CD_STYLES[CD_STATE.style] || CD_STYLES.A;
  const xN=m.xN, yN=m.yN, zN=m.zN;
  const el=document.getElementById('chCustom');
  if(el) el.style.background=S.bg;
  const brB=CD_STATE.invZ?[...S.bevRamp].reverse():S.bevRamp;
  const brP=CD_STATE.invZ?[...S.phevRamp].reverse():S.phevRamp;
  const all=bev.concat(phev);
  const exX=cdPadded(cdExtent(all,0)), exY=cdPadded(cdExtent(all,1));
  const invX=CD_STATE.invX, invY=CD_STATE.invY;
  const statB=[avg(bev,0),avg(bev,1),avg(bev,2)], statP=[avg(phev,0),avg(phev,1),avg(phev,2)];
  /* X-Y 网格分箱：每格按聚合方式求 Z（均值/中位数/样本计数） */
  const NX=44, NY=30;
  const cw=(exX[1]-exX[0])/NX || 1, chh=(exY[1]-exY[0])/NY || 1;
  const mx=exX[0]+exX[1], my=exY[0]+exY[1];
  const mkGrid=(ps)=>{
    const cells=new Array(NX*NY); let empty=0, nz=0;
    for(const p of ps){
      let xv=p.value[0], yv=p.value[1];
      if(invX) xv=mx-xv; if(invY) yv=my-yv;
      const xi=Math.min(NX-1,Math.max(0,Math.floor((xv-exX[0])/cw)));
      const yi=Math.min(NY-1,Math.max(0,Math.floor((yv-exY[0])/chh)));
      const k=yi*NX+xi;
      (cells[k]||(cells[k]=[])).push(p.value[2]);
    }
    const data=[]; let zlo=Infinity, zhi=-Infinity;
    for(let yi=0; yi<NY; yi++){
      for(let xi=0; xi<NX; xi++){
        const k=yi*NX+xi, xs=exX[0]+(xi+0.5)*cw, ys=exY[0]+(yi+0.5)*chh;
        const arr=cells[k]; let z=0;
        if(arr && arr.length){
          nz++;
          if(CD_STATE.agg==='count') z=arr.length;
          else if(CD_STATE.agg==='median'){
            const a=arr.slice().sort((a,b)=>a-b);
            z=a.length%2 ? a[(a.length-1)/2] : (a[a.length/2-1]+a[a.length/2])/2;
          } else { let s=0; for(const v of arr) s+=v; z=s/arr.length; }
          if(z<zlo)zlo=z; if(z>zhi)zhi=z;
        } else empty++;
        data.push([+(invX?(mx-xs):xs).toFixed(3), +(invY?(my-ys):ys).toFixed(3), +z.toFixed(4)]);
      }
    }
    return {data, zlo:isFinite(zlo)?zlo:0, zhi:isFinite(zhi)?zhi:1, empty, nz};
  };
  const gB=mkGrid(bev), gP=mkGrid(phev);
  const gmin=Math.min(gB.zlo,gP.zlo), gmax=Math.max(gB.zhi,gP.zhi);
  const total=all.length, totEmpty=gB.empty+gP.empty;
  const AGGN={mean:'均值',median:'中位数',count:'样本计数'}[CD_STATE.agg]||'均值';
  /* ===== 空白格/区间外一律透明(与方案B泡泡图一致的底面观感,2026-09-22 用户要求)=====
     历史:①最初 outOfRange 半透明灰泛白、且在屏幕空间盖住山峰;②改为不透明底板色后,
     我用轴标签亮度反推主题明暗,而暗色主题的刻度恰是浅色 -> 误选浅色底板,在暗背景上
     泛出一片白、压住梯度显示。
     正解:空格不铺任何实色。three.js 中 material 为 opaque(transparent:false)时关闭混合,
     alpha=0 的顶点只写深度、不改帧缓冲颜色 -> 空白处看不见,只剩 grid3D 网格线,
     与方案B的底面观感一致。地形本体保持 opacity:1:不透明 mesh 才写深度缓冲,
     山峰因 z 更高天然位于底面之前,不会被后绘制的系列罩住。 */
  const EMPTY_COLOR='rgba(0,0,0,0)';   /* 空白格/区间外:全透明 */

  const mkSf=(nm,g,fam)=>({name:nm,type:'surface',data:g.data,
    shading:S.shading,
    itemStyle:Object.assign({opacity:1},
      S.border?{borderColor:S.border(fam),borderWidth:Math.max(.5,S.bw)}:{},
      S.bevel?{bevelSize:S.bevel[0],bevelSmoothness:S.bevel[1]}:{}),
    emphasis:{itemStyle:{opacity:1}}});
  const part3=(nm,st)=>st.every(v=>v!=null)?`${nm}：<em>${xN}</em> 均值 <em>${fv(st[0])}</em> · <em>${yN}</em> 均值 <em>${fv(st[1])}</em> · <em>${zN}</em> 均值 <em>${fv(st[2])}</em>`:'';
  const seg=[part3('BEV',statB),part3('PHEV/EREV',statP)].filter(Boolean).join(' ｜ ');
  let core=`<b>梯度曲面图（方案C · 地形）</b>：X <em>${xN}</em> × Y <em>${yN}</em> · 曲面高=Z <em>${zN}</em> 的<em>${AGGN}</em> · 网格 <em>${NX}×${NY}</em> · 风格 <em>${S.name}</em> · 有效样本 <em>${fmt(total)}</em> 组（BEV <em>${fmt(bev.length)}</em> / PHEV-EREV <em>${fmt(phev.length)}</em>，缺失不计入）`;
  let detail=`颜色深度=地形海拔：曲面高度为 <em>${zN}</em> 的${AGGN}（范围 <em>${fv(gmin)} ~ ${fv(gmax)}</em>），同色系内越高颜色越深${CD_STATE.invZ?'（梯度已反向）':''}；<b>拖动右侧色带两端手柄选取数值区间</b>，区间外变为透明（镂空）。`+
    `<br><b>空白格</b>：<em>${fmt(totEmpty)}</em> 个网格无样本，高度按 0 显示为低洼，即数据覆盖盲区；`+
    `切「视图 → 三维柱状」可回到逐车型海拔视角。`;
  if(seg) detail += `<br>${seg}`;
  detail += `<br>视角操作：<b>左键拖动=旋转</b> · <b>滚轮=缩放</b> · <b>右键拖动=平移</b> · 双击图表或「视角复位」=复位（等轴测）·「俯视 X-Y」=正对平面`;
  const html=cdInsightWrap(core,detail);
  setInsight('chCustom', html);
  const panelTx=document.getElementById('insightCustomTx');
  if(panelTx) panelTx.innerHTML = html;
  const g3={boxWidth:CD_BOX.w, boxDepth:CD_BOX.d, boxHeight:CD_BOX.h,
    light:{main:{intensity:S.main,shadow:false},ambient:{intensity:S.ambient}},
    axisLine:{lineStyle:{color:S.axisLine}},
    splitLine:{lineStyle:{color:S.split}},
    axisPointer:{show:false}};
  if(ch.__cd3dFirst) g3.viewControl=Object.assign({}, CD_CAM||CD_VIEW,
    {target:[0,0,0],minDistance:60,maxDistance:1000,rotateSensitivity:1,zoomSensitivity:1,panSensitivity:1,autoRotate:false});
  const opt={
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{ const v=p.data;
        return `${p.seriesName} · 网格中心<br>${xN}：<b>${fv(v[0])} ${xuOf(xN)}</b><br>`+
          `${yN}：<b>${fv(v[1])} ${xuOf(yN)}</b><br>${zN} ${AGGN}：<b>${fv(v[2])} ${xuOf(zN)}</b>`; }},TT),
    legend:LG({data:['BEV','PHEV/EREV']}),
    visualMap:[
      {type:'continuous',seriesIndex:0,dimension:2,min:gmin,max:gmax,calculable:true,show:true,
       orient:'vertical',right:8,top:'middle',itemWidth:13,itemHeight:110,hoverLink:true,
       text:[`BEV ${fv(gmax)}`,fv(gmin)],textStyle:{color:S.label,fontSize:10},
       inRange:{color:brB}, outOfRange:{color:EMPTY_COLOR}},
      {type:'continuous',seriesIndex:1,dimension:2,min:gmin,max:gmax,calculable:true,show:true,
       orient:'vertical',right:66,top:'middle',itemWidth:13,itemHeight:110,hoverLink:true,
       text:[`PHEV/EREV ${fv(gmax)}`,fv(gmin)],textStyle:{color:S.label,fontSize:10},
       inRange:{color:brP}, outOfRange:{color:EMPTY_COLOR}}
    ],
    grid3D:g3,
    xAxis3D:cdAX3(S,xN,exX,false,invX),
    yAxis3D:cdAX3(S,yN,exY,false,invY),
    zAxis3D:cdAX3(S,zN,cdPadded([gmin,gmax]),true,false),
    series:[mkSf('BEV',gB,'bev'), mkSf('PHEV/EREV',gP,'phev')]
  };
  try{
    ch.setOption(opt, !!ch.__cd3dFirst);
    if(ch.__cd3dFirst) ch.resize();
    ch.__cd3dFirst = false;
    TREND_REGS['chCustom']={};
    return true;
  }catch(e){ return false; }
}
/* 二维热力：X/Y 网格聚合，颜色=格内 Z 均值；色带同样支持区间选取高亮 */
function renderHeat(ch, bev, phev, m, avg, rng, fv){
  const S=CD_STYLES[CD_STATE.style] || CD_STYLES.A;
  const el=document.getElementById('chCustom');
  if(el) el.style.background=S.bg;
  const xN=m.xN, yN=m.yN, zN=m.zN;
  const all=bev.concat(phev);
  const ex=cdPadded(cdExtent(all,0)), ey=cdPadded(cdExtent(all,1));
  const NX=44, NY=30;
  const cw=(ex[1]-ex[0])/NX || 1, chh=(ey[1]-ey[0])/NY || 1;
  const cells=new Map();
  for(const p of all){
    const xi=Math.min(NX-1,Math.max(0,Math.floor((p.value[0]-ex[0])/cw)));
    const yi=Math.min(NY-1,Math.max(0,Math.floor((p.value[1]-ey[0])/chh)));
    const k=xi*NY+yi; const c=cells.get(k)||{s:0,n:0}; c.s+=p.value[2]; c.n++; cells.set(k,c);
  }
  const data=[]; let dmin=Infinity,dmax=-Infinity;
  for(const [k,c] of cells){
    const xi=Math.floor(k/NY), yi=k%NY; const v=+(c.s/c.n).toFixed(1);
    if(v<dmin)dmin=v; if(v>dmax)dmax=v;
    data.push({value:[xi,yi,v], n:c.n,
      x0:Math.round(ex[0]+xi*cw), x1:Math.round(ex[0]+(xi+1)*cw),
      y0:Math.round(ey[0]+yi*chh), y1:Math.round(ey[0]+(yi+1)*chh)});
  }
  const xCats=[...Array(NX)].map((_,i)=>String(Math.round(ex[0]+(i+.5)*cw)));
  const yCats=[...Array(NY)].map((_,i)=>String(Math.round(ey[0]+(i+.5)*chh)));
  let core=`<b>二维热力图</b>：X <em>${xN}</em> × Y <em>${yN}</em> 网格聚合（${NX}×${NY} 格）· 颜色=格内 <em>${zN}</em> 均值 · 覆盖 <em>${fmt(data.length)}</em> 格 / <em>${fmt(all.length)}</em> 款`;
  let detail=`色带即梯度刻度：<b>拖动两端手柄选取均值区间</b>，区间外格子自动置灰；悬浮格子可见该格取值范围、车型数与均值。切「视图 → 三维柱状」可回到逐车型海拔视角。`;
  const html=cdInsightWrap(core,detail);
  setInsight('chCustom', html);
  const panelTx=document.getElementById('insightCustomTx');
  if(panelTx) panelTx.innerHTML = html;
  ch.setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{ const d=p.data;
        return `${xN}：<b>${fmt(d.x0)} ~ ${fmt(d.x1)} ${xuOf(xN)}</b><br>`+
          `${yN}：<b>${fmt(d.y0)} ~ ${fmt(d.y1)} ${xuOf(yN)}</b><br>`+
          `${zN} 均值：<b>${fv(d.value[2])} ${xuOf(zN)}</b> · 车型数 <b>${d.n}</b>`; }},TT),
    visualMap:{type:'continuous',min:dmin,max:dmax,calculable:true,show:true,
      orient:'vertical',right:8,top:'middle',itemWidth:13,itemHeight:120,hoverLink:true,
      text:[`${zN} ${fv(dmax)}`,fv(dmin)],textStyle:{color:S.label,fontSize:10},
      inRange:{color:S.bevRamp}, outOfRange:{color:'rgba(132,140,152,.12)'}},
    grid:Object.assign(GRID(),{right:92,bottom:60}),
    xAxis:Object.assign({type:'category',data:xCats,name:xN,nameLocation:'middle',nameGap:26,
      nameTextStyle:{color:S.label},inverse:CD_STATE.invX,
      axisLabel:{color:S.label,fontSize:10,interval:Math.max(1,Math.floor(NX/9))-1,formatter:v=>(+v).toLocaleString('zh-CN')},
      axisLine:{lineStyle:{color:S.axisLine}},axisTick:{show:false},splitLine:{show:false}},{}),
    yAxis:Object.assign({type:'category',data:yCats,name:yN,nameGap:16,inverse:CD_STATE.invY,
      nameTextStyle:{color:S.label},
      axisLabel:{color:S.label,fontSize:10,interval:Math.max(1,Math.floor(NY/8))-1,formatter:v=>(+v).toLocaleString('zh-CN')},
      axisLine:{lineStyle:{color:S.axisLine}},axisTick:{show:false},splitLine:{show:false}},{}),
    series:[{name:'全部车型',type:'heatmap',data,
      itemStyle:{borderColor:S.split,borderWidth:.5},
      emphasis:{itemStyle:{borderColor:S.label,borderWidth:1}},progressive:5000}]
  }, true);
  TREND_REGS['chCustom']={};
}
/* 固定坐标系外框：视窗缩放时按可见跨度自动放大散点，保证细节可读（可关闭） */
/* 视窗缩放状态：dataZoom 的 start/end（%），用于计算可见跨度与自动调节系数 */
let CD_ZOOM = {xStart:0, xEnd:100, yStart:0, yEnd:100};
function cdZoomFactor(axis){
  const s = axis==='x' ? CD_ZOOM.xStart : CD_ZOOM.yStart;
  const e = axis==='x' ? CD_ZOOM.xEnd   : CD_ZOOM.yEnd;
  const frac = Math.max(0.02, Math.min(1, (e - s) / 100));   // 可见比例
  return 1 / frac;                                            // 放大倍数
}
const CD_FRAME = {w: 760, h: 420};
function cdVisibleSpan(rows, xK, yK){
  let xmin=Infinity,xmax=-Infinity,ymin=Infinity,ymax=-Infinity, n=0;
  for(const r of rows){
    const x=r[xK], y=r[yK];
    if(x==null||y==null) continue;
    if(x<xmin)xmin=x; if(x>xmax)xmax=x; if(y<ymin)ymin=y; if(y>ymax)ymax=y; n++;
  }
  return n ? {x:Math.max(1,xmax-xmin), y:Math.max(1,ymax-ymin), n} : {x:1,y:1,n:0};
}
function cdAutoScale(full, vis){
  if(!CD_STATE.autoSize || !vis.n) return 1;
  const base = Math.min(CD_FRAME.w/full.x, CD_FRAME.h/full.y);
  const cur  = Math.min(CD_FRAME.w/vis.x,  CD_FRAME.h/vis.y);
  if(!isFinite(base) || !isFinite(cur) || cur<=0) return 1;
  return Math.max(0.6, Math.min(3.2, cur/base));
}
function renderCustom(rows){
  const el=document.getElementById('chCustom'); if(!el) return;
  const xK=CD_STATE.x, yK=CD_STATE.y, zK=CD_STATE.z;
  const hasZ = zK!=null;
  let mode = CD_STATE.view;
  if(mode!=='d2' && !hasZ) mode='d2';
  const bev=[], phev=[];
  let zmin=Infinity, zmax=-Infinity;
  for(const r of rows){
    const x=r[xK], y=r[yK];
    if(x==null||y==null) continue;
    let z=null;
    if(hasZ){ z=r[zK]; if(z==null||!isFinite(z)) continue; if(z<zmin)zmin=z; if(z>zmax)zmax=z; }
    /* 球径字段值（方案B 三维泡泡可选第五通道） */
    let sv=null;
    if(mode==='bub' && CD_STATE.sizeField!=null){
      sv=r[CD_STATE.sizeField]; if(sv==null||!isFinite(sv)) sv=null;
    }
    const pt={value: hasZ?[x,y,z]:[x,y], n:r[I.gn], b:r[I.bd], sv:sv};
    (r[I.t]==='BEV'?bev:phev).push(pt);
  }
  const xN=CD_LABEL(xK), yN=CD_LABEL(yK), zN=hasZ?CD_LABEL(zK):'';
  const ch3=chart('chCustom');
  if(!ch3) return;
  /* 模式切换（2D ⇄ 3D ⇄ 热力）/ 轴反向：坐标系不同或轴向变化，clear 后全量重建
     （gl 三维坐标系对 merge 的 min-max/inverse 更新不重排布局——实证选项已变画面不变）；
     重建前捕获控制器实时相机，重建后经 viewControl 下发以保留用户视角 */
  const invKey = CD_STATE.invX+','+CD_STATE.invY;
  /* 曲面聚合方式 / 泡泡球径字段 / 配色模式改变也需 clear 重建（系列数据或视觉映射整体重算） */
  const subKey = CD_STATE.agg+'|'+(CD_STATE.sizeField==null?'-':CD_STATE.sizeField)+'|'+CD_STATE.colorMode;
  if(ch3.__cdKind !== mode || ch3.__cdInvKey !== invKey || ch3.__cdSubKey !== subKey){
    if(is3dKind(ch3.__cdKind)){ const cam=cdCaptureCam(); if(cam) CD_CAM=cam; }
    ch3.clear(); ch3.__cdKind = mode; ch3.__cdInvKey = invKey; ch3.__cdSubKey = subKey; ch3.__cd3dFirst = true;
  }
  /* 顶部实时状态分析共用工具（三种模式共用） */
  const avg=(ps,k)=>{ let s=0,n=0; for(const p of ps){const v=p.value[k]; if(isFinite(v)){s+=v;n++;}} return n?s/n:null; };
  const fv=v=>v==null?'–':(Math.abs(v)>=100?Math.round(v).toLocaleString('zh-CN'):(+v.toFixed(1)));
  const rng=(ps,k)=>{ let lo=Infinity,hi=-Infinity; for(const p of ps){const v=p.value[k]; if(isFinite(v)){if(v<lo)lo=v;if(v>hi)hi=v;}} return isFinite(lo)?[lo,hi]:null; };
  if(is3dKind(mode)){
    if(renderCustom3D(ch3, mode, bev, phev, {xN,yN,zN,zmin,zmax}, avg, rng, fv)) return;
    /* WebGL 初始化失败 → 回退二维泡泡图 */
    CD_STATE.view='d2'; cdModeUI();
    ch3.__cdKind='d2'; ch3.clear(); ch3.__cd3dFirst=true;
    showToast('三维展示初始化失败（当前环境可能不支持 WebGL），已回退二维泡泡图');
  }
  if(mode==='heat'){ renderHeat(ch3, bev, phev, {xN,yN,zN,zmin,zmax}, avg, rng, fv); return; }
  /* ===== 二维散点/泡泡 ===== */
  /* 固定坐标系外框 + 视窗缩放：按当前 dataZoom 视窗计算「自动调节」系数 */
  const fullSpan = cdVisibleSpan(rows, xK, yK);
  const zx = cdZoomFactor('x'), zy = cdZoomFactor('y');
  const visSpan = {x: fullSpan.x / zx, y: fullSpan.y / zy, n: fullSpan.n};
  const autoScale = cdAutoScale(fullSpan, visSpan);
  const SCATTER_SIZE = 7;
  /* 直径写入每个数据点（实测最稳：不依赖 ECharts 回调求值；缩放/换字段即时生效）
     Z 相对大小始终保留；「自动调节」开启时按视窗整体放大，保证放大后细节可读 */
  if(hasZ){
    for(const p of bev)  p.symbolSize = cdBubbleSize(p.value[2], zmin, zmax, autoScale);
    for(const p of phev) p.symbolSize = cdBubbleSize(p.value[2], zmin, zmax, autoScale);
  } else {
    for(const p of bev)  p.symbolSize = SCATTER_SIZE * autoScale;
    for(const p of phev) p.symbolSize = SCATTER_SIZE * autoScale;
  }
  const trB=trendSeries(bev.map(p=>p.value),'#D97757','BEV趋势线');
  const trP=trendSeries(phev.map(p=>p.value),'#6B9BD1','PHEV/EREV趋势线');
  const trends=[trB,trP].filter(Boolean);
  TREND_REGS['chCustom']=Object.fromEntries(trends.map(t=>[t.name,t.__formula]));
  const mk=(nm,ps,color)=>hasZ
    ? {name:nm,type:'scatter',data:ps,itemStyle:{color:`rgba(${color},.45)`,
        borderColor:`rgba(${color},.85)`,borderWidth:1},emphasis:{focus:'series'}}
    : {name:nm,type:'scatter',data:ps,symbolSize:SCATTER_SIZE,
        itemStyle:{color:`rgba(${color},${TC().scatterA})`},large:true,largeThreshold:800};
  const series=[ mk('BEV',bev,'217,119,87'), mk('PHEV/EREV',phev,'107,155,209') ];
  trends.forEach(t=>series.push(t));
  /* 轴距等级区段辅助示意（可关闭）：挂在首个数据系列上，虚线分隔 A/B/C/D/E 区间 */
  let lvShown=false;
  if(CD_STATE.levels){
    const marks=cdLevelMarks(xN,yN);
    if(marks && series.length){ series[0].markLine=marks; lvShown=true; }
  }
  const total=bev.length+phev.length;
  const part=(nm,ps)=>ps.length?`${nm} <em>${xN}</em> 均值 <em>${fv(avg(ps,0))}</em>（${fv((rng(ps,0)||[0,0])[0])}~${fv((rng(ps,0)||[0,0])[1])}）· <em>${yN}</em> 均值 <em>${fv(avg(ps,1))}</em>（${fv((rng(ps,1)||[0,0])[0])}~${fv((rng(ps,1)||[0,0])[1])}）`:'';
  const core=`<b>${hasZ?'泡泡图':'散点图'}</b>：X <em>${xN}</em> × Y <em>${yN}</em>${hasZ?` × 高度坐标 Z <em>${zN}</em>`:''} · 有效样本 <em>${fmt(total)}</em> 组（BEV <em>${fmt(bev.length)}</em> / PHEV-EREV <em>${fmt(phev.length)}</em>，缺失不计入）`;
  let detail= hasZ
    ? `Z 轴 <em>${zN}</em> 取值范围 <em>${fv(zmin)} ~ ${fv(zmax)}</em>，直径映射 <em>${Math.round(12*CD_STATE.zScale/100)}~${Math.round(54*CD_STATE.zScale/100)} px</em>（直径 ∝ √Z，面积正比于数值，当前比例 <em>${CD_STATE.zScale}%</em>）`
    : ` ｜ 在「高度坐标 Z」中选择任一可量化字段即可切换为泡泡图；或用「视图」切换三维柱状 / 梯度泡泡 / 梯度曲面 / 二维热力`;
  const seg=[part('BEV',bev), part('PHEV/EREV',phev)].filter(Boolean).join(' ｜ ');
  if(seg) detail += `<br>${seg}`;
  /* 等级区段辅助示意状态 */
  if(xN.indexOf('轴距')===0 || yN.indexOf('轴距')===0){
    const AX = xN.indexOf('轴距')===0 ? 'X' : 'Y';
    detail += lvShown
      ? `<br>等级区段：已按《车型级别定义》在 <em>${AX} 轴（轴距）</em> 标注分级界线 —— `+
        `<span style="color:#D97757">轿车 A/B/C/D/E</span>（2650/2740/2850/3000）、`+
        `<span style="color:#6B9BD1">SUV B/C/D/E</span>（2680/2850/3000）、`+
        `<span style="color:#7FA97F">MPV C/D</span>（3000），单位 mm；可点「等级区段」关闭`
      : `<br>等级区段：<em>已关闭</em>（${AX} 轴为轴距，点「等级区段」可显示 A/B/C/D/E 分级界线）`;
  } else {
    detail += `<br>等级区段：当前 X / Y 均非轴距，无分级界线可标注（将任一轴选为「轴距(mm)」即显示）`;
  }
  /* 放大视窗与自动调节状态 */
  const zoomPct = `${(CD_ZOOM.xEnd-CD_ZOOM.xStart).toFixed(0)}% × ${(CD_ZOOM.yEnd-CD_ZOOM.yStart).toFixed(0)}%`;
  const zoomed = (CD_ZOOM.xEnd-CD_ZOOM.xStart) < 99.5 || (CD_ZOOM.yEnd-CD_ZOOM.yStart) < 99.5;
  detail += `<br>放大视窗：视窗 <em>${zoomPct}</em>` +
    (zoomed ? `（放大 <em>${cdZoomFactor('x').toFixed(1)}×</em>）` : '（完整视野）') +
    ` · 自动调节 <em>${CD_STATE.autoSize?'开':'关'}</em>` +
    (CD_STATE.autoSize ? `（当前散点大小系数 <em>${autoScale.toFixed(2)}×</em>）` : '（散点大小固定）') +
    ` ｜ 拖动滑块/框选放大，双击图表或「复位视窗」还原`;
  const html=cdInsightWrap(core,detail);
  setInsight('chCustom', html);
  const panelTx=document.getElementById('insightCustomTx');
  if(panelTx) panelTx.innerHTML = html;
  ch3.setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{
        const t=trends.find(x=>x.name===p.seriesName);
        if(t) return t.tooltip.formatter();
        const v=p.data.value;
        return `${p.seriesName} · 品牌车型：<b>${brandName(p)}</b><br>`+
          `${xN}：<b>${v[0]} ${xuOf(xN)}</b><br>${yN}：<b>${v[1]} ${xuOf(yN)}</b>`+
          (hasZ?`<br>${zN}：<b>${v[2]} ${xuOf(zN)}</b>`:'');
      }},TT),
    legend:LG({data:['BEV','PHEV/EREV',...trends.map(t=>t.name)]}),
    grid:Object.assign(GRID(), {bottom:74}),
    /* 放大视窗：坐标系外框尺寸固定（grid 不变），通过 dataZoom 改变可见范围；
       支持滑块拖动 / 框选，双击图表或「复位视窗」恢复完整视野 */
    dataZoom:[
      {type:'inside', xAxisIndex:0, start:CD_ZOOM.xStart, end:CD_ZOOM.xEnd, zoomOnMouseWheel:true, moveOnMouseWheel:false, moveOnMouseMove:true},
      {type:'inside', yAxisIndex:0, start:CD_ZOOM.yStart, end:CD_ZOOM.yEnd, zoomOnMouseWheel:true, moveOnMouseWheel:false, moveOnMouseMove:true},
      {type:'slider', xAxisIndex:0, height:16, bottom:30, start:CD_ZOOM.xStart, end:CD_ZOOM.xEnd,
       borderColor:TC().axisLine, backgroundColor:'transparent', fillerColor:'rgba(217,119,87,.14)',
       handleStyle:{color:TC().axis}, textStyle:{color:TC().axis, fontSize:10}, labelFormatter:''},
      {type:'slider', yAxisIndex:0, width:16, right:6, start:CD_ZOOM.yStart, end:CD_ZOOM.yEnd,
       borderColor:TC().axisLine, backgroundColor:'transparent', fillerColor:'rgba(107,155,209,.14)',
       handleStyle:{color:TC().axis}, textStyle:{color:TC().axis, fontSize:10}, labelFormatter:''}
    ],
    toolbox:{show:false},
    xAxis:Object.assign({type:'value',name:xN,nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true,inverse:CD_STATE.invX},AXS()),
    yAxis:Object.assign({type:'value',name:yN,nameTextStyle:{color:TC().axis},scale:true,inverse:CD_STATE.invY},AXS()),
    series
  },true);
  bindTrendClick('chCustom');
}
/* 从「名称(单位)」中取单位 */
function xuOf(name){ const m=String(name).match(/\(([^)]*)\)\s*$/); return m?m[1]:''; }
function initCustomDist(){ cdInitUI(); }

/* ==================== 数据质量（底部） ==================== */
function chFill(rows){
  if(!document.getElementById('chFill'))return;
  const entries = COL_NAMES.map((n,i)=>({n,i,f:rows.filter(r=>r[i]!=null&&r[i]!=='').length}))
    .filter(e=>e.i!==TX_IDX)
    .map(e=>[e.n,e.f])
    .sort((a,b)=>b[1]-a[1]);
  const names = entries.map(e=>e[0]).reverse();
  const vals = entries.map(e=>+(pct(e[1],rows.length)).toFixed(1)).reverse();
  const meanFill = (entries.length&&rows.length)? (entries.reduce((a,e)=>a+e[1],0)/entries.length/rows.length*100).toFixed(1):'0.0';
  setInsight('chFill', `平均字段完整度 <em>${meanFill}%</em>${entries.length?` · 最短板 <em>${entries[entries.length-1][0]}</em>（${entries[entries.length-1][1]}%）`:''}`);
  chart('chFill').setOption({
    tooltip:Object.assign({trigger:'axis',axisPointer:{type:'shadow'},
      formatter:ps=>`${ps[0].name}：非空 <b>${ps[0].value}%</b>（缺失 ${(100-ps[0].value).toFixed(1)}%）`},TT),
    grid:Object.assign({},GRID({left:10,right:60,top:8,bottom:8}),{containLabel:true}),
    xAxis:Object.assign({type:'value',max:100,name:'非空率（%）',axisLabel:{formatter:'{value}%'}},AXS({splitLine:{show:false},axisLine:{show:false}})),
    yAxis:Object.assign({type:'category',data:names,name:'字段名称'},AXS({axisLabel:{color:TC().lbl,fontSize:11.5}})),
    series:[{type:'bar',data:vals.map(v=>({value:v,itemStyle:{color:v>90?'#7FA97F':v>70?'#D4A94E':v>40?'#D97757':'#C97B7B'}})),
      barMaxWidth:14,label:{show:true,position:'right',color:TC().axis,fontSize:10.5,fontFamily:'Consolas',formatter:'{c}%'}}]
  },true);
}
function chSrc(rows){
  if(!document.getElementById('chSrc'))return;
  const es = groupCount(rows,I.src);
  const total = es.reduce((a,b)=>a+b[1],0);
  const simp = s => s.replace(/梳理补充_/g,'');
  let items = es.map(e=>({name:simp(e[0]),value:e[1]}));
  const big = items.filter(it=>total&&it.value/total>=0.015);
  const small = items.filter(it=>total&&it.value/total<0.015);
  if(small.length) big.push({name:'其他·'+small.length+'类',value:small.reduce((a,b)=>a+b.value,0)});
  setInsight('chSrc', total?`合并来源 <em>${es.length}</em> 类 · 首位 <em>${big.length?big[0].name:'-'}</em> 占 <em>${big.length?p1(big[0].value,total):0}%</em> · 长尾归并 ${small.length} 类`:'');
  chart('chSrc').setOption({
    tooltip:Object.assign({trigger:'item',formatter:p=>`${p.name}<br><b>${p.value.toLocaleString('zh-CN')}</b> 款（${p.percent}%）`},TT),
    legend:{type:'scroll',orient:'vertical',right:6,top:'middle',textStyle:{color:TC().axis,fontSize:11},
      pageIconColor:TC().axis,pageTextStyle:{color:TC().axis,fontSize:10}},
    series:[{type:'pie',radius:['42%','66%'],center:['34%','52%'],
      label:{show:true,color:TC().lbl,fontSize:11,fontFamily:'Consolas',
        formatter:p=>p.percent>=1.5?`${p.percent}%`:''},
      labelLine:{length:10,length2:8,lineStyle:{color:TC().axisLine}},
      itemStyle:{borderColor:TC().pieBd,borderWidth:2},
      data:big.map((e,i)=>({value:e.value,name:e.name,itemStyle:{color:PAL[i%PAL.length]}}))}]
  },true);
}
function chHeat(rows){
  if(!document.getElementById('chHeat'))return;
  const bs = byBatch(rows).filter((x,i)=>i%2===0);
  const labels = bs.map(x=>String(x[0]));
  const checkIdx = [I.s,I.w,I.r,I.c,I.bt,I.ed,I.ec,I.fp,I.tp,I.ms,I.tq,I.fo,I.dv,I.ep,I.es];
  const checkNames = ['细分市场','整备质量','续航','容量','电池类型','能量密度','电耗','前电机功率','总功率','电机供应商','系统扭矩','油耗','排量','发动机功率','发动机供应商'];
  const data=[];
  bs.forEach((x,yi)=>{
    checkIdx.forEach((idx,xi)=>{
      const f = x[1].filter(r=>r[idx]!=null).length;
      const v = x[1].length? +(f/x[1].length*100).toFixed(0):0;
      data.push([xi,yi,v]);
    });
  });
  let worst = null;
  data.forEach(p=>{ if(!worst||p[2]<worst[2]) worst=p; });
  setInsight('chHeat', worst?`最薄弱组合：第 <em>${labels[worst[1]]}</em> 批 · <em>${checkNames[worst[0]]}</em>（填报率 <em>${worst[2]}%</em>）`:'');
  chart('chHeat').setOption({
    tooltip:Object.assign({position:'top',
      formatter:p=>`第 ${labels[p.value[1]]} 批 · ${checkNames[p.value[0]]}<br>填报率 <b>${p.value[2]}%</b>`},TT),
    grid:Object.assign({},GRID({left:10,right:70,top:10,bottom:56}),{containLabel:true}),
    xAxis:Object.assign({type:'category',data:checkNames,axisLabel:{color:TC().axis,fontSize:10,rotate:40},name:'字段'},AXS({splitLine:{show:false}})),
    yAxis:Object.assign({type:'category',data:labels,name:'批次'},AXS({splitLine:{show:false}})),
    visualMap:{min:0,max:100,calculable:false,orient:'horizontal',left:'center',bottom:0,
      textStyle:{color:TC().axis,fontSize:10},
      inRange:{color:['#7A2E2E','#B3613C','#C9973F','#6F9E62','#4E8B5F']}},
    series:[{type:'heatmap',data,label:{show:false},
      itemStyle:{borderColor:TC().pieBd,borderWidth:1}}]
  },true);
}
function insight5(rows){
  if(!document.getElementById('insight5Tx'))return;
  const fields = COL_NAMES.map((n,i)=>({n,i,f:rows.filter(r=>r[i]!=null&&r[i]!=='').length}));
  const worst = fields.filter(f=>f.f/rows.length<0.6).sort((a,b)=>a.f-b.f).slice(0,3);
  document.getElementById('insight5Tx').innerHTML =
    `共 <em>${rows.length}</em> 条记录 × <em>${COL_NAMES.length}</em> 个关键字段。填报率不足 60% 的字段：`+
    (worst.length?worst.map(f=>`<em>${f.n}</em>（${pct(f.f,rows.length).toFixed(0)}%）`).join('、'):'无')+
    `。热力图按批次展开，可定位数据缺失集中的批次区间。业务图表统计均为已填报口径，缺失值不参与计算与百分比。`;
}

/* ==================== Panel 5 企业与品牌 ==================== */
function chEntTrend(rows){
  if(!document.getElementById('chEntTrend'))return;
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  const top = groupCount(rows,I.e,8).map(e=>e[0]);
  const colors=['#D97757','#6B9BD1','#7FA97F','#D4A94E','#A97FA9','#5BAFAF','#C97B7B','#B08968'];
  const series = top.map((ent,si)=>({
    name:ent, type:'line', stack:'total', smooth:true, symbol:'circle', symbolSize:4,
    areaStyle:{opacity:.26}, lineStyle:{width:1.4}, emphasis:{focus:'series'},
    itemStyle:{color:colors[si%colors.length]},
    data: bs.map(x=>x[1].filter(r=>r[I.e]===ent).length)
  }));
  const lastB = bs[bs.length-1];
  let leader = null, lN = 0;
  if(lastB){ top.forEach(e=>{ const c=lastB[1].filter(r=>r[I.e]===e).length; if(c>lN){lN=c;leader=e;} }); }
  setInsight('chEntTrend', top.length?`区间头部 <em>${top[0]}</em>（${fmt(groupCount(rows,I.e)[0][1])} 款）${leader?` · 最新批次领跑 ${leader}（${lN} 款）`:''}`:'暂无企业数据');
  chart('chEntTrend').setOption({
    tooltip:Object.assign({trigger:'axis'},TT),
    legend:LG({type:'scroll',pageIconColor:TC().axis}),
    grid:GRID({bottom:64,top:46}),
    xAxis:Object.assign({type:'category',data:labels,boundaryGap:false,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:Object.assign({type:'value',name:'车型数（款）',nameTextStyle:{color:TC().axis}},AXS()),
    series},true);
}
function chEntSeg(rows){
  if(!document.getElementById('chEntSeg'))return;
  const ents = groupCount(rows,I.e,10).map(e=>e[0]);
  const segs = groupCount(rows.filter(r=>ents.includes(r[I.e])),I.s,8).map(e=>e[0]);
  if(!ents.length||!segs.length){chart('chEntSeg').clear();return;}
  const data=[]; let vmax=0;
  ents.forEach((e,yi)=>segs.forEach((s,xi)=>{
    const c=rows.filter(r=>r[I.e]===e&&r[I.s]===s).length;
    if(c>0)data.push([xi,yi,c]); if(c>vmax)vmax=c;
  }));
  let bpk = null;
  data.forEach(p=>{ if(!bpk||p[2]>bpk[2]) bpk=p; });
  setInsight('chEntSeg', bpk?`最强组合 <em>${ents[bpk[1]]} × ${segs[bpk[0]]}</em>（<em>${bpk[2]}</em> 款）· 覆盖组合 ${fmt(data.length)} 组`:'暂无企业×细分市场交叉填报');
  chart('chEntSeg').setOption({
    tooltip:Object.assign({position:'top',
      formatter:p=>`<b>${ents[p.value[1]]}</b> × ${segs[p.value[0]]}<br>公告车型：<b>${p.value[2]}</b> 款`},TT),
    grid:Object.assign({},GRID({left:10,right:20,top:10,bottom:70}),{containLabel:true}),
    xAxis:Object.assign({type:'category',data:segs,axisLabel:{color:TC().axis,fontSize:10,rotate:35}},AXS({splitLine:{show:false}})),
    yAxis:Object.assign({type:'category',data:ents,axisLabel:{color:TC().lbl,fontSize:10}},AXS({splitLine:{show:false}})),
    visualMap:{min:0,max:vmax,calculable:false,orient:'horizontal',left:'center',bottom:0,
      textStyle:{color:TC().axis,fontSize:10},
      inRange:{color:['#EFE7DB','#C9973F','#B3613C']}},
    series:[{type:'heatmap',data,label:{show:true,formatter:p=>p.value[2]>=2?p.value[2]:'',color:'#FFF6EC',fontSize:9},
      itemStyle:{borderColor:TC().pieBd,borderWidth:1}}]
  },true);
}
function chNewEnt(rows){
  if(!document.getElementById('chNewEnt'))return;
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  const firstMap = new Map();
  bs.forEach((x,i)=>{ x[1].forEach(r=>{ const e=r[I.e]; if(e&&!firstMap.has(e)) firstMap.set(e,i); }); });
  const newArr = labels.map(()=>0);
  firstMap.forEach(i=>{ newArr[i]++; });
  let acc=0; const cumArr = newArr.map(v=>(acc+=v));
  setInsight('chNewEnt', `区间新进入企业 <em>${fmt(newArr.reduce((a,b)=>a+b,0))}</em> 家 · 累计在营 <em>${fmt(cumArr.length?cumArr[cumArr.length-1]:0)}</em> 家 · 最新批次 +${newArr.length?newArr[newArr.length-1]:0} 家`);
  chart('chNewEnt').setOption({
    tooltip:Object.assign({trigger:'axis',formatter(ps){
      const i=ps[0].dataIndex;
      return `<b>第 ${labels[i]} 批</b><br>新进入企业：<b>${newArr[i]}</b> 家<br>累计在营企业：<b>${cumArr[i]}</b> 家`;}},TT),
    legend:LG({data:['新进入企业数','累计在营企业数']}),
    grid:GRID({bottom:64,top:46}),
    xAxis:Object.assign({type:'category',data:labels,name:'批次',nameLocation:'middle',nameGap:44,nameTextStyle:{color:TC().axis}},AXS({axisLabel:Object.assign({},AXS().axisLabel,{rotate:50,fontSize:10})})),
    yAxis:[
      Object.assign({type:'value',name:'新进入（家）',nameTextStyle:{color:TC().axis}},AXS()),
      Object.assign({type:'value',name:'累计（家）',nameTextStyle:{color:TC().axis}},AXS({splitLine:{show:false}}))],
    series:[
      {name:'新进入企业数',type:'bar',data:newArr,itemStyle:{color:'#5BAFAF'},barMaxWidth:16},
      {name:'累计在营企业数',type:'line',yAxisIndex:1,data:cumArr,smooth:true,symbol:'none',
        lineStyle:{color:'#D4A94E',width:2}}
    ]},true);
}
function chBrandTop(rows){
  if(!document.getElementById('chBrandTop'))return;
  const es = groupCount(rows,I.bd,15);
  const filled = es.reduce((a,b)=>a+b[1],0);
  setInsight('chBrandTop', filled?`首位品牌 <em>${es[0][0]}</em>（<em>${fmt(es[0][1])}</em> 款，<em>${p1(es[0][1],filled)}%</em>）· 入榜 15 强合计 <em>${p1(filled,rows.filter(r=>r[I.bd]!=null).length)}%</em>（商标填报口径）`:'暂无商标填报');
  hbar('chBrandTop', es, '#B08968', '款', '产品商标');
}

function insightEnt(rows){
  const el=document.getElementById('insightEntTx'); if(!el)return;
  const ents=[...new Set(rows.map(r=>r[I.e]).filter(Boolean))];
  const brands=[...new Set(rows.map(r=>r[I.bd]).filter(Boolean))];
  const top=groupCount(rows,I.e,1)[0];
  const bs=byBatch(rows); const last=bs[bs.length-1];
  let newLast=null;
  if(last&&bs.length>1){
    const prev=new Set(); bs.slice(0,-1).forEach(x=>x[1].forEach(r=>{if(r[I.e])prev.add(r[I.e]);}));
    newLast=[...new Set(last[1].map(r=>r[I.e]).filter(Boolean))].filter(e=>!prev.has(e)).length;
  }
  el.innerHTML =
    `当前筛选覆盖 <em>${ents.length}</em> 家企业、<em>${brands.length}</em> 个品牌。`+
    (top?`头部企业 <em>${top[0]}</em> 公告 <em>${top[1]}</em> 款（占 <em>${rows.length?pct(top[1],rows.length).toFixed(1):0}%</em>）。`:'')+
    (newLast!=null?`最新批次（第 <em>${last[0]}</em> 批）新进入企业 <em>${newLast}</em> 家。`:'');
}

/* ==================== 车型明细查询表 ==================== */
const TBL_COLS = [
  {n:'批次',i:I.b},{n:'企业名称',i:I.e},{n:'商标',i:I.bd},{n:'通用名称',i:I.gn},{n:'产品型号',i:I.m},
  {n:'动力类型',i:I.t},{n:'细分市场',i:I.s},{n:'续航km',i:I.r},{n:'容量kWh',i:I.c},
  {n:'电池类型',i:I.bt},{n:'前电机kW',i:I.fp},{n:'后电机kW',i:I.rp},{n:'总功率kW',i:I.tp},{n:'驱动',i:I.drive},{n:'轴距mm',i:I.ab}
];
const DRIVE_TXT = d => d===1?'四驱':(d===0?'两驱':'');
const tblState={q:'',k:I.b,d:-1,page:0,per:20};
const escH=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;');
function tblData(){
  const q=tblState.q.trim().toLowerCase();
  let rs=q?filtered().filter(r=>[I.e,I.bd,I.m].some(i=>r[i]&&String(r[i]).toLowerCase().includes(q))):filtered();
  rs=[...rs].sort((a,b)=>{
    const va=a[tblState.k],vb=b[tblState.k];
    if(va==null&&vb==null)return 0;
    if(va==null)return 1; if(vb==null)return -1;
    if(typeof va==='number'&&typeof vb==='number')return(va-vb)*tblState.d;
    return String(va).localeCompare(String(vb),'zh-Hans-CN')*tblState.d;
  });
  return rs;
}
function renderTable(resetPage){
  const body=document.querySelector('#tblDetail tbody'), head=document.querySelector('#tblDetail thead tr');
  if(!body||!head)return;
  head.innerHTML=TBL_COLS.map(c=>
    `<th data-i="${c.i}" class="${c.i===tblState.k?'on':''}">${c.n}${c.i===tblState.k?(tblState.d<0?' ↓':' ↑'):''}</th>`).join('');
  head.querySelectorAll('th').forEach(th=>th.onclick=()=>{
    const k=+th.dataset.i;
    if(tblState.k===k)tblState.d=-tblState.d; else{tblState.k=k;tblState.d=-1;}
    renderTable(true);
  });
  const rs=tblData();
  const pages=Math.max(1,Math.ceil(rs.length/tblState.per));
  if(resetPage||tblState.page>=pages)tblState.page=0;
  const pg=Math.min(tblState.page,pages-1);
  const slice=rs.slice(pg*tblState.per,(pg+1)*tblState.per);
  body.innerHTML=slice.length?slice.map(r=>'<tr>'+TBL_COLS.map(c=>{
      let v=r[c.i];
      if(c.i===I.t&&v==='PHEV')v='PHEV/EREV';   // 展示层统一 PHEV/EREV
      if(c.i===I.drive)v=DRIVE_TXT(v);          // 驱动形式 0/1 → 两驱/四驱
      if(v==null||v==='')return '<td class="na">–</td>';
      if(typeof v==='number')return '<td>'+v.toLocaleString('zh-CN')+'</td>';
      return '<td title="'+escH(v)+'">'+escH(v)+'</td>';
    }).join('')+'</tr>').join('')
    :`<tr><td class="na" colspan="${TBL_COLS.length}">无匹配记录</td></tr>`;
  document.getElementById('tblInfo').innerHTML=
    `命中 <b>${rs.length.toLocaleString('zh-CN')}</b> 条 · 第 ${pg+1} / ${pages} 页`;
  document.getElementById('tblPrev').disabled=pg<=0;
  document.getElementById('tblNext').disabled=pg>=pages-1;
}
function initTable(){
  const inp=document.getElementById('tblSearch');
  if(!inp)return;
  let tm=null;
  inp.oninput=()=>{clearTimeout(tm);tm=setTimeout(()=>{tblState.q=inp.value;renderTable(true);},200);};
  document.getElementById('tblPrev').onclick=()=>{tblState.page--;renderTable(false);};
  document.getElementById('tblNext').onclick=()=>{tblState.page++;renderTable(false);};
}

/* ==================== 数据管理：导出 ==================== */
/* 导出字段清单不含「是否减免购置税」（v3.7 展示层口径：底表保留该维度） */
const EXPORT_FIELDS = COL_NAMES.map((n,i)=>({name:n,idx:i,on:true})).filter(f=>f.idx!==TX_IDX);
function initExportUI(){
  document.getElementById('fldList').innerHTML = EXPORT_FIELDS.map((f,i)=>
    `<label class="fld-chip${f.on?'':' off'}"><input type="checkbox" data-i="${i}" ${f.on?'checked':''}>${f.name}</label>`).join('');
  document.querySelectorAll('#fldList input').forEach(cb=>{
    cb.onchange = ()=>{
      EXPORT_FIELDS[+cb.dataset.i].on = cb.checked;
      cb.parentElement.classList.toggle('off',!cb.checked);
    };
  });
  document.getElementById('btnXlsx').onclick = exportXlsx;
  document.getElementById('btnCsv').onclick = exportCsv;
}
function exportRows(){
  const cols = EXPORT_FIELDS.filter(f=>f.on);
  const header = cols.map(c=>c.name);
  /* 导出口径：动力类型列展示 PHEV/EREV；驱动形式列 0/1 → 两驱/四驱 */
  const data = filtered().map(r=>cols.map(c=>{
    let v=r[c.idx];
    if(c.idx===I.t&&v==='PHEV')v='PHEV/EREV';
    if(c.idx===I.drive)v=DRIVE_TXT(v)||v;
    return v;
  }));
  return {header, data};
}
function updateDmStat(){
  const rows = filtered();
  const n = rows.length;
  const cols = EXPORT_FIELDS.filter(f=>f.on).length;
  document.getElementById('dmStat').innerHTML = `
    <div class="s"><div class="v">${n.toLocaleString('zh-CN')}</div><div class="l">待导出记录数</div></div>
    <div class="s"><div class="v">${cols}</div><div class="l">已选字段数</div></div>
    <div class="s"><div class="v">${state.bFrom}~${state.bTo}</div><div class="l">批次范围</div></div>
    <div class="s"><div class="v">${state.type==='ALL'?'全部':(state.type==='PHEV'?'PHEV/EREV':state.type)}</div><div class="l">动力类型</div></div>`;
}
function stamp(){ const d=new Date(); const p=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`; }
function exportXlsx(){
  const {header,data} = exportRows();
  if(!data.length){ alert('当前筛选无数据可导出'); return; }
  const ws = XLSX.utils.aoa_to_sheet([header,...data]);
  ws['!cols'] = header.map(h=>({wch:Math.max(10,String(h).length*1.8+4)}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'NEV筛选数据');
  XLSX.writeFile(wb, `NEV公告筛选数据_${stamp()}.xlsx`);
}
function exportCsv(){
  const {header,data} = exportRows();
  if(!data.length){ alert('当前筛选无数据可导出'); return; }
  const esc = v=>{ if(v==null) return ''; const s=String(v); return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s; };
  const csv = '\ufeff'+[header,...data].map(r=>r.map(esc).join(',')).join('\r\n');
  const blob = new Blob([csv],{type:'text/csv;charset=utf-8'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `NEV公告筛选数据_${stamp()}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ==================== 数据管理：导入 ==================== */
const NUM_RANGES_JS = {
  w:[[300,4500],[500,4500]], r:[[50,1500],[20,600]], c:[[2,250],[1,120]],
  ed:[[50,400],[30,400]], ec:[[3,40],[3,40]], pp:[[5,1500],[5,800]], tp:[[5,1500],[5,800]],
  tq:[[20,30000],[20,30000]], fo:[null,[0.1,20]], dv:[null,[300,6000]], ep:[null,[10,600]],
  ab:[[1800,4200],[1800,4200]], lg:[[2600,6600],[2600,6600]]
};
function jsCleanNum(v){
  if(v==null) return null;
  if(typeof v==='number') return isFinite(v)?v:null;
  const s=String(v).trim();
  if(!s||s.includes('不适用')||['-','—','/','None','nan'].includes(s)) return null;
  const n=parseFloat(s);
  return isFinite(n)?n:null;
}
function jsCleanStr(v){
  if(v==null) return null;
  if(typeof v==='number') return String(v);
  const s=String(v).trim();
  if(!s||s.includes('不适用')) return null;
  return s;
}
function jsTorque(v){
  if(v==null) return null;
  const s=String(v).trim();
  if(!s.includes('/')) return null;
  return jsCleanNum(s.split('/').pop());
}
/* 是否减免购置税：仅接受 是/否（与生成脚本 norm_tax 同口径） */
function jsNormTax(v){
  const s=jsCleanStr(v);
  return (s==='是'||s==='否')?s:null;
}
/* 细分市场：'/'占位符视为未填报 */
function jsCleanSeg(v){
  const s=jsCleanStr(v);
  if(s==null) return null;
  if(!s.replace(/\//g,'').trim()||['-','—','N/A'].includes(s)) return null;
  return s;
}
/* 电池类型归一化（与生成脚本同口径） */
function jsNormBt(v){
  const s=jsCleanStr(v);
  if(s==null||['未提供','未提供具体数据','无','/'].includes(s)) return null;
  const t=s.toUpperCase();
  if(s.includes('钠')) return '钠离子';
  if(s.includes('钛酸锂')) return '钛酸锂';
  if(s.includes('锰酸锂')&&!s.includes('三元')) return '锰酸锂';
  if(s.includes('磷酸铁锂')||t.includes('LFP')) return '磷酸铁锂';
  if(s.includes('三元')||s.includes('镍钴锰')||t.includes('NCM')) return '三元锂';
  if(s.includes('锂')) return '其他锂离子';
  return '其他';
}
function jsSanitize(k,val,ptype){
  if(val==null) return null;
  const pair = NUM_RANGES_JS[k];
  if(!pair) return val;
  const rng = ptype==='BEV'?pair[0]:pair[1];
  if(!rng) return null;
  return (rng[0]<=val&&val<=rng[1])?val:null;
}
function parseImportedSheet(aoa, fname){
  /* v4.9.27：表头名寻址（导入文件=权威底表 39 列同构），输出 30 位记录与 RAW_INIT 同构。
     旧版按 v4.1-era 25 位硬编码下标解析，导入后与内嵌 30 位数据错位（09-16 遗留，本版重构）。 */
  const records=[]; let cleaned=0;
  if(!aoa.length) return {records,cleaned};
  // 表头行探测：含「批次」且含「产品型号」的首行
  let hrow=-1;
  for(let i=0;i<Math.min(aoa.length,5);i++){
    const cells=(aoa[i]||[]).map(c=>String(c==null?'':c).trim());
    if(cells.includes('批次')&&cells.includes('产品型号')){ hrow=i; break; }
  }
  if(hrow<0) throw new Error('表头未识别（需含 批次/产品型号 列）');
  const H={};
  aoa[hrow].forEach((c,i)=>{ const k=String(c==null?'':c).trim(); if(k&&!(k in H)) H[k]=i; });
  const g=(row,name)=>{ const i=H[name]; return i==null?null:row[i]; };
  const ptOf=v=>{ const s=jsCleanStr(v)||''; return s||'未知'; };
  const kwFirst=v=>{ const s=String(v==null?'':v).trim(); if(!s) return null;
    const m=s.split('/')[0].match(/^[-+]?\d+(?:\.\d+)?/); return m?parseFloat(m[0]):null; };
  const kwSecond=v=>{ const s=String(v==null?'':v).trim(); if(!s||s.indexOf('/')<0) return null;
    const m=s.split('/')[1].match(/^[-+]?\d+(?:\.\d+)?/); return m?parseFloat(m[0]):null; };
  for(let i=hrow+1;i<aoa.length;i++){
    const row=aoa[i];
    if(!row||row[0]==null) continue;
    const batch=parseInt(String(row[0]).trim(),10);
    if(!isFinite(batch)) continue;
    const ptype=ptOf(g(row,'动力类型'));
    const raw={w:jsCleanNum(g(row,'整备质量(kg)')),ab:jsCleanNum(g(row,'轴距(mm)')),r:jsCleanNum(g(row,'纯电续航里程(km)')),
      c:jsCleanNum(g(row,'电池容量(kWh)')),ed:jsCleanNum(g(row,'电池能量密度(Wh/kg)')),ec:jsCleanNum(g(row,'百公里电耗(kWh/100km)')),
      fp:kwFirst(g(row,'前电机功率/扭矩')),tp:kwFirst(g(row,'电机总功率/扭矩')),tq:kwSecond(g(row,'电机总功率/扭矩')),
      rp:kwFirst(g(row,'后电机功率/扭矩')),rt2:kwSecond(g(row,'后电机功率/扭矩')),
      fo:jsCleanNum(g(row,'综合油耗(L/100km)')),dv:jsCleanNum(g(row,'发动机排量(mL)')),ep:jsCleanNum(g(row,'发动机功率(kW)')),
      lg:(H['车长(mm)']!=null?jsCleanNum(g(row,'车长(mm)')):null),sn:(H['月销量(辆)']!=null?jsCleanNum(g(row,'月销量(辆)')):null)};
    const val={};
    for(const k in raw){
      const s=(k==='lg'||k==='sn'||k==='fp'||k==='tp'||k==='tq'||k==='rp'||k==='rt2')?(raw[k]==null?null:raw[k]):jsSanitize(k,raw[k],ptype);
      if(raw[k]!=null&&s==null) cleaned++;
      val[k]=s;
    }
    // 驱动形式派生：后电机有值 或 总>前 ⇒ 四驱；否则两驱
    let drive='两驱';
    if(val.rp!=null||(val.tp!=null&&val.fp!=null&&val.tp>val.fp)) drive='四驱';
    records.push([batch,ptype,jsCleanSeg(g(row,'细分市场')),jsCleanStr(g(row,'企业名称')),val.w,val.r,val.c,jsNormBt(g(row,'电池类型')),
      val.ed,val.ec,val.fp,val.tp,jsCleanStr(g(row,'电机生产企业')),val.tq,val.fo,val.dv,val.ep,jsCleanStr(g(row,'发动机生产企业')),jsCleanStr(g(row,'数据来源')),
      jsCleanStr(g(row,'产品型号')),jsCleanStr(g(row,'产品商标')),jsNormTax(g(row,'是否减免购置税')),val.ab,val.lg,jsCleanStr(g(row,'通用名称')),
      val.rt2,val.rp,drive,jsCleanStr(g(row,'产品名称')),val.sn,
      /* 31st 核验复合字段（与生成器 rec.append(f'{ann_body}|{raw_pn}') 同构） */
      jsCleanStr(g(row,'动力类型'))+'|'+jsCleanStr(g(row,'产品名称'))]);
  }
  return {records,cleaned};
}
function applyNewData(records,fname){
  RAW = records;
  const batches=[...new Set(records.map(r=>r[0]))].sort((a,b)=>a-b);
  META_CUR = {
    total:records.length, batches,
    batchMin:batches[0]||0, batchMax:batches[batches.length-1]||0, batchCount:batches.length,
    segments:[...new Set(records.map(r=>r[2]).filter(Boolean))].sort((a,b)=>{
      const ka=segKey(a), kb=segKey(b);
      return ka[0]-kb[0]||ka[1]-kb[1]||String(ka[2]).localeCompare(String(kb[2]),'zh-Hans-CN');
    }),
    types:[...new Set(records.map(r=>r[1]))].sort(),
    generatedAt:new Date().toLocaleString('zh-CN'), sourceFile:fname
  };
  // 刷新页头
  document.getElementById('hSrc').textContent = fname;
  document.getElementById('hBRange').textContent = `${META_CUR.batchMin} ~ ${META_CUR.batchMax}`;
  document.getElementById('hBCnt').textContent = META_CUR.batchCount;
  document.getElementById('hGen').textContent = META_CUR.generatedAt;
  // 重置筛选并重建选项（initFilters 内会按新数据重建细分市场并全选）
  state.bFrom=META_CUR.batchMin; state.bTo=META_CUR.batchMax; state.type='ALL';
  initFilters();
  updateAll();
}
function handleFile(file){
  const st = document.getElementById('importStatus');
  st.textContent = `正在解析 ${file.name} …`;
  const reader = new FileReader();
  reader.onload = e => {
    try{
      const wb = XLSX.read(new Uint8Array(e.target.result),{type:'array'});
      const ws = wb.Sheets['NEV公告参数汇总']||wb.Sheets[wb.SheetNames[0]];
      if(!ws) throw new Error('未找到可用工作表');
      const aoa = XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
      const {records,cleaned} = parseImportedSheet(aoa,file.name);
      if(!records.length) throw new Error('未解析到有效数据行，请确认表结构与原始汇总表一致');
      applyNewData(records,file.name);
      st.innerHTML = `✓ 导入成功：<b>${file.name}</b> · ${records.length.toLocaleString('zh-CN')} 条记录`+
        (cleaned?` · 剔除异常值 ${cleaned} 个`:'')+` · 全局图表已刷新`;
      st.style.color = 'var(--sage)';
    }catch(err){
      st.textContent = '✗ 导入失败：'+err.message;
      st.style.color = 'var(--rose)';
    }
  };
  reader.readAsArrayBuffer(file);
}
function initImport(){
  const dz=document.getElementById('dropzone'), fi=document.getElementById('fileInput');
  dz.onclick = ()=>fi.click();
  fi.onchange = ()=>{ if(fi.files[0]) handleFile(fi.files[0]); fi.value=''; };
  ['dragover','dragenter'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('over');}));
  ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('over');}));
  dz.addEventListener('drop',e=>{ if(e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  document.getElementById('btnImport').onclick = ()=>{
    document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
    document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
    document.querySelector('[data-p="p6"]').classList.add('on');
    document.getElementById('p6').classList.add('on');
    setTimeout(()=>{Object.values(CHARTS).forEach(c=>{try{c.resize();}catch(e){}});},60);
    document.getElementById('p6').scrollIntoView({behavior:'smooth',block:'start'});
  };
}

/* ==================== 刷新 ==================== */
function updateAll(){
  resetGap();
  const rows = filtered();
  document.getElementById('fCount').innerHTML = `已选 <b>${rows.length.toLocaleString('zh-CN')}</b> / ${META_CUR.total.toLocaleString('zh-CN')} 条`;
  updateKPIs(rows);
  insight1(rows); chBatch(rows); chType(rows); chSeg(rows); chEnt(rows); chEvo(rows);
  insight2(rows); chPwDist(rows); chMsTop(rows); chPtq(rows); chPwTrend(rows); chPwr(rows);
  insight3(rows); chDv(rows); chDvEp(rows); chEsTop(rows); chFo(rows); chDvTrend(rows);
  insight4(rows); chRgB(rows); chRgP(rows); chCr(rows); chBtTrend(rows); chEd(rows); chEcTrend(rows); chEc(rows);
  insightD(rows); renderDists(rows); renderCustom(rows);
  insightEnt(rows); chEntTrend(rows); chEntSeg(rows); chNewEnt(rows); chBrandTop(rows);
  insight5(rows); chFill(rows); chSrc(rows); chHeat(rows);
  updateDmStat(); renderTable(false);
  ['chPtq','chDvEp','chCr'].forEach(bindTrendClick);
  updateGapBtn();
}

/* ==================== 筛选器 ==================== */
function rebuildSegPanel(){
  const segList = document.getElementById('segList');
  // 各细分市场记录计数（只统计已规范的具体值）
  const cnt = new Map();
  for(const r of RAW){ if(r[I.s]!=null) cnt.set(r[I.s],(cnt.get(r[I.s])||0)+1); }
  const items = [...cnt.entries()];
  items.sort((a,b)=>{
    const ka=segKey(a[0]), kb=segKey(b[0]);
    if(ka[0]!==kb[0]) return ka[0]-kb[0];
    if(ka[1]!==kb[1]) return ka[1]-kb[1];
    return b[1]-a[1];
  });
  // 初始：全部勾选（segAll=true → 不过滤）
  state.segAll = true; state.segs = new Set(items.map(x=>x[0]));
  segList.innerHTML = items.map(([s,c])=>
    `<label class="ms-item"><input type="checkbox" value="${s}" checked><span>${s}</span><span class="c">${c}</span></label>`).join('');
  segList.querySelectorAll('input').forEach(cb=>{
    cb.onchange = ()=>{
      if(cb.checked) state.segs.add(cb.value); else state.segs.delete(cb.value);
      state.segAll = state.segs.size === items.length;
      syncSegBtn();
      updateAll();
    };
  });
  syncSegBtn();
}
function setAllSegs(on){
  const segList = document.getElementById('segList');
  segList.querySelectorAll('input').forEach(cb=>{ cb.checked = on; });
  state.segs = new Set(on ? [...segList.querySelectorAll('input')].map(cb=>cb.value) : []);
  state.segAll = on;
  syncSegBtn();
  updateAll();
}
function syncSegBtn(){
  const segList = document.getElementById('segList');
  const boxes = [...segList.querySelectorAll('input')];
  const n = state.segs.size, total = boxes.length;
  const btn = document.getElementById('segBtn');
  if(n===0) btn.innerHTML = `细分市场：<span class="n">未选择</span><span class="caret">▾</span>`;
  else if(state.segAll) btn.innerHTML = `全部细分市场<span class="n">(${total})</span><span class="caret">▾</span>`;
  else btn.innerHTML = `细分市场：<span class="n">${n}</span>/${total}<span class="caret">▾</span>`;
  // state.segs 语义：已勾选（纳入统计）的细分市场
}
function initFilters(){
  const bF = document.getElementById('bFrom'), bT = document.getElementById('bTo');
  bF.innerHTML=''; bT.innerHTML='';
  META_CUR.batches.forEach(b=>{
    bF.insertAdjacentHTML('beforeend',`<option value="${b}">${b}</option>`);
    bT.insertAdjacentHTML('beforeend',`<option value="${b}">${b}</option>`);
  });
  bF.value = state.bFrom; bT.value = state.bTo;
  bF.onchange = ()=>{ state.bFrom = Math.min(+bF.value,state.bTo); bF.value=state.bFrom; updateAll(); };
  bT.onchange = ()=>{ state.bTo = Math.max(+bT.value,state.bFrom); bT.value=state.bTo; updateAll(); };

  document.querySelectorAll('.quick button').forEach(btn=>{
    btn.onclick = ()=>{
      const n = META_CUR.batches.length;
      const q = btn.dataset.q;
      if(q==='all'){ state.bFrom=META_CUR.batches[0]; state.bTo=META_CUR.batches[n-1]; }
      else if(q==='latest'){ state.bFrom=META_CUR.batches[n-1]; state.bTo=META_CUR.batches[n-1]; }
      else { const k=+q.replace('last',''); state.bFrom=META_CUR.batches[Math.max(0,n-k)]; state.bTo=META_CUR.batches[n-1]; }
      bF.value=state.bFrom; bT.value=state.bTo; updateAll();
    };
  });

  document.querySelectorAll('#typeGroup button').forEach(btn=>{
    btn.classList.toggle('on', btn.dataset.t===state.type);
    btn.onclick = ()=>{
      document.querySelectorAll('#typeGroup button').forEach(b=>b.classList.remove('on'));
      btn.classList.add('on');
      state.type = btn.dataset.t;
      updateAll();
    };
  });

  rebuildSegPanel();

  // 多选下拉开合
  const sBtn=document.getElementById('segBtn'), sPanel=document.getElementById('segPanel');
  sBtn.onclick = e=>{ e.stopPropagation(); sBtn.classList.toggle('open'); sPanel.classList.toggle('open'); };
  sPanel.onclick = e=>e.stopPropagation();
  document.addEventListener('click',()=>{ sBtn.classList.remove('open'); sPanel.classList.remove('open'); });
  document.getElementById('segSelAll').onclick = ()=>setAllSegs(true);
  document.getElementById('segClrAll').onclick = ()=>setAllSegs(false);

  const sF = document.getElementById('salesF');
  sF.value = state.sales;
  sF.onchange = ()=>{ state.sales = sF.value; updateAll(); };

  document.getElementById('fReset').onclick = ()=>{
    state.bFrom=META_CUR.batches[0]; state.bTo=META_CUR.batches[META_CUR.batches.length-1];
    state.type='ALL'; state.sales='ALL';
    bF.value=state.bFrom; bT.value=state.bTo; sF.value='ALL';
    document.querySelectorAll('#typeGroup button').forEach(b=>b.classList.toggle('on',b.dataset.t==='ALL'));
    setAllSegs(true);
  };
}

/* ==================== Tabs / 折叠 / 主题 / 回顶 ==================== */
function initTabs(){
  document.querySelectorAll('.tab').forEach(tab=>{
    tab.onclick = ()=>{
      document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
      document.querySelectorAll('.panel').forEach(p=>p.classList.remove('on'));
      tab.classList.add('on');
      document.getElementById(tab.dataset.p).classList.add('on');
      setTimeout(()=>{ Object.values(CHARTS).forEach(c=>{try{c.resize();}catch(e){}}); updateAll(); },50);
      setTimeout(()=>{ Object.values(CHARTS).forEach(c=>{try{c.resize();}catch(e){}}); },350);
    };
  });
}
function initQuality(){
  const head=document.getElementById('qHead'), body=document.getElementById('qBody');
  if(!head||!body)return;
  head.onclick = ()=>{
    head.classList.toggle('open');
    body.classList.toggle('open');
    setTimeout(()=>{ Object.values(CHARTS).forEach(c=>{try{c.resize();}catch(e){}}); updateAll(); },60);
  };
}
function initToTop(){
  const btn=document.getElementById('toTop');
  window.addEventListener('scroll',()=>{ btn.classList.toggle('show',window.scrollY>600); });
  btn.onclick = ()=>window.scrollTo({top:0,behavior:'smooth'});
}

/* ==================== 启动 ==================== */
initKPIs();
initFilters();
initTabs();
initQuality();
initExportUI();
initImport();
initDist();
initCustomDist();
initCardInsights();
initGapBtn();
initToTop();
initTable();
document.getElementById('btnTheme').onclick = toggleTheme;
updateAll();
</script>
</body>
</html>'''


def main():
    pos_args = [a for a in sys.argv[1:] if not a.startswith('--')]
    release = '--release' in sys.argv
    in_path = pos_args[0] if len(pos_args) > 0 else DEFAULT_IN
    out_path = pos_args[1] if len(pos_args) > 1 else (DEFAULT_RELEASE_OUT if release else DEFAULT_OUT)

    for p, tip in [(in_path, '输入文件'), (ECHARTS_PATH, 'echarts.min.js'),
                   (ECHARTSGL_PATH, 'echarts-gl.min.js'), (XLSXLIB_PATH, 'xlsx.full.min.js')]:
        if not os.path.exists(p):
            print(f'[错误] 找不到{tip}: {p}')
            sys.exit(1)

    print(f'[1/4] 读取数据: {in_path}')
    records, skipped, cleaned, seg_stat, seg_unresolved = load_records(in_path)
    print(f'      共 {len(records)} 条记录' + (f'（跳过 {skipped} 条无效行）' if skipped else ''))
    print(f'      剔除物理范围外异常值 {cleaned} 个（已置空，避免污染统计）')
    # 细分市场规范化结果
    SEG_STAT_LABEL = {
        'keep': '已是具体值', 'name-history': '同名称历史', 'wheelbase': '轴距分级',
        'name-wheelbase': '同名称轴距', 'lux': 'Lux 保留类别', 'length': '车长兜底',
        'body-only': '仅车身形式(未分级)', 'unresolved': '无法落位(置空)',
    }
    print('      细分市场规范化：' + ' · '.join(
        f'{SEG_STAT_LABEL.get(k, k)} {v}' for k, v in
        sorted(seg_stat.items(), key=lambda x: -x[1])))
    if seg_unresolved:
        print(f'      未能落位 {len(seg_unresolved)} 条（无尺寸且无同名称参照，置空不臆测）')

    print('[2/4] 构建数据与元信息...')
    meta = build_meta(records, in_path)
    data_json = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    with open(ECHARTS_PATH, 'r', encoding='utf-8') as f:
        echarts_lib = f.read()
    with open(XLSXLIB_PATH, 'r', encoding='utf-8') as f:
        xlsx_lib = f.read()
    for lib in (echarts_lib, xlsx_lib):
        pass
    with open(ECHARTSGL_PATH, 'r', encoding='utf-8') as f:
        gl_lib = f.read()
    echarts_lib = echarts_lib.replace('</script', '<\\/script')
    gl_lib = gl_lib.replace('</script', '<\\/script')
    xlsx_lib = xlsx_lib.replace('</script', '<\\/script')

    print('[3/4] 生成看板HTML...')
    html = HTML_HEAD
    if release:
        # 发布版：整体移除【数据质量检查】区块（独立 <section>，无导航入口）
        html = re.sub(r'<!-- ============ 数据质量（页面底部） ============ -->.*?</section>\n\n', '', html, flags=re.S)
        # 发布版：取消【导出当前筛选数据】与【导入最新数据表】（含工具栏导入按钮），保留车型明细查询。
        # 用 CSS 隐藏而非删除 DOM，数据管理页的 JS 绑定无需判空。
        html = html.replace(
            '<div class="tx">导出当前筛选数据供进一步细化分析；导入最新公告数据表后，<b>全局数据与图表将实时刷新</b>（无需重新运行生成脚本）。</div>',
            '<div class="tx">车型明细查询支持在当前筛选范围内检索、排序与翻页。</div>')
        html = html.replace('</head>',
                            '<style>#p6 .dm-grid .card.c-s7,#p6 .dm-grid .card.c-s5,#btnImport{display:none!important}</style></head>')
    html = html.replace('__ECHARTS_LIB__', echarts_lib)
    html = html.replace('__ECHARTSGL_LIB__', gl_lib)
    html = html.replace('__XLSX_LIB__', xlsx_lib)
    html = html.replace('__META_JSON__', json.dumps(meta, ensure_ascii=False, separators=(',', ':')))
    html = html.replace('__DATA_JSON__', data_json)
    html = html.replace('__SRC_FILE__', meta['sourceFile'])
    html = html.replace('__TOTAL__', f"{meta['total']:,}")
    html = html.replace('__BMIN__', str(meta['batchMin']))
    html = html.replace('__BMAX__', str(meta['batchMax']))
    html = html.replace('__BCNT__', str(meta['batchCount']))
    html = html.replace('__GEN_AT__', meta['generatedAt'])
    html = html.replace('__COPYRIGHT__', COPYRIGHT)
    html = html.replace('__REL_NAME__', ' · 发布版' if release else '')
    html = html.replace('__TITLE_SUFFIX__', ' · 发布版' if release else '')
    rel_tag = (' <span class="tag" style="background:rgba(212,169,78,.14);'
               'border-color:rgba(212,169,78,.5);color:#D4A94E;margin-left:6px">RELEASE 发布版</span>') if release else ''
    html = html.replace('__REL_TAG__', rel_tag)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_mb = os.path.getsize(out_path) / 1024 / 1024
    print(f'\n✓ 看板已生成: {out_path}（{size_mb:.1f} MB）')
    print(f'  批次范围: {meta["batchMin"]}~{meta["batchMax"]} · 动力类型: {"/".join(meta["types"])}')
    if release:
        print('  发布版：隐藏【数据质量检查】区块 · 取消导出/导入功能（保留车型明细查询）')
    print('  v3.7：新增【分布格局】页（EV Range vs Length 等五组散点，DIST_DEFS 配置可增删）·')
    print('        全图表坐标轴标注名称与单位 · 展示层移除免征购置税信息 · 底表新增车长(mm)占位列')
    print('  v3.5：企业与品牌分析页（企业演进/企业×细分市场/新进入企业/品牌Top）')
    print('        车型明细查询表（搜索/排序/翻页）· 记录扩展 产品型号/商标/免购置税 字段')
    print('  v3：PHEV/EREV 展示名统一 / 细分市场"/"不显示 / 电池类型归一化 / 数据质量底部化')
    print('          导出Excel·CSV / 页内导入刷新 / 白天黑夜模式')
    print('  双击HTML文件即可在浏览器中离线打开使用')


if __name__ == '__main__':
    main()
