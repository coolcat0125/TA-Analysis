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
  python3 generate_dashboard.py --release                # 离线发布版（隐藏数据质量区块，含版权页脚）

v3.7.1 更新：
  1. 分布格局五图标题中文化（英文名保留为副标题注记）
  2. 细分市场下拉按 Car/SUV/MPV 分组、字母级别→数字级别→未分级排序（Python/JS 双侧同键）
  3. 全部图表卡片新增「一句话核心要点」行（.ci），随全局筛选实时联动

v3.7.2 更新：
  1. 所有散点图悬浮窗补上车型【通用名称】（记录新增 gn 字段，页内导入同步）
  2. 车型明细查询表新增【通用名称】列；导出字段同步
  3. 坐标轴标签完整显示：网格自适应留白（containLabel），取消类目标签截断

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
DEFAULT_RELEASE_OUT = os.path.join(HERE, 'NEV公告数据看板_离线发布版.html')
COPYRIGHT = 'Copyright (c) 2026 David YEAH'
ECHARTS_PATH = os.path.join(HERE, 'echarts.min.js')
XLSXLIB_PATH = os.path.join(HERE, 'xlsx.full.min.js')
SHEET_NAME = 'NEV公告参数汇总'

FIELDS = ['b', 't', 's', 'e', 'w', 'r', 'c', 'bt', 'ed', 'ec',
          'pp', 'tp', 'ms', 'tq', 'fo', 'dv', 'ep', 'es', 'src',
          'm', 'bd', 'tx', 'ab', 'lg', 'gn']


def clean_num(v):
    """数值清洗：'不适用(BEV)'、空值、异常文本 → None"""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if '不适用' in s or s in ('-', '—', '/', 'None', 'nan'):
        return None
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


def clean_seg(v):
    """细分市场清洗：'/' 等占位符视为未填报，不计入统计与分类展示"""
    s = clean_str(v)
    if s is None or s.strip('/') == '' or s in ('-', '—', 'N/A'):
        return None
    return s


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
    'pp': ((5, 1500), (5, 800)),         # 电机峰值功率 kW
    'tp': ((5, 1500), (5, 800)),         # 电机总功率 kW
    'tq': ((20, 30000), (20, 30000)),    # 扭矩 Nm
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
    """读取Excel并转为紧凑记录数组"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    records = []
    skipped = 0
    cleaned = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            batch = int(str(row[0]).strip())
        except (ValueError, TypeError):
            skipped += 1
            continue
        ptype = clean_str(row[9]) or '未知'
        raw = {
            'w':  clean_num(row[11]),    # 整备质量kg
            'ab': clean_num(row[12]),    # 轴距mm
            'r':  clean_num(row[13]),    # 纯电续航km
            'c':  clean_num(row[14]),    # 电池容量kWh
            'ed': clean_num(row[16]),    # 能量密度Wh/kg
            'ec': clean_num(row[17]),    # 百公里电耗
            'pp': clean_num(row[18]),    # 电机峰值功率kW
            'tp': clean_num(row[19]),    # 电机总功率kW
            'tq': parse_torque(row[21]), # 扭矩Nm
            'fo': clean_num(row[23]),    # 综合油耗
            'dv': clean_num(row[25]),    # 排量mL
            'ep': clean_num(row[26]),    # 发动机功率kW
        }
        val = {}
        for k, v in raw.items():
            s = sanitize(k, v, ptype)
            if v is not None and s is None:
                cleaned += 1
            val[k] = s
        rec = [
            batch, ptype,
            clean_seg(row[8]),          # s 细分市场（'/'占位符视为未填报）
            clean_str(row[3]),           # e 企业名称
            val['w'], val['r'], val['c'],
            norm_bt(row[15]),            # bt 电池类型（归一化口径）
            val['ed'], val['ec'], val['pp'], val['tp'],
            clean_str(row[20]),          # ms 电机生产企业
            val['tq'], val['fo'], val['dv'], val['ep'],
            clean_str(row[27]),          # es 发动机生产企业
            clean_str(row[29]),          # src 数据来源
            clean_str(row[1]),           # m 产品型号
            clean_str(row[2]),           # bd 产品商标
            norm_tax(row[10]),           # tx 是否减免购置税
            val['ab'],                   # ab 轴距mm（分布格局 Length 视角现用口径）
            clean_num(row[30]) if len(row) > 30 else None,  # lg 车长mm（占位列）
            clean_str(row[6]),           # gn 通用名称（散点悬浮/明细表展示）
        ]
        records.append(rec)
    wb.close()
    return records, skipped, cleaned


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
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--border);margin-bottom:24px;flex-wrap:wrap}
.tab{background:transparent;border:none;color:var(--dim);font-size:14.5px;padding:11px 20px;cursor:pointer;
  position:relative;transition:color .18s;font-family:var(--sans);letter-spacing:.5px}
.tab:hover{color:var(--text)}
.tab.on{color:var(--clay);font-weight:600}
.tab.on::after{content:'';position:absolute;left:14px;right:14px;bottom:-1px;height:2px;background:var(--clay)}
.tab i{font-style:normal;font-family:var(--mono);font-size:11px;color:var(--faint);margin-left:5px}

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
.chart.short{height:280px}
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
        <div>数据源：<b id="hSrc">__SRC_FILE__</b> · 共 <b id="hTotal">__TOTAL__</b> 条公告记录</div>
        <div>批次范围：<b id="hBRange">__BMIN__ ~ __BMAX__</b>（<b id="hBCnt">__BCNT__</b> 个批次）<span class="tag">OFFLINE READY</span>__REL_TAG__</div>
        <div>生成时间：<b id="hGen">__GEN_AT__</b> · 筛选联动全部图表</div>
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
  <button class="tab on" data-p="p1">总览<i>Overview</i></button>
  <button class="tab" data-p="p2">电机系统<i>Motor</i></button>
  <button class="tab" data-p="p3">发动机系统<i>Engine</i></button>
  <button class="tab" data-p="p4">电池系统<i>Battery</i></button>
  <button class="tab" data-p="pDist">分布格局<i>Distribution</i></button>
  <button class="tab" data-p="p5">企业与品牌<i>Enterprise / Brand</i></button>
  <button class="tab" data-p="p6">数据管理<i>Export / Import</i></button>
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
  <div>
    <b>数据刷新方法</b>：① 页面内导入——"数据管理"页拖入最新Excel，全局实时刷新；<br>
    ② 脚本重生成——更新Excel后运行 <code>python3 generate_dashboard.py</code>；亦可指定文件 <code>python3 generate_dashboard.py 新数据.xlsx 看板.html</code>
  </div>
  <div style="text-align:right">
    <b>NEV 公告数据看板__REL_NAME__</b> · 自包含离线HTML · 内嵌 ECharts 5 + SheetJS<br>
    生成于 <span id="ftGen">__GEN_AT__</span> · 数据截止批次 <span id="ftBMax">__BMAX__</span><br>
    <span style="font-size:12px;color:var(--dim)">__COPYRIGHT__ · 保留所有权利</span>
  </div>
</footer>

</div>
<button id="toTop" title="返回顶部">↑</button>

<script>__ECHARTS_LIB__</script>
<script>__XLSX_LIB__</script>
<script>
const META = __META_JSON__;
const RAW_INIT = __DATA_JSON__;
</script>
<script>
'use strict';
/* ==================== 索引与全局 ==================== */
const I = {b:0,t:1,s:2,e:3,w:4,r:5,c:6,bt:7,ed:8,ec:9,pp:10,tp:11,ms:12,tq:13,fo:14,dv:15,ep:16,es:17,src:18,
  m:19,bd:20,tx:21,ab:22,lg:23,gn:24};
const TX_IDX = I.tx;   // 免征购置税：数据层保留，展示层全量隐藏（v3.7口径）
const COL_NAMES = ['批次','动力类型','细分市场','企业名称','整备质量(kg)','纯电续航(km)','电池容量(kWh)','电池类型',
  '能量密度(Wh/kg)','百公里电耗(kWh/100km)','电机峰值功率(kW)','电机总功率(kW)','电机生产企业','峰值扭矩(Nm)',
  '综合油耗(L/100km)','发动机排量(mL)','发动机功率(kW)','发动机生产企业','数据来源',
  '产品型号','产品商标','是否减免购置税','轴距(mm)','车长(mm)','通用名称'];
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
  'dRvsLen','dCvsRg','dEcVsRg','dEdVsRg','dMassLen',
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
/* 皮尔逊相关系数（分布格局要点用） */
function pearson(pts){
  const n=pts.length; if(n<3) return null;
  const mx=pts.reduce((a,p)=>a+p[0],0)/n, my=pts.reduce((a,p)=>a+p[1],0)/n;
  let sxy=0,sxx=0,syy=0;
  for(const p of pts){ const dx=p[0]-mx, dy=p[1]-my; sxy+=dx*dy; sxx+=dx*dx; syy+=dy*dy; }
  return (sxx&&syy)? sxy/Math.sqrt(sxx*syy) : null;
}

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
const state = {bFrom:META.batchMin, bTo:META.batchMax, type:'ALL', segs:new Set()};

/* ==================== 数据工具（缺失不计入统计） ==================== */
function filtered(){
  return RAW.filter(r=>{
    if(r[I.b] < state.bFrom || r[I.b] > state.bTo) return false;
    if(state.type!=='ALL' && r[I.t]!==state.type) return false;
    if(state.segs.size>0){
      const s = r[I.s];
      if(s==null || !state.segs.has(s)) return false;   // 未填报细分市场在选择具体市场时排除
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
  return {name:name||'趋势线',type:'line',data:[[x1,+(slope*x1+icept).toFixed(1)],[x2,+(slope*x2+icept).toFixed(1)]],
    showSymbol:false,smooth:false,lineStyle:{color,width:2,type:'dashed',opacity:.65},
    itemStyle:{color},tooltip:{formatter:()=>`${name||'趋势线'} · 线性趋势：y = ${slope.toFixed(3)}x + ${icept.toFixed(1)}`},z:9};
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
  // 按动力类型拆分散点，各组独立趋势线
  const bevP=[],phevP=[];
  for(const r of rows){
    const tp=r[I.tp], tq=r[I.tq];
    if(tp&&tq){ (r[I.t]==='BEV'?bevP:phevP).push({value:[tp,tq], n:r[I.gn]}); }
  }
  const trB = trendSeries(bevP.map(p=>p.value),'#D97757','BEV趋势线');
  const trP = trendSeries(phevP.map(p=>p.value),'#6B9BD1','PHEV/EREV趋势线');
  const trends = [trB,trP].filter(Boolean);
  setInsight('chPtq', (bevP.length+phevP.length)?`散点样本 <em>${fmt(bevP.length+phevP.length)}</em> 组（BEV <em>${fmt(bevP.length)}</em> / PHEV/EREV <em>${fmt(phevP.length)}</em>）· 虚线为分组线性趋势`:'暂无功率×扭矩匹配样本');
  const fin = [
    {name:'BEV',type:'scatter',data:bevP,symbolSize:6,itemStyle:{color:`rgba(217,119,87,${TC().scatterA})`},large:true,largeThreshold:800},
    {name:'PHEV/EREV',type:'scatter',data:phevP,symbolSize:6,itemStyle:{color:`rgba(107,155,209,${TC().scatterA})`},large:true,largeThreshold:800}
  ];
  trends.forEach(t=>fin.push(t));
  chart('chPtq').setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>{
        const t=trends.find(x=>x.name===p.seriesName);
        if(t) return t.tooltip.formatter();
        return `${p.seriesName} · 通用名称：<b>${p.data.n||'–'}</b><br>总功率 <b>${p.data.value[0]}</b> kW<br>峰值扭矩 <b>${p.data.value[1]}</b> Nm`;
      }},TT),
    legend:LG({data:['BEV','PHEV/EREV',...trends.map(t=>t.name)]}),
    grid:GRID(),
    xAxis:Object.assign({type:'value',name:'总功率 kW',nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true},AXS()),
    yAxis:Object.assign({type:'value',name:'扭矩 Nm',nameTextStyle:{color:TC().axis},scale:true},AXS()),
    series:fin
  },true);
}
function chPwTrend(rows){
  const bs = byBatch(rows);
  const labels = bs.map(x=>String(x[0]));
  const defs = [
    ['BEV平均总功率', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='BEV'),I.tp);return v?+v.toFixed(0):null;}), '#D97757'],
    ['PHEV/EREV平均总功率', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='PHEV'),I.tp);return v?+v.toFixed(0):null;}), '#6B9BD1'],
    ['BEV平均峰值功率', bs.map(x=>{const v=avg(x[1].filter(r=>r[I.t]==='BEV'),I.pp);return v?+v.toFixed(0):null;}), '#D4A94E', {lineStyle:{width:1.5,type:'dashed'},symbol:'none'}]
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
  const pts = rows.filter(r=>r[I.dv]!=null&&r[I.ep]!=null).map(r=>({value:[+(r[I.dv]/1000).toFixed(2),r[I.ep]], n:r[I.gn]}));
  const tr = trendSeries(pts.map(p=>p.value),'#D4A94E','PHEV/EREV趋势线');
  setInsight('chDvEp', pts.length?`样本 <em>${fmt(pts.length)}</em> 组 · 平均排量 <em>${(pts.reduce((a,p)=>a+p.value[0],0)/pts.length).toFixed(2)} L</em> · 平均功率 <em>${(pts.reduce((a,p)=>a+p.value[1],0)/pts.length).toFixed(0)} kW</em>`:'暂无排量×功率样本');
  const series=[{type:'scatter',name:'PHEV/EREV',data:pts,symbolSize:7,itemStyle:{color:`rgba(169,127,169,${TC().scatterA})`}}];
  if(tr) series.push(tr);
  chart('chDvEp').setOption({
    tooltip:Object.assign({trigger:'item',
      formatter:p=>p.seriesName==='PHEV/EREV趋势线'?tr.tooltip.formatter():`PHEV/EREV · 通用名称：<b>${p.data.n||'–'}</b><br>排量 <b>${p.data.value[0]}</b> L · 功率 <b>${p.data.value[1]}</b> kW`},TT),
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
    if(r[I.c]!=null&&r[I.r]!=null){ (r[I.t]==='BEV'?bevP:phevP).push({value:[r[I.c],r[I.r]], n:r[I.gn]}); }
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
        return `${p.seriesName} · 通用名称：<b>${p.data.n||'–'}</b><br>容量 <b>${p.data.value[0]}</b> kWh · 续航 <b>${p.data.value[1]}</b> km`;
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
  {id:'dRvsLen',  t:'纯电续航 × 尺寸',                 sub:'EV Range vs Length · 散点 · BEV/PHEV 分列趋势线', yK:I.r, xK:I.ab, yN:'纯电续航里程(km)', xN:'轴距(mm)',
   note:'Length 现以公告轴距(mm)呈现；底表已增设车长(mm)列，补数后将 xKey 换为 lg 即可切换'},
  {id:'dCvsRg',   t:'电池容量 × 纯电续航',             sub:'Energy vs EV Range · 容量-续航匹配关系散点',      yK:I.c, xK:I.r, yN:'电池容量(kWh)',    xN:'纯电续航里程(km)'},
  {id:'dEcVsRg',  t:'百公里电耗 × 纯电续航',           sub:'Energy Consumption vs EV Range · 能效水平散点',   yK:I.ec,xK:I.r, yN:'百公里电耗(kWh/100km)', xN:'纯电续航里程(km)'},
  {id:'dEdVsRg',  t:'能量密度 × 纯电续航',             sub:'Energy Density vs EV Range · 技术水平散点',       yK:I.ed,xK:I.r, yN:'电池能量密度(Wh/kg)', xN:'纯电续航里程(km)'},
  {id:'dMassLen', t:'整备质量 × 尺寸',                 sub:'Curb Mass vs Length · 重量-尺寸分布散点',         yK:I.w, xK:I.ab, yN:'整备质量(kg)',     xN:'轴距(mm)',
   note:'纵轴为整备质量(Curb Mass)，横轴与「纯电续航 × 尺寸」同用轴距口径'}
];
function initDist(){
  const g=document.getElementById('distGrid'); if(!g)return;
  g.innerHTML = DIST_DEFS.map(d=>
    `<div class="card c-s6"><div class="c-h"><div class="c-t">${d.t}</div><div class="c-s">${d.sub}</div></div>`+
    (d.note?`<div style="color:var(--faint);font-size:11px;margin-bottom:2px">${d.note}</div>`:'')+
    `<div class="chart" id="${d.id}"></div></div>`).join('');
}
function renderDists(rows){
  for(const d of DIST_DEFS){
    if(!document.getElementById(d.id)) continue;
    const bevP=[], phevP=[];
    for(const r of rows){
      const x=r[d.xK], y=r[d.yK];
      if(x!=null&&y!=null){ (r[I.t]==='BEV'?bevP:phevP).push({value:[x,y], n:r[I.gn]}); }
    }
    const trB=trendSeries(bevP.map(p=>p.value),'#D97757','BEV趋势线');
    const trP=trendSeries(phevP.map(p=>p.value),'#6B9BD1','PHEV/EREV趋势线');
    const trends=[trB,trP].filter(Boolean);
    const series=[
      {name:'BEV',type:'scatter',data:bevP,symbolSize:6,itemStyle:{color:`rgba(217,119,87,${TC().scatterA})`},large:true,largeThreshold:800},
      {name:'PHEV/EREV',type:'scatter',data:phevP,symbolSize:6,itemStyle:{color:`rgba(107,155,209,${TC().scatterA})`},large:true,largeThreshold:800}
    ];
    trends.forEach(t=>series.push(t));
    const all = bevP.concat(phevP).map(p=>p.value);
    const cr = pearson(all);
    setInsight(d.id, all.length?`样本 <em>${fmt(all.length)}</em> 组（BEV ${fmt(bevP.length)} / PHEV·EREV ${fmt(phevP.length)}）· 相关系数 r=<em>${cr==null?'-':cr.toFixed(2)}</em>${cr==null?'':(cr>=0?'，正相关':'，负相关')}`:'当前筛选无有效样本');
    chart(d.id).setOption({
      tooltip:Object.assign({trigger:'item',
        formatter:p=>{
          const t=trends.find(x=>x.name===p.seriesName);
          if(t) return t.tooltip.formatter();
          const [xn,xu]=spl(d.xN), [yn,yu]=spl(d.yN);
          return `${p.seriesName} · 通用名称：<b>${p.data.n||'–'}</b><br>${xn}：<b>${p.data.value[0]} ${xu}</b><br>${yn}：<b>${p.data.value[1]} ${yu}</b>`;
        }},TT),
      legend:LG({data:['BEV','PHEV/EREV',...trends.map(t=>t.name)]}),
      grid:GRID(),
      xAxis:Object.assign({type:'value',name:d.xN,nameLocation:'middle',nameGap:26,nameTextStyle:{color:TC().axis},scale:true},AXS()),
      yAxis:Object.assign({type:'value',name:d.yN,nameTextStyle:{color:TC().axis},scale:true},AXS()),
      series
    },true);
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
    `本页由 <b>DIST_DEFS</b> 配置驱动，后续增加或下线分析维度只需增删配置条目。`;
}

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
  const checkIdx = [I.s,I.w,I.r,I.c,I.bt,I.ed,I.ec,I.pp,I.tp,I.ms,I.tq,I.fo,I.dv,I.ep,I.es];
  const checkNames = ['细分市场','整备质量','续航','容量','电池类型','能量密度','电耗','峰值功率','总功率','电机供应商','扭矩','油耗','排量','发动机功率','发动机供应商'];
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
  {n:'电池类型',i:I.bt},{n:'总功率kW',i:I.tp},{n:'轴距mm',i:I.ab}
];
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
      if(v==null)return '<td class="na">–</td>';
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
  /* 导出口径：动力类型列展示 PHEV/EREV */
  const data = filtered().map(r=>cols.map(c=>(c.idx===I.t&&r[c.idx]==='PHEV')?'PHEV/EREV':r[c.idx]));
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
  const records=[]; let cleaned=0;
  for(let i=1;i<aoa.length;i++){
    const row=aoa[i];
    if(!row||row[0]==null) continue;
    const batch=parseInt(String(row[0]).trim(),10);
    if(!isFinite(batch)) continue;
    const ptype = jsCleanStr(row[9])||'未知';
    const raw = {w:jsCleanNum(row[11]),ab:jsCleanNum(row[12]),r:jsCleanNum(row[13]),c:jsCleanNum(row[14]),ed:jsCleanNum(row[16]),
      ec:jsCleanNum(row[17]),pp:jsCleanNum(row[18]),tp:jsCleanNum(row[19]),tq:jsTorque(row[21]),
      fo:jsCleanNum(row[23]),dv:jsCleanNum(row[25]),ep:jsCleanNum(row[26]),
      lg:(row.length>30?jsCleanNum(row[30]):null)};
    const val={};
    for(const k in raw){
      const s=(k==='lg')?(raw[k]==null?null:raw[k]):jsSanitize(k,raw[k],ptype);
      if(raw[k]!=null&&s==null) cleaned++;
      val[k]=s;
    }
    records.push([batch,ptype,jsCleanSeg(row[8]),jsCleanStr(row[3]),val.w,val.r,val.c,jsNormBt(row[15]),
      val.ed,val.ec,val.pp,val.tp,jsCleanStr(row[20]),val.tq,val.fo,val.dv,val.ep,jsCleanStr(row[27]),jsCleanStr(row[29]),
      jsCleanStr(row[1]),jsCleanStr(row[2]),jsNormTax(row[10]),val.ab,val.lg,jsCleanStr(row[6])]);
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
  // 刷新页头/页脚
  document.getElementById('hSrc').textContent = fname;
  document.getElementById('hTotal').textContent = records.length.toLocaleString('zh-CN');
  document.getElementById('hBRange').textContent = `${META_CUR.batchMin} ~ ${META_CUR.batchMax}`;
  document.getElementById('hBCnt').textContent = META_CUR.batchCount;
  document.getElementById('hGen').textContent = META_CUR.generatedAt;
  document.getElementById('ftGen').textContent = META_CUR.generatedAt;
  document.getElementById('ftBMax').textContent = META_CUR.batchMax;
  // 重置筛选并重建选项
  state.bFrom=META_CUR.batchMin; state.bTo=META_CUR.batchMax; state.type='ALL'; state.segs.clear();
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
  const rows = filtered();
  document.getElementById('fCount').innerHTML = `已选 <b>${rows.length.toLocaleString('zh-CN')}</b> / ${META_CUR.total.toLocaleString('zh-CN')} 条`;
  updateKPIs(rows);
  insight1(rows); chBatch(rows); chType(rows); chSeg(rows); chEnt(rows); chEvo(rows);
  insight2(rows); chPwDist(rows); chMsTop(rows); chPtq(rows); chPwTrend(rows); chPwr(rows);
  insight3(rows); chDv(rows); chDvEp(rows); chEsTop(rows); chFo(rows); chDvTrend(rows);
  insight4(rows); chRgB(rows); chRgP(rows); chCr(rows); chBtTrend(rows); chEd(rows); chEcTrend(rows); chEc(rows);
  insightD(rows); renderDists(rows);
  insightEnt(rows); chEntTrend(rows); chEntSeg(rows); chNewEnt(rows); chBrandTop(rows);
  insight5(rows); chFill(rows); chSrc(rows); chHeat(rows);
  updateDmStat(); renderTable(false);
}

/* ==================== 筛选器 ==================== */
function rebuildSegPanel(){
  const segList = document.getElementById('segList');
  // 各细分市场记录计数
  const cnt = new Map();
  for(const r of RAW){ if(r[I.s]!=null) cnt.set(r[I.s],(cnt.get(r[I.s])||0)+1); }
  const items = [...cnt.entries()];
  items.sort((a,b)=>{
    const ka=segKey(a[0]), kb=segKey(b[0]);
    if(ka[0]!==kb[0]) return ka[0]-kb[0];
    if(ka[1]!==kb[1]) return ka[1]-kb[1];
    return b[1]-a[1];
  });
  segList.innerHTML = items.map(([s,c])=>
    `<label class="ms-item"><input type="checkbox" value="${s}" checked><span>${s}</span><span class="c">${c}</span></label>`).join('');
  segList.querySelectorAll('input').forEach(cb=>{
    cb.onchange = ()=>{
      if(cb.checked) state.segs.delete(cb.value); else state.segs.add(cb.value);
      syncSegBtn();
      updateAll();
    };
  });
}
function syncSegBtn(){
  const segList = document.getElementById('segList');
  const boxes = [...segList.querySelectorAll('input')];
  const off = boxes.filter(b=>!b.checked).length;
  const btn = document.getElementById('segBtn');
  if(off===0) btn.innerHTML = `全部细分市场<span class="caret">▾</span>`;
  else if(off===boxes.length) btn.innerHTML = `细分市场：无<span class="n">(0)</span><span class="caret">▾</span>`;
  else btn.innerHTML = `细分市场：<span class="n">${boxes.length-off}</span>/${boxes.length}<span class="caret">▾</span>`;
  // state.segs 语义：存储"排除项"（保持选择具体市场）
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
  state.segs.clear();
  syncSegBtn();

  // 多选下拉开合
  const sBtn=document.getElementById('segBtn'), sPanel=document.getElementById('segPanel');
  sBtn.onclick = e=>{ e.stopPropagation(); sBtn.classList.toggle('open'); sPanel.classList.toggle('open'); };
  sPanel.onclick = e=>e.stopPropagation();
  document.addEventListener('click',()=>{ sBtn.classList.remove('open'); sPanel.classList.remove('open'); });
  document.getElementById('segSelAll').onclick = ()=>{ state.segs.clear(); sPanel.querySelectorAll('input').forEach(b=>b.checked=true); syncSegBtn(); updateAll(); };
  document.getElementById('segClrAll').onclick = ()=>{
    sPanel.querySelectorAll('input').forEach(b=>{ b.checked=false; state.segs.add(b.value); });
    syncSegBtn(); updateAll();
  };

  document.getElementById('fReset').onclick = ()=>{
    state.bFrom=META_CUR.batches[0]; state.bTo=META_CUR.batches[META_CUR.batches.length-1];
    state.type='ALL'; state.segs.clear();
    bF.value=state.bFrom; bT.value=state.bTo;
    document.querySelectorAll('#typeGroup button').forEach(b=>b.classList.toggle('on',b.dataset.t==='ALL'));
    sPanel.querySelectorAll('input').forEach(b=>b.checked=true);
    syncSegBtn(); updateAll();
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
initCardInsights();
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

    for p, tip in [(in_path, '输入文件'), (ECHARTS_PATH, 'echarts.min.js'), (XLSXLIB_PATH, 'xlsx.full.min.js')]:
        if not os.path.exists(p):
            print(f'[错误] 找不到{tip}: {p}')
            sys.exit(1)

    print(f'[1/4] 读取数据: {in_path}')
    records, skipped, cleaned = load_records(in_path)
    print(f'      共 {len(records)} 条记录' + (f'（跳过 {skipped} 条无效行）' if skipped else ''))
    print(f'      剔除物理范围外异常值 {cleaned} 个（已置空，避免污染统计）')

    print('[2/4] 构建数据与元信息...')
    meta = build_meta(records, in_path)
    data_json = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    with open(ECHARTS_PATH, 'r', encoding='utf-8') as f:
        echarts_lib = f.read()
    with open(XLSXLIB_PATH, 'r', encoding='utf-8') as f:
        xlsx_lib = f.read()
    for lib in (echarts_lib, xlsx_lib):
        pass
    echarts_lib = echarts_lib.replace('</script', '<\\/script')
    xlsx_lib = xlsx_lib.replace('</script', '<\\/script')

    print('[3/4] 生成看板HTML...')
    html = HTML_HEAD
    if release:
        # 离线发布版：整体移除【数据质量检查】区块（独立 <section>，无导航入口）
        html = re.sub(r'<!-- ============ 数据质量（页面底部） ============ -->.*?</section>\n\n', '', html, flags=re.S)
    html = html.replace('__ECHARTS_LIB__', echarts_lib)
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
    html = html.replace('__REL_NAME__', ' · 离线发布版' if release else '')
    html = html.replace('__TITLE_SUFFIX__', ' · 离线发布版' if release else '')
    rel_tag = (' <span class="tag" style="background:rgba(212,169,78,.14);'
               'border-color:rgba(212,169,78,.5);color:#D4A94E;margin-left:6px">RELEASE 离线发布版</span>') if release else ''
    html = html.replace('__REL_TAG__', rel_tag)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    size_mb = os.path.getsize(out_path) / 1024 / 1024
    print(f'\n✓ 看板已生成: {out_path}（{size_mb:.1f} MB）')
    print(f'  批次范围: {meta["batchMin"]}~{meta["batchMax"]} · 动力类型: {"/".join(meta["types"])}')
    if release:
        print('  离线发布版：隐藏【数据质量检查】区块 · 其余功能全部保留（筛选/导出/导入/主题）')
    print('  v3.7：新增【分布格局】页（EV Range vs Length 等五组散点，DIST_DEFS 配置可增删）·')
    print('        全图表坐标轴标注名称与单位 · 展示层移除免征购置税信息 · 底表新增车长(mm)占位列')
    print('  v3.5：企业与品牌分析页（企业演进/企业×细分市场/新进入企业/品牌Top）')
    print('        车型明细查询表（搜索/排序/翻页）· 记录扩展 产品型号/商标/免购置税 字段')
    print('  v3：PHEV/EREV 展示名统一 / 细分市场"/"不显示 / 电池类型归一化 / 数据质量底部化')
    print('          导出Excel·CSV / 页内导入刷新 / 白天黑夜模式')
    print('  双击HTML文件即可在浏览器中离线打开使用')


if __name__ == '__main__':
    main()
