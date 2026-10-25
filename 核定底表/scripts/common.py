# -*- coding: utf-8 -*-
"""核定底表公共层：路径常量 + 解析工具（自 15 OEM Portfolio/scripts/_ress_extract_oem.py 移植）。

职责边界
--------
- 本目录（01-主仓库/TA-Analysis/核定底表/，2026-10-25 自 02-数据底表/核定底表/ 迁入并纳入 git）只新建文件，不改 17 任何旧文件。
- 权威源 = 01-主仓库 合并版 xlsx（AGENTS 2026-09-04：派生库滞后，以仓库底表为准）。
- 411 正式/412 公示 来自 主仓内 公告批次扫描_411-412_20261017/ 的扫描底表（字段更全，含来源URL）。
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

# ── 路径 ──────────────────────────────────────────────
ROOT17 = r"D:\#AI\00 Project\17 TA analysis"
HERE = os.path.dirname(os.path.abspath(__file__))          # ...\核定底表\scripts
BASE = os.path.dirname(HERE)                               # ...\核定底表
DATA = os.path.join(BASE, 'data')

SRC_XLSX = os.path.join(ROOT17, '01-主仓库', 'TA-Analysis',
                        'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SCAN_DIR = os.path.join(ROOT17, '01-主仓库', 'TA-Analysis',
                        '公告批次扫描_411-412_20261017')
SCAN_XLSX = os.path.join(SCAN_DIR, 'outputs',
                         '01a147a8-938a-7ff3-a15c-94b7fcac7a5e',
                         '411-412公告NEV车型扫描底表.xlsx')
DERIVED_DB = os.path.join(ROOT17, '02-数据底表', 'data', 'nev-announcements',
                          'db', 'master_vehicles.json')    # 交叉对账用（派生库，仅对账不改）

# 产出物
XLSX_OUT = os.path.join(DATA, 'NEV公告核定底表（341~411批+412公示）_v1.xlsx')
CSV_OUT = os.path.join(DATA, 'NEV公告核定底表_v1.csv')
JSON_OUT = os.path.join(DATA, 'NEV公告核定底表_v1.json')
CAR_LEN_JSON = os.path.join(DATA, 'car_len_master.json')
ALIAS_JSON = os.path.join(DATA, 'series_alias_master.json')
DIAG_MD = os.path.join(DATA, 'diag_report.md')
DIAG_JSON = os.path.join(DATA, 'diag_data.json')
VERIFY_MD = os.path.join(DATA, 'verify_report.md')


# ── 解析工具（语义占位为公告原生合法值，AGENTS 2026-09-07）──
def num(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s or s in ('不适用(BEV)', '不适用(PHEV)', '-', '—', '/', 'N/A'):
        return None
    m = re.findall(r'-?\d+(?:\.\d+)?', s.replace(',', ''))
    if not m:
        return None
    try:
        return float(m[0])
    except Exception:
        return None


def ranges(v):
    if v is None:
        return []
    return [float(x) for x in re.findall(r'\d+(?:\.\d+)?', str(v))]


def norm_mat(v):
    s = str(v or '').strip()
    if not s or s in ('不适用(BEV)', '不适用(PHEV)'):
        return '待核实'
    if '磷酸铁' in s or '铁锂' in s or s.upper() == 'LFP':
        return '磷酸铁锂'
    if '三元' in s or '镍钴' in s or s.upper() == 'NCM':
        return '三元锂'
    if '锰酸' in s:
        return '锰酸锂'
    if '锂离子' in s:
        return '锂离子(未细分)'
    return s


# 「通用名称」实为产品类型（C 问题）——动力类型泄漏占位值（AGENTS 2026-09-17 同源教训）
PTYPE_RE = re.compile(r'^(纯电动|插电式|换电式|增程式?)[^，,]*?(轿车|乘用车|多用途乘用车|'
                      r'运动型乘用车|越野乘用车|交叉型乘用车|客车|货车)?$')


def is_ptype_name(v):
    return bool(PTYPE_RE.fullmatch(str(v or '').strip()))


def series_of(row, I):
    """车系名：优先通用名称，退化到车型名称；跳过产品型占位值。"""
    for k in ('通用名称', '车型名称', '产品名称'):
        v = str(row[I[k]] or '').strip()
        if v and not is_ptype_name(v):
            return v
    return ''      # 全部为占位 → 车系缺失（master 侧置 null + 标志，不再拿产品名称顶替）


def year_of_batch(b):
    """341-352→2021 … 401-412→2026（约 12 批/年）。"""
    if b <= 352:
        return 2021
    if b <= 364:
        return 2022
    if b <= 376:
        return 2023
    if b <= 388:
        return 2024
    if b <= 400:
        return 2025
    return 2026


def rel_batch_month(b):
    """批次序号 → 大致年月（仅诊断参考，341 批≈2021-01，每月一批）。"""
    y = year_of_batch(b)
    return y, ((b - 341) % 12) + 1
