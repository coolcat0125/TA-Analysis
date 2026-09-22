# -*- coding: utf-8 -*-
"""dcd_common.py — DCD 自动优化工作流共享模块

被 dcd_queue.py / dcd_refill_core.py / dcd_pipeline.py / render_dashboard.py 引用。
集中定义：底表读写、空值判定、占位符过滤、品牌令牌、车系键解析、字段映射。

纪律来源(仓库既有约定,勿改)：
  - PLACEHOLDER_GENERIC_RE 逐字复制自 fill_consensus_v5.py(v4.8.1 假家族污染事故后固化)
  - FIELD_MAP 与 integrate_dcd_refill.py 的 9 字段一致(门禁 Agent 已执行口径)
"""
import hashlib
import os
import re

# ---------------------------------------------------------------- 路径
HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)                      # 仓库根(本目录的上一级)
WB_NAME = 'NEV公告参数汇总表_合并版（341~410批）.xlsx'
WB = os.path.join(REPO, WB_NAME)
SHEET = 'NEV公告参数汇总'
LEDGER = '变更记录'
RAW_DCD = os.path.join(REPO, 'raw', 'dcd_refill')
STATE = os.path.join(HERE, 'state')
SEARCH_DIR = os.path.join(STATE, 'search')
DECIDED_DIR = os.path.join(STATE, 'decided')
BATCH_DIR = os.path.join(STATE, 'batches')
BACKUP_DIR = os.path.join(STATE, 'backups')
QUEUE_FILE = os.path.join(STATE, 'queue.json')
PIPELINE_STATE = os.path.join(STATE, 'pipeline_state.json')
DASHBOARD = os.path.join(REPO, 'DCD优化监控.html')
STOP_FLAG = os.path.join(STATE, 'STOP')

# ---------------------------------------------------------------- 字段映射(9 核心字段,与 integrate_dcd_refill.py 一致)
FIELD_MAP = {
    '车长(mm)': 'length',
    '轴距(mm)': 'wheelbase',
    '整备质量(kg)': 'curb_weight',
    '纯电续航里程(km)': 'cltc_recharge_mileage',
    '百公里电耗(kWh/100km)': 'power_consumption',
    '电池容量(kWh)': 'battery_capacity',
    '电池类型': 'battery_type',
    '电池能量密度(Wh/kg)': 'battery_energy_density',
}
# 可选扩展(默认关):电机功率/扭矩文本字段,需 parse_power 解析,风险高于 9 核心字段
MOTOR_FIELD_MAP = {
    '电机总功率/扭矩': 'total_electric_power+total_electric_torque',
    '前电机功率/扭矩': 'front_electric_max_power+front_electric_max_torque',
    '后电机功率/扭矩': 'rear_electric_max_power+rear_electric_max_torque',
}
BT_NORM = {'三元锂电池': '三元锂', '磷酸铁锂电池': '磷酸铁锂', '锰酸锂电池': '锰酸锂', '钛酸锂电池': '钛酸锂'}
DCD_URL = 'https://www.dongchedi.com/auto/series/{sid}（懂车帝参数页，检索 {date}）'

# ---------------------------------------------------------------- 空值/占位符(复制自 fill_consensus_v5.py,勿改)
EMPTY_PLACEHOLDER = {'', '/', '-', '—', '0', '未填报', '无', '无信息', '待查', '待核实',
                     '待确认', '未知', '？', '?', 'N/A', 'NA', 'None', '未提供'}
SEMANTIC_PLACEHOLDER = {'不适用', '不适用(BEV)', '不适用(PHEV)', '不适用(HEV)', '不适用(EREV)'}
PLACEHOLDER_GENERIC_RE = re.compile(
    r'^(纯电动|插电式|插电混动|增程式|增程|混合动力|燃料电池|混动|纯|双层)?'
    r'(轿车|SUV|SUV车|MPV|MPV车|多用途乘用车|乘用车|运动型乘用车|客车|货车|卡车|底盘|用车)$')

# 上式为仓库既有口径(fill_consensus_v5.py),只挡精确单层形态。
# 实测(2026-09-22 队列扫描):'插电式增程混合动力多用途乘用车'、'换电式纯电动轿车' 等
# 复合形态漏网 91 条(与 D6 占位符残留 179 行同源),故补令牌全覆盖判定:
# 整串可被【动力类型词+车身类型词】完全切分 → 占位符,不可作为车系名检索。
_POWER_TOKENS = ['纯电动', '插电式', '插电混动', '插电', '增程式', '增程', '混合动力',
                 '混动', '燃料电池', '换电式', '换电', '电动', '超混动', '双层', '纯']
_BODY_TOKENS = ['多用途乘用车', '运动型乘用车', '越野乘用车', '跨界乘用车', '乘用车',
                '轿车', 'SUV', 'MPV', '客车', '货车', '卡车', '底盘', '用车', '跨界']
_ALL_TOKENS = sorted(set(_POWER_TOKENS + _BODY_TOKENS), key=len, reverse=True)


def is_placeholder_name(s):
    """整串仅由动力类型词+车身类型词构成 → True(占位符,不可检索)。"""
    s = str(s or '').strip()
    if not s:
        return True
    if PLACEHOLDER_GENERIC_RE.match(s):
        return True
    i, n = 0, len(s)
    while i < n:
        for t in _ALL_TOKENS:
            if s.startswith(t, i):
                i += len(t)
                break
        else:
            return False
    return True


def is_empty(v):
    """可补空:真空值或占位符。语义占位(不适用(BEV))不可补,返回 False。"""
    if v is None:
        return True
    s = str(v).strip()
    if s in EMPTY_PLACEHOLDER:
        return True
    return False


def is_semantic_placeholder(v):
    return str(v or '').strip() in SEMANTIC_PLACEHOLDER


def fnum(v):
    """数值解析:'35kW(48Ps)'→35;'130/270'→130;带逗号千分位。"""
    try:
        return float(str(v).split('/')[0].replace(',', '').replace('km', '').strip())
    except Exception:
        return None


def parse_power(v):
    m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*kW', str(v), re.I)
    return float(m.group(1)) if m else None


# ---------------------------------------------------------------- 车系键解析
GENERIC_SPLIT_RE = re.compile(r'[,，/、;；]')


def generic_parts(gn):
    """通用名称拆分为车系名候选(过滤占位符/空/过短)。"""
    out = []
    for p in GENERIC_SPLIT_RE.split(str(gn or '')):
        p = p.strip()
        if not p or len(p) < 2:
            continue
        if is_placeholder_name(p):
            continue
        if p in EMPTY_PLACEHOLDER:
            continue
        out.append(p)
    return out


def strip_brand(s):
    """'埃安(AION)牌'→('埃安', {'aion'});'欧拉牌'→('欧拉', set())"""
    s = str(s or '').strip()
    latin = set()
    for m in re.findall(r'[（(]([^)）]+)[)）]', s):
        for t in re.findall(r'[A-Za-z]+', m):
            if len(t) >= 2:
                latin.add(t.lower())
    s = re.sub(r'[（(][^)）]*[)）]', '', s)
    s = s.replace('牌', '').strip()
    return s, latin


def brand_tokens(*vals):
    """从商标/车型名称/企业名提取品牌令牌(小写),用于车系名校验。"""
    toks = set()
    for v in vals:
        core, latin = strip_brand(v)
        if latin:
            toks |= latin
        if not core:
            continue
        toks.add(core.lower())
        for t in re.findall(r'[A-Za-z]+', core):
            if len(t) >= 2:
                toks.add(t.lower())
        # 中文品牌取前 2~4 字,覆盖 '北京新能源汽车' → '北京'
        if re.search(r'[一-鿿]', core) and len(core) > 2:
            toks.add(core[:2].lower())
    toks.discard('')
    return toks


def row_family(batch, model, brand, corp):
    """车系族键:优先车型名称,空/占位符则商标核心,再空则企业名前 6 字。"""
    for v in (batch, model):
        s = str(v or '').strip()
        if s and not is_placeholder_name(s):
            return s
    core, _ = strip_brand(brand)
    if core:
        return core
    return str(corp or '').strip()[:6]


def norm_name(s):
    """车系名/关键词归一:小写、去空格与常见分隔符。"""
    s = str(s or '').lower()
    return re.sub(r'[\s\-_·．.]+', '', s)


def qkey(family, part):
    """队列条目稳定 ID。"""
    return hashlib.md5(f'{family}|{part}'.encode('utf-8')).hexdigest()[:10]


def ensure_dirs():
    for d in (STATE, SEARCH_DIR, DECIDED_DIR, BATCH_DIR, BACKUP_DIR):
        os.makedirs(d, exist_ok=True)


def load_rows():
    """读底表全部行 + 表头索引。read_only 模式,调用方负责 close。"""
    import openpyxl
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb[SHEET]
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    return hdr, hi, rows


def today():
    import datetime
    return datetime.date.today().isoformat()
