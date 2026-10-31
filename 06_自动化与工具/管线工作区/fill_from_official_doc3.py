# -*- coding: utf-8 -*-
"""
官方购置税目录(doc3)交叉补全
============================
从 data/nev-announcements/raw/b<批>/doc3.txt（工信部《减免车辆购置税的新能源汽车
车型目录》原文）解析官方参数，按 (批次, 产品型号) 交叉补全底表空位。
优先级：官方原文 > master 主链 > 共识 > 公式（本脚本只补空、不覆盖已有值）。

字段映射（BEV）：通用名称→通用名称*；产品名称→产品类型*；纯电动续驶里程→纯电续航里程；
              整车整备质量→整备质量；动力蓄电池组总能量→电池容量
字段映射（PHEV）：+ 燃料消耗量→综合油耗；发动机排量→发动机排量
* 通用名称/产品类型仅在底表为空时补。

用法：
  python fill_from_official_doc3.py <底表.xlsx> <raw根目录>
"""
import sys
import os
import re
import glob
import openpyxl
from collections import defaultdict

GREEN = openpyxl.styles.PatternFill('solid', fgColor='C6EFCE')
MODEL_RE = re.compile(r'^[A-Z0-9][A-Z0-9\-]{3,20}$')
BADP = ('未提供', '无有效', '未找到', '无符合', '请提供', '无法')

# 官方字段 → 底表字段
MAP = {
    '纯电动续驶里程(km)': '纯电续航里程(km)',
    '整车整备质量(kg)': '整备质量(kg)',
    '动力蓄电池组总能量（kWh）': '电池容量(kWh)',
    '燃料消耗量(L/100km)': '综合油耗(L/100km)',
    '发动机排量(mL)': '发动机排量(mL)',
    '通用名称': '通用名称',
    '产品名称': '产品类型',
}
NUM_TFIELDS = {'纯电续航里程(km)', '整备质量(kg)', '电池容量(kWh)',
               '综合油耗(L/100km)', '发动机排量(mL)'}
# 合理范围（用于识别拼接脏值，如 285302 = "285/302"）
RANGES = {'纯电续航里程(km)': (20, 1500), '整备质量(kg)': (500, 4000),
          '电池容量(kWh)': (0, 300), '综合油耗(L/100km)': (0.1, 30),
          '发动机排量(mL)': (500, 8000)}


def clean(v):
    s = str(v).strip()
    s = re.sub(r'（[^）]*）', '', s).strip()
    s = re.sub(r'\([^)]*\)', '', s).strip()
    return s


def fillable(v):
    if v is None:
        return True
    s = str(v).strip()
    if s == '' or s in ('/', '-', '—', '0', '未填报', '无', '未提供', '无信息', '未知'):
        return True
    return False


def is_bad_numeric(v, field):
    """当前值是否为范围外脏值（拼接/错误值），可被官方值覆盖"""
    s = str(v).strip()
    if not re.match(r'^[\d.]+$', s):
        return False  # 多值文本/非数值不自动覆盖
    if field not in RANGES:
        return False
    try:
        lo, hi = RANGES[field]
        return not (lo <= float(s) <= hi)
    except ValueError:
        return False


def parse_doc3(path):
    """解析一份 doc3.txt → {型号: {官方字段: 值}}（字段逐行+空行分隔）"""
    data = open(path, encoding='utf-8-sig', errors='replace').read()
    # 按空行聚合连续字段行 → 每条记录 [字段值...]
    records = []
    cur = []
    for line in data.splitlines():
        cells = [c.strip() for c in line.split('\x07') if c.strip()]
        if not cells:
            if cur:
                records.append(cur)
                cur = []
            continue
        cur.extend(cells)
    if cur:
        records.append(cur)
    out = {}
    header_seq = None
    cur_enterprise = None
    HEADER_START = ('序号', '汽车生产企业名称', '车辆型号')
    for cells in records:
        if '车辆型号' in cells:
            # 表头记录可能混入标题行，从首个已知表头字段截取
            for i, c in enumerate(cells):
                if c in HEADER_START:
                    header_seq = cells[i:]
                    break
            cur_enterprise = None
            continue
        if header_seq is None or '车辆型号' not in header_seq:
            continue
        # 企业名只出现在每条企业的首条记录中，需继承
        if not MODEL_RE.match(cells[0]):
            cur_enterprise = cells[0]
            eff = ([''] + cells) if header_seq[0] == '序号' else cells
        else:
            eff = (['', ''] + cells) if header_seq[0] == '序号' else ([''] + cells)
        if eff and cur_enterprise and eff[1] == '':
            eff[1] = cur_enterprise
        m_idx = header_seq.index('车辆型号')
        if m_idx >= len(eff):
            continue
        model = eff[m_idx]
        if not MODEL_RE.match(model):
            continue
        rec = {}
        for i, h in enumerate(header_seq):
            if i < len(eff):
                rec[h] = eff[i]
        out[model] = rec
    return out


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return
    src_path, raw_root = sys.argv[1], sys.argv[2]
    print(f'[1/3] 解析官方 doc3 目录文件: {raw_root}')
    records = {}
    files = sorted(glob.glob(os.path.join(raw_root, 'b*', 'doc3.txt')))
    files += [os.path.join(raw_root, 'doc3.txt')]
    for f in files:
        m = re.search(r'b(\d+)', f)
        batch = int(m.group(1)) if m else None
        try:
            parsed = parse_doc3(f)
        except Exception as e:
            print(f'  跳过 {f}: {e}')
            continue
        for model, rec in parsed.items():
            key = (batch, model)
            if key not in records:
                records[key] = rec
    print(f'      解析 {len(files)} 个文件, 得到记录 {len(records)} 条')

    print(f'[2/3] 加载底表: {os.path.basename(src_path)}')
    wb = openpyxl.load_workbook(src_path)
    ws = wb['NEV公告参数汇总']
    hdr = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(hdr)}
    cr = wb['变更记录']
    last_seq = 0
    for r in range(cr.max_row, 1, -1):
        v = cr.cell(r, 1).value
        if isinstance(v, int):
            last_seq = v
            break

    stats = defaultdict(int)
    ledger = []  # (batch, model, ptype, tfield, val)
    for r in range(2, ws.max_row + 1):
        batch = ws.cell(r, idx['批次'] + 1).value
        model = str(ws.cell(r, idx['产品型号'] + 1).value or '').strip()
        ptype = ws.cell(r, idx['产品类型'] + 1).value
        rec = records.get((batch, model))
        if rec is None:
            continue
        for ofield, tfield in MAP.items():
            if ofield not in rec:
                continue
            tcol = idx[tfield] + 1
            cur = ws.cell(r, tcol).value
            if not fillable(cur) and not is_bad_numeric(cur, tfield):
                continue
            val = clean(rec[ofield])
            if not val or any(b in val for b in BADP):
                continue
            if tfield in NUM_TFIELDS:
                if not re.match(r'^[\d.]+(\s*/\s*[\d.]+)*$', val):
                    continue
                try:
                    vals = [float(x) for x in re.split(r'\s*/\s*', val)]
                    if any(not (0 <= x < 1e6) for x in vals):
                        continue
                except ValueError:
                    continue
            cell = ws.cell(r, tcol)
            old = '(空)' if fillable(cur) else str(cur)
            cell.value = val
            cell.fill = GREEN
            stats[tfield] += 1
            ledger.append((batch, model, ptype, tfield, val, old))

    print(f'[3/3] 写入 {len(ledger)} 格')
    for f, n in sorted(stats.items(), key=lambda x: -x[1]):
        print(f'      {f}: +{n}')
    for batch, model, ptype, tfield, val, old in ledger:
        last_seq += 1
        cr.append([last_seq, batch, model, ptype, tfield, '官方目录交叉补全',
                   old, val, '购置税减免目录(官方doc3)交叉补全'])
    wb.save(src_path)
    print(f'      变更记录末序号 {last_seq}')


if __name__ == '__main__':
    main()
