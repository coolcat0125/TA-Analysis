# -*- coding: utf-8 -*-
"""
底表空缺「同车型共识补全」v4.0
=============================
需求：尽可能补全表格空缺项，参考已有数据（同车型共识）或公开资料。
本版仅采用「同车型共识补全」：在 (企业名称, 车型名称, 通用名称, 动力类型) 维度下，
若该字段已有非空取值且全部一致（唯一共识值），则将空缺行补为该值。
原则：仅当 100% 共识才补，不造值、不覆盖已有值、不把 BEV/PHEV 混填
（动力类型纳入维度，杜绝发动机字段误填进纯电、电机字段跨动力误填）。
绿色底纹( C6EFCE )标注为「共识补全」，逐格记入「变更记录」。

用法：
  python fill_consensus_v4.py
Copyright © 2026 David YE
"""
import openpyxl
import os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SHEET = 'NEV公告参数汇总'
GREEN = openpyxl.styles.PatternFill('solid', fgColor='C6EFCE')

FIELDS = ['电机型号', '电机生产企业', '电机峰值功率(kW)', '电机总功率(kW)',
          '电池类型', '电池容量(kWh)', '纯电续航里程(km)', '电池能量密度(Wh/kg)',
          '百公里电耗(kWh/100km)', '整备质量(kg)', '轴距(mm)', '车长(mm)',
          '综合油耗(L/100km)', 'B状态油耗(L/100km)', '发动机排量(mL)',
          '发动机功率(kW)', '发动机生产企业', '发动机型号', '产品商标']


def norm(v):
    if v is None:
        return None
    s = str(v).strip()
    if s == '' or s in ('/', '-', '—', '0', '未填报', '无'):
        return None
    return s


def cast(v):
    s = str(v).strip()
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return s


def main():
    print(f'[1/4] 加载工作簿: {os.path.basename(SRC)}')
    wb = openpyxl.load_workbook(SRC)
    ws = wb[SHEET]
    hdr = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(hdr)}
    fi = {f: idx[f] for f in FIELDS}
    key_cols = ['企业名称', '车型名称', '通用名称', '动力类型']
    ki = [idx[c] for c in key_cols]

    cr = wb['变更记录']
    last_seq = 0
    for r in range(cr.max_row, 1, -1):
        v = cr.cell(r, 1).value
        if isinstance(v, int):
            last_seq = v
            break

    # 每字段：key -> 该字段出现过的 {规范化值}
    groups = {f: defaultdict(set) for f in FIELDS}
    rows = []  # (row, keytuple)
    for r in range(2, ws.max_row + 1):
        key = tuple(norm(ws.cell(r, ki[c]).value) for c in range(4))
        if None not in key:
            rows.append((r, key))
        else:
            rows.append((r, None))
        for f in FIELDS:
            v = norm(ws.cell(r, fi[f] + 1).value)
            if v is not None and None not in key:
                groups[f][key].add(v)

    ledger = []
    for ri, (r, key) in enumerate(rows, start=2):
        if key is None:
            continue
        for f in FIELDS:
            cell = ws.cell(r, fi[f] + 1)
            if norm(cell.value) is not None:
                continue
            s = groups[f].get(key)
            if s and len(s) == 1:
                val = cast(next(iter(s)))
                batch = ws.cell(r, idx['批次'] + 1).value
                ptype = ws.cell(r, idx['产品类型'] + 1).value
                cell.value = val
                cell.fill = GREEN
                ledger.append((batch, key, f, ptype, val))

    # 记入变更记录
    for (batch, key, f, ptype, val) in ledger:
        last_seq += 1
        src = '同车型共识补全：%s«%s»(%s) 唯一共识值' % (
            key[2], key[0], key[3])
        cr.append([last_seq, batch, None, ptype, f, '共识补全',
                   '(空)', f'{val}', src])
    wb.save(SRC)
    return ledger, len(rows), last_seq


if __name__ == '__main__':
    ledger, nrows, last_seq = main()
    # 汇总
    from collections import Counter
    cnt = Counter(l[2] for l in ledger)
    print(f'[2/4] 共识补全合计 {len(ledger)} 格，分布：')
    for f, n in cnt.most_common():
        print(f'      {f:22s} {n}')
    print(f'[3/4] 已写入并记入变更记录（末序号 {last_seq}）')
    print('[4/4] 完成')