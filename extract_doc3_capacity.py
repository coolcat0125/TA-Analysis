# -*- coding: utf-8 -*-
"""extract_doc3_capacity.py v3 — 购置税目录(doc3.txt)官方额定容量证据提取器

兼容两种转换格式：
  A) 单行 SEP(chr7) 连接全部字段
  B) 逐字段一行（SEP 前缀，按表头顺序成组）
"""
import json
import os
import re
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements', 'raw')
OUT = os.path.join(HERE, 'audit-output', 'doc3_capacity_evidence.json')
SEP = chr(7)
NL = chr(10)
TAB = chr(9)

KW = {
    'model': ('车辆型号',),
    'enterprise': ('企业',),
    'generic': ('通用名称',),
    'product': ('产品名称',),
    'range': ('续驶里程',),
    'curb': ('整备质量',),
    'energy_kwh': ('总能量',),
}


def col_index(header_cells, *keywords):
    for i, c in enumerate(header_cells):
        for kw in keywords:
            if kw in c:
                return i
    return None


def scan_all(lines, batch, path):
    """Token 流状态机：段标志→乘用车子段→表头→数据（SEP 前缀单字段行 + 行内多值行兼容）。"""
    out = []
    n = len(lines)
    i = 0
    pt = None
    in_psg = False
    header = None
    buf = []
    while i < n:
        raw = lines[i]
        s = raw.strip().lstrip(SEP).strip()
        if s.startswith(('一、', '二、', '三、', '四、')) and '乘用车' not in s:
            pt = 'BEV' if '纯电动汽车' in s else ('PHEV' if '插电式混合动力' in s else None)
            in_psg = False
            header = None
            buf = []
            i += 1
            continue
        if (s.startswith('（') or s.startswith('(')) and '）' in s[:8]:
            in_psg = '乘用车' in s
            i += 1
            continue
        if pt is None or not in_psg:
            i += 1
            continue
        if s == '序号':
            header = ['序号']
            i += 1
            while i < n:
                t = lines[i].strip().lstrip(SEP).strip()
                if not t:
                    i += 1
                    continue
                if t[0].isdigit():
                    break
                header.append(t)
                i += 1
            continue
        if (SEP in raw or TAB in raw) and s:
            vals = [c.strip().lstrip(SEP) for c in re.split('[' + SEP + TAB + ']', raw)]
            vals = [v for v in vals if v != '']
            buf.extend(vals)
            if len(buf) >= len(header):
                rec = build_rec(header, buf[:len(header)], pt, batch, path)
                if rec:
                    out.append(rec)
                buf = buf[len(header):]
        i += 1
    return out


def build_rec(header, values, pt, batch, path):
    """按表头位置映射字段值列表 → 记录。"""
    if header is None or len(header) < 2 or len(values) < len(header):
        return None
    rec = {}
    ok = False
    for j, h in enumerate(header):
        v = values[j]
        if v:
            ok = True
        for key, kws in KW.items():
            if any(kw in h for kw in kws):
                rec[key] = v
    if not ok or not rec.get('model'):
        return None
    rec['batch'] = batch
    rec['pt'] = pt
    rec['source'] = os.path.basename(path) + '（购置税目录原文，公告附件3）'
    return rec


def main():
    all_rows = []
    rawdir = RAW
    for d in sorted(os.listdir(rawdir)):
        p = os.path.join(rawdir, d, 'doc3.txt')
        if not (os.path.isfile(p) and re.match(r'^b\d{3}$', d)):
            continue
        lines = open(p, encoding='utf-8').read().split(NL)
        all_rows.extend(scan_all(lines, d[1:], p))
    idx = defaultdict(list)
    for r in all_rows:
        if r.get('model'):
            idx[r['model']].append(r)
    json.dump({'total_rows': len(all_rows), 'unique_models': len(idx),
               'by_model': dict(idx)},
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'parsed rows={len(all_rows)} unique_models={len(idx)} -> {OUT}')


if __name__ == '__main__':
    main()
