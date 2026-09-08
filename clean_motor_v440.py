# -*- coding: utf-8 -*-
"""v4.4.0 电机列残留脏值清理
================================
迁移后审计发现三类残留：
  1. '0'/'1' 等占位格（旧数值 0/1 被格式化为 P）——清空
  2. 功率段超出物理范围 (5,1500) 的格（如 '2118/543'、'2003/343'）——清空（U9 Xtreme 2220kW 保留例外）
  3. 清理后若 前 存在且 后 为空且 总≠前 → 总=前（用户口径：总=前+后）
全部逐格台账（类型=数据清洗更正，浅红 FFC7CE）。
"""
import sys

import openpyxl
from openpyxl.styles import PatternFill

sys.stdout.reconfigure(encoding="utf-8")
SRC = "NEV公告参数汇总表_合并版（341~410批）.xlsx"
RED = PatternFill("solid", fgColor="FFC7CE")
KEEP_TOTAL = {"QCJ7000AFBEV3"}   # 仰望U9 Xtreme 2220kW 真实值


def parse_pt(v):
    if v is None:
        return None, None, ""
    s = str(v).strip()
    if not s:
        return None, None, ""
    parts = s.split("/")
    try:
        p = float(parts[0]) if parts[0] else None
    except ValueError:
        p = None
    try:
        t = float(parts[1]) if len(parts) > 1 and parts[1] else None
    except ValueError:
        t = None
    return p, t, s


def pt_cell(p, t):
    if p is None:
        return ""
    a = str(int(p)) if float(p) == int(p) else str(p)
    if t is None:
        return a
    b = str(int(t)) if float(t) == int(t) else str(t)
    return f"{a}/{b}"


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb["NEV公告参数汇总"]
    hdr = [str(c.value).strip() if c.value else "" for c in ws[1]]
    fi = {h: i for i, h in enumerate(hdr)}
    iF, iT, iR = fi["前电机功率/扭矩"], fi["电机总功率/扭矩"], fi["后电机功率/扭矩"]
    iM = fi["产品型号"]
    cr = wb["变更记录"]
    seq = 0
    for rr in range(cr.max_row, 1, -1):
        v = cr.cell(rr, 1).value
        if isinstance(v, int):
            seq = v
            break

    led = []
    stats = {"clear": 0, "retotal": 0}

    def clear(r, col, old, field, note):
        cell = ws.cell(r, col + 1)
        cell.value = None
        cell.fill = RED
        stats["clear"] += 1
        led.append([ws.cell(r, 1).value, ws.cell(r, iM + 1).value, ws.cell(r, fi["动力类型"] + 1).value,
                    field, "数据清洗更正", str(old)[:40], "(空)", note])

    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            continue
        model = str(ws.cell(r, iM + 1).value or "").strip()
        fp, ft, fs = parse_pt(ws.cell(r, iF + 1).value)
        tp, tt, ts = parse_pt(ws.cell(r, iT + 1).value)
        rp, rt, rs = parse_pt(ws.cell(r, iR + 1).value)
        # 1) 占位/越界清理
        if fs and (fp is None or not (5 <= fp <= 1500)):
            clear(r, iF, fs, "前电机功率/扭矩", "占位或功率超物理范围(5~1500)")
            fp = ft = None
        if rs and (rp is None or not (5 <= rp <= 1500)):
            clear(r, iR, rs, "后电机功率/扭矩", "占位或功率超物理范围(5~1500)")
            rp = rt = None
        if ts and (tp is None or not (5 <= tp <= 1500)) and model not in KEEP_TOTAL:
            clear(r, iT, ts, "电机总功率/扭矩", "占位或功率超物理范围(5~1500)；U9 Xtreme 2220kW 保留例外")
            tp = tt = None
        # 2) 总=前+后 口径回填
        fp2, ft2, _ = parse_pt(ws.cell(r, iF + 1).value)
        rp2, rt2, _ = parse_pt(ws.cell(r, iR + 1).value)
        tp2, tt2, ts2 = parse_pt(ws.cell(r, iT + 1).value)
        if fp2 is not None and rp2 is None:
            if tp2 is None or abs(tp2 - fp2) > 2:
                val = pt_cell(fp2, ft2)
                cell = ws.cell(r, iT + 1)
                old = cell.value
                cell.value = val
                cell.fill = RED
                stats["retotal"] += 1
                led.append([ws.cell(r, 1).value, ws.cell(r, iM + 1).value,
                            ws.cell(r, fi["动力类型"] + 1).value, "电机总功率/扭矩", "数据清洗更正",
                            "" if old is None else str(old)[:40], val, "两驱口径：总=前电机（原值不符已修正）"])
    for row in led:
        seq += 1
        cr.append([seq] + row)
    wb.save(SRC)
    print(f"清理 {stats['clear']} 格，总=前 回填 {stats['retotal']} 格，台账 {len(led)} 条（末序号 {seq}）")


if __name__ == "__main__":
    main()
