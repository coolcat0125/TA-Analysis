# -*- coding: utf-8 -*-
"""fix_miniev_gen.py — 宏光MINIEV 代际错配修复（dry-run / apply）

代际真值表（汽车之家 series 5714 逐款取证 + 懂车帝 dcd_s4499 双源互证）：
  2020款初代   2917/1940   spec 45757  (2917*1493*1621, 1940)
  2021款马卡龙 2920/1940   spec 50352  (2920*1493*1621, 1940)
  2022款马卡龙 2920/1940   spec 56033  (2920*1493*1621, 1940)
  GAMEBOY200   3061/2010   spec 55822  (3061*1520*1665, 2010, 17.3kWh/200km/CW772, 单电机30kW)
  GAMEBOY300   3061/2010   spec 55810  (3061*1520*1659, 2010, CW822, 单电机30kW)
  敞篷版       3059/2010   spec 55267  (3059*1521*1614, 2010, 26.5kWh/280km/CW925, 单电机30kW)
  第三代       3064/2010   spec 68994  (3064, 2010, 17.3kWh/215km) + spec 66044(120km) + DCD
  第四代四门   3256/2190   spec 71113  (3256*1510*1578, 2190, 16.2kWh/205km/CW780) + DCD
  第五代       3268/2190   spec 76733  (3268, 2190) + DCD
  2025两门版   ?           spec 69931 paramisshow=0 无数据 → HOLD

行→代际判定规则（批次日期 + 容量/续航/整备签名）：
  batch<=342                → 2020款 (2917/1940)   [马卡龙2021-04-08上市, 342批=04-01更早]
  343<=batch<=367 且 3门签名 → 2021/2022马卡龙 (2920/1940)
  CAP≈17.3/17.4 & RNG≈200   → GAMEBOY200 (3061/2010)
  CAP≈26.5/26.7 & RNG≈280~300 & CW>=822 → GAMEBOY300 (3061/2010)；RNG=280 → 敞篷版 (3059/2010)
  batch>=373 且 A系列代码    → 第三代 (3064/2010)
  CAP≈16.2/16.3 & RNG≈205   → 第四代 (3256/2190)
  CAP≈25.1 & RNG≈301 & CW≈845 → 第四代 (3256/2190)
  CW≈790/860                 → 第五代 (3268/2190)
  其余（CW=750/16.2kWh/205km）→ HOLD

电机列修复：单电机车（全族均单电机后驱）后电机=150/310 系共识污染 → 置空；总=前（公告值）。
"""
import json
import os
import sys

import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
RED = PatternFill('solid', fgColor='FFC7CE')
GREEN = PatternFill('solid', fgColor='C6EFCE')

# 代际真值: (车长, 轴距, 代际名, 证据)
GEN = {
    'g2020': (2917, 1940, '2020款初代', '汽车之家宏光MINIEV spec45757(2020款轻松款): 2917×1493×1621/轴距1940'),
    'g2021': (2920, 1940, '2021款马卡龙', '汽车之家宏光MINIEV spec50352(2021款马卡龙时尚款): 2920×1493×1621/轴距1940'),
    'g2022': (2920, 1940, '2022款马卡龙', '汽车之家宏光MINIEV spec56033(2022款马卡龙时尚款): 2920×1493×1621/轴距1940'),
    'gboy2': (3061, 2010, 'GAMEBOY 200km', '汽车之家宏光MINIEV spec55822(2022款GAMEBOY200km玩乐款): 3061×1520×1665/轴距2010'),
    'gboy3': (3061, 2010, 'GAMEBOY 300km', '汽车之家宏光MINIEV spec55810(2022款GAMEBOY300km玩咖款): 3061×1520×1659/轴距2010'),
    'gcab':  (3059, 2010, '敞篷版', '汽车之家宏光MINIEV spec55267(2022款敞篷版): 3059×1521×1614/轴距2010'),
    'g3rd':  (3064, 2010, '2024款第三代', '汽车之家spec68994(215km青春版)+DCD dcd_s4499三门版: 3064/2010'),
    'g4th':  (3256, 2190, '2025款第四代四门版', '汽车之家spec71113(四门版进阶款)+DCD: 3256/2190'),
    'g5th':  (3268, 2190, '2026款第五代', '汽车之家spec76733(第五代205km进阶款)+DCD: 3268/2190'),
}

AUTOHOME_URL = 'https://car.autohome.com.cn/5714/'  # 车系页;具体 spec 见台账证据


def fnum(v):
    try:
        return float(str(v).split('/')[0].split('（')[0].split('(')[0].replace(',', '').strip())
    except Exception:
        return None


def classify(batch, code, cap, rng, cw):
    """返回代际 key 或 None(HOLD)。"""
    b = int(batch) if str(batch).isdigit() else 0
    # 第五代: CW 790/860 (2025-12 起)
    if cw in (790, 860) and cap and abs(cap - 16.2) < 0.3 or (cw in (790, 860) and cap and abs(cap - 25.1) < 0.3):
        return 'g5th'
    # 第四代: 16.2/16.3kWh 205km CW760-780; 25.1kWh 301km CW845
    if cap and abs(cap - 25.1) < 0.3 and rng and rng >= 290:
        return 'g4th'
    if cap and 15.9 <= cap <= 16.6 and rng and 200 <= rng <= 210 and cw and 755 <= cw <= 785:
        return 'g4th'
    # 敞篷版: 26.5kWh 280km CW925
    if cap and abs(cap - 26.5) < 0.3 and rng and 275 <= rng <= 285:
        return 'gcab'
    # GAMEBOY 300km: 26.5/26.7kWh 300km CW822-832
    if cap and 26.0 <= cap <= 27.0 and rng and rng >= 290 and cw and cw >= 815:
        return 'gboy3'
    # GAMEBOY 200km: 17.3/17.4kWh 200km CW750-772
    if cap and 17.0 <= cap <= 17.6 and rng and 195 <= rng <= 210:
        return 'gboy2'
    # 第三代 A系列 (373+): 17.3kWh 215km 或 120/170km A系列代码
    if b >= 373 and code.startswith('LZW7004EVA'):
        return 'g3rd'
    if cap and abs(cap - 17.3) < 0.3 and rng and 210 <= rng <= 220:
        return 'g3rd'
    # 2020款: batch<=342
    if b <= 342:
        return 'g2020'
    # 2021/2022款马卡龙: 343-367 三门 120/170km
    if 343 <= b <= 372 and rng and rng <= 175:
        return 'g2022'
    return None


def main():
    apply = '--apply' in sys.argv
    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in ws[1]]
    hi = {h: i for i, h in enumerate(hdr)}
    ledger = wb['变更记录']
    seq = max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
               if isinstance(r[0], int)), default=0)

    rows = list(ws.iter_rows(min_row=2))
    plan = []   # (row, gen, L, W, changes)
    holds = []
    for row in rows:
        code = str(row[hi['产品型号']].value or '').strip()
        if not (code.startswith('LZW7004EV') or code.startswith('LZW7001EV') or code.startswith('LZW7001BEV')):
            continue
        batch = str(row[hi['批次']].value or '').strip()
        cap = fnum(row[hi['电池容量(kWh)']].value)
        rng = fnum(row[hi['纯电续航里程(km)']].value)
        cw = fnum(row[hi['整备质量(kg)']].value)
        # 电机列清理（与代际判定独立）：单电机车后电机=150/310 系共识污染（前=公告值）
        r_motor = str(row[hi['后电机功率/扭矩']].value or '').strip()
        f_motor = str(row[hi['前电机功率/扭矩']].value or '').strip()
        t_motor = str(row[hi['电机总功率/扭矩']].value or '').strip()
        motor_changes = []
        if r_motor == '150/310' and f_motor:
            motor_changes.append(('后电机功率/扭矩', r_motor, ''))
            if t_motor != f_motor:
                motor_changes.append(('电机总功率/扭矩', t_motor, f_motor))
        gk = classify(batch, code, cap, rng, cw)
        cur_l = fnum(row[hi['车长(mm)']].value)
        cur_w = fnum(row[hi['轴距(mm)']].value)
        if gk is None:
            holds.append((batch, code, cur_l, cur_w, cap, rng, cw))
            if motor_changes:
                plan.append((row, None, '（代际未判定，仅电机列清理）', '', motor_changes))
            continue
        L, W, gname, ev = GEN[gk]
        changes = []
        if cur_l != L:
            changes.append(('车长(mm)', cur_l, L))
        if cur_w != W:
            changes.append(('轴距(mm)', cur_w, W))
        changes.extend(motor_changes)
        if changes:
            plan.append((row, gk, gname, ev, changes))

    # 汇总
    from collections import Counter
    gc = Counter(p[2] for p in plan)
    print(f'=== 修复计划（dry-run）: {len(plan)} 行 ===')
    for g, n in gc.most_common():
        gk0 = [k for k, v in GEN.items() if v[2] == g]
        if gk0:
            L, W, _, _ = GEN[gk0[0]]
            print(f'  {g}: {n} 行 → L={L} W={W}')
        else:
            print(f'  {g}: {n} 行')
    cells = sum(len(p[4]) for p in plan)
    print(f'合计 {cells} 格')
    print(f'HOLD（无代际匹配，保持原值）: {len(holds)} 行')
    for h in holds:
        print(f'  hold {h[0]} {h[1]} L={h[2]} W={h[3]} CAP={h[4]} RNG={h[5]} CW={h[6]}')
    # 明细抽样
    print('\n--- 明细（前12行）---')
    for row, gk, gname, ev, changes in plan[:12]:
        code = str(row[hi['产品型号']].value or '').strip()
        ch = '; '.join(f'{f}: {o}→{n}' for f, o, n in changes)
        print(f'  {code} [{gname}] {ch}')

    if not apply:
        print('\nDRY-RUN（未写入）。加 --apply 落库。')
        return

    # 落库
    stats = Counter()
    for row, gk, gname, ev, changes in plan:
        batch = row[hi['批次']].value
        code = str(row[hi['产品型号']].value or '').strip()
        pt = str(row[hi['动力类型']].value or '')
        for fld, old, new in changes:
            cell = row[hi[fld]]
            cell.value = new if new != '' else None
            cell.fill = RED
            seq += 1
            ledger.append((seq, batch, code, pt, fld, '代际错配更正', str(old), str(new),
                           f'汽车之家宏光MINIEV(sid5714)逐代取证+懂车帝互证：{gname}；{ev}'))
            stats['cells'] += 1
    wb.save(WB)
    print(f'\nsaved: {stats["cells"]} 格，台账 seq→{seq}')


if __name__ == '__main__':
    main()
