# -*- coding: utf-8 -*-
"""fill_410_official.py — 410 批正式公告对齐入库（S1 触发轮，2026-09-18 公告发布）

数据源：02-数据底表/data/nev-announcements/raw/b410-official/parsed_official_410.json
  （购置税目录第三十四批 + 车船税目录第八十九批，410 批公告附件 2/3，官方 A 级）

流程（对齐 align_official.py 409 先例）：
  1. 占位行配对：410 批 29 行"无型号公示占位"按 人工复核配对表 挂到官方型号
     （整备/续航/能量多重印证；3 行目录未收录维持留档）
  2. 占位行参数对齐：型号补空(绿)；续航/整备/容量 与官方不一致→官方值更正(浅红+台账双值留痕)
  3. 新增行：官方行中未入库的 48 款（跨批同码合法双批，直接作 410 行插入，绿底）
  4. 台账逐格 + 410 公告日期(2026-09-18)写入批次时间表

用法：python fill_410_official.py [--apply]  （默认干跑）
"""
import json
import os
import sys
import datetime as dt

import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SRC = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                   'raw', 'b410-official', 'parsed_official_410.json')

GREEN = PatternFill('solid', fgColor='C6EFCE')   # 官方补空
RED = PatternFill('solid', fgColor='FFC7CE')     # 官方更正（公示占位媒体值→官方值）
ANNOUNCE_DATE = '2026-09-18'
SRC_TAG = '410批正式公告附件（购置税34批/车船税89批）'

# 人工复核配对表（占位行号→官方型号；依据：整备/续航/能量≥2 重叠印证，2026-09-18 逐行复核）
PAIRING = {
    4342: 'AHC7000BEVE2B',   # 埃安Ray7 整备1855+续航~700
    4343: 'JL7001BEV111',    # smart精灵#2 整备1130+能量35.8≈35.7
    4344: 'BJ7000T37SEV',    # 极狐T1 唯一 T1（北汽新能源）
    4347: 'BYD7002BXBEV1',   # 海豹07 EV 整备1975
    4348: 'BYD7002BXBEV2',   # 海豹07 EV 第二变体（原行为重复占位，官方为两个配置）
    4349: 'CSA7005TBEV2',    # MG 07 唯一
    4350: 'BJ6480A50ABEV',   # 极狐αT5 唯一
    4351: 'SVW646191AEV',    # ID.ERA 5X 整备1685
    4352: 'SKE6522DBEVR1',   # 问界M8纯电 续航708≈705
    4353: 'SQR6500BEVEHX6',  # 智界RX 续航702+整备2285
    4354: 'SQR6500BEVEHX7',  # 智界RX 整备2400
    4355: 'SQR6450BEVT1VA',  # 捷途环游者 续航550 精确
    4356: 'CC6530DC21HABEV', # 魏牌V9X 整备2805
    4358: 'BJ6500C311RPHEV', # 星钽5X PHEV 续航232+能量50.4
    4363: 'BYD7152BX6HEV2',  # 海豹07 DM-i 三值全等
    4364: 'HQ6483DCHEV102',  # 银河M7 续航168+能量30.71
    4367: 'CC6542BD24CPHEV', # 魏牌高山9 整备2960+能量53.9
    4368: 'CC6481DE00HBBEV', # 欧拉5 能量58.3
    4369: 'BJ6500C31RPHEV',  # 星钽5X PHEV 续航232（另一变体）
    4370: 'HQ6472DCHEV108',  # 银河星舰7 续航178+能量30.71
    4371: 'DFL6480NAPJ2PHEV',# 东风日产NX7 续航225+能量40.95
    4372: 'EQ6500MS5F3PHEV', # 猛士X700 三值全等
    4373: 'SKE6522DREEVR1',  # 问界M8增程 续航256+能量56.036
    4374: 'SKE6522DREEVR2',  # 问界M8增程 续航365+能量75.408
    4375: 'NEQ6510REEVS5DL', # iCAR V27 续航200+整备2455（唯一 V27；能量 34.3 系媒体值待官方）
    4376: 'EQ6500MS5F0REEV', # 猛士X700 续航244+能量51.46
}
# 目录未收录维持留档：4357 宝马iX3 40L / 4360 奔驰VLE / 4362 欧拉5·430km（及既有型号行）

PTYPE_OF = {'bev': '纯电动轿车', 'phev': '插电式混合动力轿车'}


def ptype_of_row(o, is_bev):
    name = str(o.get('产品名称') or '')
    if is_bev:
        return '纯电动轿车' if '轿车' in name else '纯电动轿车' if '轿车' in str(o.get('通用名称')) else '纯电动轿车' if '轿' in name else ('换电式纯电动轿车' if '换电' in name else '纯电动多用途乘用车' if '多用途' in name else '纯电动轿车')
    return '插电式混合动力轿车' if '轿' in name else '插电式混合动力多用途乘用车' if '多用途' in name else '插电式增程混合动力乘用车' if '增程' in name or 'REEV' in o.get('型号', '') else '插电式混合动力轿车'


def main():
    apply = '--apply' in sys.argv
    d = json.load(open(SRC, encoding='utf-8'))
    officials = {}
    for key, is_bev in (('bev', True), ('phev', False)):
        for o in d[key]:
            o['_bev'] = is_bev
            officials.setdefault(o['型号'], o)  # doc2/doc3 重复型号取先到（并集字段后补）
    # 并集补字段：doc2 有油耗/排量，doc3 有整备/能量
    for key in ('bev', 'phev'):
        for o in d[key]:
            t = officials[o['型号']]
            for f in ('续航', '整备', '能量', '电池质量', '油耗', '排量'):
                if not t.get(f) and o.get(f):
                    t[f] = o[f]

    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in ws[1]]
    hi = {h: i for i, h in enumerate(hdr)}
    ledger = wb['变更记录']
    seq = max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
               if isinstance(r[0], int)), default=0)

    # 行号→行对象
    rows = {r[0].row: r for r in ws.iter_rows(min_row=2)}

    def L(rowno, row, field, ctype, old, new, note):
        nonlocal seq
        seq += 1
        ledger.append((seq, row[hi['批次']].value, row[hi['产品型号']].value or '', row[hi['动力类型']].value or '',
                       field, ctype, str(old)[:60] if old is not None else '(空)',
                       new if new is not None else '(空)', note))

    n_code = n_fix = n_new = 0
    used = set(PAIRING.values())

    # --- 1) 410 已有型号行：官方补空/更正 ---
    for o in officials.values():
        if o['型号'] in used:
            continue
    for rowno, row in rows.items():
        code = str(row[hi['产品型号']].value or '').strip()
        if str(row[hi['批次']].value) != '410' or code not in officials or code in used:
            continue
        o = officials[code]
        for fld, col, off in (('续航', '纯电续航里程(km)', '续航'), ('整备', '整备质量(kg)', '整备'),
                              ('容量', '电池容量(kWh)', '能量')):
            newv = o.get(off)
            cur = row[hi[col]].value
            if newv and str(cur or '').strip() != str(newv).strip():
                if cur in (None, ''):
                    row[hi[col]].value = newv
                    row[hi[col]].fill = GREEN
                else:
                    row[hi[col]].value = newv
                    row[hi[col]].fill = RED
                L(rowno, row, col, '官方目录更正' if cur not in (None, '') else '官方目录补全',
                  cur, newv, SRC_TAG)
                n_fix += 1
        n_code += 1

    # --- 2) 占位行配对 ---
    for rowno, code in PAIRING.items():
        row = rows.get(rowno)
        o = officials[code]
        if row is None:
            print(f'!! 行{rowno} 不存在')
            continue
        cur_code = str(row[hi['产品型号']].value or '').strip()
        if not cur_code:
            row[hi['产品型号']].value = code
            row[hi['产品型号']].fill = GREEN
            L(rowno, row, '产品型号', '公告对齐补全', None, code, SRC_TAG + ' 占位行人工复核配对')
        for fld, col, off in (('纯电续航里程(km)', '纯电续航里程(km)', '续航'),
                              ('整备质量(kg)', '整备质量(kg)', '整备'),
                              ('电池容量(kWh)', '电池容量(kWh)', '能量')):
            newv = o.get(off)
            cur = row[hi[col]].value
            if newv and str(cur or '').strip() != str(newv).strip():
                was_empty = cur in (None, '')
                row[hi[col]].value = newv
                row[hi[col]].fill = GREEN if was_empty else RED
                L(rowno, row, col, '公告对齐更正' if not was_empty else '公告对齐补全',
                  cur, newv, f'{SRC_TAG} 官方型号={code}')
                n_fix += 1
        # 动力类型校准（星钽5X 等 PHEV 占位可能标错）
        want_pt = ptype_of_row(o, o['_bev'])
        cur_pt = str(row[hi['动力类型']].value or '').strip()
        if not cur_pt:
            row[hi['动力类型']].value = want_pt
            row[hi['动力类型']].fill = GREEN
            L(rowno, row, '动力类型', '公告对齐补全', None, want_pt, SRC_TAG)
        elif '纯电动' in cur_pt and not o['_bev']:
            row[hi['动力类型']].value = want_pt
            row[hi['动力类型']].fill = RED
            L(rowno, row, '动力类型', '公告对齐更正', cur_pt, want_pt, SRC_TAG + ' 官方目录为插混')
        n_code += 1

    # --- 3) 新增行 ---
    # 仅排除 410 批已有该码的行；其他批次的同码（跨批同码=合法双批，397/402 先例）不阻塞 410 新行
    existing_410_codes = {str(r[hi['产品型号']].value or '').strip() for r in rows.values()
                          if str(r[hi['批次']].value) == '410'}
    new_officials = [o for code, o in officials.items()
                     if code not in used and code not in existing_410_codes]
    added_start = ws.max_row + 1
    for o in new_officials:
        r = [None] * len(hdr)
        r[hi['批次']] = 410
        r[hi['产品型号']] = o['型号']
        r[hi['企业名称']] = o.get('企业') or None
        r[hi['通用名称']] = o.get('通用名称') or None
        r[hi['产品名称']] = o.get('产品名称') or None
        r[hi['动力类型']] = ptype_of_row(o, o['_bev'])
        r[hi['纯电续航里程(km)']] = o.get('续航')
        r[hi['整备质量(kg)']] = o.get('整备')
        r[hi['电池容量(kWh)']] = o.get('能量')
        r[hi['综合油耗(L/100km)']] = o.get('油耗')
        r[hi['发动机排量(mL)']] = o.get('排量')
        r[hi['数据来源']] = SRC_TAG
        ws.append(r)
        for c in range(1, len(hdr) + 1):
            ws.cell(row=ws.max_row, column=c).fill = GREEN
        seq += 1
        ledger.append((seq, 410, o['型号'], r[hi['动力类型']], '整行', '公告新增行', '(新)', None,
                       f'{SRC_TAG}；企业={o.get("企业")}；通用={o.get("通用名称")}'))
        n_new += 1

    # --- 4) 批次时间表 410 公告日期 ---
    if '批次时间表' in wb.sheetnames:
        ts = wb['批次时间表']
        thdr = [str(c.value or '').strip() for c in ts[1]]
        for row in ts.iter_rows(min_row=2):
            vals = [c.value for c in row]
            if any(str(v) == '410' for v in vals):
                for c, h in zip(row, thdr):
                    if '公告' in h and '日期' in h and not c.value:
                        c.value = ANNOUNCE_DATE
                        c.fill = GREEN

    print(f'codes_aligned={n_code} cells_fixed={n_fix} new_rows={n_new} ledger_seq->{seq}')
    if not apply:
        print('DRY-RUN（未写入）')
        return
    wb.save(WB)
    print('saved')


if __name__ == '__main__':
    main()
