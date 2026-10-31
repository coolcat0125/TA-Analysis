# -*- coding: utf-8 -*-
"""NEV 公告核定底表构建器（Phase 3，幂等一键重建）。

输入：合并版 xlsx（341~410 主表 + 409-410 新增清单）+ 411/412 扫描底表 + 车长核定台账 + 别名台账
输出：xlsx（8 sheet）/ csv（明细）/ json（全量）

铁律（15 交接 H2 + 17 口径双原则）内化：
- HEV 归 ICE 类（不可外接充电）；EREV 独立统计类；多标签按 EREV>PHEV>BEV>ICE
- 未披露一律 null，禁止填 0 冒充无数据
- 合资留行打标（jv_flag），不入 12 家统计
- 412 公示单列（status=公示），默认排除统计
- 电压等级=补值非公告，独立来源标注
- 车长：核定值优先（A公告>B媒体>C待复核），未核定置 null 不猜；公告原值并列保留
"""
import collections
import csv
import json
import os
import sys
from datetime import date

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (SRC_XLSX, SCAN_XLSX, DATA, XLSX_OUT, CSV_OUT, JSON_OUT,
                    CAR_LEN_JSON, ALIAS_JSON, num, ranges, norm_mat,
                    is_ptype_name, year_of_batch)
from oem_groups import GROUPS, BOARDS, SERIES_BOARD, MAIN_BOARDS, board_of, attr_group
from series_norm import norm_series, load_alias
import volt_rules as V

SPREAD_FLAG = 400            # 系级公告车长极差 → 打「混乱」标
KGM_LO, KGM_HI = 200, 700    # 整备质量/核定车长 kg/m 合理区间
KWM_MAX = 28                 # 电量密度 kWh/m 上限

# H1 行数口径对照（15 交接 H1：用户给数；实测以本表「有效行」口径复核）
H1_USER = {'nio': 17, 'xiaomi': 24, 'tesla': 31, 'liauto': 83, 'xpeng': 100,
           'gwm': 250, 'changan': 300, 'gac': 362, 'chery': 386, 'geely': 464, 'saic': 548}


def norm_power(raw, erev, pname=''):
    s = str(raw or '')
    up = s.upper()
    if erev or '增程' in s or '增程' in str(pname or ''):
        return 'EREV'
    if 'BEV' in up or '纯电' in s or '换电' in s:
        return 'BEV'
    if 'PHEV' in up or '插电' in s:
        return 'PHEV'
    if 'HEV' in up or '混合' in s:
        return 'HEV-ICE'
    if 'FCEV' in up or '燃料' in s:
        return 'FCEV'
    return '其他'


def load_main_and_extra():
    """主表 + 409-410 清单 → 行 dict 列表。"""
    wb = openpyxl.load_workbook(SRC_XLSX, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    I = {h: i for i, h in enumerate(hdr)}
    need = ['批次', '产品型号', '产品商标', '企业名称', '车型名称', '产品名称', '通用名称',
            '产品类型', '细分市场', '动力类型', '整备质量(kg)', '轴距(mm)',
            '纯电续航里程(km)', '电池容量(kWh)', '电池类型', '电池能量密度(Wh/kg)',
            '车长(mm)', '电芯供应商', '电池包供应商', '是否增程(EREV)']
    for k in need:
        assert k in I, f'主表缺列 {k}'
    out = []
    for r in rows:
        if not r[I['产品型号']]:
            continue
        erev = str(r[I['是否增程(EREV)']] or '').strip() in ('是', 'Y', 'y', 'true', 'True')
        out.append({
            'batch': int(num(r[I['批次']]) or 0),
            'status': '公示' if str(num(r[I['批次']])) == '412' else '正式',
            'model': str(r[I['产品型号']]).strip(), 'mark': str(r[I['产品商标']] or ''),
            'ent': str(r[I['企业名称']] or ''), 'name': str(r[I['车型名称']] or ''),
            'common_raw': str(r[I['通用名称']] or ''), 'pname': str(r[I['产品名称']] or ''),
            'ptype': str(r[I['产品类型']] or ''), 'segment': str(r[I['细分市场']] or ''),
            'power_raw': str(r[I['动力类型']] or ''), 'power': norm_power(r[I['动力类型']], erev, r[I['产品名称']]),
            'is_erev': erev,
            'range_raw': r[I['纯电续航里程(km)']], 'cap': num(r[I['电池容量(kWh)']]),
            'cap_raw': r[I['电池容量(kWh)']], 'mat': norm_mat(r[I['电池类型']]),
            'mat_raw': r[I['电池类型']], 'edens': num(r[I['电池能量密度(Wh/kg)']]),
            'cell_sup': r[I['电芯供应商']], 'pack_sup': r[I['电池包供应商']],
            'weight': num(r[I['整备质量(kg)']]), 'wb_raw': num(r[I['轴距(mm)']]),
            'len_raw': num(r[I['车长(mm)']]), 'src_url': '', 'year': year_of_batch(int(num(r[I['批次']]) or 0)),
        })
    # 409-410 清单（主表未收录的新申报；无车长/能量密度/供应商）
    ws2 = wb['409-410批新增清单']
    r2 = ws2.iter_rows(values_only=True)
    h2 = list(next(r2))
    I2 = {h: i for i, h in enumerate(h2)}
    have = {(r['common_raw'], r['cap']) for r in out}
    n_add = 0
    for r in r2:
        if not r or not r[I2['产品商标']]:
            continue
        # 清单表无「通用名称/产品名称」列：车系名取「车型名称」，动力推断用「产品类型」
        common = str(r[I2['车型名称']] or '').strip()
        cap = num(r[I2['电池容量(kWh)']])
        if (common, cap) in have:
            continue
        b = int(num(r[I2['批次']]) or 0)
        pname = str(r[I2['产品类型']] or '')
        out.append({
            'batch': b, 'status': '清单补充', 'model': f"NEW-{str(r[I2['产品商标']]).strip()}-{common}-{cap}",
            'mark': str(r[I2['产品商标']] or ''), 'ent': str(r[I2['企业名称']] or ''),
            'name': str(r[I2['车型名称']] or ''), 'common_raw': common, 'pname': pname,
            'ptype': str(r[I2['产品类型']] or ''), 'segment': '',
            'power_raw': str(r[I2['动力类型']] or ''),
            'power': norm_power(r[I2['动力类型']], False, pname), 'is_erev': False,
            'range_raw': r[I2['纯电续航(km)']], 'cap': cap,
            'cap_raw': r[I2['电池容量(kWh)']], 'mat': norm_mat(r[I2['电池类型']]),
            'mat_raw': r[I2['电池类型']], 'edens': None, 'cell_sup': None, 'pack_sup': None,
            'weight': num(r[I2['整备质量(kg)']]), 'wb_raw': num(r[I2['轴距(mm)']]),
            'len_raw': None,
            'src_url': '', 'year': year_of_batch(b),
        })
        n_add += 1
    wb.close()
    print(f'主表 {len(out) - n_add} 行 + 清单补充 {n_add} 行')
    return out


def load_scan():
    """411 正式 + 412 公示（扫描底表，含官方车长/来源URL）。"""
    wb = openpyxl.load_workbook(SCAN_XLSX, read_only=True, data_only=True)
    out = []
    for sn, status in (('411正式车型', '正式'), ('412公示车型', '公示')):
        ws = wb[sn]
        it = ws.iter_rows(values_only=True)
        next(it); next(it)
        hdr = list(next(it))
        J = {h: i for i, h in enumerate(hdr)}
        for r in it:
            if not r or not r[J['产品型号']]:
                continue
            b = int(num(r[J['批次']]) or (411 if status == '正式' else 412))
            pname = str(r[J['产品名称']] or '')
            erev = str(r[J['是否增程']] or '').strip() in ('是', 'Y')
            out.append({
                'batch': b, 'status': status, 'model': str(r[J['产品型号']]).strip(),
                'mark': str(r[J['产品商标']] or ''), 'ent': str(r[J['企业名称']] or ''),
                'name': '', 'common_raw': str(r[J['通用名称']] or ''), 'pname': pname,
                'ptype': pname, 'segment': '',
                'power_raw': str(r[J['技术路线']] or ''),
                'power': norm_power(r[J['技术路线']], erev, pname), 'is_erev': erev,
                'range_raw': r[J['纯电续航(km)']], 'cap': num(r[J['电池容量(kWh)']]),
                'cap_raw': r[J['电池容量(kWh)']], 'mat': norm_mat(r[J['电池类型']]),
                'mat_raw': r[J['电池类型']], 'edens': None,
                'cell_sup': r[J['电芯供应商']], 'pack_sup': r[J['电池包供应商']],
                'weight': num(r[J['整备质量(kg)']]), 'wb_raw': num(r[J['轴距(mm)']]),
                'len_raw': num(r[J['车长(mm)']]),
                'src_url': str(r[J['来源URL']] or ''), 'year': year_of_batch(b),
            })
    wb.close()
    print(f'扫描底表 {len(out)} 行（411 正式 + 412 公示）')
    return out


def enrich(rows):
    """归属 / 归一化 / 台账 / 电压 / dq 标志。"""
    ledger = json.load(open(CAR_LEN_JSON, encoding='utf-8')) if os.path.exists(CAR_LEN_JSON) else {}
    alias = load_alias()
    # 系级公告车长极差（打标用）
    spread_bad = set()
    g = collections.defaultdict(list)
    for r in rows:
        key, jv, why = attr_group(r['mark'], r['ent'])
        r['group'], r['jv_flag'], r['jv_reason'] = key, jv, (why if jv else '')
        board = None
        if key:
            for sub, nm in SERIES_BOARD.get(key, []):
                if sub and sub in r['common_raw']:
                    board = nm
                    break
            if board is None:
                board = board_of(r['mark'], r['ent'], BOARDS[key])
        r['board'] = board
        r['in_main'] = bool(key and board in MAIN_BOARDS.get(key, []))
        # 车系主名：占位型产品类型 → null + 标志
        sr = ''
        for k in ('common_raw', 'name', 'pname'):
            v = str(r[k] or '').strip()
            if v and not is_ptype_name(v):
                sr = v
                break
        r['series_raw'] = sr
        r['name_is_ptype'] = bool(sr == '' and r['common_raw'] and is_ptype_name(r['common_raw']))
        r['series_main'] = norm_series(sr, key, board) if (key and sr) else None
        if r['name_is_ptype'] or not sr:
            r['series_main'] = None
        if key and sr:
            g[(key, r['series_main'])].append(r['len_raw'])
    for (key, s), lens in g.items():
        vals = [x for x in lens if x is not None]
        if vals and max(vals) - min(vals) > SPREAD_FLAG:
            spread_bad.add((key, s))
    # 电压 + 台账 + dq
    for r in rows:
        V.set_oem(r['group'])
        r['volt'], r['volt_conf'] = V.volt_of(f"{r['common_raw']} {r['mark']} {r['name']}")
        r['volt_source'] = '补值(非公告)' if r['group'] else ''
        ent = ledger.get(f"{r['group']}|{r['series_main']}") if (r['group'] and r['series_main']) else None
        if ent and ent.get('len'):
            r['len_cur'] = int(ent['len'])
            r['len_src'] = ent['src']
            r['len_conf'] = ent['conf']
            r['len_note'] = ent.get('cross') or ''
            r['wb_cur'] = int(ent['wb']) if ent.get('wb') else None
            r['len_url'] = ent.get('url') or ''
        else:
            r.update(len_cur=None, len_src='未核定', len_conf=None, len_note='',
                     wb_cur=None, len_url='')
        flags = []
        if r['cap'] is None:
            flags.append('无容量')
        else:
            rng = ranges(r['range_raw'])
            if r['cap'] < 5 and rng and max(rng) > 150:
                flags.append('容量续航矛盾')
        if r['weight'] and r['len_cur']:
            kgm = r['weight'] / (r['len_cur'] / 1000)
            if not (KGM_LO <= kgm <= KGM_HI):
                flags.append(f'质量车长比出界({kgm:.0f}kg/m)')
        if r['cap'] and r['len_cur']:
            if r['cap'] / (r['len_cur'] / 1000) > KWM_MAX:
                flags.append('电量密度出界')
        if (r['group'], r['series_main']) in spread_bad:
            flags.append('系公告车长混乱')
        if r['name_is_ptype']:
            flags.append('通用名为产品类型')
        r['dq_flag'] = ';'.join(flags)
        r['alias_used'] = alias.get(r['series_main'], '') if r['series_main'] else ''
    return rows, spread_bad


def valid_row(r):
    """「有效行」口径：12 家内非合资 × 新能源类 × 有容量 × 无矛盾。"""
    return (r['group'] and not r['jv_flag'] and r['status'] != '公示'
            and r['power'] in ('BEV', 'PHEV', 'EREV')
            and r['cap'] is not None and '容量续航矛盾' not in r['dq_flag'])


def build_series_master(rows):
    agg = {}
    for r in rows:
        if not (r['group'] and r['series_main']):
            continue
        k = (r['group'], r['series_main'])
        a = agg.setdefault(k, {'group': r['group'], 'series': r['series_main'],
                               'boards': set(), 'rows': 0, 'rows_valid': 0,
                               'b0': 9999, 'b1': 0, 'powers': set(),
                               'raw_names': collections.Counter(),
                               'len_cur': r['len_cur'], 'len_src': r['len_src'],
                               'len_conf': r['len_conf'], 'len_note': r['len_note'],
                               'jv_any': False})
        a['boards'].add(r['board'])
        a['rows'] += 1
        a['rows_valid'] += 1 if valid_row(r) else 0
        a['b0'] = min(a['b0'], r['batch'])
        a['b1'] = max(a['b1'], r['batch'])
        a['powers'].add(r['power'])
        if r['series_raw']:
            a['raw_names'][r['series_raw']] += 1
        a['jv_any'] = a['jv_any'] or r['jv_flag']
    out = []
    for a in agg.values():
        out.append({**{k: v for k, v in a.items() if k not in ('boards', 'powers', 'raw_names')},
                    'boards': '/'.join(sorted(x for x in a['boards'] if x)),
                    'powers': '/'.join(sorted(a['powers'])),
                    'raw_names': ' | '.join(f'{k}×{n}' for k, n in a['raw_names'].most_common(4))})
    return sorted(out, key=lambda z: (z['group'], -z['rows']))


def write_xlsx(rows, series_master, ledger, alias, spread_bad):
    wb = openpyxl.Workbook()
    thin = openpyxl.styles.Font(size=10)

    def sheet(ws, header, data_rows, widths=None):
        ws.append(header)
        for c in ws[1]:
            c.font = openpyxl.styles.Font(bold=True, size=10)
        for row in data_rows:
            ws.append(row)
        if widths:
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.freeze_panes = 'A2'

    # ① 明细全量
    ws = wb.active
    ws.title = '明细全量'
    hdr = ['批次', '年份', '状态', '产品型号', '产品商标', '企业名称', '集团', '板块', '合资剔除',
           '主口径内', '车系主名', '车系原始名', '通用名称', '产品类型', '细分市场',
           '动力类型(原)', '统计动力类', '增程', '续航(原)', '容量kWh', '容量(原)', '电池类型',
           '能量密度', '电芯供应商', '电池包供应商', '整备质量', '轴距(公告)', '车长(公告)',
           '车长(核定)', '车长来源', '车长置信', '车长核验', '轴距(核定)', '车长证据URL',
           '电压等级(补值)', '电压置信', '数据质量标志', '来源URL']
    keys = ['batch', 'year', 'status', 'model', 'mark', 'ent', 'group', 'board', 'jv_flag',
            'in_main', 'series_main', 'series_raw', 'common_raw', 'ptype', 'segment',
            'power_raw', 'power', 'is_erev', 'range_raw', 'cap', 'cap_raw', 'mat',
            'edens', 'cell_sup', 'pack_sup', 'weight', 'wb_raw', 'len_raw',
            'len_cur', 'len_src', 'len_conf', 'len_note', 'wb_cur', 'len_url',
            'volt', 'volt_conf', 'dq_flag', 'src_url']
    sheet(ws, hdr, ([r.get(k) for k in keys] for r in rows),
          widths=[6, 6, 8, 20, 10, 22, 8, 8, 8, 8, 14, 16, 16, 14, 10, 10, 10, 6, 12, 9, 12,
                  12, 9, 14, 14, 9, 10, 10, 10, 11, 8, 12, 10, 26, 12, 8, 22, 24])

    # ② 12 家统计视图（含 H1 对照）
    ws2 = wb.create_sheet('12家统计视图')
    stat_rows = []
    sum_path = r"D:\#AI\00 Project\15 OEM Portfolio\out\oem\_summary.json"
    prev = json.load(open(sum_path, encoding='utf-8')) if os.path.exists(sum_path) else {}
    for key, (cn, grp, *_rest) in GROUPS.items():
        rs = [r for r in rows if r['group'] == key]
        ok = [r for r in rs if valid_row(r)]
        pw = collections.Counter(r['power'] for r in ok)
        lens_ok = sum(1 for r in ok if r['len_cur'])
        stat_rows.append([key, cn, grp, len(rs), len(ok), pw.get('BEV', 0), pw.get('PHEV', 0),
                          pw.get('EREV', 0),
                          len({r['series_main'] for r in ok if r['series_main']}),
                          lens_ok, round(lens_ok / len(ok), 3) if ok else None,
                          prev.get(key, {}).get('n'), H1_USER.get(key)])
    sheet(ws2, ['key', '集团', '分组', '全行(含合资/公示)', '有效行', 'BEV', 'PHEV', 'EREV',
                '车系数', '核定覆盖行', '覆盖率', 'RESS线n(对账)', 'H1用户给数'],
          stat_rows)

    # ③ 车系主数据
    ws3 = wb.create_sheet('车系主数据')
    sheet(ws3, ['集团', '车系主名', '板块', '统计动力类', '全行', '有效行', '首批', '末批',
                '车长核定', '来源', '置信', '核验', '公告原始名(去重)'],
          ([s['group'], s['series'], s['boards'], s['powers'], s['rows'], s['rows_valid'],
            s['b0'], s['b1'], s['len_cur'], s['len_src'], s['len_conf'], s['len_note'],
            s['raw_names']] for s in series_master))

    # ④ 车长核定台账
    ws4 = wb.create_sheet('车长核定台账')
    ledger_rows = ([v['group'], v['series'], v['len'], v['wb'], v['conf'], v['src'],
                    v['url'], v['cross'], v.get('note', ''),
                    '/'.join(str(x) for x in v.get('alts') or []), v['verified']]
                   for v in ledger.values())
    sheet(ws4, ['集团', '车系主名', '车长mm', '轴距mm', '置信', '来源', '证据URL', '交叉核验',
                '备注', '代际多值alts', '核定期'], ledger_rows,
          widths=[8, 14, 9, 9, 6, 12, 40, 18, 24, 18, 11])

    # ⑤ 别名台账
    ws5 = wb.create_sheet('别名台账')
    sheet(ws5, ['公告异写(归一化后)', '车系主名'], alias.items())

    # ⑥ 剔除与存疑
    ws6 = wb.create_sheet('剔除与存疑')
    sus = [[r['batch'], r['group'] or '', r['mark'], r['model'], r['common_raw'],
            r['dq_flag'], r['cap'], r['len_raw'], r['len_cur']]
           for r in rows if r['dq_flag']]
    sheet(ws6, ['批次', '集团', '商标', '产品型号', '通用名称', '数据质量标志', '容量kWh',
                '车长(公告)', '车长(核定)'], sus)

    # ⑦ 批次覆盖
    ws7 = wb.create_sheet('批次覆盖')
    bt = collections.defaultdict(lambda: [0, 0, set()])
    for r in rows:
        b = bt[r['batch']]
        b[0] += 1
        if valid_row(r):
            b[1] += 1
        b[2].add(r['status'])
    sheet(ws7, ['批次', '年份', '状态', '全行', '12家有效行'],
          [[k, year_of_batch(k), '/'.join(sorted(v[2])), v[0], v[1]] for k, v in sorted(bt.items())])

    # ⑧ 口径说明
    ws8 = wb.create_sheet('口径说明')
    lines = [
        ['NEV 公告核定底表 口径说明', f'构建 {date.today()} · 范围 341~411 正式 + 412 公示(默认排除统计) + 409-410 清单补充'],
        ['源表提示', '合并版 xlsx 文件名为旧名「341~410批」，实际已含 411 正式 46 行与 412 公示 58 行（2026-10-19 冻结前并入）'],
        ['行口径', '明细全量=合并版主表5470+清单补充+411正式47+412公示58；一行=一条公告型号记录'],
        ['有效行', '12家集团内 × 非合资 × 统计动力类∈{BEV,PHEV,EREV} × 有容量 × 无容量续航矛盾'],
        ['HEV铁律', 'HEV/混合动力(不可外接充电)归 ICE 类，不进新能源统计（power=HEV-ICE）'],
        ['EREV', '增程行单列 EREV 统计类（多标签按 EREV>PHEV>BEV>ICE 取主类）'],
        ['412公示', 'status=公示，默认排除一切统计；仅供车长核定参考（其尺寸为官方公示值）'],
        ['null铁律', '未披露/未核定一律空值，禁止填 0 冒充无数据'],
        ['合资', '留行打标 jv_flag=是，不入12家统计；排除词精确到「上汽通用汽车」不误伤五菱'],
        ['车长', '车长(核定)=台账值：A=公告官方(411/412 art页) B=汽车之家参数 C=待复核；未核定=空'],
        ['车长局限', '核定值为车系当前在售代际主长度；换代差异记在台账 alts；行级逐型号长度待17侧官方通道扩展'],
        ['电压', '公告无此字段，volt=按平台+公开配置的补值，置信A/B/C，未命中=待核实'],
        ['双版本', '400V/800V并存车系：明细按公告行保留；统计单元=(动力类×容量×电压)拆分由下游执行'],
        ['H1行数', '「H1用户给数」为15交接文档中需求方原口径，与源表实测差异未对账，不反向凑数；'],
        ['', '对账列 RESS线n 为 15 out/oem/_summary.json 口径（含主表+411+清单，剔除无容量/矛盾行）'],
        ['证据', 'A/B级均带证据URL；台账 alts=各来源出现过的代际长度集合'],
    ]
    for row in lines:
        ws8.append(row)
    ws8.column_dimensions['A'].width = 12
    ws8.column_dimensions['B'].width = 130

    wb.save(XLSX_OUT)
    print('SAVED', XLSX_OUT)


def main():
    os.makedirs(DATA, exist_ok=True)
    rows = load_main_and_extra()
    scan = load_scan()
    # 主表文件名 341~410 为旧名，实际已含 411 正式 46 行 + 412 公示 58 行（10-19 冻结前并入）。
    # 扫描底表按型号回填来源URL/车长等官方字段；主表没有的型号才整行补入。
    by_model = {r['model']: r for r in rows}
    n_url = n_new = 0
    for s in scan:
        tgt = by_model.get(s['model'])
        if tgt is None:
            rows.append(s)
            by_model[s['model']] = s
            n_new += 1
            continue
        if not tgt['src_url'] and s['src_url']:
            tgt['src_url'] = s['src_url']
            n_url += 1
        if tgt['len_raw'] is None and s['len_raw'] is not None:
            tgt['len_raw'] = s['len_raw']
        if tgt['weight'] is None and s['weight'] is not None:
            tgt['weight'] = s['weight']
        if tgt['cap'] is None and s['cap'] is not None:
            tgt['cap'], tgt['cap_raw'] = s['cap'], s['cap_raw']
    print(f'合并后总行 {len(rows)}（扫描底表补入 {n_new} 行，回填来源URL {n_url} 行）')

    rows, spread_bad = enrich(rows)
    ledger = json.load(open(CAR_LEN_JSON, encoding='utf-8')) if os.path.exists(CAR_LEN_JSON) else {}
    alias = load_alias()
    series_master = build_series_master(rows)

    write_xlsx(rows, series_master, ledger, alias, spread_bad)

    # csv / json 镜像
    keys = ['batch', 'year', 'status', 'model', 'mark', 'ent', 'group', 'board', 'jv_flag',
            'in_main', 'series_main', 'series_raw', 'common_raw', 'ptype', 'segment',
            'power_raw', 'power', 'is_erev', 'range_raw', 'cap', 'mat', 'edens', 'weight',
            'wb_raw', 'len_raw', 'len_cur', 'len_src', 'len_conf', 'len_note', 'wb_cur',
            'len_url', 'volt', 'volt_conf', 'dq_flag', 'src_url']
    with open(CSV_OUT, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(keys)
        for r in rows:
            w.writerow([r.get(k) for k in keys])
    json.dump({'meta': {'built': str(date.today()),
                        'rows': len(rows),
                        'valid_12oem': sum(1 for r in rows if valid_row(r)),
                        'ledger': len(ledger)},
               'rows': [{k: r.get(k) for k in keys} for r in rows],
               'series_master': series_master,
               'car_len_ledger': ledger,
               'alias': alias},
              open(JSON_OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('SAVED', CSV_OUT, JSON_OUT)
    nv = sum(1 for r in rows if valid_row(r))
    print(f'12 家有效行 {nv} | 车系主数据 {len(series_master)} 系')


if __name__ == '__main__':
    main()
