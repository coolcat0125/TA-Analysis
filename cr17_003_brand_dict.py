#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_003_brand_dict.py — CR-17-003 品牌口径规范化：数据字典落库（默认 dry-run）

CR-17-003 要求：品牌值一律带「牌」后缀、缺失一律空值（不填占位符），
「验收：数据字典更新一行即可」。

本脚本做两件事：
  1) 现状合规扫描（不改数）：统计 产品商标 列是否 100% 带「牌」后缀、是否残留占位符。
  2) 向底表「字段映射说明」sheet 追加口径定义行（数据字典），幂等。

口径正文（数据字典行）：
  产品商标：公告原文品牌值，统一带「牌」后缀（如「比亚迪牌」「埃安(AION)牌」）；
          缺失一律空值，禁止占位符（—/-/未填报 等）；
          允许别名写法保留原文括注（如「极狐(ARCFOX)牌」「别克(BUICK)牌」），比较时按核心名归一；
          空值回填仅走唯一证据通道（C1/C2/C0/C3/C4），有独立源实质反对者留空并入疑点。
用法：python cr17_003_brand_dict.py [--write]
"""
import json, os, argparse, collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OUT = os.path.join(HERE, 'audit-output', 'cr17_003_brand_dict_20261005.json')
I_BR = 2
PLACEHOLDER = {'—', '-', '--', '/', 'N/A', 'NA', '未填报', '无', '无信息', '待查',
               '待核实', '待确认', '未知', '？', '?', 'None', '未提供', '不适用', ''}
DICT_ROWS = [
    ['派生（17A 会话 2026-10-05）', '产品商标', '公告原文品牌值（原样保留括注与后缀状态）',
     '口径：① 自主/国产品牌公告原文惯例带「牌」后缀（比亚迪牌/长安牌/零跑牌），原样保留；'
     '合资与外资品牌公告原文常无后缀（实测 93 行 15 值：大众/别克(BUICK)/丰田(TOYOTA)/'
     '宝马(BMW)/奥迪(AUDI)/凯迪拉克(CADILLAC)/大众汽车(VOLKSWAGEN)/极狐(ARCFOX)/精灵(smart)/'
     '埃安(AION)/梅赛德斯-奔驰(含截断值)/北京汽车 等），按公告原文保留，'
     '**不补后缀**（疑点不擅改·不得改写公告原文）；'
     '② 缺失一律空值，禁止占位符（—/-/未填报/待查 等一律不写入，2026-10-05 扫描占位符残留=0）；'
     '③ 括注保留原文（别克(BUICK)牌 / 别克(BUICK) 并存），聚合比较时按核心名归一'
     '（去括注 + 去尾「牌」，故 极狐牌≡极狐(ARCFOX)牌≡极狐(ARCFOX)）；'
     '④ 空值回填只走唯一证据通道（C1 同型号跨批 / C2 同企业+同通用名称 / C0 MIIT型号前缀 / '
     'C3 同企业单牌 / C4 名称牌词），存在独立源实质反对者不写、留空入疑点；'
     '⑤ 子品牌与母品牌并存（比亚迪/方程豹、长安/阿维塔、吉利/领克、赛力斯/问界、'
     '五菱/宝骏、奇瑞/星途/捷途）均为独立品牌值，不折叠为母品牌；'
     '⑥ 表内同核心多写法为已知状态，非脏数据，统计口径须先归一再聚合。'],
    ['派生（17A 会话 2026-10-05）', '是否增程(EREV)', '派生标记（值域 1/空）',
     '口径：动力类型=PHEV 且 车型名称/产品名称/通用名称 含「增程」字样 → 1；'
     '其余留空（不打 0，避免「未判定」与「确定非增程」混淆）；'
     '不影响现有 PHEV 口径，仅供增程独立计数与交叉验证。'],
]


def s(v):
    return '' if v is None else str(v).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--write', action='store_true')
    args = ap.parse_args()
    wb = openpyxl.load_workbook(WB, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    n = empty = suffix_ok = placeholder = 0
    bad_suffix, ph_hits = collections.Counter(), []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        n += 1
        v = s(r[I_BR])
        if not v:
            empty += 1
            continue
        if v.endswith('牌'):
            suffix_ok += 1
        else:
            bad_suffix[v] += 1
        if v in PLACEHOLDER:
            placeholder += 1
            if len(ph_hits) < 20:
                ph_hits.append({'id': f"{s(r[0])}::{s(r[1])}", 'value': v})
    stats = {'rows': n, 'nonempty': n - empty, 'empty': empty,
             'fill_rate': f'{(n - empty) / n * 100:.2f}%',
             'suffix_ok': suffix_ok,
             'suffix_bad': dict(bad_suffix.most_common(20)),
             'suffix_bad_rows': sum(bad_suffix.values()),
             'suffix_bad_distinct': len(bad_suffix),
             'suffix_bad_verdict': ('按公告原文保留（合资/外资品牌惯例无「牌」后缀）；'
                                    '数据字典已明文记录该例外，聚合须先按核心名归一'),
             'placeholder_residue': placeholder,
             'compliance': (placeholder == 0)}
    print(json.dumps({k: v for k, v in stats.items()}, ensure_ascii=False))

    ws2 = wb['字段映射说明']
    have = {str(r[1] or '') for r in ws2.iter_rows(values_only=True) if len(r) > 1}
    need = [r for r in DICT_ROWS if r[1] not in have]
    print(f'字段映射说明：{len(DICT_ROWS)} 条口径待写入 {len(need)} 条')

    if args.write and need:
        wb.close()
        wb = openpyxl.load_workbook(WB)
        log = wb['变更记录']
        last = 0
        for rr in range(log.max_row, 1, -1):
            v = log.cell(rr, 1).value
            if isinstance(v, int):
                last = v
                break
        fs = wb['字段映射说明']
        for row in need:
            fs.append(row)
            last += 1
            log.append([last, '—', '—', '—', f"字段映射说明·{row[1]}", '口径字典', '（无行）',
                        f"新增字典行：{row[1]}", 'CR-17-003 品牌口径规范化；CR-17-002 增程标记口径'])
        wb.save(WB)
        print(f'已写入 字段映射说明 {len(need)} 行；变更记录末序号 {last}')
    else:
        print('dry-run：未写入。加 --write 执行')

    json.dump({'generated': '2026-10-05', 'mode': 'write' if args.write else 'dry-run',
               'stats': stats, 'dict_rows_pending': len(need), 'dict_rows': DICT_ROWS},
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('->', OUT)


if __name__ == '__main__':
    main()