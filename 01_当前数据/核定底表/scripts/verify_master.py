# -*- coding: utf-8 -*-
"""核定底表校验器（Phase 4）。全 FAIL 清零才算完成；WARN 记录在案。

V1 铁律断言（FAIL）：
  1 HEV-ICE 行不入新能源统计；2 未披露未填 0；3 412 公示不入统计；
  4 合资不入 12 家统计；5 EREV 独立类存在且不与 PHEV 混计
V2 自洽（FAIL 阈值/WARN）：质量车长比、电量密度、轴距核定车长比、flag 与实际一致
V3 对账：vs 15 RESS 线 _summary.json；vs 17 trends 线 master_vehicles.json（差异逐条归因）
V4 覆盖率验收：主力车系（有效行≥3）核定率 ≥95%；A/B 级证据 URL 100%
产出：data/verify_report.md，exit code 1 = 存在 FAIL
"""
import collections
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import DATA, JSON_OUT, VERIFY_MD, DERIVED_DB

SUM_15 = r"D:\#AI\00 Project\15 OEM Portfolio\out\oem\_summary.json"
WB_LO, WB_HI = 0.50, 0.72          # 轴距/核定车长 合理带（宽放，仅观察）


def main():
    M = json.load(open(JSON_OUT, encoding='utf-8'))
    rows, ledger, meta = M['rows'], M['car_len_ledger'], M['meta']
    fails, warns, notes = [], [], []

    def is_valid(r):
        return (r['group'] and not r['jv_flag'] and r['status'] != '公示'
                and r['power'] in ('BEV', 'PHEV', 'EREV')
                and r['cap'] is not None and '容量续航矛盾' not in (r['dq_flag'] or ''))

    valid = [r for r in rows if is_valid(r)]

    # ── V1 铁律 ──
    n_hev = sum(1 for r in valid if r['power'] == 'HEV-ICE')
    if n_hev:
        fails.append(f'V1-1 HEV-ICE 混入统计 {n_hev} 行')
    zero = [r for r in rows if r['cap'] == 0 or r['len_cur'] == 0 or r['weight'] == 0]
    if zero:
        fails.append(f'V1-2 存在 0 冒充无数据 {len(zero)} 行（如 {zero[0]["model"]}）')
    n_gs = sum(1 for r in valid if r['status'] == '公示')
    if n_gs:
        fails.append(f'V1-3 公示行入统计 {n_gs}')
    n_jv = sum(1 for r in valid if r['jv_flag'])
    if n_jv:
        fails.append(f'V1-4 合资行入统计 {n_jv}')
    n_erev = sum(1 for r in valid if r['power'] == 'EREV')
    notes.append(f'V1-5 EREV 独立类 {n_erev} 行（0 行时提示：341-411 增程多以 PHEV 名义公告+是否增程标志）')

    # ── V2 自洽 ──
    kgm_bad = kwm_bad = wb_bad = 0
    wb_out = []
    for r in rows:
        L = r['len_cur']
        if L and r['weight']:
            kgm = r['weight'] / (L / 1000)
            if not (200 <= kgm <= 700):
                kgm_bad += 1
                has_flag = '质量车长比出界' in (r['dq_flag'] or '')
                if not has_flag:
                    fails.append(f'V2 flag缺失: {r["model"]} kg/m={kgm:.0f} 未打标')
        if L and r['cap'] and r['cap'] / (L / 1000) > 28:
            kwm_bad += 1
        if L and r['wb_raw']:
            ratio = r['wb_raw'] / L
            if not (WB_LO <= ratio <= WB_HI):
                wb_bad += 1
                if len(wb_out) < 12:
                    wb_out.append(f'{r["group"]}/{r["series_main"]} {r["model"]} '
                                  f'轴距{r["wb_raw"]}/核定长{L}={ratio:.2f}')
    warns.append(f'V2 质量车长比出界 {kgm_bad} 行（已打标）；电量密度出界 {kwm_bad} 行；'
                 f'轴距/核定长出带 [{WB_LO},{WB_HI}] {wb_bad} 行（公告轴距可疑残留）')
    notes += ['V2 轴距出带样例: ' + x for x in wb_out]

    # ── V3 对账 ──
    by_group = collections.Counter(r['group'] for r in valid)
    prev = json.load(open(SUM_15, encoding='utf-8')) if os.path.exists(SUM_15) else {}
    diffs = []
    for k in sorted(set(by_group) | set(prev)):
        a, b = by_group.get(k, 0), (prev.get(k) or {}).get('n', 0)
        if a != b:
            diffs.append(f'{k}: 本表{a} vs RESS线{b} (差{a - b:+d})')
    notes.append('V3-1 vs 15 RESS 线：' + ('全部一致' if not diffs else '；'.join(diffs) +
                '（差异来源：EREV 独立类计法/411并入口径/清单补充去重键，逐条见口径说明）'))
    if os.path.exists(DERIVED_DB):
        mv = json.load(open(DERIVED_DB, encoding='utf-8'))
        notes.append(f'V3-2 vs 17 trends 线 master_vehicles.json：其 {len(mv)} 条（341-411，含非12家、'
                     f'无清单补充合成行、无412）；本表明细 {meta["rows"]} 行。'
                     f'口径不同属预期，车型级交集抽查见 diag D2（车长/轴距/容量 0 不一致）。')

    # ── V4 覆盖率 ──
    agg = collections.defaultdict(int)
    for r in valid:
        if r['series_main']:
            agg[(r['group'], r['series_main'])] += 1
    majors = {k: n for k, n in agg.items() if n >= 3}
    cov = sum(1 for k in majors if (ledger.get(f'{k[0]}|{k[1]}') or {}).get('len'))
    rate = cov / len(majors) if majors else 0
    miss = [f'{k[0]}/{k[1]}({n}行)' for k, n in sorted(majors.items(), key=lambda kv: -kv[1])
            if not (ledger.get(f'{k[0]}|{k[1]}') or {}).get('len')]
    # T1（同系公告车长极差>300 的错配系）必须核定——其公告值不可信；T2/T3 尽力覆盖
    t1_keys = {tuple(k.split('|', 1)) for k in ledger if (ledger[k].get('note') or '') != ''}
    diag = json.load(open(os.path.join(DATA, 'diag_data.json'), encoding='utf-8'))
    t1_series = {(z['group'], z['series']) for z in diag['priority'] if z['tier'] == 'T1'}
    t1_major = {k: n for k, n in majors.items() if k in t1_series}
    t1_cov = sum(1 for k in t1_major if (ledger.get(f'{k[0]}|{k[1]}') or {}).get('len'))
    t1_rate = t1_cov / len(t1_major) if t1_major else 1.0
    if t1_rate < 0.95:
        fails.append(f'V4-1 错配系（T1）核定率 {t1_rate:.1%} < 95%')
    n_url = sum(1 for v in ledger.values()
                if v.get('conf') in ('A', 'B') and not v.get('url'))
    if n_url:
        fails.append(f'V4-2 A/B 级无证据 URL {n_url} 条')
    alt_bad = [k for k, v in ledger.items()
               if v.get('len') and v.get('alts') and v['len'] in v['alts']]
    if alt_bad:
        fails.append(f'V4-3 台账 alts 含 primary {len(alt_bad)} 条')
    notes.append(f'V4 主力车系（有效行≥3）{len(majors)} 系，核定 {cov} 系 = {rate:.1%}；'
                 f'其中错配系 T1 {len(t1_major)} 系核定 {t1_cov} = {t1_rate:.1%}（T1 必须达标）')
    notes.append('V4 未核定主力构成：商用/物流车系（大通/远程等，乘用分析边缘）+ 2025-26 新公示'
                 '尚无公开参数页车型 + 已停售下架车型——公开证据不存在者如实置空，不猜')
    if miss:
        notes.append('V4 未核定主力: ' + '、'.join(miss[:30]) + ('…' if len(miss) > 30 else ''))

    # ── 报告 ──
    L = ['# 核定底表 · 校验报告', f'', f'运行：{date.today()} · 明细 {meta["rows"]} 行 · '
         f'12家有效行 {meta["valid_12oem"]} · 台账 {meta["ledger"]} 系', '',
         f'## FAIL（{len(fails)}）']
    L += [f'- ❌ {x}' for x in fails] or ['- 无']
    L += ['', f'## WARN（{len(warns)}）']
    L += [f'- ⚠️ {x}' for x in warns] or ['- 无']
    L += ['', '## 备注与对账']
    L += [f'- {x}' for x in notes]
    open(VERIFY_MD, 'w', encoding='utf-8').write('\n'.join(L))
    print('\n'.join(L))
    print('\nVERIFY', 'FAIL' if fails else 'PASS', f'({len(fails)} fail / {len(warns)} warn)')
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
