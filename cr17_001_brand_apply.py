#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_001_brand_apply.py — CR-17-001 底表「产品商标」回填执行器（默认 dry-run，只读）

证据通道（强→弱，全部只补空、唯一值）：
  C1 同产品型号跨批已有公告原值        n = 该型号非空行数
  C0 MIIT 型号前缀（厂标代号）唯一牌     n = 该前缀非空行数（仅前缀下牌唯一时才用）
  C2 同企业+同通用名称 唯一牌核心       n = 同族非空行数
  C3 同企业唯一牌核心                  n = 该企业非空行数
  C4 名称牌词命中企业已用牌词（唯一）   n = 命中牌在该企业的非空行数
双源可辩护门槛（AGENTS 09-12「更正须双源可辩护」的收窄版）：
  C1/C2 公告原文级证据（同型号/同企业同族唯一原值）：本身即一手证据，
        仅当**无任何独立源明确反对**才写；有反对但别名等价视作同意。
  C0/C3/C4 推导级证据（型号前缀 / 同企业单牌 / 名称牌词）：
        必须至少 1 个独立源（15号映射表推导 或 车型名牌词）别名等价，方可写。
  三源全不一致 或 存在实质反对且无别名 → HOLD（疑点不擅改），进疑点清单。
支持度门槛：n=1 的 C2/C0 视为薄证据，须另有独立源支持才写。
用法：python cr17_001_brand_apply.py [--write]
"""
import json, os, re, sys, time, argparse, collections
import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from media_fill import BRAND_ALIAS  # noqa: E402

WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
LEDGER = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements',
                      'db', 'brand_backfill_ledger_2026-10-04.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_001_brand_apply_20261005.json')

I_B, I_M, I_BR, I_ENT, I_MN, I_PN, I_GN, I_PT = 0, 1, 2, 3, 4, 5, 6, 9
ALIAS = set()
for a, b in BRAND_ALIAS:
    ALIAS.add((a, b)); ALIAS.add((b, a))


def core(b):
    t = re.sub(r'[（(].*?[)）]', '', str(b or '').strip())
    return t[:-1] if t.endswith('牌') else t


def equiv(a, b):
    if not a or not b:
        return False
    if a == b or a in b or b in a:
        return True
    return (a, b) in ALIAS or (b, a) in ALIAS


def s(v):
    return '' if v is None else str(v).strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--write', action='store_true', help='真正写入 xlsx（默认 dry-run）')
    args = ap.parse_args()

    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    model_brand = collections.defaultdict(collections.Counter)
    ent_fam = collections.defaultdict(collections.Counter)
    ent_brand = collections.defaultdict(collections.Counter)
    pref_brand = collections.defaultdict(collections.Counter)
    raw_of, empties = {}, []
    for idx, r in enumerate(rows, start=2):
        b, m, e = s(r[I_BR]), s(r[I_M]), s(r[I_ENT])
        gn = s(r[I_GN])
        rec = {'row': idx, 'id': f"{s(r[I_B])}::{m}", 'batch': s(r[I_B]), 'model': m,
               'enterprise': e, 'modelName': s(r[I_MN]), 'productName': s(r[I_PN]),
               'genericName': gn, 'powertrain': s(r[I_PT])}
        if not b:
            empties.append(rec); continue
        c = core(b); raw_of.setdefault(c, b)
        if m: model_brand[m][c] += 1
        if e and gn: ent_fam[(e, gn)][c] += 1
        if e: ent_brand[e][c] += 1
        mm = re.match(r'^([A-Za-z]{2,4})', m or '')
        if mm: pref_brand[mm.group(1).upper()][c] += 1

    derived = {r['id']: r.get('brand') or '' for r in json.load(open(LEDGER, encoding='utf-8'))['rows']}

    plan, hold = [], []
    for rec in empties:
        rid, m, e, gn = rec['id'], rec['model'], rec['enterprise'], rec['genericName']
        names = f"{rec['modelName']} {rec['productName']} {gn}"
        src = {}
        # --- 证据通道（公告原文级优先）---
        ev = ch = n = None
        if len(model_brand.get(m) or {}) == 1:
            c, n = next(iter(model_brand[m].items())); ch, ev = 'C1', c
        elif e and gn and len(ent_fam.get((e, gn)) or {}) == 1:
            c, n = next(iter(ent_fam[(e, gn)].items())); ch, ev = 'C2', c
        elif len(ent_brand.get(e, {})) == 1:
            c, n = next(iter(ent_brand[e].items())); ch, ev = 'C3', c
        else:
            used = set(ent_brand.get(e, {}))
            hits = [k for k in used if k and k in names]
            if len(hits) == 1:
                ev = hits[0]; ch = 'C4'; n = ent_brand[e][ev]
        if ch is None:
            mm = re.match(r'^([A-Za-z]{2,4})', m or '')
            if mm and len(pref_brand.get(mm.group(1).upper()) or {}) == 1:
                c, n = next(iter(pref_brand[mm.group(1).upper()].items()))
                ch, ev = 'C0', c
        if ch is None:
            hold.append({**rec, 'verdict': 'HOLD', 'reason': '四通道均无唯一证据（多牌企业/无同族/名称无牌词）',
                         'enterprise_brands': sorted(ent_brand.get(e, {}))})
            continue
        # --- 独立源 A：15号映射表推导 ---
        d = core(derived.get(rid, ''))
        src['map_derivation'] = d or None
        src['agrees_map'] = equiv(ev, d)
        # --- 独立源 B：车型/产品名含牌词 ---
        toks = [k for k in ent_brand.get(e, {}) if k and k in names]
        src['name_token'] = toks[0] if len(toks) == 1 else (toks or None)
        src['agrees_name'] = (len(toks) == 1 and equiv(ev, toks[0]))
        agree = [k for k in ('agrees_map', 'agrees_name') if src[k]]
        # 反对源：有值但与证据不别名等价
        if d and not src['agrees_map']:
            src['oppose_map'] = d
        toks_n = len(toks)
        if toks_n == 1 and not src['agrees_name']:
            src['oppose_name'] = toks[0]
        oppose = [k for k in ('oppose_map', 'oppose_name') if k in src]
        src['oppose_count'] = len(oppose)
        thin = (ch in ('C2', 'C0') and n == 1)
        rec2 = {**rec, 'channel': ch, 'evidence_core': ev, 'evidence_raw': raw_of.get(ev, ev + '牌'),
                'support_n': n, 'derived': derived.get(rid, ''), 'sources': src,
                'agree_count': len(agree), 'oppose_count': len(oppose)}
        if oppose:
            rec2.update(verdict='HOLD', reason=f'存在实质反对源 {"+".join(oppose)}（疑点不擅改）')
            hold.append(rec2)
        elif thin and not agree:
            rec2.update(verdict='HOLD', reason='薄证据（证据域仅1行支撑）且无独立源佐证')
            hold.append(rec2)
        else:
            srcs = (['公告原文'] if ch in ('C1', 'C2') else ['推导级证据']) + \
                   ([f'map:{src["agrees_map"] and "同" or "空"}'] if src.get('agrees_map') else []) + \
                   (['name:同'] if src.get('agrees_name') else [])
            rec2.update(verdict='WRITE', reason=f'{ch} n={n} 佐证[{"/".join(srcs) or "单一公告原文源"}]')
            plan.append(rec2)

    stats = {'empty_total': len(empties), 'write': len(plan), 'hold': len(hold),
             'write_by_channel': dict(collections.Counter(p['channel'] for p in plan)),
             'hold_by_channel': dict(collections.Counter(h.get('channel', 'none') for h in hold))}
    print(json.dumps(stats, ensure_ascii=False))

    if args.write and plan:
        fill = PatternFill('solid', fgColor='FFF2CC')
        cs = wb['颜色说明']
        have = {str(row[1] or '') for row in cs.iter_rows(values_only=True) if len(row) > 1}
        if '品牌证据回填' not in have:
            cs.append(['浅黄底纹', '品牌证据回填',
                       '「产品商标」空值按四通道证据回填：C1 同型号跨批 / C2 同企业+同通用名称 / '
                       'C3 同企业单牌 / C4 名称牌词 / C0 MIIT型号前缀，均只补唯一值；'
                       '存在独立源实质反对者不写，列疑点（详见变更记录与 audit-output 逐格台账）。'])
        log = wb['变更记录']
        last_seq = 0
        for rr in range(log.max_row, 1, -1):
            v = log.cell(rr, 1).value
            if isinstance(v, int):
                last_seq = v
                break
        for p in plan:
            cell = ws.cell(p['row'], I_BR + 1)
            cell.value = p['evidence_raw']
            cell.fill = fill
            last_seq += 1
            log.append([last_seq, p['batch'], p['model'], p.get('powertrain', ''),
                        '产品商标', '证据回填', '(空)', p['evidence_raw'], p['reason']])
        saved = None
        try:
            wb.save(WB); saved = WB
            print('      已就地保存')
        except PermissionError:
            tmp = WB.replace('（341~410批）', '（341~410批）v42tmp')
            wb.save(tmp)
            print(f'      原文件被占用，先写 {os.path.basename(tmp)}，重试替换…')
            for i in range(6):
                time.sleep(20)
                try:
                    os.replace(tmp, WB); saved = WB
                    print(f'      替换成功（重试 {i+1}）'); break
                except PermissionError:
                    pass
            if not saved:
                kept = WB.replace('（341~410批）', '（341~410批·CR17-001版）')
                os.replace(tmp, kept); saved = kept
                print(f'      持续被占用，保留为 {os.path.basename(kept)}')
        print(f'已写入：{len(plan)} 行；变更记录追加 {len(plan)} 条（末序号 {last_seq}）')
    else:
        print('dry-run：未写入。确认后加 --write')

    json.dump({'generated': '2026-10-05', 'mode': 'write' if args.write else 'dry-run',
               'stats': stats, 'plan': plan, 'hold': hold},
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('HOLD 样例：')
    for h in hold[:20]:
        print(' ', h.get('channel', '-'), h['id'], h['enterprise'], '|', h['reason'],
              '| 证据=', h.get('evidence_raw', ''), '| 推导=', h.get('derived', ''))
    print('->', OUT)


if __name__ == '__main__':
    main()