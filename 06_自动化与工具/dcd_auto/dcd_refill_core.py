# -*- coding: utf-8 -*-
"""dcd_refill_core.py — 批次回填规则核心(护栏全自动模式的执行层)

规则血统(勿改语义,改必记 CHANGELOG):
  匹配/打分/只补空/唯一值/动力守卫/电池类型归一 —— 逐条继承 integrate_dcd_refill.py
    (门禁 Agent 2026-09-20/21 已执行口径:302 绿填 + 8 双源红修)
  W5 成对铁律 —— 继承 Update/DCD参数回填_进度跟踪.md 门禁协调记录(2026-09-21 00:45):
    "任何触及 车长/轴距 的回填必须成对评估,缺一项证据整行 hold"
    (09-20 夜 est 红修 153 格因只修车长不修轴距致 W5 爆增 66 → 整批回退)
  比值判据 —— 逐字继承 verify_consistency.py:C4=车长<轴距(CRITICAL);W5=比值出界[1.4,2.4]

红修路径(双源/估算)默认关闭:
  --allow-t4-redfix   开启 T4 补丁∩DCD 双源红修(T4 76 格补丁本身待 Orchestrator/人工批准)
  est 红修不提供开关 —— 该路径 09-20 已实证失控,保留在 integrate_dcd_refill.py 由门禁手动管理
"""
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dcd_common import (BT_NORM, DCD_URL, FIELD_MAP, fnum, is_empty,
                        is_semantic_placeholder, norm_name, parse_power, today)

GREEN_FILL = 'C6EFCE'
RED_FILL = 'FFC7CE'
W5_LO, W5_HI = 1.4, 2.4          # verify_consistency.py W5 车长/轴距区间
W7_LO, W7_HI = 0.18, 0.70        # verify_consistency.py W7 整备/车长区间(车长>2500 时检验)
LEN_FIELD, WB_FIELD = '车长(mm)', '轴距(mm)'
MASS_FIELD = '整备质量(kg)'


def ratio_violations(vals):
    """按 verify_consistency.py 判据检查一组(字段→值)的比值违规,返回 [(pair, reason)]。

    vals 为模拟值(计划值或现值)。C4=车长<轴距(CRITICAL);W5=车长/轴距出界;
    W7=整备/车长出界且车长>2500。
    """
    out = []
    L, A = fnum(vals.get(LEN_FIELD)), fnum(vals.get(WB_FIELD))
    if L is not None and A is not None:
        if L < A:
            out.append(((LEN_FIELD, WB_FIELD), f'C4预防:车长{L:.0f}<轴距{A:.0f}'))
        elif not (W5_LO <= L / A <= W5_HI):
            out.append(((LEN_FIELD, WB_FIELD), f'W5预防:车长/轴距={L / A:.2f} 出界[{W5_LO},{W5_HI}]'))
    M, L2 = fnum(vals.get(MASS_FIELD)), fnum(vals.get(LEN_FIELD))
    if M is not None and L2 is not None and L2 > 2500 and not (W7_LO <= M / L2 <= W7_HI):
        out.append(((MASS_FIELD, LEN_FIELD), f'W7预防:整备/车长={M / L2:.3f} 出界[{W7_LO},{W7_HI}]'))
    return out


# ---------------------------------------------------------------- 搜索→定链
def decide_series(kw, candidates, expect_brand, shared_name):
    """从搜索候选中定链。candidates: [{'sid','series_name'}]。

    规则(对仓库既有 15 例人审结果全量复现,见 test 自测):
      - 共享名(同通用名称跨车系族):必须命中断期品牌令牌(严格)
      - 非共享名:车系名与关键词互为包含(任一方向),或命中断期品牌令牌
    返回 {status: accept|reject|hold, sid, series_name, reason}
    """
    nkw = norm_name(kw)
    if not candidates:
        return {'status': 'reject', 'sid': None, 'series_name': None, 'reason': '搜索无候选车系'}
    accepted = []
    for c in candidates:
        nm = norm_name(c.get('series_name') or '')
        if not nm:
            continue
        hit_brand = any(t and t in nm for t in (expect_brand or []))
        # 包含判定(双向)都要求"被包含的一侧"足够长,防短名误配:
        #   kw⊂series 需 len(kw)>=3 或 kw 为单个 CJK 字(车系'唐'⊂'比亚迪唐')
        #   series⊂kw 需 len(series)>=3 或 series 为单个 CJK 字
        # 反例(2026-09-22 实跑抓到):kw='U5'(爱驰飞碟U5) ⊂ series='北京EU5' → 误配,已拦
        kw_ok = len(nkw) >= 3 or bool(re.search(r'[一-鿿]', nkw))
        nm_ok = len(nm) >= 3 or bool(re.search(r'[一-鿿]', nm))
        contain = bool(nkw) and ((kw_ok and nkw in nm) or (nm_ok and nm in nkw))
        kw_has_brand = any(t and t in nkw for t in (expect_brand or []))
        if shared_name:
            # 共享名:品牌令牌优先;但子品牌(欧拉∈长城、AION∈广汽)不在商标令牌里,
            # 故放行"车系名与关键词实质对应"的形态;短名/部分包含仍须品牌令牌
            # (2026-09-22 实跑:欧拉黑猫/比亚迪唐 曾被误 hold,人审结论为接受)
            ok = (hit_brand or nkw == nm
                  or (len(nm) >= 3 and nm in nkw)
                  or (len(nkw) >= 3 and nkw in nm)
                  or (len(nm) == 1 and re.search(r'[一-鿿]', nm) and kw_has_brand))
        else:
            ok = hit_brand or contain
        if ok:
            accepted.append(c)
    if len(accepted) == 1:
        c = accepted[0]
        return {'status': 'accept', 'sid': str(c['sid']), 'series_name': c.get('series_name'),
                'reason': '品牌/包含校验通过'}
    if not accepted:
        names = '、'.join((c.get('series_name') or '?') for c in candidates[:5])
        toks = '、'.join(t for t in (expect_brand or []) if t) or '(空)'
        if shared_name:
            # 共享名下令牌未命中 ≠ 确认错误:商标可能未收录子品牌(如"几何"),留人工
            return {'status': 'hold', 'sid': None, 'series_name': None,
                    'reason': f'共享名且品牌令牌[{toks}]未命中候选[{names}],疑令牌集缺子品牌,待人工'}
        return {'status': 'reject', 'sid': None, 'series_name': None,
                'reason': f'无候选通过品牌/包含校验(令牌[{toks}] vs 候选[{names}]),疑误配,按仓库人审先例拒收'}
    sids = {str(c['sid']) for c in accepted}
    if len(sids) == 1:
        c = accepted[0]
        return {'status': 'accept', 'sid': str(c['sid']), 'series_name': c.get('series_name'),
                'reason': '多候选中 only 同一 sid'}
    return {'status': 'hold', 'sid': None, 'series_name': None,
            'reason': f'{len(accepted)} 个候选通过校验,车系歧义待人工'}


# ---------------------------------------------------------------- 行匹配
def entry_row_indices(val_rows, hi, entry):
    """队列条目拥有的底表行:通用名称含该 part(占位符过滤后)且车系族一致。

    val_rows: values_only 行元组列表(勿传单元格对象)。
    """
    from dcd_common import generic_parts, row_family
    part = entry['part']
    fam = entry['family']
    out = []
    for idx, r in enumerate(val_rows):
        if part not in generic_parts(r[hi['通用名称']]):
            continue
        if row_family(r[hi['车型名称']], r[hi['产品名称']], r[hi['产品商标']], r[hi['企业名称']]) != fam:
            continue
        out.append(idx)
    return out


# ---------------------------------------------------------------- 款型筛选与打分(继承 integrate)
def select_trims(dcd, fuel_ff, sid, is_bev):
    trims = []
    for c in (dcd.get('cars') or []):
        v = c.get('vals') or {}
        ff = fuel_ff.get((str(sid), str(c.get('id'))), '')
        if ff == '汽油':
            continue
        if is_bev and not v.get('battery_capacity'):
            continue
        if (not is_bev) and not v.get('battery_capacity') and 'DM' not in str(c.get('name')) \
                and 'EV' not in str(c.get('name')):
            continue
        trims.append(c)
    return trims


def score_trims(trims, row_rng, row_cw):
    """续航接近(±15/±30)+ 整备接近(±40/±90)打分,返回 (best_score, tied[])。"""
    scored = []
    for c in trims:
        v = c.get('vals') or {}
        import re as _re
        nm_km = _re.search(r'(\d{3,4})\s*km', str(c.get('name') or ''))
        cltc = fnum(v.get('cltc_recharge_mileage')) or (float(nm_km.group(1)) if nm_km else None)
        cw = fnum(v.get('curb_weight'))
        s = 0
        if row_rng and cltc:
            d = abs(row_rng - cltc)
            s += 0 if d <= 15 else (1 if d <= 30 else 3)
        elif row_rng and not cltc:
            s += 1
        if row_cw and cw:
            d = abs(row_cw - cw)
            s += 0 if d <= 40 else (1 if d <= 90 else 3)
        scored.append((s, c))
    if not scored:
        return None, []
    scored.sort(key=lambda t: t[0])
    best = scored[0][0]
    return best, [c for s, c in scored if s == best]


# ---------------------------------------------------------------- 计划与护栏
def plan_row_fills(row, hi, tied, fields):
    """对一行计算计划绿填:{字段: 值} + 歧义字段列表。只补空;唯一候选值才写。"""
    plan, ambiguous = {}, []
    for fld in fields:
        cell_val = row[hi[fld]]
        if is_semantic_placeholder(cell_val):
            continue
        if not is_empty(cell_val):
            continue
        key = FIELD_MAP[fld]
        vals = set()
        for c in tied:
            v = (c.get('vals') or {}).get(key)
            if fld == '电池类型':
                vals.add(BT_NORM.get(str(v).strip(), str(v).strip()))
            else:
                fv = fnum(v)
                vals.add(fv if fv is not None else str(v or '').strip())
        vals.discard('')
        if len(vals) == 1:
            plan[fld] = vals.pop()
        elif len(vals) > 1:
            ambiguous.append(fld)
    return plan, ambiguous


def current_violation(val_row, hi):
    """行现值是否已违反 C4/W5/W7(疑污染锚点)。返回 reason 或 None。

    val_row: values_only 行元组。
    """
    vals = {f: val_row[hi[f]] for f in (LEN_FIELD, WB_FIELD, MASS_FIELD)}
    v = ratio_violations(vals)
    return v[0][1] if v else None


def w5_pair_guard(val_row, hi, plan):
    """写入前护栏,两层:
      1) W5 成对铁律:计划触碰 车长/轴距 → 成对后两侧都必须有值(缺一项证据该对整体 hold)
      2) 比值预防:按计划值模拟 C4/W5/W7,违规则该对字段 hold(其余字段不受影响)
    返回 (plan, held[]) — held: [{fields, reason}]
    val_row: values_only 行元组。
    """
    held = []
    touches = LEN_FIELD in plan or WB_FIELD in plan
    if not touches:
        return plan, held
    L = fnum(plan.get(LEN_FIELD, val_row[hi[LEN_FIELD]]))
    A = fnum(plan.get(WB_FIELD, val_row[hi[WB_FIELD]]))
    if L is None or A is None:
        for f in (LEN_FIELD, WB_FIELD):
            plan.pop(f, None)
        held.append({'fields': [LEN_FIELD, WB_FIELD],
                     'reason': 'W5成对铁律:车长/轴距缺一项证据,该对整体 hold'})
        return plan, held
    sim = {LEN_FIELD: L, WB_FIELD: A,
           MASS_FIELD: fnum(plan.get(MASS_FIELD, val_row[hi[MASS_FIELD]]))}
    for pair, reason in ratio_violations(sim):
        for f in pair:
            plan.pop(f, None)
        held.append({'fields': list(pair), 'reason': reason})
    return plan, held


# ---------------------------------------------------------------- 应用一个车系
def apply_entry(wb, ws, ledger, hi, seq, entry, dcd, fuel_ff, sid, fields,
                allow_t4_redfix=False, t4=None, dry_run=True):
    """把单个队列条目(车系)的 DCD 数据应用到底表。返回 (report, seq)。

    wb/ws/ledger 为已 load 的工作簿对象(调用方负责 save/回退)。
    """
    import openpyxl
    from openpyxl.styles import PatternFill
    green = PatternFill('solid', fgColor=GREEN_FILL)
    red = PatternFill('solid', fgColor=RED_FILL)

    rows = list(ws.iter_rows(min_row=2))
    val_rows = [tuple(c.value for c in r) for r in rows]   # 匹配用值视图
    idxs = entry_row_indices(val_rows, hi, entry)
    rep = {'qkey': entry['qkey'], 'family': entry['family'], 'part': entry['part'], 'sid': sid,
           'rows_total': len(idxs), 'green_fill': 0, 'red_fix': 0,
           'ambiguous_skip': 0, 'held': [], 'fills': [], 'errors': []}
    date = today()
    for idx in idxs:
        vrow = val_rows[idx]          # 值视图:供所有行级函数
        row = rows[idx]               # 单元格视图:仅供写入
        pt = str(vrow[hi['动力类型']] or '')
        if not ('纯电' in pt or '插电' in pt or '增程' in pt or 'BEV' in pt or 'PHEV' in pt):
            continue
        is_bev = '纯电' in pt or 'BEV' in pt
        # 污染锚点预检:现值已违反 C4/W5/W7 的行整体 hold,不在污染值旁盖新值
        cv = current_violation(vrow, hi)
        if cv:
            rep['held'].append({'row': idx + 2, 'code': str(vrow[hi['产品型号']] or ''),
                                'fields': ['(整行)'], 'reason': f'现存比值违规-疑污染锚点:{cv}'})
            continue
        trims = select_trims(dcd, fuel_ff, sid, is_bev)
        if not trims:
            rep['errors'].append({'row': idx + 2, 'reason': '无可用款型(动力守卫过滤后为空)'})
            continue
        best, tied = score_trims(trims, fnum(vrow[hi['纯电续航里程(km)']]),
                                 fnum(vrow[hi['整备质量(kg)']]))
        if not tied:
            rep['errors'].append({'row': idx + 2, 'reason': '打分后无候选款型'})
            continue
        plan, ambiguous = plan_row_fills(vrow, hi, tied, fields)
        if ambiguous:
            rep['ambiguous_skip'] += len(ambiguous)
        plan, held = w5_pair_guard(vrow, hi, plan)
        for h in held:
            h['row'] = idx + 2
            h['code'] = str(vrow[hi['产品型号']] or '')
            rep['held'].append(h)
        code = str(vrow[hi['产品型号']] or '').strip()
        batch = vrow[hi['批次']]
        for fld, val in plan.items():
            cell = row[hi[fld]]
            old = cell.value
            if not dry_run:
                cell.value = val
                cell.fill = green
                seq += 1
                ledger.append((seq, batch, code, pt, fld, '媒体参数补全', '(空)', str(val),
                               f'懂车帝参数页(sid={sid})：{DCD_URL.format(sid=sid, date=date)}'))
            rep['green_fill'] += 1
            rep['fills'].append({'row': idx + 2, 'code': code, 'field': fld,
                                 'old': '(空)', 'new': str(val), 'kind': 'green'})
        # 双源红修(默认关):T4 补丁∩DCD 唯一值一致才更正污染值
        if allow_t4_redfix and t4:
            for fld in (LEN_FIELD, WB_FIELD):
                if fld in plan:
                    continue
                k4 = (str(batch or '').strip(), code, fld.split('(')[0])
                cur = fnum(vrow[hi[fld]])
                if k4 in t4 and cur is not None and cur != fnum(t4[k4]):
                    cand = {fnum(c['vals'].get(FIELD_MAP[fld])) for c in tied}
                    cand.discard(None)
                    if cand == {fnum(t4[k4])}:
                        old = row[hi[fld]].value
                        if not dry_run:
                            row[hi[fld]].value = fnum(t4[k4])
                            row[hi[fld]].fill = red
                            seq += 1
                            ledger.append((seq, batch, code, pt, fld, '双源更正', str(old), t4[k4],
                                           f'懂车帝参数页(sid={sid}) 与 T4补丁 双源一致：'
                                           f'{DCD_URL.format(sid=sid, date=date)}'))
                        rep['red_fix'] += 1
                        rep['fills'].append({'row': idx + 2, 'code': code, 'field': fld,
                                             'old': str(old), 'new': t4[k4], 'kind': 'red'})
    return rep, seq


def ledger_max_seq(ledger):
    return max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
                if isinstance(r[0], int)), default=0)
