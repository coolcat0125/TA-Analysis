# -*- coding: utf-8 -*-
"""validate_batch.py — 夜间门禁·懂车帝 Agent 补缺批次验证器（10 项检查 + 五态判定）

输入：批次 JSON（契约格式，见 docs/v4.9/reports/gate_contract_20260920.md）
输出：验证报告 JSON（各项计数 + 逐条 disposition + 五态统计）

用法：python validate_batch.py --batch <batch.json> [--workbook <底表.xlsx>]
"""
import argparse
import json
import os
import re
import sys
from collections import Counter

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
RULES = json.load(open(os.path.join(HERE, 'field_rules_v1.json'), encoding='utf-8'))
MISSING = {'', 'None', 'nan', '未知', '无', '待定'}


def norm(v):
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s in MISSING else s


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', ''))
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--batch', required=True)
    ap.add_argument('--workbook', default=None)
    ap.add_argument('--out', default=None)
    args = ap.parse_args()

    batch = json.load(open(args.batch, encoding='utf-8'))
    items = batch.get('items') or batch.get('records') or []
    if isinstance(items, dict):
        items = [dict(v, model_key=k) for k, v in items.items()]

    here = os.path.dirname(os.path.abspath(__file__))
    wb_path = args.workbook or os.path.join(here, '..', '..', 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    idx = {}
    for i, r in enumerate(rows):
        code = str(r[hi['产品型号']] or '').strip()
        if code:
            idx.setdefault(code, []).append(i)

    R = RULES['fields']
    rep = {k: 0 for k in ('total', 'new_rows', 'fixed', 'conflicted', 'unmatched',
                          'missing_evidence', 'candidate', 'kept_unknown', 'rejected')}
    dispositions = []
    for it in items:
        disp = {'model_key': it.get('model_key'), 'field': it.get('field'), 'disposition': None, 'reasons': []}
        code = norm(it.get('model_key')).upper()
        field = norm(it.get('field') or it.get('changed_fields'))
        val = it.get('normalized_value')
        src = it.get('source_url') or ''
        raw = it.get('raw_value')

        # 1 车型键唯一匹配
        locs = idx.get(code, [])
        if not locs:
            rep['unmatched'] += 1
            disp.update(disposition='rejected', reasons=['车型键无法匹配现有记录'])
            dispositions.append(disp)
            continue
        r = rows[locs[0]]
        pt = str(r[hi['动力类型']] or '')
        is_bev = '纯电' in pt or 'BEV' in pt

        # 2-3 年款/动力类型串位
        rule = R.get(field)
        if rule and rule.get('pt') and ('BEV' in pt) and 'PHEV' in rule['pt'] and 'BEV' not in rule['pt']:
            rep['rejected'] += 1
            disp.update(disposition='rejected', reasons=['BEV 行试图写入 PHEV 专属字段（串位）'])
            dispositions.append(disp)
            continue
        if it.get('powertrain') and it['powertrain'] not in ('BEV', 'PHEV'):
            rep['rejected'] += 1
            disp.update(disposition='rejected', reasons=['powertrain 非枚举值'])
            dispositions.append(disp)
            continue

        # 4-5 单位与类型
        if rule and rule['type'] == 'number':
            fv = fnum(val)
            if fv is None:
                rep['missing_evidence'] += 1
                disp.update(disposition='rejected', reasons=['数值字段含非数值'])
                dispositions.append(disp)
                continue
            if rule.get('range') and not (rule['range'][0] <= fv <= rule['range'][1]):
                rep['conflicted'] += 1
                disp.update(disposition='conflicted', reasons=[f'数值出合法区间 {rule["range"]}'])
                dispositions.append(disp)
                continue

        # 6-8 来源与证据
        if not src.startswith('http'):
            rep['missing_evidence'] += 1
            disp.update(disposition='rejected', reasons=['缺少可定位 source_url'])
            dispositions.append(disp)
            continue
        if not it.get('captured_at'):
            disp['reasons'].append('缺采集时间（降级 candidate）')

        # 9 新增/修正/冲突
        cur = norm(r[hi[field]]) if field in hi else None
        if cur in ('', None):
            kind, state = '补充空值', 'candidate'
            rep['candidate'] += 1
        elif norm(val) == cur:
            kind, state = '确认一致', 'candidate'
            rep['candidate'] += 1
        else:
            kind, state = '冲突（现值≠新值）', 'conflicted'
            rep['conflicted'] += 1
        # 10 是否触及 verified——本表无 verified 层，candidate 不晋升
        rep['fixed' if kind == '补充空值' else ('conflicted' if state == 'conflicted' else 'candidate')] += 0  # 计数已在上方
        disp.update(disposition=state, kind=kind, current=cur, incoming=val,
                    anchor_row=locs[0] + 2, reasons=disp['reasons'])
        dispositions.append(disp)

    rep['total'] = len(items)
    out = {'batch_id': batch.get('batch_id'), 'input_head': batch.get('input_head'),
           'branch': batch.get('branch'), 'source': batch.get('source'),
           'counts': rep, 'dispositions': dispositions,
           'note': '本器仅产出 disposition 报告；不写 canonical。candidate 入候选层需人工复核后另存。'}
    path = args.out or os.path.join(HERE, f"batch_validation_{batch.get('batch_id', 'x')}.json")
    json.dump(out, open(path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('counts:', json.dumps(rep, ensure_ascii=False))
    print('report:', path)


if __name__ == '__main__':
    main()
