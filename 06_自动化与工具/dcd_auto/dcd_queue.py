# -*- coding: utf-8 -*-
"""dcd_queue.py — 缺口优先队列生成器

扫底表 9 个 DCD 可填字段(默认)的缺口,按 (车型名称族, 通用名称) 聚成车系,
按缺口格数降序生成优先队列,供抽取驱动按批次消费。

用法:
    python dcd_queue.py                    # 重建队列
    python dcd_queue.py --include-motor    # 电机功率/扭矩缺口也计入优先级(默认不计)
    python dcd_queue.py --force            # 不过滤已抽取车系(默认跳过 mapping 里已人审接受的)

输出: dcd_auto/state/queue.json
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dcd_common import (FIELD_MAP, MOTOR_FIELD_MAP, QUEUE_FILE, RAW_DCD, STATE,
                        brand_tokens, ensure_dirs, generic_parts, is_empty,
                        load_rows, norm_name, qkey, row_family, today)

# 同一通用名称出现在多个车系族下 → 共享名,抽取时须强制品牌令牌匹配(见 dcd_refill_core.decide_series)
def build(fields, force=False):
    hdr, hi, rows = load_rows()
    fam_of = lambda r: row_family(r[hi['车型名称']], r[hi['产品名称']],
                                  r[hi['产品商标']], r[hi['企业名称']])
    # 已人审接受的车系:不再排除(它们可能仍有未填缺口,如欧拉黑猫 21 格),
    # 改为打 already_extracted 标记——next-batch 普通模式跳过(避免重复搜索),
    # from-mapping 模式专门消化这批(数据已在 raw/dcd_refill/)
    done_kw = set()
    mp = os.path.join(RAW_DCD, 'series_mapping.json')
    if os.path.exists(mp):
        try:
            m = json.load(open(mp, encoding='utf-8'))
            for e in m.get('reviewed', []):
                if e.get('ok'):
                    done_kw.add(norm_name(e.get('kw') or e.get('series') or ''))
        except Exception as e:
            print('warn: mapping 读取失败(%s),按无标记处理' % e)

    entries = {}
    unqueueable_rows = 0      # 有缺口但通用名称全为占位符/空的行(无车系名可检索)
    unqueueable_gap_cells = 0
    for idx, r in enumerate(rows):
        fam = fam_of(r)
        # 品牌令牌只从 商标/企业名 取——绝不能把车系族名(车型名称)当品牌:
        # 2026-09-22 实跑抓到 U5(爱驰飞碟U5) 因族名'U5'是'北京EU5'子串而误过校验
        brands = {str(r[hi['产品商标']] or '').strip()}
        corps = {str(r[hi['企业名称']] or '').strip()}
        g = sum(1 for f in fields if is_empty(r[hi[f]]))
        parts = generic_parts(r[hi['通用名称']])
        if not parts:
            if g:
                unqueueable_rows += 1
                unqueueable_gap_cells += g
            continue
        for part in parts:
            k = (fam, part)
            e = entries.setdefault(k, {'rows': 0, 'rows_with_gap': 0, 'gap_cells': 0,
                                       'gap_fields': {f: 0 for f in fields},
                                       'brands': set(), 'corps': set(), 'row_idx': []})
            e['rows'] += 1
            e['row_idx'].append(idx)
            e['brands'] |= {b for b in brands if b}
            e['corps'] |= {c for c in corps if c}
            if g:
                e['rows_with_gap'] += 1
                e['gap_cells'] += g
                for f in fields:
                    if is_empty(r[hi[f]]):
                        e['gap_fields'][f] += 1

    # 共享名检测:同一 part 出现在多个族下
    part_fams = {}
    for (fam, part) in entries:
        part_fams.setdefault(part, set()).add(fam)

    out = []
    for (fam, part), e in entries.items():
        if e['gap_cells'] == 0:
            continue
        shared = len(part_fams[part]) > 1
        toks = sorted(brand_tokens(*e['brands'], *e['corps']))
        out.append({
            'qkey': qkey(fam, part),
            'family': fam,
            'part': part,
            'kw': part,
            'kw_alt': f'{fam} {part}' if shared else None,
            'shared_name': shared,
            'expect_brand': toks,
            'already_extracted': norm_name(part) in done_kw,
            'rows': e['rows'],
            'rows_with_gap': e['rows_with_gap'],
            'gap_cells': e['gap_cells'],
            'gap_fields': {f: c for f, c in e['gap_fields'].items() if c},
            'brands': sorted(e['brands']),
            'corps': sorted(e['corps']),
        })
    out.sort(key=lambda x: (-x['gap_cells'], -x['rows_with_gap'], x['qkey']))
    for i, e in enumerate(out):
        e['priority_rank'] = i + 1

    doc = {
        'generated_at': today(),
        'fields': list(fields),
        'total_series_with_gaps': len(out),
        'total_gap_cells': sum(e['gap_cells'] for e in out),
        'unqueueable': {'rows': unqueueable_rows, 'gap_cells': unqueueable_gap_cells,
                        'note': '通用名称为占位符/空,无车系名可检索——属 D6 占位符残留,走官方源通道,不属本工作流范围'},
        'entries': out,
    }
    ensure_dirs()
    json.dump(doc, open(QUEUE_FILE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    return doc


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--include-motor', action='store_true')
    ap.add_argument('--force', action='store_true')
    a = ap.parse_args()
    fields = dict(FIELD_MAP)
    if a.include_motor:
        fields.update(MOTOR_FIELD_MAP)
    doc = build(fields, force=a.force)
    print(f"队列已生成: {QUEUE_FILE}")
    print(f"有缺口车系 {doc['total_series_with_gaps']} 个,缺口格 {doc['total_gap_cells']} 个")
    print('前 10:')
    for e in doc['entries'][:10]:
        sh = ' [共享名]' if e['shared_name'] else ''
        print(f"  #{e['priority_rank']:<4} {e['family']} / {e['part']}{sh}  rows={e['rows']} gap={e['gap_cells']}")
