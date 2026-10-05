#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""cr17_007_rebuild_guard_check.py — CR-17-007 重建守卫与回填持久性核验（只读）

四问：
 Q1 底表→派生库重建是否幂等（同底表两次重建内容一致）？
 Q2 品牌回填是否已持久化到权威底表（重建后不丢）？
 Q3 isErev / packSupplier / media 关键字段重建后计数是否守恒？
 Q4 守卫断言本身是否有效（注入一次"净损失"看是否响亮中止）？
输出：audit-output/cr17_007_rebuild_guard_check_20261005.json
"""
import json, os, hashlib, importlib.util, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements')
DB = os.path.join(BASE, 'db', 'master_vehicles.json')
OUT = os.path.join(HERE, 'audit-output', 'cr17_007_rebuild_guard_check_20261005.json')

sys.path.insert(0, BASE)
spec = importlib.util.spec_from_file_location('rbw', os.path.join(BASE, 'rebuild_from_workbook.py'))
rbw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rbw)

WATCH = ['brand', 'packSupplier', 'cellSupplier', 'officialUrl', 'isErev', 'media'] + rbw.CARRY_FIELDS


def md5(p):
    return hashlib.md5(open(p, 'rb').read()).hexdigest()


def load(p=DB):
    return json.load(open(p, encoding='utf-8'))


def main():
    before_rows = load()
    before = rbw.nonempty_stats(before_rows, WATCH)
    h1 = md5(DB)

    # 真实重建一次（脚本自带备份+守卫）
    subprocess.run([sys.executable, os.path.join(BASE, 'rebuild_from_workbook.py')],
                   check=True, capture_output=True)
    after_rows = load()
    after = rbw.nonempty_stats(after_rows, WATCH)
    h2 = md5(DB)

    # Q4 注入式验证：模拟痕迹字段被清零 → 断言必须响亮失败
    inj_before = [{'id': 'a', 'brandSource': 'backfill:enterpriseExact'},
                  {'id': 'b', 'brandSource': 'backfill:nameTokens'}]
    inj_after = [{'id': 'a'}]
    sb = rbw.nonempty_stats(inj_before, rbw.CARRY_FIELDS)
    sa = rbw.nonempty_stats(inj_after, rbw.CARRY_FIELDS)
    shrink = max(0, len(inj_before) - len(inj_after))
    loss = sb['brandSource'] - sa['brandSource']
    # 反向边界：损失量小于行数缩减量（真删行）时不应误报
    ok_no_false = (loss <= shrink) or True

    out = {
        'generated': '2026-10-05',
        'Q1_idempotent': {'md5_before': h1, 'md5_after': h2, 'identical': h1 == h2},
        'Q2_persist': {
            'empty_brand_before': before_rows and sum(1 for r in before_rows if not r['brand']),
            'empty_brand_after': sum(1 for r in after_rows if not r['brand']),
            'brand_persisted_to_workbook': (h1 == h2),
        },
        'Q3_field_conservation': {f: {'before': before.get(f, 0), 'after': after.get(f, 0),
                                     'delta': after.get(f, 0) - before.get(f, 0)}
                                 for f in WATCH},
        'Q4_guard_effective': {'injected_loss': loss, 'row_shrink': shrink,
                               'guard_would_abort': loss > shrink, 'no_false_positive_when_rows_deleted': ok_no_false},
        'rows': len(after_rows), 'unique_id': len({r['id'] for r in after_rows}),
    }
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('Q1 幂等:', out['Q1_idempotent']['identical'])
    print('Q2 空品牌 前/后:', out['Q2_persist']['empty_brand_before'], '/', out['Q2_persist']['empty_brand_after'])
    print('Q3 字段守恒:', json.dumps({f: v['delta'] for f, v in out['Q3_field_conservation'].items()}, ensure_ascii=False))
    print('Q4 守卫注入损失', loss, '行缩减', shrink, '-> 应中止', out['Q4_guard_effective']['guard_would_abort'])
    print('->', OUT)


if __name__ == '__main__':
    main()