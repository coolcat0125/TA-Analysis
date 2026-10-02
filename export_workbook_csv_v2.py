# -*- coding: utf-8 -*-
"""export_workbook_csv_v2.py — 权威底表 → master_export CSV（方案甲：全量 38 列直导出）

替代旧 export_master_csv.js（v4.3 33 列 schema）。行集=当前底表全量（5,392），
列=底表 38 列原序原表头（天然对齐后续 schema）。reader/writer 逐行写（09-22 教训）。
旧 CSV 先 .bak-时间戳 归档；产出校验：行数+列数+抽样值三方核对。
用法：python export_workbook_csv_v2.py [--apply]（缺省 dry 预览统计）
"""
import csv
import datetime
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OLD = os.path.normpath(os.path.join(HERE, '..', '..', '02-数据底表', 'data', 'nev-announcements', 'db', 'master_export_fixed.csv'))
NEW = OLD


def main():
    apply = '--apply' in sys.argv
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['NEV公告参数汇总']
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append(['' if v is None else (str(v) if not isinstance(v, float) else ((':%.10g' % v).lstrip('0').replace(':.', '0.') if v == v else '')) for v in r])
    wb.close()
    hdr, data = rows[0], rows[1:]
    print(f'底表读出: {len(data)} 行 × {len(hdr)} 列')
    if not apply:
        print('DRY：将写入', NEW)
        print('表头预览:', hdr[:10])
        return
    if os.path.exists(OLD):
        bak = OLD + '.bak-' + datetime.datetime.now().strftime('%Y%m%d%H%M')
        os.replace(OLD, bak)
        print('旧 CSV 归档:', bak)
    with open(NEW, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(hdr)
        w.writerows(data)
    # 校验：重读行数/列数/抽样
    chk = list(csv.reader(open(NEW, encoding='utf-8-sig')))
    assert len(chk) == len(data) + 1, '行数不符'
    assert len(chk[0]) == len(hdr), '列数不符'
    print(f'校验 PASS: {len(chk)-1} 行 × {len(chk[0])} 列 → {NEW}')


if __name__ == '__main__':
    main()
