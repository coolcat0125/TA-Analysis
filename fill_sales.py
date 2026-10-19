# -*- coding: utf-8 -*-
"""fill_sales.py — 【已废弃 OBSOLETE · 2026-10-19】易车车系销量榜 → 底表月销量参考列补全

⚠️ 本脚本已被 link_sales.py（汽车之家 2026-09 口径）取代，**禁止运行**。
   详见文件末尾的硬停守卫与 Update/销量口径规范_20261019.md §10 禁止运行清单。

数据源：audit-output/yiche_sales_rank_202608_raw.json
  （car.yiche.com/newcar/salesrank/?energy=6，零售量=行业综合销量，2026年08月，车系级）

口径原则（用户指令 2026-09）：
  1. 仅参考项，不参与任何官方字段口径
  2. 车系级月销量（同车系跨动力类型变体为分类汇总口径，无法按年份切分）
  3. 只补空；匹配唯一才填；多候选/无匹配一律留空
  4. 逐格台账 + 底纹（F8CBAD 媒体通道色）

匹配算法：
  norm(小写/去空格连字符/去"新能源") 后双向包含：
  workbook键 K 与车系名 Y 互相包含即候选；候选唯一才采纳。
  通用名称优先，车型名称兜底。

用法：
  python fill_sales.py --dry-run   # 只出报告
  python fill_sales.py --apply     # 写入底表（media_fill 等写管线运行期间禁用！）
"""
import json
import os
import re
import sys
import datetime as dt
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill

# ============================================================================
# 【硬停守卫】OBSOLETE WRITER —— 2026-10-19 由 Lead 加装
# 本脚本已于 2026-10-19 被 link_sales.py（汽车之家 2026-09 口径）取代，禁止运行：
#   ① 源 audit-output/yiche_sales_rank_202608_raw.json 是「易车 2026-08」旧口径，
#      而底表 col35 已统一为「汽车之家 2026-09」——运行会把旧口径值回灌已统一列；
#   ② MONTH_TAG='2026-08'、SOURCE_TAG='易车车系销量榜' 均为硬编码；
#   ③ 只补空、不写新增的 col39(车系ID) / col40(销量覆盖状态)，
#      会造成「数值已统一、覆盖状态未写」的不一致状态；
#   ④ 绕过 link_sales.py 的口径自检与写入后校验。
# 依据：AGENTS 经验教训（过期写入器在本轮统一口径后再运行 = 数据回灌污染，
#       同族事故见 2026-09-11「工作簿清洗必须同步源链」11 格回灌）。
# 唯一放行方式（仅限离线取证 / 历史复现）：
#       PowerShell:  $env:ALLOW_OBSOLETE_FILL_SALES='1'
# ============================================================================
if os.environ.get("ALLOW_OBSOLETE_FILL_SALES") != "1":
    raise SystemExit(
        "fill_sales.py 已废弃（OBSOLETE）：源为易车 2026-08 旧口径，运行会污染已统一为\n"
        "汽车之家 2026-09 的底表销量列，并绕开 col39/col40 写入与口径自检。\n"
        "请改用 link_sales.py。确需离线取证时设置 ALLOW_OBSOLETE_FILL_SALES=1 强制放行，\n"
        "并在变更记录中注明（参见 Update/销量口径规范_20261019.md §10 禁止运行清单）。"
    )

HERE = os.path.dirname(os.path.abspath(__file__))
WB_PATH = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
SRC_JSON = os.path.join(HERE, 'audit-output', 'yiche_sales_rank_202608_raw.json')

COL_GENERIC = 6      # 通用名称（0-based，与 main sheet 列索引一致）
COL_BATCH = 0
COL_CODE = 1         # 产品型号
COL_POWER = 9        # 动力类型
COL_SALES = 35       # 月销量(辆)
COL_SMONTH = 36      # 销量数据月份
COL_SDATE = 37       # 销量更新日期

MIN_OVERLAP = 4      # 包含匹配的最短重叠长度（防品牌短词误配，如"荣威"→"荣威D6"）

FILL_MEDIA = PatternFill('solid', fgColor='F8CBAD')
MONTH_TAG = '2026-08'
SOURCE_TAG = '易车车系销量榜（零售量，行业综合销量，车系级分类汇总）'


def norm(s):
    s = str(s or '').strip().lower()
    s = re.sub(r'[\s\-_·．.（）()]+', '', s)
    return s.replace('新能源', '')


def load_yiche():
    d = json.load(open(SRC_JSON, encoding='utf-8'))
    return {norm(i['name']): i for i in d['items'] if i.get('sales')}


def main():
    dry = '--dry-run' in sys.argv
    ymap = load_yiche()
    ykeys = sorted(ymap.keys(), key=len, reverse=True)

    wb = openpyxl.load_workbook(WB_PATH)
    ws = wb['NEV公告参数汇总']
    rows = list(ws.iter_rows(min_row=2))
    ledger = wb['变更记录']
    seq = max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
               if isinstance(r[0], int)), default=0)

    filled = matched = ambiguous = unmatched = 0
    examples = []
    ledger_rows = []

    for row in rows:
        sales_cell = row[COL_SALES]
        if sales_cell.value not in (None, ''):
            continue  # 只补空
        generic = row[COL_GENERIC].value
        k1 = norm(generic)
        cands = set()
        if k1:
            if k1 in ymap:                      # 整名精确命中：直接采纳，不做包含扩展
                cands.add(k1)
            else:
                parts = re.split(r'[,，、]', str(generic or ''))
                for part in {norm(p) for p in parts} - {''}:
                    if part in ymap:
                        cands.add(part)
                        continue
                    for y in ykeys:
                        if min(len(y), len(part)) >= MIN_OVERLAP and (y in part or part in y):
                            cands.add(y)
        if not cands:
            unmatched += 1
            continue
        if len(cands) > 1:
            ambiguous += 1
            if len(examples) < 400:
                examples.append(('AMBIG', row[COL_BATCH].value, row[COL_CODE].value,
                                 str(generic), sorted(cands)[:4]))
            continue
        y = next(iter(cands))
        matched += 1
        sales = ymap[y]['sales']
        if not dry:
            sales_cell.value = sales
            sales_cell.fill = FILL_MEDIA
            row[COL_SMONTH].value = MONTH_TAG
            row[COL_SMONTH].fill = FILL_MEDIA
            row[COL_SDATE].value = dt.date.today().isoformat()
            row[COL_SDATE].fill = FILL_MEDIA
            seq += 1
            ledger_rows.append((seq, row[COL_BATCH].value, row[COL_CODE].value,
                                row[COL_POWER].value, '月销量(辆)', '销量参考补全', '(空)',
                                sales, f'{SOURCE_TAG} 车系={ymap[y]["name"]}；销量数据月份={MONTH_TAG}/更新日期同轮填充'))
        if len(examples) < 15:
            examples.append(('OK', row[COL_BATCH].value, row[COL_CODE].value,
                             str(generic), f'{ymap[y]["name"]}={sales}'))

    print(f'rows_with_empty_sales_filled_or_candidate={matched} unmatched={unmatched} ambiguous={ambiguous}')
    for e in examples[:15]:
        print(' ', e)
    if dry:
        print('DRY-RUN（未写入）')
        return
    for r in ledger_rows:
        ledger.append(r)
    wb.save(WB_PATH)
    print(f'saved. ledger+{len(ledger_rows)} (seq {ledger_rows[0][0]}~{ledger_rows[-1][0]})'
          if ledger_rows else 'saved. ledger+0')


if __name__ == '__main__':
    main()
