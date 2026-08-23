# 参数证据采集与补全

Copyright © 2026 David YE

## 证据规则

| 等级 | 可写入条件 | 例子 |
|---|---|---|
| A | 监管、企业或权威机构原始资料；OCR 必须人工复核 | 工信部公告附件、购置税目录、企业技术资料 |
| B | 权威数据库/专业媒体转引，并保留原始链接 | 需标注转引限制 |
| C | 可解释计算或行业推断；默认仅生成候选，不写回主表 | 电耗=容量×100÷续航；密度=能量×1000÷电池组质量 |

未公开字段必须留空。C 级仅可通过 `--include-c-inference` 写入，且绝不覆盖原始申报值；每条补充或候选都会写入 `【证据补充审计】`。

## OCR 专项流程

1. 将公告参数图/PDF的每条识别结果填入 `evidence_ledger_template.csv`；记录来源 ID、URL、发布日期和页码/截图位置。
2. `evidence_type=OCR` 的记录须二次核对后改为 `review_status=verified`。
3. A/B 级扭矩证据可以写入缺失字段；字段名称必须与源表表头包含关系一致，例如 `电机最大扭矩`。

## 运行

```bash
pip install openpyxl
python3 enrich_nev_parameters.py \
  --input 'NEV公告参数汇总表_合并版（341~409批）.xlsx' \
  --ledger evidence_ledger_template.csv \
  --output 'NEV公告参数汇总表_补充候选.xlsx'
```

默认只写 A/B 级、已复核的外部证据；C 级计算仅进入审计表，便于复核。确认后才使用 `--include-c-inference`。
