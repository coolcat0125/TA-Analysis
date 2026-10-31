# 数据质量与版本审计

Copyright © 2026 David YE

`audit_data.py` 为公告源表提供可复跑的质量与版本审计。每次新增公告批次后执行审计，并将 JSON 与 Markdown 报告同数据文件、看板产物一起归档。

## 运行

```bash
pip install openpyxl
python3 audit_data.py --gates quality_gates.json --strict
```

默认自动发现当前目录下最新命名的 `NEV公告参数汇总表_合并版*.xlsx`。也可显式指定：

```bash
python3 audit_data.py \
  --input 'NEV公告参数汇总表_合并版（341~410批）.xlsx' \
  --gates quality_gates.json \
  --output-dir audit-output \
  --strict
```

## 产物与用途

| 产物 | 用途 |
|---|---|
| `data_audit_report.json` | 机器可读：输入文件 SHA-256、表结构、记录数、批次分布、字段覆盖率、异常样本和门禁结果。 |
| `data_audit_report.md` | 人工复核：可直接随版本说明提交。 |

## 审计口径

- 固定记录输入文件 SHA-256，保证每个结论可追溯至具体数据版本。
- 检查必需字段、最小记录数和关键字段覆盖率门禁。
- 对整备质量、纯电续航、电池容量、能量密度、电耗、电机功率和发动机排量实施物理范围检查，并列出最多 10 条异常样本。
- 默认仅在 `--strict` 模式下将未通过门禁作为非零退出码，便于未来接入 GitHub Actions。

## 更新规则

当新增公告批次或修改字段口径时，同时更新：

1. `quality_gates.json` 的记录数与覆盖率门槛；
2. `CHANGELOG.md` 中的数据范围、来源与审计结论；
3. 本次输出的 JSON/Markdown 审计报告。

本工具只记录客观检查结果；补数、推断与预测应在原始数据与审计表中标明方法、来源及置信度。
