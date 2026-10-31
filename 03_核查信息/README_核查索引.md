# 核查信息索引

> 本目录只有索引说明。核查产物实物因**被脚本硬引用**，保留在管线工作区内。

## 一、实际的核查产物位置

| 类型 | 路径 | 内容 |
|---|---|---|
| 审计报告 | `06_自动化与工具/管线工作区/audit-output/data_audit_report.json/.md` | 字段覆盖率、物理范围、批次分布、门禁结论 |
| 一致性报告 | `…/audit-output/consistency_report.json` | 跨表一致性校验 |
| 证据与快照 | `…/audit-output/*.json` | 容量取证、车长核定、W2/W4 仲裁等 |
| QA 截图 | `…/audit-output/*.png` | 报告/演示文稿目检样页（33 张） |
| 门禁产物 | `…/audit-output/gate/` | 门禁规则、批次校验 |
| 质量门禁阈值 | `…/管线工作区/quality_gates.json` | `expected_min_records` 等 |
| 同步比对 | `06_自动化与工具/_cloud_scan/` | 云端差异清单、核对表、下载日志、结构快照 |

## 二、常用核查命令

```bash
cd "06_自动化与工具/管线工作区"

# 数据审计 + 门禁
python audit_data.py --gates quality_gates.json --strict

# 跨表一致性
python verify_consistency.py

# 品牌维度审计
python audit_brand_v2.py

# 销量源质检
python sales_source_qc.py
```

## 三、核查口径要点

- **基准**：`quality_gates.json`（云端权威，勿改）
- **比对方法**：git blob SHA（文本按 LF 归一化），不用文件大小判断
- **门禁基线**：verify CRITICAL=0 / WARN 208（2026-09-27 云端状态）

> ⚠️ `_baseline/`（旧质检基准）已于 2026-09-29 删除。质检请直接以管线工作区内的云端权威底表为数据源。
