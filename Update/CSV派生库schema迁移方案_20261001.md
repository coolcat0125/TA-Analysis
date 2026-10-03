# CSV 派生库 schema 迁移方案（v4.3 → v4.4 电机列 · 设计稿 2026-10-01）

## 现状
- `master_export_fixed.csv`：4,459 行 × 33 列，v4.3 旧 schema；电机域为旧三列
  `电机峰值功率(kW)` / `电机总功率(kW)` / `电机功率/扭矩`，另有
  `电芯供应商（垂媒校验拓展）`、`动力类型细分` 两个 CSV 独有列。
- 权威底表现为 v4.4 schema：`前电机功率/扭矩` / `电机总功率/扭矩` / `后电机功率/扭矩`
  （P/T 文本）+ `车长(mm)` + `电芯供应商` + 电池包供应商/URL/销量三列。
- 消费方：`build-master.js`（→master_vehicles.json，本机已由 rebuild_from_workbook.py
  直接重建，暂不经 CSV）与历史分析脚本。

## 迁移映射（v4.3 列 → 新值）
| 旧列 | 处置 | 新列 |
|---|---|---|
| 电机峰值功率(kW) | 语义≈前电机功率（数值），升格为 P/T 文本 `值/扭矩`（扭矩空缺则裸值） | 前电机功率/扭矩 |
| 电机功率/扭矩 | 原本就是 P/T 文本，多为总口径 → 归入 | 电机总功率/扭矩 |
| 电机总功率(kW) | 数值 → `值/扭矩`（扭矩取旧 电机功率/扭矩 第二段若一致） | 电机总功率/扭矩 |
| 电芯供应商（垂媒校验拓展） | 改名对齐 | 电芯供应商 |
| 动力类型细分 | 保留为拓展列（不动） | — |
| 车长(mm)/后电机功率/扭矩 | 新增空列（不回填——历史 CSV 为 4,459 行子集，与现底表 5,392 不同源） | 车长(mm) 等 |

## 关键决策点（执行前须确认）
1. **行集差异**：CSV 4,459 行 vs 底表 5,392——迁移是否顺带重置行集为当前底表全量？
   - 方案甲（推荐）：CSV 直接由底表全量重导出（列序按现 38 列），旧 CSV 归档退役
     ——一行代码等价于 export_master_csv.js 的 Python 版，schema 天然对齐；
   - 方案乙：保 4,459 行子集仅迁列（保留旧语料边界，但长期双轨）。
2. **消费方改造**：build-master.js/parse-master.js 的列名引用需同步
   （grep 旧列名 5 处）；若走方案甲，字段名以底表 38 列为准。
3. **回滚**：迁移前 .bak-时间戳 备份（09-22 教训：reader/writer 逐行，禁 DictWriter）。

## 建议排期
方案甲成本≈30 分钟（脚本+验证+消费方 grep 修）；乙≈1.5 小时且留双轨债。
待确认后执行。

## 执行记录（2026-10-01~02）
- **方案甲已落地**：export_workbook_csv_v2.py 底表全量重导出（5,392×38，旧 33 列/4,459 行 CSV → .bak-202610022248 归档）。
- **消费方核验**：sync_master_vehicles.py 列映射已适配（新增 P/T 三列+lengthMm/monthSales/cellSupplier 等；旧列键保留为 Legacy 缺省）；**build-master.js/parse-master.js 实为另一条 legacy 链**（输入 source-sheets/NEV公告参数汇总.csv，不读 master_export_fixed.csv）——零破坏；build-master.js 电机域已预防性适配 v4.4 列（ptFirst/ptSecond helper）。
- **键名约定（重要）**：双构建器统一采用 rebuild_from_workbook.py 键名——motorPeakKw/motorTotalKw/motorRearKw = 前后总电机 P/T 文本；**勿再引入 motorPeakPT 等第二套键名**（10-02 曾引发双构建器 schema 冲突，已即时修复）。
- 派生库现状：master_vehicles.json 5,392 条（rebuild 权威态），lengthMm 4,628 / monthSales 806 / motorPeakKw 4,648 / cellSupplier 1,630。
