# A08 Evidence Audit · r3 候选包与 QA（基于 main@fd075b5 / 分支 agent/night-20260920）

- **r3 范围**：将 r2 分支（agent/a08-evidence-audit-r2, fedf4b5）候选层整体并入夜分支并复跑独立 QA。r2 与 fd075b5 的差异**仅 candidate/ 层与 QA 报告**（10 文件 +1011 行），与 v4.9.13~15 的看板/生成器变更零重叠——候选包可直接叠加于当前 main。
- **五态计数**：verified **0** ／ candidate **78**（车长 55 + A2_A5 契约 23）／ pending **22**（车长待取证，构成见 findings）／ conflicted **0** ／ unknown **1**（A4 BYD 380 批维持）。
- **证据范围**：candidate 78 项全部带 source_name/source_url/quoted_value/retrieved_at（车长 9 车型族多源交叉：五菱官网/汽车之家/易车/官方上市信息/工信部目录转述；A2_A5 为契约化 23 条）。
- **与旧 PR #5 的差异**：findings 结论行口径修正（58/B44/剩19 → 55/B41/22）+ 新增 population_inclusion（77⊂103）+ r2/r3 溯源块；ledger/candidate 数值层与 PR #5 原子级一致（EOL 幻影已判别）。
- **是否具备单独开 PR 条件**：**具备**（候选层自洽、QA 全过、不触 canonical、基于当前 main）——是否开 PR 由 Orchestrator 决定。
- **QA**：`candidate/v4.9/vehicle_length/qa_a08_r3.json`（9/9 PASS）。
- **verified=0 声明**：无任何 candidate 自动晋升；canonical 零改动。
