# 看板回归报告 · v4.9.15（真实数据 5,392×38 · 分支 agent/night-20260920）

- **回归范围定谳**：任务书基线 a61bc6e 实际已前移至 fd075b5（v4.9.15）——增量回归覆盖 v4.9.13~15 变更面（视图模式重构+导航+发布版三维恢复）；镜像/主题/全屏处理器经 `git diff a61bc6e..fd075b5` 确认未触碰，沿用 R1/R2（3d7cc99）证据。
- **方案 C（三维曲面）**：cdView 枚举仅 4 项（d2/bar/bub/heat）——**确认已随 v4.9.15 下线**，"双 visualMap 双曲面"特性随之移除，无从回归（上游删除，非遗漏）。

| 项 | 结果 | 证据 |
|---|---|---|
| 真实数据 | **PASS**：d2=2,212 点；方案 A bar3D BEV 2,031 + PHEV/EREV 1,317（X/Y/Z 三字段齐备的真实行）；KPI 5,392 ✓；无随机/演示数据 | getOption 采样 |
| 方案 A 柱状三维 | **PASS**：bar3D:BEV + bar3D:PHEV/EREV 双系列独立图例；**双 visualMap（dim=2，全局 Z 域 7.68~144.4 共享，5 色带）** | night_r3_barA_c*.png |
| 方案 B 梯度泡泡 | **PASS**：scatter3D 双系列同构 | night_r3_bubB_c*.png |
| 热力 + 风格往返 | **PASS**：heatmap 类型；A→E→A 色带长度 5→5 稳定 | evaluate 采样 |
| 发布版三维恢复 | **PASS**：bar3D 双系列可切、质量区块隐藏、导入隐藏、零 JS 错误（v4.9.15 起发布版含三维——旧"剥离"预期作废） | evaluate 采样 |
| JS 错误 | **0**（全流程 window.__errs 双钩子） | — |
| ECharts/ECharts-GL 依赖 | **在位**（echarts-gl.min.js 本地内嵌） | 文件存在性 |
| 镜像/主题/全屏 | 沿用 R1/R2 证据（处理器 diff 为空未变更）+ 本夜零错误旁证 | 3d7cc99 |

- **可视化填补声明**：三维/热力的空值仅影响绘图点位（缺字段行不入图），不回写 canonical；无隐式伪造数据。
