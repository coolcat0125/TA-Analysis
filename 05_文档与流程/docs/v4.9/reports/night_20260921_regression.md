# 看板回归报告 · v4.9.17~v4.9.19（真实数据 5,392×38）

## 发现并修复 1 个 P1 级缺陷
- **方案 C 梯度曲面完全失效**：v4.9.17 重建时系列注册名误用 `type:'surface3D'`，而 echarts-gl 2.x 注册名是 `type:'surface'`——未知类型被 setOption **静默丢弃**（series=[]、GL 画布空白，AGENTS.md 2026-09-20 条目原样复发）
- **修复**：generate_dashboard.py mkSf 系列类型 1 处替换 + 两版看板重生成
- **修复后验证**：surface:BEV + surface:PHEV/EREV 双曲面注册成功；双 visualMap（全局 Z 域 8.3~144.4 共享、5 色带）；GL 画布渲染出曲面网格（证据 night_surfc_C_c0.png）；发布版同步生效（v4.9.15 起发布版含三维）；零 JS 错误

## 回归矩阵（修复后）
| 项 | 结果 |
|---|---|
| 视图枚举 5 项（d2/bar/bub/surf/heat） | ✅ |
| 方案 A bar3D 双系列 | ✅ bar3D:BEV + bar3D:PHEV/EREV |
| 方案 B scatter3D 双系列 | ✅ |
| 方案 C surface 双曲面（修复后） | ✅ 双 visualMap 全局 Z 域 |
| 热力 | ✅（前轮） |
| 真实数据 | ✅（KPI 5,392；X/Y/Z 齐备行入图） |
| 发布版对齐 | ✅（三维/曲面可用、质量与导入隐藏） |
| JS 错误 | 0 |

## 备注
- 曲面相机默认视野较近（网格特写），属视角参数非缺陷
- 空值行不入曲面网格，可视化填补无 canonical 回写
