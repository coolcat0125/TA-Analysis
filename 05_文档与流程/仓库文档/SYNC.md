# SYNC · 多终端/多会话同步协议

> 目标：任何电脑、任何 agent 会话，clone/pull 本仓库即获得**完全一致**的权威数据与工具链，且推送互不覆盖。
> 原则：**GitHub main = 唯一事实源**。权威底表 xlsx、两版看板、管线脚本、审计产物、CHANGELOG、PROJECT_STATUS 全部入 git；git 历史即云端版本史。

## 一、一键同步（每台电脑、每次开工前）

```powershell
cd <仓库目录>
./sync.ps1          # fetch → ff-only 快进 → EOL 幻影自检修复 → audit 门禁 → 状态汇报
```

脚本行为：只做**快进**（本地领先/分叉时拒绝并提示人工处理），绝不自动 push——推送始终是显式动作。

## 二、工作循环（每台终端一致）

1. **开工**：`./sync.ps1`（确保基线最新）
2. **干活**：遵守 PROJECT_STATUS【协同约定】——只补空、唯一共识、逐格台账、底纹标注、版本号递增
3. **收工**：底表/两版看板/audit-output 三位一体刷新 → CHANGELOG 登记 → `git add -A && git commit && git pull --rebase origin main && git push`
4. **冲突预防**：push 前必 pull；长任务跨会话时先看远端 HEAD 是否被他人推进（对方 PROJECT_STATUS 迭代记录）

## 三、换行符幻影（已知问题，脚本自动处置）

`.gitattributes` 规定文本 LF 入库。不同 Windows 工具链可能以 CRLF 提交，导致 pull 后整文件显示"已修改"。
判别：`git diff --ignore-cr-at-eol --stat` 为**空**即纯幻影。处置：`git add --renormalize . && git commit -m "chore: EOL renormalize"`（零内容变化）。`sync.ps1` 已内置此自检。

## 四、仓库外目录（不入 git，各终端自管）

| 目录 | 性质 | 一致性策略 |
|---|---|---|
| `02-数据底表/`（工作区） | 应用侧派生库 + raw 官方 .doc 原文 | 权威数据=仓库底表 xlsx；派生库按需用管线重导，滞后不影响主线 |
| `03-应用工具/` | Web 应用/离线看板 | nev-hub 未入版本管理（NA-2 待办） |
| `04-历史归档/`（工作区） | 按日期归档的旧文件 | 本机追溯；归档夹内放 README 索引说明来源与原因 |

## 五、版本与台账

- 版本号只增不减，变更一律先记 `CHANGELOG.md` 再提交
- 数据修复类变更：底纹标注（清洗=浅红 FFC7CE）+【参数对齐记录】逐格留痕 + CHANGELOG 说明证据来源
- 多线并行若发生同名版本冲突：以"对方底表为基座 + 重跑己方管线"无损合并（v4.2.0 先例）
