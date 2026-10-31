# 云端仓库维护手册（_scripts/ 工具链）

> 本目录**不入 git**（见 `.gitignore`），属本地运维工具。云端仓库为
> `coolcat0125/TA-Analysis`（私有），根目录只保留最新交付，历史版本在 `archive/`。

---

## 一、本机网络约束（关键，务必先读）

| 通道 | 状态 | 说明 |
|---|---|---|
| **PowerShell / .NET**（`Invoke-RestMethod`、`HttpWebRequest`） | ✅ 可用 | 走系统代理并使用当前登录用户的凭据 |
| **Python**（urllib / requests） | ❌ 不可用 | 企业代理要求 Windows 集成认证，Python 一律 407 |
| **git 传输**（clone/fetch/push） | ⚠️ 时好时坏 | 代理 407 或无法解析 `proxy.sos.shanghaigm.com`；大包推送易被中断 |

**结论：所有网络操作交给 PowerShell，Python 只做本地文本/编码处理。**

`GITHUB_TOKEN` 需对仓库有 **Contents: Read and write**；仅在环境变量中传递，勿写盘。

```powershell
$env:GITHUB_TOKEN = 'github_pat_...'
```

---

## 二、两个必须避开的坑

### 坑 1：Windows PowerShell 5.1 以 GBK 读取 `.ps1`
脚本源码中**禁止出现中文**（否则乱码 → 路径错误）。含中文的数据一律经
UTF-8 文件传递（`tsvlib.ps1` 的 `Read-Utf8Lines` / `Write-Utf8Lines`）。
判断脚本是否安全：

```powershell
(([IO.File]::ReadAllBytes('script.ps1')) | Where-Object { $_ -gt 127 }).Count   # 应为 0
```

### 坑 2：`git/trees` API 会把中文文件名降级为 `?`
**这是本项目最危险的坑。** `GET /git/trees/{sha}?recursive=1` 返回的 CJK 路径
会变成 `NEV??????.html`。若以该结果重建树并提交，`?` 会成为仓库**真实文件名**
（本次已发生 17 处，见提交 `9ad07e1d`）。

- **枚举文件一律用 contents API 递归遍历**（`net_walk.ps1`，UTF-8 正确）
- 同理 `ConvertTo-Json` 在 PS 5.1 也会把 CJK 变 `?`（本次未触发，但仍需避免）
- 提交前自检：清单中若出现 `?` 立即中止

---

## 三、常用操作

### 1. 拉取云端最新（生成清单）
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File _scripts\net_walk.ps1
# 产出 C:\tmp\remote_files.tsv（sha \t size \t path）与 C:\tmp\remote_head.txt
```

### 2. 云端 → 本地 同步
```powershell
python _tools\py311\python.exe _scripts\sync_local.py plan      # 算差异、出下载清单
powershell -NoProfile -ExecutionPolicy Bypass -File _scripts\download_raw.ps1
python _tools\py311\python.exe _scripts\sync_local.py apply     # 校验 sha → 备份 → 落位
```
- 比对**按 `.gitattributes` 归一化**（文本 CRLF→LF）后算 blob id，避免换行符误判
- 覆盖前自动备份到 `_archive\sync_backup_<时间戳>/`
- 默认不含 `archive/`；需要时加 `--with-archive`

### 3. 上传单个/批量文件（blob → 树 → 提交）
```powershell
# 生成清单（Python 只读本地，不联网）
python _tools\py311\python.exe _scripts\make_manifest_local.py
powershell -NoProfile -ExecutionPolicy Bypass -File _scripts\push_blobs.ps1
# push_blobs.ps1 产出 C:\tmp\blob_shas.json；再据此建树提交（见 commit_tree.ps1 / fix_updates.ps1）
```

### 4. 校验云端路径是否被损坏
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File _scripts\net_walk.ps1
python _tools\py311\python.exe _scripts\list_broken_paths.py   # 含 '?' 的路径应恒为 0
```

---

## 四、并发安全

另一会话会同时向本仓库推送（曾于 2026-09-10 一天内推进 v4.4.3→v4.5.2）。
因此：

- 建树/提交**必须以「执行瞬间读到的 HEAD」为 parent**，且推进 ref 时 `force=false`；
  并发推送会导致 ref 更新失败并中止 —— 这是保护，不是故障，重跑即可
- **绝不**用本地旧副本覆盖云端同名文件。本次同步前，本地 `CHANGELOG.md` 仅 37 版、
  云端 42 版；若反向覆盖将永久丢失 v4.4.3~v4.5.2 记录
- 动手前先跑一次差异比对，确认本地确实落后再同步

---

## 五、脚本清单

| 脚本 | 用途 |
|---|---|
| `tsvlib.ps1` | UTF-8 TSV 读写（中文安全的数据通道） |
| `jsonlib.ps1` | UTF-8 JSON 辅助（**注意**：PS 5.1 下不可靠，优先用 TSV） |
| `net_walk.ps1` | 经 contents API 递归枚举云端文件 → TSV（**枚举的唯一正确方式**） |
| `net_get.ps1` | 经 `git/trees` 取树（⚠️ CJK 会损坏，仅供无中文路径时参考） |
| `push_blobs.ps1` | 按清单上传 blob（断点续传、重试） |
| `download_raw.ps1` | 以 raw 媒体类型流式下载文件 |
| `commit_tree.ps1` / `fix_updates.ps1` / `fix_paths.ps1` | 建树 + 提交 + 推进 main |
| `cloud_admin.py` | 凭据诊断 `doctor`、文件树 `tree`（Python，仅在代理可用时） |
| `sync_local.py` | 云端→本地差异计算与落位（`plan` / `apply`） |
| `restructure_cloud.py` | 架构重构映射表（archive 归档映射的**权威定义**） |
| `make_manifest_local.py` | 由本地归档生成上传清单 |
| `make_rename_map.py` | 生成损坏路径→正确路径修复映射（尺寸唯一匹配 + 定向核对） |
| `list_broken_paths.py` | 列出含 `?` 的损坏路径 |

---

*最后更新：2026-09-11 · 对应云端 HEAD `9ad07e1d`*
