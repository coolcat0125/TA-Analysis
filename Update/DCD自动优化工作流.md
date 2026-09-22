# DCD 懂车帝参数自动优化工作流

> 目标:以**懂车帝官网车型参数页**为源,按缺口优先顺序分批、持续地回填底表空缺字段,全程护栏自检 + 实时监控看板。
> 2026-09-22 建立,同日重构抓取层:**从 Tabbit 抽离为浏览器直抓**(Tabbit 降级为可选备选)。

## 0. 架构总览

```
① 缺口队列        dcd_queue.py          扫底表 8 字段缺口 → 车系粒度优先队列(845 个/3,744 格)
② 抓取目标        export_targets.py     队列 → targets.json(给浏览器工具)+ targets.csv(人审)
③ 官网直抓        extract_browser.js    在懂车帝官网页面控制台运行,同源 fetch 搜索页+参数页
                                         → 每 15 车系下载一个 bundle(断点续跑/限速/探查处)
④ 导入校验        import_bundle.py      bundle → raw/dcd_refill/dcd_s<sid>.json + fuel_form
                                         → decide_series 品牌校验 → series_mapping.json(review:auto)
⑤ 护栏回填        dcd_pipeline.py       备份 → 只补空/唯一值/W5成对/比值预防/污染锚点预检 → 落库+台账
⑥ 三件套自检      verify_consistency + audit_data    CRITICAL=0 / W4-W8 不增 / strict PASS,不过即回退
⑦ 实时监控        render_dashboard.py   DCD优化监控.html(15s 自动刷新 + ETA 面板)
```

**为什么抓取必须发生在浏览器里**:懂车帝边缘 WAF 对非浏览器客户端返回空体(实测:服务端 fetch / urllib 直连全部拿到 `200 + text/plain + 空体`;API 另需 msToken/a_bogus 签名)。所以抓取用**同源 fetch**在真实浏览器会话内完成——但这不等同于依赖 Tabbit:`extract_browser.js` 是一段独立 JS,零安装、零第三方工具,贴进控制台即可跑,也可随时改写替换。

## 1. 目录与文件

| 文件 | 角色 | 运行环境 |
|---|---|---|
| `dcd_queue.py` | 缺口优先队列生成 | Python |
| `export_targets.py` | 导出抓取目标(targets.json / targets.csv) | Python |
| `dcd_pipeline.py` | 编排:baseline/next-batch/decide/apply/status/resume/eta | Python |
| `dcd_refill_core.py` | 回填规则核心(匹配打分/只补空/护栏) | Python |
| `dcd_common.py` | 共享常量与工具 | Python |
| `import_bundle.py` | 浏览器 bundle → 数据文件 + 品牌校验 + mapping | Python |
| `render_dashboard.py` | 生成监控看板(含 ETA) | Python |
| **`extract_browser.js`** | **官网直抓工具(主抓取层)** | **浏览器控制台** |
| `run_full.ps1` | 全量四步编排 | Windows PowerShell |
| `run_loop.ps1` / `hook_nightly.ps1` | 循环/夜间轮挂接 | Windows PowerShell |
| `dcd_extract.js` / `dcd_extract.ps1` | Tabbit 通道(已降级为备选,见 §6) | Tabbit 运行时 |
| `dcd_auto/state/` | 队列/状态/bundle/备份(建议 gitignore) | 运行时产物 |

## 2. 快速开始(全量四步)

```powershell
# 第一步:建队列 + 导目标(也可由 run_full.ps1 自动做)
python dcd_auto\dcd_queue.py
python dcd_auto\export_targets.py

# 第二步:浏览器抓取(约 1 小时,845 车系)
#   1) 打开 https://www.dongchedi.com/ 停在该页面
#   2) F12 → Console → 粘贴 dcd_auto\extract_browser.js 全文 → 回车
#   3) 首次运行弹框选 dcd_auto\state\targets.json
#   4) 每 15 车系自动下载 dcd_bundle_NNN.json;中断后重跑自动续抓
#   5) 把下载目录里所有 dcd_bundle_*.json 拷到 dcd_auto\state\bundles\

# 第三步 + 第四步:一条命令跑完 导入→校验→批量回填→自检→看板
powershell -File dcd_auto\run_full.ps1
```

首次使用**务必先跑探查**:把 `extract_browser.js` 第一行 `const DCD_PROBE = false;` 改成 `true` 再粘贴,它会抓 1 个搜索页 + 1 个参数页并下载 `dcd_probe.json`。核对其中 `info_keys` 是否含 15 个目标字段、`fuel_form` 是否有值;不符则把 probe.json 反馈给维护者修正 SELECTOR 后再正式跑。

## 3. 护栏规则(自动执行)

| 护栏 | 判据 | 出处 |
|---|---|---|
| 只补空 | 仅写空格;唯一非空共识才补;`不适用(BEV)` 语义占位不可补 | 协同约定 |
| 动力守卫 | BEV 行只匹配有电池的款型;`ff=汽油` 款型排除 | integrate_dcd_refill.py |
| 修剪匹配 | 续航 ±15/±30 + 整备 ±40/±90 打分,取最优并列款型 | 同上 |
| W5 成对铁律 | 触碰车长/轴距 → 成对后两侧都必须有值,缺一项该对整体 hold | DCD 进度跟踪·门禁记录 |
| 比值预防 | 写入前模拟 C4(车长<轴距)/ W5(比值出 [1.4,2.4])/ W7(整备/车长出 [0.18,0.70] 且车长>2500) | verify_consistency.py |
| 污染锚点预检 | 现值已违反 C4/W5/W7 的行整体 hold | 本工作流新增 |
| 品牌校验 | 搜索候选与预期品牌令牌/包含关系校验;共享名歧义留人工 | 复现 15 例人审结论 |
| 写库时段 | 22:40-06:00 拒绝写库(`--force` 越权);时区按 `DCD_TZ_OFFSET`(默认 +8) | 白天分派纪律 |
| 三件套自检 | apply 后 CRITICAL 必须=0;W4/W5/W6/W7/W8 不得增;audit `--strict` 必须 PASS | 协同约定 |

## 4. 残差类型与处置(全量跑完也不会清零的三类)

按本地数据实测产出率(126 格 → 32 格,25%),全量跑完预计自动落库 900~1,400 格,其余分三类:

1. **多款型歧义**:同字段多个款型取值不一致,唯一值规则拒写。处置:人工按车型年款裁定,或在 targets 里把 kw 细化到款型级再抓。
2. **污染锚点行**:车长=4470 等历史估算污染值在位(实测波及 100 行,文档仅记 17 条)。处置:等车长修复包(T4)批准后成对修复,本工作流不单独动。
3. **DCD 无对应款型**:如比亚迪唐的款型经动力守卫后无剩余。处置:维持留档,或换官方详情页通道。

看板与 `进度与ETA报告.md` 会持续给出这三类的当前计数。

## 5. 监控与 ETA

`DCD优化监控.html`:15 秒自动刷新,数据内嵌,无外部依赖。含循环状态、基线(HEAD/底表 SHA/CRITICAL/WARN)、总体进度条、**进度与预计完成时间面板**、8 字段覆盖率、批次历史(含回退标记)、当前批次与暂停原因、回退/阻塞记录、队列预览、最近 25 条填充。

ETA 模型:`每批 = apply 实测均值 + 抽取假设 90s`;剩余批次 = 剩余车系 ÷ 每批车系数。假设明示在面板注释里,前 3 个真实批次后应以实测值校正(`EXTRACT_SEC_PER_BATCH` 环境变量)。

## 6. Tabbit 通道(已降级为备选)

`dcd_extract.js` + `dcd_extract.ps1` 保留为**可选的无人值守抓取通道**:若需在无人看管时长时间抓取,可改用它(按仓库既有配方调用 `tabbit-cli.exe nodejs`)。它与主管道的数据契约完全一致(都产 `raw/dcd_refill/dcd_s<sid>.json`),可混用。默认推荐 `extract_browser.js`:零安装、不依赖第三方运行时、用你自己的浏览器会话。

## 7. 边界(明确不做)

- **不推送 GitHub**:按 SYNC 协议,推送始终是显式动作;底表动→看板重生成→CHANGELOG 登记→提交,由主线终端执行
- **不动 CHANGELOG/版本号**:版本登记留主线
- **不替代 23:00 夜间轮**:`hook_nightly.ps1` 只在其后串联
- **不处理占位符通用名称行**(队列 `unqueueable` 计数可见):属 D6 残留,走官方源通道
- **电芯供应商 3,706 格**:DCD 无此标准列,维持公示详情页通道
- **百公里电耗默认不填**:DCD 为 CLTC 工况,与底表既有口径混填会制造 W2 矛盾;`--include-consumption` 可纳入

## 8. 故障处理

| 现象 | 处置 |
|---|---|
| 浏览器脚本提示"不在懂车帝域名下" | 先打开 dongchedi.com 任意页面再粘贴 |
| probe 显示 `info_keys` 不含目标字段 | 页面结构已变,把 probe.json 给维护者修正 SELECTOR |
| 抓取中断 | 直接重跑,localStorage 会跳过已完成的 |
| apply 返回 3 | 看 `paused_reason` 与看板阻塞区;底表已自动回退。处理完 `python dcd_auto\dcd_pipeline.py resume` |
| apply 返回 2 | 撞禁写窗,本批未写库;到窗口后重跑 |
| 品牌校验 reject/hold | 看 `state/import_report.json`;数据已落盘但不回填,人工复核后可改 mapping 的 `ok` |
| 队列跑空 | `dcd_queue.py` 重建(底表更新后) |

## 9. 建议的 .gitignore 追加

```
dcd_auto/state/
__pycache__/
```
