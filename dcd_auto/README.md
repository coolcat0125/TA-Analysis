# dcd_auto — 懂车帝参数自动优化工作流

按缺口优先顺序,分批、持续地把**懂车帝官网车型参数**回填进底表空缺字段;全程护栏自检,失败自动回退;实时监控看板跟踪状态与 ETA。

**完整运行手册:[`Update/DCD自动优化工作流.md`](../Update/DCD自动优化工作流.md)**

## 快速开始(全量四步)

```powershell
python dcd_auto\dcd_queue.py                    # 1. 建缺口优先队列
python dcd_auto\export_targets.py               #    导出抓取目标 targets.json

# 2. 浏览器抓取:打开 https://www.dongchedi.com/ → F12 控制台
#    粘贴 dcd_auto\extract_browser.js 全文 → 选 targets.json → 自动下载 bundle
#    (首次把第一行 DCD_PROBE 改 true 先跑探查)

# 3+4. 导入校验 + 批量回填 + 自检 + 看板,一条命令:
powershell -File dcd_auto\run_full.ps1
```

## 抓取层为什么是浏览器直抓(不是 Tabbit)

懂车帝边缘 WAF 对非浏览器客户端返回空体(服务端 fetch / urllib 直连实测全部拿到 `200 + text/plain + 空体`,API 另需 msToken/a_bogus 签名)。因此抓取用**同源 fetch** 在真实浏览器会话内完成。`extract_browser.js` 是独立的一段 JS——零安装、零第三方工具,不绑定任何 agent 运行时;Tabbit 通道(`dcd_extract.js`/`.ps1`)保留为可选的无人值守备选,数据契约一致可混用。

## 文件职责

| 文件 | 职责 |
|---|---|
| `dcd_queue.py` | 扫底表 8 字段缺口 → 车系粒度优先队列(占位符过滤/共享名标记/品牌令牌) |
| `export_targets.py` | 队列 → `state/targets.json`(浏览器工具用)+ `targets.csv`(人审) |
| `extract_browser.js` | **官网直抓**:同源 fetch 搜索页定 sid → 参数页取 `__NEXT_DATA__` → 抽 15 字段+能源类型 → bundle 下载;断点续跑/限速/probe |
| `import_bundle.py` | bundle → `raw/dcd_refill/dcd_s<sid>.json` + fuel_form 合并 + `decide_series` 品牌校验 → mapping(review:auto) |
| `dcd_pipeline.py` | 编排:baseline / next-batch / decide / apply / status / resume / eta |
| `dcd_refill_core.py` | 回填规则:匹配打分、只补空、唯一值、W5 成对铁律、比值预防、污染锚点预检 |
| `dcd_common.py` | 共享常量与工具(路径、空值、占位符、品牌令牌、字段映射) |
| `render_dashboard.py` | 生成 `DCD优化监控.html`(15s 自动刷新 + ETA 面板) |
| `run_full.ps1` | 全量四步编排(队列→抓取提示→导入→批量回填) |
| `run_loop.ps1` / `hook_nightly.ps1` | 循环/夜间轮挂接 |

## 三条红线

1. **只补空 + 唯一共识**——不覆盖已有值,多候选取值不一致就跳过并计数
2. **车长/轴距成对**——缺一项证据整对 hold;写入前模拟 C4/W5/W7 比值,违规则拦
3. **22:40-06:00 不写库**,每批三件套自检(verify CRITICAL=0 / audit strict PASS / W4-W8 不增),不过就整批回退并暂停

## 残差(全量跑完也不会清零的三类)

多款型歧义(唯一值规则拒写)、污染锚点行(等车长修复包批准)、DCD 无对应款型。当前计数看板与 `进度与ETA报告.md` 持续给出。

Copyright © 2026 David YE
