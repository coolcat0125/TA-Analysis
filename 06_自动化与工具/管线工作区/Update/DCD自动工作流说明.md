# DCD 自动优化工作流说明（懂车帝参数通道 · 分批持续查缺补漏）

> 建立日期：2026-09-21。本工作流把既有「Tabbit 浏览器通道 + 门禁纪律」固化为**可分批、可断点续跑、可实时监控**的自动管线。
> 上游依据：`Update/DCD参数回填_进度跟踪.md`（通道配方/字段映射）、`Update/白天分派_20260921.md`（Agent B 车道任务）。
> 用户指令（09-20）：以懂车帝车型参数页为源刷新底表，分阶段执行、阶段间可接力。优先级：缺失 > 不确定推测 > 存疑 > 逻辑异常。

## 1. 工作流全景

```
┌──────────┐   ┌──────────┐   ┌──────────┐   ┌──────────┐   ┌──────────┐
│ scan_gaps│ → │  plan    │ → │  jobgen  │ → │  ingest  │ → │  apply   │
│ 缺口盘点  │   │ 批次计划  │   │ 抽取任务  │   │ 结果登记  │   │ 回填写库 │
└──────────┘   └──────────┘   └──────────┘   └──────────┘   └────┬─────┘
                                                                 ↓
                              ┌─────────────────────────── monitor 监控看板
                              │  status.json ←（每步刷新）    DCD监控看板.html
```

| 环节 | 命令 | 产物 |
|---|---|---|
| ① 缺口盘点 | `python scan_gaps.py` | `raw/dcd_refill/gap_inventory.json` + `audit-output/gap_summary.txt` |
| ② 批次计划 | `python dcd_batch_run.py plan --batch-size 5 --top 40` | `raw/dcd_refill/batch_state.json`（断点续跑依据） |
| ③ 任务生成 | `python dcd_batch_run.py jobgen --batch b01` | `raw/dcd_refill/jobs/b01/*.js` + `manifest.json` |
| ④ 通道抓取 | 主力终端按配方执行（Tabbit 浏览器） | `raw/dcd_refill/dcd_s<sid>.json`（落盘后回步骤⑤） |
| ⑤ 结果登记 | `python dcd_batch_run.py ingest --batch b01` | 批次状态 → ingested/partial/waiting |
| ⑥ 回填执行 | `python dcd_batch_run.py apply --batch b01`（dry-run 复核）→ 加 `--apply` 落库 | 底表绿填 + 变更记录台账 |
| ⑦ 状态刷新 | `python dcd_batch_run.py status && python build_monitor.py` | `status.json` + `DCD监控看板.html` |
| ⑧ 实时监控 | `python serve_monitor.py` → http://localhost:8797 | 看板 3s 自动刷新 |

## 2. 缺口优先级（scan_gaps 分类口径）

| 级别 | 含义 | 处置 |
|---|---|---|
| P1 缺失 | 字段为空（语义占位 `不适用(BEV)` 不算缺失） | DCD 通道可补 |
| P2 占位 | `未提供`/`0` 等无效占位 | DCD 通道可补 |
| P3 存疑 | BEV 行带油耗/发动机值、电耗与容量/续航不自洽（>20%） | 人工核证，DCD 不作来源 |
| P4 越界 | 物理范围外（如整备>4000kg、续航>1500km） | 甄别后处置 |

## 3. 首批盘点结果（2026-09-21 实测）

- 底表 5,392 行（341~410 批，38 列）
- **DCD 可补缺口 9,765 格**：后电机功率/扭矩 3,709、能量密度 1,151、前电机功率/扭矩 994、车长 937、轴距 922、整备 527、电池容量 474、电池类型 448、电耗 277、电机总功率 183、续航 143
- 存疑/越界 1,972 处（P3/P4，不进 DCD 通道）
- 批次计划：Top40 车系 → **8 批 × 5 车系，覆盖缺口约 1,374 格**（b01: V919/欧拉R1/瑞驰EC75/菱智M5EV/比亚迪唐 …）

## 4. 通道配方（主力终端执行）

Tabbit 浏览器（用户会话，无登录墙；urllib 直连不可用——登录墙 + msToken/a_bogus 签名）：

```
%LOCALAPPDATA%\Tabbit\LocalAgent\bin\tabbit-cli.exe nodejs --task dcd-params --request-id <RID> --timeout-ms 180000 < jobs\b01\<RID>_xxx.js
```

- 每批 5 车系（180s 限额内），页间等待 1.5~3s
- JS 逻辑：搜索页取 series_id → 参数页取 `__NEXT_DATA__.props.pageProps.rawData` → `car_info[].info[key].value` 九字段 + 款型目录
- 结果 JSON 落盘 `raw/dcd_refill/dcd_s<sid>.json`（结构：`{kw, series_url, cars:[{id,name,vals:{...}}]}`）

## 5. 回填纪律（apply 强制，继承门禁铁律）

1. **只补空**：不覆盖已有值（`skip_filled` 计数可见）
2. **唯一值**：并列款型候选值不一致 → 放弃该格（`ambiguous_skip`）
3. **品牌校验**：series_mapping.json 人审 `ok` 才接入；搜索关键词与底表车系名一致性核验
4. **能源对齐**：BEV 行只匹配带电池款型；`ff=汽油` 款型排除
5. **W5 成对铁律**：车长/轴距必须双双有唯一值才写（缺一项整行 hold）
6. **修剪匹配**：续航±15km + 整备±40kg 打分取最优
7. **逐格台账**：变更记录追加 `媒体参数补全`，来源注明 `懂车帝参数页(sid=…)`
8. **底纹**：绿 C6EFCE（补充）/ 红 FFC7CE（双源更正，T4 锚点场景）

## 6. 实时监控看板

- `DCD监控看板.html`：覆盖率条 / 批次进度 / 台账流水 / KPI 四卡
- 实时模式：`python serve_monitor.py` 后访问，每 3s fetch `status.json`，有更新即整页刷新
- 离线模式：双击显示快照，把最新 `status.json` 拖入页面刷新

## 7. 每轮收尾（三件套纪律）

apply 落库后必须：① 重跑 `python generate_dashboard.py`（两版看板）② `python audit_data.py --strict` ③ `python verify_consistency.py` ④ 提交注明"DCD 通道（懂车帝）"⑤ 回写 `Update/DCD参数回填_进度跟踪.md` 阶段状态。

## 8. 已知边界

- 电芯供应商/电池包供应商不在 DCD 标准参数范围——维持公示详情页通道
- 车长/轴距受 W5 成对约束，单字段证据不足时该行 hold（历史教训：估算红修曾致比值爆增 66）
- 411 批公示数据未正式发布前不入库；410 正式已对齐（v4.9.0）
