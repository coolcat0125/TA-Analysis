# DCD 懂车帝参数回填 · 进度跟踪（多 agent 接力中枢）

> 用户指令（09-20）：以懂车帝车型参数页为源刷新底表，分阶段执行、阶段间可接力。
> 优先级：缺失字段 > 不确定推测 > 存疑 > 逻辑异常。只补空、逐格台账、BEV 禁填油耗/发动机、不绕验证码。
> **本文件是接力中枢：任何 agent 接手前先读本文件与 raw/dcd_refill/series_mapping.json。**

## 通道配方（已验证，复用）

1. Tabbit 浏览器（用户会话，无登录墙）：
   - 批处理包装：`%LOCALAPPDATA%\Tabbit\LocalAgent\bin\tabbit-cli.exe nodejs --task dcd-params --request-id <ID> --timeout-ms 180000 < <js文件>`
   - Git Bash 调用需 `MSYS_NO_PATHCONV=1 cmd.exe /d /c <包装cmd> <ID> <js路径>`
   - JS 程序模板：`raw/dcd_refill/` 系列抽取脚本逻辑（搜索页取 series_id → params 页取 `__NEXT_DATA__.props.pageProps.rawData`）
2. **urllib 直连不可用**（登录墙 + API 需 msToken/a_bogus 签名）——必须走浏览器
3. 速率：每页 1.5~3s 等待；批 5 车系/次（180s 限额内）

## 字段映射（DCD key → 底表列）

| 底表列 | DCD key | 备注 |
|---|---|---|
| 车长(mm) | length | |
| 轴距(mm) | wheelbase | |
| 整备质量(kg) | curb_weight | |
| 纯电续航里程(km) | cltc_recharge_mileage | 工信部版 recharge_mileage 备选 |
| 百公里电耗(kWh/100km) | power_consumption | |
| 电池容量(kWh) | battery_capacity | |
| 电池类型 | battery_type | |
| 电池能量密度(Wh/kg) | battery_energy_density | |
| 电机总功率/扭矩 | total_electric_power / total_electric_torque | 拼接 P/T 文本 |
| 前/后电机功率/扭矩 | front/rear_electric_max_power + torque | |
| **电芯供应商** | **无标准列** | 继续走公示详情页通道 |

值位置：`rawData.car_info[].info[key].value`；车型目录：`rawData.properties[]`（308 项，样例存 `~/AppData/Local/Temp/dcd_catalog_9660.json`）。

## 阶段状态

| 阶段 | 内容 | 状态 |
|---|---|---|
| 阶段0 | 通道探测与结构闭环 | ✅ 09-20 夜（urllib 墙→Tabbit 通） |
| 阶段1 | 首批 15 车系抽取 | ✅ 09-20 夜（13 有效+2 拒收，`raw/dcd_refill/dcd_s*.json`，映射表 series_mapping.json） |
| 阶段2 | 回填脚本 + 首批 apply | 🔶 **脚本就绪，dry-run=354 格**（能源对齐过滤后；整备质量已排除）；**待安全窗执行**：`python _dcd_apply_refill.py --dry` 复核 → 去 --dry → 三件套 → 提交（信息注明 DCD 通道） |
| 阶段3 | 扩展映射（缺口Top40 → 全量） | ⬜ 未开始（首击命中率 13/15，apply 带品牌校验后可放量） |
| 阶段4 | 存疑/异常字段复核（优先级 2~4） | ⬜ 未开始 |

## 抽取清单（阶段1产物）

`raw/dcd_refill/dcd_s*.json` 14 文件（含 2 个拒收映射的 sid 留档）+ `dcd_fuel_form.json`（能源对齐）+ `series_mapping.json`。

## 首批映射审核表

见 `raw/dcd_refill/series_mapping.json`。规则：**series_name 与关键词/底表企业品牌一致性校验，不符即拒**（首击 88% 命中，不可盲取）。
拒收记录：菱智M5EV→远志M1（品牌不符）、比亚迪e3→AION S（搜索未中）。

## 下一步（接力指引）

1. 自动轮静默窗（无 python 进程 ≥5 分钟）或 06:30 后执行 `_dcd_apply_refill.py --dry` → 复核 → 去掉 --dry 落库 → 三件套 → 提交
2. 阶段3 扩展映射：按"缺口Top40"清单继续搜索页映射（脚本同阶段1），每批 5~8 车系
3. 全部 apply 须在提交信息注明"DCD 通道（懂车帝）"，台账 数据来源=`懂车帝参数页 · <series_url> · dcd_refill`
4. 电芯供应商缺口（3706）不在本通道范围——维持公示详情页通道
