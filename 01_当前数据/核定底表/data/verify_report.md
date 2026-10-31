# 核定底表 · 校验报告

运行：2026-10-25 · 明细 5524 行 · 12家有效行 2976 · 台账 392 系

## FAIL（0）
- 无

## WARN（1）
- ⚠️ V2 质量车长比出界 3 行（已打标）；电量密度出界 4 行；轴距/核定长出带 [0.5,0.72] 69 行（公告轴距可疑残留）

## 备注与对账
- V1-5 EREV 独立类 124 行（0 行时提示：341-411 增程多以 PHEV 名义公告+是否增程标志）
- V2 轴距出带样例: gwm/白猫 CC7001ZM00EBEV 轴距2650.0/核定长3625=0.73
- V2 轴距出带样例: saic/NANOEV LZW7002EVUHAN 轴距2600.0/核定长2497=1.04
- V2 轴距出带样例: saic/NANOEV LZW7002EVA1EAN 轴距2600.0/核定长2497=1.04
- V2 轴距出带样例: saic/NANOEV LZW7002EVB1EAN 轴距2600.0/核定长2497=1.04
- V2 轴距出带样例: saic/NANOEV LZW7002EVBHANA 轴距2600.0/核定长2497=1.04
- V2 轴距出带样例: saic/NANOEV LZW7002EVCHANA 轴距2600.0/核定长2497=1.04
- V2 轴距出带样例: saic/NANOEV LZW7002EVB1PAN 轴距2600.0/核定长2497=1.04
- V2 轴距出带样例: gwm/白猫 CC7001ZM02ABEV 轴距2650.0/核定长3625=0.73
- V2 轴距出带样例: saic/KIWIEV LZW7003EVB3EAB 轴距2600.0/核定长2894=0.90
- V2 轴距出带样例: saic/KIWIEV LZW7003EVB3KAB 轴距2600.0/核定长2894=0.90
- V2 轴距出带样例: saic/KIWIEV LZW7003EVB3KBE 轴距2600.0/核定长2894=0.90
- V2 轴距出带样例: saic/KIWIEV LZW7003EVGMAB 轴距2600.0/核定长2894=0.90
- V3-1 vs 15 RESS 线：xiaomi: 本表27 vs RESS线28 (差-1)（差异来源：EREV 独立类计法/411并入口径/清单补充去重键，逐条见口径说明）
- V3-2 vs 17 trends 线 master_vehicles.json：其 5392 条（341-411，含非12家、无清单补充合成行、无412）；本表明细 5524 行。口径不同属预期，车型级交集抽查见 diag D2（车长/轴距/容量 0 不一致）。
- V4 主力车系（有效行≥3）335 系，核定 281 系 = 83.9%；其中错配系 T1 70 系核定 68 = 97.1%（T1 必须达标）
- V4 未核定主力构成：商用/物流车系（大通/远程等，乘用分析边缘）+ 2025-26 新公示尚无公开参数页车型 + 已停售下架车型——公开证据不存在者如实置空，不猜
- V4 未核定主力: geely/远程SV(15行)、gwm/猛龙V7PLUS(10行)、geely/吉利E(7行)、geely/星享V6E(7行)、saic/大通1(7行)、saic/名爵MGS5(6行)、chery/星途ET7(6行)、geely/吉利EVPRO(5行)、chery/奇瑞EQ5(5行)、saic/大通MIFA9(5行)、changan/深蓝S7(5行)、geely/路特斯EMEYA(5行)、changan/长安E07(5行)、gac/埃安N60(5行)、chery/奇瑞EC1(4行)、gwm/圆梦(4行)、saic/大通MIFA7(4行)、saic/大通EV70(4行)、saic/大通MILA(4行)、chery/山海L7C-DM(4行)、chery/风云T10C-DM(4行)、saic/飞凡RC7(4行)、saic/宏光(4行)、chery/风云T8C-DM(4行)、changan/阿维塔6(4行)、geely/远程V7E(4行)、geely/星舰7EV(4行)、geely/星舰7EF(4行)、chery/旅行者(4行)、geely/吉利TT(4行)…