# -*- coding: utf-8 -*-
"""官方锚点落地 + 核定覆盖率清单（Phase 2a）。

官方通道结论（2026-10-25 实测）：
- EIDC doCpQueryByeOne：可查型号→批次/cpid，无验证码；但 queryCpData / queryCpParamPage
  详情页均挂「访问行为验证」滑块——按 17 AGENTS 铁律不绕验证，放弃该通道。
- www.miit.gov.cn api-gateway（jpaas）列表 + art 详情页直读：仅 411/412 公示查询列可用
  （栏目 pageId 由公示通知页获得）；历史批次栏目未被站点搜索索引，发现需浏览器遍历，
  属 17 主力终端待办，本轮不依赖。
- 因此官方行级尺寸 = 411 正式 47 条 + 412 公示 58 条（扫描底表已带 art URL）。

本脚本：
1. 读扫描底表 411/412 → 归属+归一化 → official_anchors.json（conf=A，逐型号 art URL）
2. 汇总核定范围覆盖率：15 迁移种子 / 官方锚点 / 待网络取证（Track B 工作清单）
"""
import collections
import json
import os
import sys
from datetime import date

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import SCAN_XLSX, DATA, DIAG_JSON
from oem_groups import GROUPS, BOARDS, SERIES_BOARD, board_of, attr_group
from series_norm import norm_series

OFFICIAL_JSON = os.path.join(DATA, 'official_anchors.json')
COVERAGE_MD = os.path.join(DATA, 'coverage_report.md')

# 15 迁移种子（scripts/_ress_car_len.py CAR_LEN，37 条，描述性来源、暂无 URL → Track B 补证）
CAR_LEN_SEED = {
    '唐': (4870, 'B', '唐EV 2025款 4870×1950×1725，轴距2820'),
    '比亚迪D1': (4390, 'B', '比亚迪D1 4390mm，轴距2800'),
    '元': (4360, 'B', '元Pro/元EV 4360mm，轴距2535'),
    '海豹': (4800, 'B', '海豹 4800×1875×1460，轴距2920'),
    '汉': (4995, 'B', '汉EV 4995×1910×1495，轴距2920'),
    '秦': (4765, 'B', '秦PLUS EV 4765×1837×1515，轴距2718'),
    '宋L': (4840, 'B', '宋L EV 4840×1950×1560，轴距2930'),
    '秦PLUS': (4765, 'B', '秦PLUS EV 4765mm'),
    '海狮07EV': (4830, 'B', '海狮07 EV 4830×1925×1620，轴距2930'),
    '海豚': (4290, 'B', '海豚 4290×1770×1570，轴距2700'),
    '唐L': (5040, 'A', '唐L 5040×1996×1760，轴距2950（比亚迪官网/百科）'),
    '海狮06': (4810, 'A', '海狮06 EV 4810×1920×1675，轴距2820（比亚迪官网）'),
    '海狮05': (4520, 'B', '海狮05 EV 4520×1860×1630，轴距2720（2026款多源一致）'),
    '海狮08': (5115, 'B', '海狮08 5115×1999×1800，轴距3030'),
    '大汉': (5256, 'B', '大汉 5256×1999×1510，轴距3130（工信部申报）'),
    '比亚迪E6': (4695, 'C', '比亚迪E6 4695mm（早期车型，待复核）'),
    '宋PLUS': (4785, 'B', '宋PLUS新能源 4785×1890×1660，轴距2765'),
    '海豹06': (4830, 'B', '海豹06 4830×1875×1495，轴距2790'),
    '汉L': (5050, 'B', '汉L EV 5050×1960×1505，轴距2970'),
    '海豹08': (5150, 'B', '海豹08 EV 5150×1999×1505，轴距3030'),
    '秦L': (4830, 'C', '秦L 4830×1900×1495，轴距2790（待复核）'),
    '元PLUS': (4455, 'A', '元PLUS 4455×1875×1615，轴距2720（比亚迪官网）'),
    '极氪001': (4970, 'B', '极氪001 4970×1999×1560，轴距3005（2025款5030）'),
    '极氪007': (4865, 'B', '极氪007 4865×1900×1450，轴距2928'),
    '极氪007GT': (4865, 'B', '极氪007GT 4865×1900×1445，轴距2928'),
    '极氪7X': (4825, 'B', '极氪7X 4825×1930×1666，轴距2925'),
    '极氪X': (4450, 'B', '极氪X 4450×1836×1572，轴距2750'),
    '极氪MIX': (4688, 'B', '极氪MIX 4688×1995×1755，轴距3008'),
    '吉利E5': (4615, 'B', '银河E5 4615×1901×1670，轴距2750'),
    '吉利E8': (5010, 'B', '银河E8 5010×1920×1465，轴距2925'),
    '星愿': (4135, 'B', '星愿 4135×1805×1570，轴距2650'),
    '领克Z20': (4460, 'B', '领克Z20 4460×1845×1573，轴距2755'),
    '领克Z10': (5028, 'B', '领克Z10 5028×1966×1468，轴距3005'),
    '吉利A': (4736, 'B', '几何A 4736×1804×1503，轴距2700'),
    '吉利60S': (4706, 'B', '嘉际/60S 4706×1909×1690，轴距2805'),
    '吉利80V': (4820, 'C', '吉利80V MPV 4820mm（待复核）'),
    '萌虎': (4380, 'C', '待复核'),
}


def load_scan_rows():
    wb = openpyxl.load_workbook(SCAN_XLSX, read_only=True, data_only=True)
    out = []
    for sn in ('411正式车型', '412公示车型'):
        ws = wb[sn]
        it = ws.iter_rows(values_only=True)
        next(it); next(it)
        hdr = list(next(it))
        J = {h: i for i, h in enumerate(hdr)}
        for r in it:
            if not r or not r[J['产品型号']]:
                continue
            out.append({'sheet': sn, 'batch': r[J['批次']], 'model': str(r[J['产品型号']]).strip(),
                        'mark': str(r[J['产品商标']] or ''), 'ent': str(r[J['企业名称']] or ''),
                        'common': str(r[J['通用名称']] or ''), 'len': r[J['车长(mm)']],
                        'wb': r[J['轴距(mm)']], 'wt': r[J['整备质量(kg)']],
                        'url': str(r[J['来源URL']] or '')})
    wb.close()
    return out


def main():
    os.makedirs(DATA, exist_ok=True)
    anchors = {}          # (group, series) → [ {model,batch,len,wb,wt,url,status} ]
    unmatched = []
    for r in load_scan_rows():
        key, jv, _why = attr_group(r['mark'], r['ent'])
        if not key or jv:
            continue
        bd = None
        for sub, nm in SERIES_BOARD.get(key, []):
            if sub and sub in r['common']:
                bd = nm
                break
        if bd is None:
            bd = board_of(r['mark'], r['ent'], BOARDS[key])
        s = norm_series(r['common'], key, bd) if r['common'] else None
        if not s or s == '未命名':
            unmatched.append(r)
            continue
        anchors.setdefault((key, s), []).append({
            'model': r['model'], 'batch': str(r['batch']),
            'status': '正式' if r['sheet'] == '411正式车型' else '公示',
            'len': r['len'], 'wb': r['wb'], 'wt': r['wt'], 'url': r['url']})

    json.dump({f'{k[0]}|{k[1]}': v for k, v in anchors.items()},
              open(OFFICIAL_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    # ── 覆盖率：核定范围 = diag priority 的 T1+T2+T3 ──
    diag = json.load(open(DIAG_JSON, encoding='utf-8'))
    scope = [p for p in diag['priority'] if p['tier'] in ('T1', 'T2', 'T3')]
    seed_keys = {s for s in CAR_LEN_SEED}
    rows = []
    for p in scope:
        k = (p['group'], p['series'])
        a = anchors.get(k)
        cov, val, url, conf = 'need_web', None, None, None
        has_seed = p['series'] in seed_keys
        if has_seed:
            L0, c0, _note = CAR_LEN_SEED[p['series']]
            cov, val, conf = 'seed15', L0, c0
        if a:
            vals = sorted({x['len'] for x in a if x['len'] is not None})
            # 官方（411 正式 / 412 公示）优先于种子；与种子冲突时官方覆盖、备注留痕
            cov = 'official' if any(x['status'] == '正式' for x in a) else 'official_draft'
            val = vals[0] if vals else None
            url = a[0]['url']
            conf = 'A'
        rows.append({**p, 'coverage': cov, 'cur_len': val, 'url': url, 'conf': conf,
                     'has_seed': has_seed,
                     'official_lens': sorted({x['len'] for x in (a or []) if x['len'] is not None})})

    cnt = collections.Counter(r['coverage'] for r in rows)
    by_group = collections.defaultdict(collections.Counter)
    for r in rows:
        by_group[r['group']][r['coverage']] += 1
    json.dump(rows, open(os.path.join(DATA, 'coverage_worklist.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)

    L = ['# 核定覆盖率清单', f'', f'运行：{date.today()} · 范围：T1+T2+T3 共 {len(rows)} 系', '',
         f"- 官方锚点(411正式)：{cnt['official']} 系",
         f"- 官方锚点(412公示)：{cnt['official_draft']} 系",
         f"- 15 迁移种子(待补URL)：{cnt['seed15']} 系",
         f"- 待网络取证(Track B)：{cnt['need_web']} 系", '',
         '| 集团 | 官方 | 公示锚点 | 种子 | 待取证 |', '|---|---|---|---|---|']
    for g in GROUPS:
        c = by_group[g]
        L.append(f"| {GROUPS[g][0]} | {c['official']} | {c['official_draft']} | "
                 f"{c['seed15']} | {c['need_web']} |")
    L.append('')
    L.append('## 待网络取证清单（按 T1 优先）')
    for r in rows:
        if r['coverage'] == 'need_web':
            lens = ' '.join(f'{a}×{b}' for a, b in r['lens'])
            L.append(f"- [{r['tier']}] {r['group']}/{r['series']} 行{r['rows']} "
                     f"批{r['batches'][0]}~{r['batches'][1]} 现值:{lens}")
    open(COVERAGE_MD, 'w', encoding='utf-8').write('\n'.join(L))
    print('锚点车系:', len(anchors), '| 未匹配行:', len(unmatched))
    print('覆盖率:', dict(cnt), '→', OFFICIAL_JSON)
    print('工作清单:', COVERAGE_MD)


if __name__ == '__main__':
    main()
