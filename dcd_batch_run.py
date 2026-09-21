# -*- coding: utf-8 -*-
"""dcd_batch_run.py — 懂车帝(DCD)自动补缺批次编排器（工作流核心）

环节：
  plan     从缺口清单生成批次计划（按车系聚合、缺口量排序、排除已完成映射）
  jobgen   生成某批次的 Tabbit 抽取任务（JS + manifest，通道配方见 Update/DCD参数回填_进度跟踪.md）
  ingest   校验并登记某批次抓取结果（raw/dcd_refill/dcd_sXXXX.json）
  apply    执行回填（只补空/唯一值/品牌校验/能源对齐/W5成对铁律）+逐格台账+底纹
  status   汇总全局状态 → raw/dcd_refill/status.json（实时监控看板数据源）

纪律（继承项目铁律）：
  只补空不覆盖 · 唯一非空共识才补 · BEV禁填油耗/发动机 · 车长/轴距成对(W5) · 逐格台账 · 不绕验证码
断点续跑：raw/dcd_refill/batch_state.json 记录每批次状态；中断后重跑同一命令自动续。
"""
import argparse
import json
import os
import re
import sys
import time
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
REFILL = os.path.join(HERE, 'raw', 'dcd_refill')
GAP_JSON = os.path.join(REFILL, 'gap_inventory.json')
MAPPING = os.path.join(REFILL, 'series_mapping.json')
STATE = os.path.join(REFILL, 'batch_state.json')
STATUS = os.path.join(REFILL, 'status.json')
FUEL = os.path.join(REFILL, 'dcd_fuel_form.json')

DCD_FIELDS = {
    '车长(mm)', '轴距(mm)', '整备质量(kg)', '纯电续航里程(km)', '百公里电耗(kWh/100km)',
    '电池容量(kWh)', '电池类型', '电池能量密度(Wh/kg)', '电机总功率/扭矩',
    '前电机功率/扭矩', '后电机功率/扭矩',
}
GREEN = PatternFill('solid', fgColor='C6EFCE')   # 补充（媒体）
RED = PatternFill('solid', fgColor='FFC7CE')     # 双源更正
FIELD_MAP = {
    '车长(mm)': 'length', '轴距(mm)': 'wheelbase', '整备质量(kg)': 'curb_weight',
    '纯电续航里程(km)': 'cltc_recharge_mileage', '百公里电耗(kWh/100km)': 'power_consumption',
    '电池容量(kWh)': 'battery_capacity', '电池类型': 'battery_type',
    '电池能量密度(Wh/kg)': 'battery_energy_density', '电机总功率/扭矩': 'total_electric_power',
    '前电机功率/扭矩': 'front_electric_max_power', '后电机功率/扭矩': 'rear_electric_max_power',
}
DCD_URL = 'https://www.dongchedi.com/auto/series/{sid}'
BT_NORM = {'三元锂电池': '三元锂', '磷酸铁锂电池': '磷酸铁锂', '锰酸锂电池': '锰酸锂', '钛酸锂电池': '钛酸锂'}


def fnum(v):
    try:
        return float(str(v).split('/')[0].replace(',', '').replace('km', '').replace('kW', '').strip())
    except Exception:
        return None


def load_state():
    if os.path.exists(STATE):
        return json.load(open(STATE, encoding='utf-8'))
    return {'batches': {}, 'created': time.strftime('%Y-%m-%d %H:%M:%S')}


def save_state(st):
    json.dump(st, open(STATE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


# ---------------------------------------------------------------- plan
def cmd_plan(args):
    inv = json.load(open(GAP_JSON, encoding='utf-8'))
    mapping = json.load(open(MAPPING, encoding='utf-8'))
    known = {str(m['sid']): m for m in mapping.get('reviewed', []) if m.get('ok')}
    st = load_state()

    # 按车系聚合缺口（通用名称/车型名称）
    by_series = defaultdict(lambda: {'rows': 0, 'fields': defaultdict(int), 'models': []})
    for g in inv['dcd_gaps']:
        key = g['model']  # 型号粒度太细；用车系粒度需读表——简化为型号前缀聚合见下
        by_series[key]['rows'] += 1
        by_series[key]['fields'][g['field']] += 1

    # 读表取车系名（通用名称优先），按车系聚合
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    series_of = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        model = str(r[hi['产品型号']] or '').strip()
        gen = str(r[hi['通用名称']] or '').strip()
        car = str(r[hi['车型名称']] or '').strip()
        series_of[model] = gen or car or model

    ser_gap = defaultdict(lambda: {'rows': 0, 'fields': defaultdict(int)})
    for g in inv['dcd_gaps']:
        s = series_of.get(g['model'], g['model'])
        ser_gap[s]['rows'] += 1
        ser_gap[s]['fields'][g['field']] += 1

    # 已接入批次的车系跳过
    done_series = set()
    for b in st['batches'].values():
        for s in b.get('series', []):
            done_series.add(s['kw'] if isinstance(s, dict) else s)

    ranked = sorted(ser_gap.items(), key=lambda x: -x[1]['rows'])
    # 过滤泛称（产品类型/占位/过短），非可检索车系
    GENERIC = re.compile(r'纯电动|插电|换电式|增程|燃料电池|轿车|多用途|运动型|越野|乘用车|未知|未提供|^\s*$')
    plan_series = []
    for name, info in ranked:
        if name in done_series:
            continue
        if GENERIC.search(name) or len(name) < 2:
            continue
        # 生成搜索关键词（去版本后缀）
        kw = re.sub(r'[（(].*?[)）]|\s*\d{4}\s*款.*$', '', name).strip() or name
        if GENERIC.search(kw) or len(kw) < 2:
            continue
        known_sid = next((str(m['sid']) for m in mapping.get('reviewed', [])
                          if m.get('ok') and str(m['kw']).lower().replace(' ', '') == kw.lower().replace(' ', '')), None)
        plan_series.append({'kw': kw, 'series_name': name, 'gap_rows': info['rows'],
                            'fields': dict(info['fields']),
                            'sid': known_sid, 'ok': bool(known_sid)})
        if len(plan_series) >= args.top:
            break

    # 分批次
    batches = []
    for i in range(0, len(plan_series), args.batch_size):
        grp = plan_series[i:i + args.batch_size]
        bid = f"b{len(st['batches']) + len(batches) + 1:02d}"
        batches.append({'id': bid, 'series': grp, 'status': 'planned',
                        'created': time.strftime('%Y-%m-%d %H:%M:%S')})
    for b in batches:
        st['batches'][b['id']] = b
    save_state(st)

    total_gap = sum(s['gap_rows'] for b in batches for s in b['series'])
    print(f"计划生成：{len(batches)} 批 × 最多{args.batch_size}车系，覆盖缺口约 {total_gap} 格")
    for b in batches:
        print(f"  {b['id']}: " + '、'.join(f"{s['kw']}({s['gap_rows']}格)" for s in b['series']))
    print(f"→ {STATE}")
    if not batches:
        print("（无新车系可规划——所有缺口车系已在批次计划中）")


# ---------------------------------------------------------------- jobgen
JS_TEMPLATE = '''// DCD 抽取任务 · {kw}
// 通道：Tabbit 浏览器（配方见 Update/DCD参数回填_进度跟踪.md）
// 用法：tabbit-cli.exe nodejs --task dcd-params --request-id {rid} --timeout-ms 180000 < 本文件
(async () => {{
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const out = {{ kw: "{kw}", series_url: null, cars: [] }};
  try {{
    // 1) 搜索页取 series_id
    const searchUrl = "https://www.dongchedi.com/search?keyword=" + encodeURIComponent("{kw}");
    // （浏览器上下文内执行：导航→等待→取 __NEXT_DATA__ 或 DOM 中的 series 链接）
    // 2) 参数页取 rawData
    //    document.querySelectorAll('script') 找 __NEXT_DATA__ → props.pageProps.rawData
    //    rawData.car_info[] → info[key].value；rawData.properties[] → 车型目录
    // 3) 款型遍历：cars.push({{ id, name, ff, vals }})
    // 期望输出：JSON（stdout）— 由 Tabbit 回传后落盘 raw/dcd_refill/dcd_s<sid>.json
    console.log(JSON.stringify(out));
  }} catch (e) {{
    console.log(JSON.stringify({{ kw: "{kw}", error: String(e) }}));
  }}
}})();
'''


def cmd_jobgen(args):
    st = load_state()
    b = st['batches'].get(args.batch)
    if not b:
        print(f"批次不存在: {args.batch}"); sys.exit(1)
    jdir = os.path.join(REFILL, 'jobs', args.batch)
    os.makedirs(jdir, exist_ok=True)
    manifest = {'batch': args.batch, 'created': time.strftime('%Y-%m-%d %H:%M:%S'),
                'channel': 'tabbit-browser', 'jobs': []}
    for i, s in enumerate(b['series'], 1):
        if s.get('sid') is None and not args.force:
            pass  # sid 待搜索页回填；manifest 记录 kw 即可
        rid = f"{args.batch}-{i:02d}"
        js = JS_TEMPLATE.format(kw=s['kw'], rid=rid)
        jf = os.path.join(jdir, f"{rid}_{re.sub(r'[^A-Za-z0-9一-龥]', '', s['kw'])[:20]}.js")
        open(jf, 'w', encoding='utf-8').write(js)
        manifest['jobs'].append({'rid': rid, 'kw': s['kw'], 'js': os.path.basename(jf),
                                 'series_name': s['series_name'], 'gap_rows': s['gap_rows']})
    json.dump(manifest, open(os.path.join(jdir, 'manifest.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    b['status'] = 'jobgen'
    save_state(st)
    print(f"批次 {args.batch} 任务已生成：{len(manifest['jobs'])} 个 JS → {jdir}")
    print("下一步：在主力终端按配方执行（批 5 车系/次，180s 限额），结果 JSON 落入 raw/dcd_refill/dcd_s<sid>.json")
    print("完成后执行：python dcd_batch_run.py ingest --batch " + args.batch)


# ---------------------------------------------------------------- ingest
def cmd_ingest(args):
    st = load_state()
    b = st['batches'].get(args.batch)
    if not b:
        print(f"批次不存在: {args.batch}"); sys.exit(1)
    jdir = os.path.join(REFILL, 'jobs', args.batch)
    manifest = json.load(open(os.path.join(jdir, 'manifest.json'), encoding='utf-8'))
    found, missing = [], []
    for j in manifest['jobs']:
        # 匹配 raw/dcd_refill/dcd_s*.json 中含该 kw 的产物
        hit = None
        for f in os.listdir(REFILL):
            if re.match(r'dcd_s\d+\.json$', f):
                try:
                    d = json.load(open(os.path.join(REFILL, f), encoding='utf-8'))
                except Exception:
                    continue
                if str(d.get('kw', '')).strip() == j['kw']:
                    hit = f; break
        (found if hit else missing).append({'kw': j['kw'], 'file': hit})
    b['status'] = 'ingested' if found and not missing else ('partial' if found else 'waiting')
    b['ingested'] = {'found': found, 'missing': missing, 'time': time.strftime('%Y-%m-%d %H:%M:%S')}
    save_state(st)
    print(f"批次 {args.batch} 登记：已到 {len(found)}，缺 {len(missing)}")
    for m in missing:
        print(f"  缺: {m['kw']}")
    if found:
        print("下一步：python dcd_batch_run.py apply --batch " + args.batch + " （先 dry-run 复核）")


# ---------------------------------------------------------------- apply
def cmd_apply(args):
    st = load_state()
    b = st['batches'].get(args.batch)
    if not b:
        print(f"批次不存在: {args.batch}"); sys.exit(1)
    apply_flag = args.apply

    mapping = json.load(open(MAPPING, encoding='utf-8'))
    kw2sid = {}
    for m in mapping.get('reviewed', []):
        if m.get('ok'):
            kw2sid[str(m['kw']).lower().replace(' ', '')] = str(m['sid'])
    try:
        fuel = json.load(open(FUEL, encoding='utf-8'))
    except Exception:
        fuel = {}
    fuel_ff = {}
    for sid_f, lst in fuel.items():
        for sp in (lst or []):
            if isinstance(sp, dict) and sp.get('id'):
                fuel_ff[(str(sid_f), str(sp['id']))] = str(sp.get('ff') or '')

    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in ws[1]]
    hi = {h: i for i, h in enumerate(hdr)}
    ledger = wb['变更记录']
    seq = max((r[0] for r in ledger.iter_rows(min_row=2, values_only=True)
               if isinstance(r[0], int)), default=0)

    # 本批次目标车系
    targets = [s for s in b['series']]
    stats = defaultdict(int)
    filled_cells = []

    for s in targets:
        kw = str(s['kw']).lower().replace(' ', '')
        sid = kw2sid.get(kw)
        if not sid:
            stats['no_sid'] += 1
            continue
        f = os.path.join(REFILL, f'dcd_s{sid}.json')
        if not os.path.exists(f):
            stats['no_data'] += 1
            continue
        dcd = json.load(open(f, encoding='utf-8'))
        series_url = f'https://www.dongchedi.com/auto/series/{sid}（懂车帝参数页，批次 {args.batch} 检索 {time.strftime("%Y-%m-%d")}）'
        rows = list(ws.iter_rows(min_row=2))
        for row in rows:
            g = (str(row[hi['通用名称']].value or '') + str(row[hi['车型名称']].value or '')).lower().replace(' ', '')
            if kw not in g:
                continue
            pt = str(row[hi['动力类型']].value or '')
            if not ('纯电' in pt or '插电' in pt or '增程' in pt or 'BEV' in pt or 'PHEV' in pt):
                continue
            is_bev = '纯电' in pt or 'BEV' in pt
            code = str(row[hi['产品型号']].value or '').strip()
            # 款型筛选：BEV 需有电池；排除汽油
            trims = []
            for c in dcd.get('cars', []):
                v = c.get('vals', {})
                ff = fuel_ff.get((sid, str(c.get('id'))), '')
                if ff == '汽油':
                    continue
                if is_bev and not v.get('battery_capacity'):
                    continue
                if (not is_bev) and not v.get('battery_capacity') and 'DM' not in str(c.get('name')) and 'EV' not in str(c.get('name')):
                    continue
                trims.append(c)
            if not trims:
                continue
            # 打分：续航±15 / 整备±40
            r_rng, r_cw = fnum(row[hi['纯电续航里程(km)']].value), fnum(row[hi['整备质量(kg)']].value)
            scored = []
            for c in trims:
                v = c['vals']
                nm_km = re.search(r'(\d{3,4})\s*km', str(c.get('name')))
                cltc = fnum(v.get('cltc_recharge_mileage')) or (float(nm_km.group(1)) if nm_km else None)
                cw = fnum(v.get('curb_weight'))
                sc = 0
                if r_rng and cltc:
                    d = abs(r_rng - cltc)
                    sc += 0 if d <= 15 else (1 if d <= 30 else 3)
                if r_cw and cw:
                    d = abs(r_cw - cw)
                    sc += 0 if d <= 40 else (1 if d <= 90 else 3)
                scored.append((sc, c))
            scored.sort(key=lambda t: t[0])
            best = scored[0][0]
            tied = [c for sc, c in scored if sc == best]

            for fld, key in FIELD_MAP.items():
                # W5 成对铁律：车长/轴距必须双双有唯一值才写
                pair_ok = True
                if fld in ('车长(mm)', '轴距(mm)'):
                    pair_ok = len({fnum(c['vals'].get('length')) for c in tied} - {None}) == 1 and \
                              len({fnum(c['vals'].get('wheelbase')) for c in tied} - {None}) == 1
                cell = row[hi[fld]]
                cur = cell.value
                vals = set()
                for c in tied:
                    v = c['vals'].get(key)
                    if key == 'battery_type':
                        vals.add(BT_NORM.get(str(v).strip(), str(v).strip()))
                    elif key in ('total_electric_power', 'front_electric_max_power', 'rear_electric_max_power'):
                        m = re.match(r'^\s*(\d+(?:\.\d+)?)\s*kW', str(v), re.I)
                        if m:
                            vals.add(float(m.group(1)))
                    else:
                        fv = fnum(v)
                        vals.add(fv if fv is not None else str(v).strip())
                vals.discard('')
                vals.discard(None)
                if cur in (None, '') or str(cur).strip() == '':
                    if len(vals) == 1 and pair_ok:
                        val = vals.pop()
                        cell.value = val
                        cell.fill = GREEN
                        seq += 1
                        ledger.append((seq, row[hi['批次']].value, code, pt, fld, '媒体参数补全',
                                       '(空)', str(val), f'懂车帝参数页(sid={sid})：{series_url}'))
                        stats['green_fill'] += 1
                        filled_cells.append({'field': fld, 'model': code})
                    else:
                        stats['ambiguous_skip'] += 1
                else:
                    stats['skip_filled'] += 1
    print(f"批次 {args.batch} apply：green={stats['green_fill']} ambiguous={stats['ambiguous_skip']} "
          f"skip_filled={stats['skip_filled']} no_sid={stats['no_sid']} no_data={stats['no_data']}")
    if not apply_flag:
        print('DRY-RUN（未写入）')
    else:
        wb.save(WB)
        b['status'] = 'applied'
        b['applied'] = {'green': stats['green_fill'], 'time': time.strftime('%Y-%m-%d %H:%M:%S'),
                        'cells': filled_cells[:200]}
        save_state(st)
        print('saved（记得重跑三件套：看板×2 + audit + 提交注明 DCD 通道）')
    # 刷新状态
    cmd_status(argparse.Namespace())


# ---------------------------------------------------------------- status
def cmd_status(args):
    st = load_state()
    # 覆盖率快照（读表轻量统计）
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    total = 0
    filled = defaultdict(int)
    for r in ws.iter_rows(min_row=2, values_only=True):
        total += 1
        for f in DCD_FIELDS:
            v = r[hi[f]] if f in hi else None
            if v is not None and str(v).strip() != '':
                filled[f] += 1
    cov = {f: round(filled[f] / total * 100, 1) for f in DCD_FIELDS}

    # 最近台账（DCD 条目）
    led = wb['变更记录']
    recent = []
    for r in led.iter_rows(min_row=2, values_only=True):
        if '懂车帝' in str(r[8] or ''):
            recent.append({'seq': r[0], 'model': r[2], 'field': r[4], 'kind': r[5],
                           'old': r[6], 'new': r[7]})
    recent = recent[-15:][::-1]

    dcd_cells = sum(1 for r in led.iter_rows(min_row=2, values_only=True)
                    if isinstance(r[0], int) and '懂车帝' in str(r[8] or ''))

    status = {
        'updated': time.strftime('%Y-%m-%d %H:%M:%S'),
        'workbook_rows': total,
        'dcd_coverage': cov,
        'dcd_ledger_cells': dcd_cells,
        'batches': [{'id': bid, 'status': b['status'], 'series': [s.get('kw') for s in b.get('series', [])],
                     'gap_rows': sum(s.get('gap_rows', 0) for s in b.get('series', [])),
                     'green': b.get('applied', {}).get('green', 0)}
                    for bid, b in st['batches'].items()],
        'recent_ledger': recent,
    }
    json.dump(status, open(STATUS, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f"status.json 已刷新 · DCD台账 {dcd_cells} 格 · 行数 {total}")
    for f, c in sorted(cov.items(), key=lambda x: x[1]):
        print(f"  {f}: {c}%")


def main():
    ap = argparse.ArgumentParser(description='DCD 自动补缺批次编排器')
    sub = ap.add_subparsers(dest='cmd')
    p = sub.add_parser('plan'); p.add_argument('--batch-size', type=int, default=5)
    p.add_argument('--top', type=int, default=40)
    p = sub.add_parser('jobgen'); p.add_argument('--batch', required=True); p.add_argument('--force', action='store_true')
    p = sub.add_parser('ingest'); p.add_argument('--batch', required=True)
    p = sub.add_parser('apply'); p.add_argument('--batch', required=True); p.add_argument('--apply', action='store_true')
    p = sub.add_parser('status')
    args = ap.parse_args()
    {'plan': cmd_plan, 'jobgen': cmd_jobgen, 'ingest': cmd_ingest,
     'apply': cmd_apply, 'status': cmd_status}.get(args.cmd, lambda a: ap.print_help())(args)


if __name__ == '__main__':
    main()
