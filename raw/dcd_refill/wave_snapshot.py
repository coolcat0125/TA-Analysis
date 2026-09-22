# -*- coding: utf-8 -*-
"""wave_snapshot.py — 监控快照（供动态工作流轮询 / wave_status 的 JSON 模式）

输出一份 JSON：批次阶段 + 全局指标 + ETA。--sleep N 先睡 N 秒再采（轮询节奏）。
done 条件：本地时间 ≥ 次日 06:30，或存在 monitor_stop.txt。
"""
import io
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
STOP = os.path.join(HERE, 'monitor_stop.txt')

if '--sleep' in sys.argv:
    time.sleep(int(sys.argv[sys.argv.index('--sleep') + 1]))

out = io.StringIO()
_old = sys.stdout
sys.stdout = out
try:
    # ---- 批次状态 ----
    st = json.load(open(os.path.join(HERE, 'batch_state.json'), encoding='utf-8'))
    mapping = json.load(open(os.path.join(HERE, 'series_mapping.json'),
                             encoding='utf-8'))
    ok_kws = {str(m['kw']).lower().replace(' ', '')
              for m in mapping.get('reviewed', []) if m.get('ok')}
    rej = sum(1 for m in mapping.get('reviewed', [])
              if not m.get('ok')
              and str(m['kw']).lower().replace(' ', '') not in ok_kws)
    import openpyxl
    wb = openpyxl.load_workbook(
        [os.path.join(ROOT, f) for f in os.listdir(ROOT)
         if f.endswith('.xlsx') and '410' in f and '汇总' in f][0], read_only=True)
    led = wb['变更记录']
    dcd_cells = 0
    last_seq = 0
    for r in led.iter_rows(min_row=2, values_only=True):
        if isinstance(r[0], int):
            last_seq = max(last_seq, r[0])
            src = r[8] if len(r) > 8 else None
            if isinstance(src, str) and '懂车帝' in src:
                dcd_cells += 1
    # rear_only 池（gap_reclass.json 超 2h 则重算）
    gr_path = os.path.join(HERE, 'gap_reclass.json')
    stale = (not os.path.exists(gr_path)) or \
            (time.time() - os.path.getmtime(gr_path) > 7200)
    if stale:
        subprocess.run([sys.executable, os.path.join(HERE, 'reclass_rear_motor.py')],
                       capture_output=True, timeout=300)
    gr = json.load(open(gr_path, encoding='utf-8')) if os.path.exists(gr_path) else {}
    rear_only = gr.get('counts', {}).get('rear_only', 0)
    wb.close()
finally:
    sys.stdout = _old

status = json.load(open(os.path.join(HERE, 'status.json'), encoding='utf-8'))
cov = status.get('dcd_coverage', {})
fields = ['整备质量(kg)', '轴距(mm)', '纯电续航里程(km)', '电池容量(kWh)', '电池类型',
          '电池能量密度(Wh/kg)', '百公里电耗(kWh/100km)', '前电机功率/扭矩',
          '电机总功率/扭矩', '后电机功率/扭矩', '车长(mm)']
vals = [cov[f] for f in fields if f in cov]
coverage = round(sum(vals) / len(vals), 1) if vals else 0

warn = '?'
cr_path = os.path.join(ROOT, 'audit-output', 'consistency_report.json')
if os.path.exists(cr_path):
    try:
        cr = json.load(open(cr_path, encoding='utf-8'))
        warn = cr.get('warnTotal', '?')
    except Exception:
        pass

now = datetime.now()
nightly = now.hour >= 23 or now.hour < 6
base = datetime(2026, 9, 22, 16, 26)
hours = max((now - base).total_seconds() / 3600, 0.1)
rate = max((dcd_cells - 1036) / hours, 30)
remaining = rear_only + rej
eta_h = remaining / rate
eta_done = now + timedelta(hours=eta_h + 1)
eta = (f"剩余池≈{remaining}（后驱{rear_only}+无源{rej}），速率≈{rate:.0f}格/h → "
       f"后驱定向回填预计 {eta_done.strftime('%m-%d %H:%M')} 前"
       + ("（夜间自动轮作业中）" if nightly else "（日间/待自动轮接续）"))

batches = []
for bid in sorted(st['batches']):
    b = st['batches'][bid]
    ss = b.get('series', [])
    mapped = sum(1 for s in ss if s.get('ok') and s.get('sid'))
    ap = b.get('applied') or {}
    green = ap.get('green')
    status_b = b.get('status', 'planned')
    if status_b == 'applied':
        stage = '已回填'
    elif status_b in ('ingested', 'partial'):
        stage = '数据就绪'
    elif mapped > 0:
        stage = '映射中'
    else:
        stage = '待处理'
    batches.append({'id': bid, 'stage': stage, 'mapped': mapped,
                    'total': len(ss), 'green': green})

snap = {
    'ts': now.strftime('%H:%M:%S'),
    'nightly': nightly,
    'totals': {'series': sum(b[4] if False else len(st['batches'][b]['series'])
                             for b in st['batches']),
               'mapped': sum(x['mapped'] for x in batches),
               'ledger_cells': dcd_cells, 'last_seq': last_seq,
               'warn': warn, 'coverage': coverage,
               'rear_only': rear_only, 'rejected': rej},
    'batches': batches,
    'eta': eta,
    'done': os.path.exists(STOP) or (now.hour >= 6 and now.hour < 23 and
                                     now.date() > base.date()),
}
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
print(json.dumps(snap, ensure_ascii=False))
