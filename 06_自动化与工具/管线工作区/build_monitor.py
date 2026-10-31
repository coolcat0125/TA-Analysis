# -*- coding: utf-8 -*-
"""build_monitor.py — 生成 DCD 实时监控看板（DCD监控看板.html）

数据源：raw/dcd_refill/status.json + gap_inventory.json（由 dcd_batch_run.py status 刷新）
实时模式：经 serve_monitor.py 本地服务打开时，每 3s fetch 最新 status.json 自动刷新
离线模式：双击打开显示生成时快照，支持把最新 status.json 拖入页面刷新
"""
import json
import os
import time

HERE = os.path.dirname(os.path.abspath(__file__))
REFILL = os.path.join(HERE, 'raw', 'dcd_refill')
STATUS = os.path.join(REFILL, 'status.json')
GAP = os.path.join(REFILL, 'gap_inventory.json')
OUT = os.path.join(HERE, 'DCD监控看板.html')

status = json.load(open(STATUS, encoding='utf-8'))
gap = json.load(open(GAP, encoding='utf-8'))

FIELD_LABELS = {
    '车长(mm)': '车长', '轴距(mm)': '轴距', '整备质量(kg)': '整备质量',
    '纯电续航里程(km)': '纯电续航', '百公里电耗(kWh/100km)': '百公里电耗',
    '电池容量(kWh)': '电池容量', '电池类型': '电池类型',
    '电池能量密度(Wh/kg)': '能量密度', '电机总功率/扭矩': '电机总功率',
    '前电机功率/扭矩': '前电机功率', '后电机功率/扭矩': '后电机功率',
}
STATUS_STYLE = {'planned': ('#8B7E5E', '已计划'), 'jobgen': ('#0055A5', '任务已生成'),
                'waiting': ('#808080', '待抓取'), 'partial': ('#F2A900', '部分到货'),
                'ingested': ('#0055A5', '已登记'), 'applied': ('#7FA641', '已回填')}

cov = status['dcd_coverage']
gap_by_field = gap.get('dcd_gap_by_field', {})
gap_total = gap.get('dcd_gap_total', 0)

# 覆盖率条
cov_rows = ''.join(f'''
<div class="cov-row">
  <span class="cov-name">{FIELD_LABELS.get(f, f)}</span>
  <div class="cov-bar"><div class="cov-fill" style="width:{c}%"></div></div>
  <span class="cov-num">{c}%</span>
  <span class="cov-gap">缺{gap_by_field.get(f, 0)}</span>
</div>''' for f, c in sorted(cov.items(), key=lambda x: x[1]))

# 批次表
batch_rows = ''.join(f'''
<tr>
  <td class="mono">{b['id']}</td>
  <td><span class="badge" style="background:{STATUS_STYLE.get(b['status'], ('#808080', b['status']))[0]}">{STATUS_STYLE.get(b['status'], ('', b['status']))[1]}</span></td>
  <td>{'、'.join(str(s) for s in b['series'][:5])}</td>
  <td class="mono">{b['gap_rows']}</td>
  <td class="mono" style="color:#7FA641;font-weight:700">{b['green']}</td>
</tr>''' for b in status['batches'])

# 最近台账
led_rows = ''.join(f'''
<tr>
  <td class="mono">#{r['seq']}</td>
  <td class="mono">{r['model']}</td>
  <td>{r['field']}</td>
  <td>{r['kind']}</td>
  <td class="mono">{r['old']} → <b style="color:#7FA641">{r['new']}</b></td>
</tr>''' for r in status['recent_ledger'])

html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DCD 实时监控看板 · 懂车帝补缺进度</title>
<style>
:root {{ color-scheme: light; }}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ font-family: "Microsoft YaHei", "PingFang SC", sans-serif; background: #F5F8FC; color: #333; padding: 20px; }}
.wrap {{ max-width: 1200px; margin: 0 auto; }}
header {{ display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 6px; }}
h1 {{ font-size: 22px; color: #003366; }}
#live {{ font-size: 12px; color: #F2A900; font-weight: 700; }}
#live.off {{ color: #808080; }}
.meta {{ font-size: 12px; color: #808080; margin-bottom: 16px; }}
.kpis {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px; }}
.kpi {{ background: #fff; border: 1px solid #D0D9E5; border-left: 4px solid #0055A5; padding: 14px 16px; }}
.kpi .v {{ font-size: 24px; font-weight: 700; color: #0055A5; font-family: Consolas, monospace; }}
.kpi .l {{ font-size: 12px; color: #808080; margin-top: 2px; }}
.kpi.warn {{ border-left-color: #F2A900; }} .kpi.warn .v {{ color: #F2A900; }}
.kpi.ok {{ border-left-color: #7FA641; }} .kpi.ok .v {{ color: #7FA641; }}
.grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
.card {{ background: #fff; border: 1px solid #D0D9E5; padding: 16px; }}
.card h2 {{ font-size: 14px; color: #003366; margin-bottom: 12px; border-bottom: 2px solid #0055A5; padding-bottom: 6px; }}
.cov-row {{ display: flex; align-items: center; gap: 8px; margin-bottom: 7px; }}
.cov-name {{ width: 64px; font-size: 12px; }}
.cov-bar {{ flex: 1; height: 12px; background: #EDF1F7; }}
.cov-fill {{ height: 100%; background: linear-gradient(90deg, #0055A5, #7FA8D9); }}
.cov-num {{ width: 44px; font-size: 12px; font-family: Consolas, monospace; text-align: right; }}
.cov-gap {{ width: 52px; font-size: 11px; color: #F2A900; text-align: right; }}
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
th {{ background: #0055A5; color: #fff; padding: 7px 8px; text-align: left; font-weight: 600; }}
td {{ padding: 6px 8px; border-bottom: 1px solid #EDF1F7; }}
tr:nth-child(even) td {{ background: #F8FAFD; }}
.mono {{ font-family: Consolas, monospace; font-size: 11.5px; }}
.badge {{ color: #fff; font-size: 11px; padding: 2px 8px; border-radius: 2px; }}
.tip {{ margin-top: 14px; font-size: 12px; color: #808080; background: #FFF9E6; border: 1px solid #F2A900; padding: 10px 12px; }}
.drop {{ border: 2px dashed #F2A900 !important; }}
</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>DCD 实时监控看板 · 懂车帝补缺进度</h1>
  <span id="live">● LIVE（3s 自动刷新）</span>
</header>
<div class="meta">数据源：status.json · 刷新时间 <span id="ts">{status['updated']}</span> · 工作流：dcd_batch_run.py</div>

<div class="kpis">
  <div class="kpi"><div class="v">{status['workbook_rows']:,}</div><div class="l">底表行数（341~410批）</div></div>
  <div class="kpi ok"><div class="v">{status['dcd_ledger_cells']:,}</div><div class="l">DCD 台账累计格数</div></div>
  <div class="kpi warn"><div class="v">{gap_total:,}</div><div class="l">DCD 可补缺口（剩余）</div></div>
  <div class="kpi"><div class="v">{len(status['batches'])}</div><div class="l">计划批次数</div></div>
</div>

<div class="grid">
  <div class="card">
    <h2>字段覆盖率（DCD 可映射 11 字段）</h2>
    {cov_rows}
  </div>
  <div class="card">
    <h2>批次进度</h2>
    <table>
      <tr><th>批次</th><th>状态</th><th>车系</th><th>缺口</th><th>已填</th></tr>
      {batch_rows}
    </table>
  </div>
</div>

<div class="card" style="margin-top:12px">
  <h2>最近台账流水（懂车帝通道）</h2>
  <table>
    <tr><th>Seq</th><th>产品型号</th><th>字段</th><th>类型</th><th>旧值 → 新值</th></tr>
    {led_rows}
  </table>
</div>

<div class="tip" id="tip">
  <b>使用方式</b>：① 实时模式——在仓库目录运行 <code>python serve_monitor.py</code> 后访问 <code>http://localhost:8797/DCD监控看板.html</code>，每 3 秒自动刷新；
  ② 离线模式——双击打开显示当前快照，将最新的 <code>status.json</code> 拖入本页即可刷新。工作流命令：<code>python dcd_batch_run.py plan/jobgen/ingest/apply/status</code>。
</div>
</div>

<script>
const $ = (s) => document.querySelector(s);
async function poll() {{
  try {{
    const r = await fetch('raw/dcd_refill/status.json?t=' + Date.now());
    if (!r.ok) throw 0;
    const d = await r.json();
    if (d.updated !== $('#ts').textContent) location.reload();
  }} catch (e) {{
    $('#live').textContent = '● 离线快照（拖入 status.json 刷新）';
    $('#live').classList.add('off');
  }}
}}
setInterval(poll, 3000);
document.addEventListener('dragover', e => e.preventDefault());
document.addEventListener('drop', async e => {{
  e.preventDefault();
  const f = e.dataTransfer.files[0];
  if (f && f.name.endsWith('.json')) {{
    const d = JSON.parse(await f.text());
    document.body.classList.remove('drop');
    if (d.updated) {{ $('#ts').textContent = d.updated; location.reload(); }}
  }}
}});
</script>
</body>
</html>'''

open(OUT, 'w', encoding='utf-8').write(html)
print(f"监控看板已生成: {OUT} ({os.path.getsize(OUT)//1024} KB)")
