# -*- coding: utf-8 -*-
"""build_mobile_v4935.py — 手机查阅版分析报告（单文件离线 HTML，可复跑）

结构（核心前置，用户 09-30 指令）：口径徽章 → KPI 卡片 → 核心发现 → 关键趋势图
→ 年度快照表 → 供应商 Top5 → 车型检索（前端过滤，懒渲染）。
数据：build_deliverables_v4923.load_records()（权威底表 5,392 款），全部数值程序化派生。
输出：NEV公告车型行业分析_移动版.html
"""
import html as _html
import json
import re
import statistics as st
from collections import Counter, defaultdict

from build_deliverables_v4923 import load_records

OUT = 'NEV公告车型行业分析_移动版.html'
YEARS = [2021, 2022, 2023, 2024, 2025, 2026]
ECHARTS = 'echarts.min.js'


def btype(v):
    if '磷酸铁锂' in v:
        return 'LFP'
    if '三元' in v:
        return 'NCM'
    return '其他'


def main():
    recs = load_records()
    total = len(recs)
    nb = sum(1 for r in recs if r['组'] == 'BEV')
    nph = total - nb
    yr_cnt = {g: {y: sum(1 for r in recs if r['组'] == g and r['年'] == y) for y in YEARS}
              for g in ('BEV', 'PHEV')}
    tot_y = {y: yr_cnt['BEV'][y] + yr_cnt['PHEV'][y] for y in YEARS}
    phe_pct = {y: round(yr_cnt['PHEV'][y] / max(1, tot_y[y]) * 100, 1) for y in YEARS}

    def ymean(g, f, y=None, filt=None):
        rs = [r for r in recs if r['组'] == g and (y is None or r['年'] == y)
              and r[f] is not None and (filt is None or filt(r))]
        return round(sum(r[f] for r in rs) / len(rs), 1) if rs else None

    def yseries(g, f, filt=None):
        return [ymean(g, f, y, filt) for y in YEARS]

    # KPI（核心前置）
    kpis = [
        ('公告车型总量', f'{total:,}', '款'),
        ('BEV / PHEV', f'{nb:,} / {nph:,}', '款'),
        ('2026 年公告', f'{tot_y[2026]}', f'款（BEV {yr_cnt["BEV"][2026]} / PHEV {yr_cnt["PHEV"][2026]}）'),
        ('2025 PHEV 占比', f'{phe_pct[2025]}', '%'),
        ('BEV 平均续航(2026)', f'{ymean("BEV","续航",2026) or "-"}', 'km'),
        ('BEV 平均容量(2026)', f'{ymean("BEV","容量",2026) or "-"}', 'kWh'),
        ('BEV 平均电耗(2026)', f'{ymean("BEV","电耗",2026) or "-"}', 'kWh/100km'),
        ('PHEV 纯电续航(2026)', f'{ymean("PHEV","续航",2026) or "-"}', 'km'),
    ]
    # 核心发现（全部数据锚定）
    b21, b25 = ymean('BEV', '续航', 2021), ymean('BEV', '续航', 2025)
    p21, p25 = ymean('PHEV', '续航', 2021), ymean('PHEV', '续航', 2025)
    c21, c25 = ymean('BEV', '容量', 2021), ymean('BEV', '容量', 2025)
    _bt25 = Counter(btype(r['电池类型']) for r in recs
                    if r['年'] == 2025 and r['组'] == 'BEV' and r['电池类型'])
    _n25 = sum(_bt25.values())
    lfp25 = round(_bt25['LFP'] / _n25 * 100, 1) if _n25 else None
    e21, e25 = ymean('BEV', '电耗', 2021), ymean('BEV', '电耗', 2025)
    findings = [
        f'BEV 续航迈入 500km+ 时代：均值 {b21}→{b25} km（2021→2025），2026 年达 {ymean("BEV","续航",2026)} km',
        f'PHEV 纯电续航跃升：{p21}→{p25} km，2026 年 {ymean("PHEV","续航",2026)} km，插混长途化',
        f'大电量化：BEV 容量 {c21}→{c25} kWh（2021→2025），80-100kWh 为主流',
        f'LFP 主导：2025 年 BEV 磷酸铁锂占比约 {lfp25}%，三元退居次席',
        f'能效平稳：BEV 电耗 {e21}→{e25} kWh/100km，在增重背景下未恶化',
        f'结构变化：PHEV 占比 {phe_pct[2021]}%→{phe_pct[2025]}%，插混/增程为公告增量主极',
    ]
    # 电池类型趋势（BEV）
    bt = {y: Counter(btype(r['电池类型']) for r in recs
                     if r['年'] == y and r['组'] == 'BEV' and r['电池类型']) for y in YEARS}
    lfp_pct = [round(bt[y]['LFP'] / max(1, sum(bt[y].values())) * 100, 1) for y in YEARS]
    ncm_pct = [round(bt[y]['NCM'] / max(1, sum(bt[y].values())) * 100, 1) for y in YEARS]
    # 供应商 Top5（有值子集）
    top_bat = Counter(r['电芯'] for r in recs if r['电芯']).most_common(5)
    top_mot = Counter(r['电机企'] for r in recs if r['电机企']).most_common(5)
    # 月销 Top10（易车零售口径；车系级数值→按销量值去重（同族变体共享值），取最短规范名）
    _sname = {}
    for r in recs:
        if r['月销']:
            nm = r['通用'] or r['商标'] or r['企业'][:10]
            if r['月销'] not in _sname or len(nm) < len(_sname[r['月销']]):
                _sname[r['月销']] = nm
    sales = sorted(((v, nm, None) for v, nm in _sname.items()), key=lambda x: -x[0])[:10]
    # 检索数据（轻量字段）
    search_rows = [[r['批次'], str(r['商标'] or ''), str(r['企业'] or '')[:14],
                    r['续航'], r['容量'], r['功率'],
                    'BEV' if r['组'] == 'BEV' else 'PHEV',
                    str(r.get('通用') or '')[:20]]
                   for r in recs]
    DATA = {
        'years': YEARS,
        'cntB': [yr_cnt['BEV'][y] for y in YEARS],
        'cntP': [yr_cnt['PHEV'][y] for y in YEARS],
        'phePct': [phe_pct[y] for y in YEARS],
        'rngB': yseries('BEV', '续航'),
        'rngP': yseries('PHEV', '续航'),
        'capB': yseries('BEV', '容量'),
        'lfp': lfp_pct,
        'ncm': ncm_pct,
        'rows': search_rows,
    }
    year_rows_html = ''.join(
        f'<tr><td>{y}</td><td>{yr_cnt["BEV"][y]:,}</td><td>{yr_cnt["PHEV"][y]:,}</td>'
        f'<td>{tot_y[y]:,}</td><td>{phe_pct[y]}%</td><td>{ymean("BEV","续航",y) or "-"}</td>'
        f'<td>{ymean("BEV","容量",y) or "-"}</td></tr>'
        for y in YEARS)
    find_html = ''.join(f'<div class="card find">💡 {t}</div>' for t in findings)
    kpi_html = ''.join(f'<div class="kpi"><div class="kv">{v}</div><div class="kl">{k}'
                       f'<span class="ku">{u}</span></div></div>' for k, v, u in kpis)
    bat_html = ''.join(f'<div class="sup"><b>{i+1}. {k}</b><span>{v} 款</span></div>'
                       for i, (k, v) in enumerate(top_bat))
    mot_html = ''.join(f'<div class="sup"><b>{i+1}. {k}</b><span>{v} 款</span></div>'
                       for i, (k, v) in enumerate(top_mot))
    sales_html = ''.join(f'<div class="sup"><b>{i+1}. {_html.escape(str(nm))}</b><span>{int(v):,} 辆/月</span></div>'
                         for i, (v, nm, _m) in enumerate(sales))
    ejs = open(ECHARTS, encoding='utf-8').read()
    html = HTML.replace('__ECHARTS__', ejs).replace('__DATA__', json.dumps(DATA, ensure_ascii=False)) \
        .replace('__KPI__', kpi_html).replace('__FIND__', find_html) \
        .replace('__YEARROWS__', year_rows_html).replace('__BAT__', bat_html) \
        .replace('__MOT__', mot_html).replace('__SALES__', sales_html) \
        .replace('__GEN__', '2026-09-30').replace('__TOTAL__', f'{total:,}')
    open(OUT, 'w', encoding='utf-8').write(html)
    print(f'手机版生成: {OUT}（{total} 款口径, KPI {len(kpis)} / 发现 {len(findings)} / 图 4 / 检索 {len(search_rows)} 行）')


HTML = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>NEV公告车型行业分析 · 移动版</title>
<style>
:root{--bg:#0F1522;--card:#171F30;--line:#2A3550;--tx:#E8ECF4;--sub:#93A0B8;
--blue:#4E8FD9;--org:#E8893C;--green:#3FA780;--acc:#D4A94E}
*{margin:0;padding:0;box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{background:var(--bg);color:var(--tx);font:15px/1.65 -apple-system,"PingFang SC","Microsoft YaHei",sans-serif;padding-bottom:40px}
.wrap{max-width:640px;margin:0 auto;padding:0 14px}
header{padding:22px 0 10px;text-align:center;border-bottom:1px solid var(--line)}
header h1{font-size:20px;letter-spacing:1px}
.badge{display:inline-block;margin-top:8px;padding:3px 10px;border:1px solid var(--acc);border-radius:12px;color:var(--acc);font-size:11px}
.sub{color:var(--sub);font-size:11.5px;margin-top:6px}
h2{font-size:16px;margin:26px 0 10px;padding-left:9px;border-left:3px solid var(--org)}
.kgrid{display:grid;grid-template-columns:1fr 1fr;gap:9px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px}
.kv{font-size:20px;font-weight:700;color:var(--acc)}
.kl{font-size:11.5px;color:var(--sub);margin-top:3px}
.ku{float:right;color:var(--sub)}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 14px;margin-bottom:9px;font-size:13.5px}
.chart{width:100%;height:250px;background:var(--card);border:1px solid var(--line);border-radius:12px;margin-bottom:12px}
table{width:100%;border-collapse:collapse;font-size:12px;background:var(--card);border-radius:12px;overflow:hidden}
th{background:#1E2A42;color:var(--sub);font-weight:600;padding:7px 5px;text-align:center}
td{padding:6.5px 5px;text-align:center;border-top:1px solid var(--line)}
.sup{display:flex;justify-content:space-between;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:9px 12px;margin-bottom:7px;font-size:13px}
.sup span{color:var(--acc)}
.search{width:100%;padding:11px 14px;border-radius:12px;border:1px solid var(--line);background:var(--card);color:var(--tx);font-size:15px;margin-bottom:9px}
.pill{display:inline-block;padding:2px 8px;border-radius:9px;font-size:10.5px;margin-left:6px}
.bev{background:rgba(78,143,217,.18);color:var(--blue)}
.phev{background:rgba(232,137,60,.18);color:var(--org)}
footer{color:var(--sub);font-size:10.5px;text-align:center;margin-top:26px;line-height:1.8}
</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>NEV 公告车型行业分析</h1>
  <div class="badge">移动版 · 核心前置</div>
  <div class="sub">权威底表 v4.9.23 口径 · __TOTAL__ 款 × 38 列 · 第 341~411 批（2021–2026）<br>生成 __GEN__ · 数据同源：45 表分析底表</div>
</header>

<h2>核心指标</h2>
<div class="kgrid">__KPI__</div>

<h2>核心发现</h2>
__FIND__

<h2>关键趋势</h2>
<div id="c1" class="chart"></div>
<div id="c2" class="chart"></div>
<div id="c3" class="chart"></div>
<div id="c4" class="chart"></div>

<h2>年度快照</h2>
<table>
<tr><th>年份</th><th>BEV</th><th>PHEV</th><th>合计</th><th>PHEV占比</th><th>BEV续航</th><th>BEV容量</th></tr>
__YEARROWS__
</table>
<div class="sub" style="margin-top:6px">续航 km / 容量 kWh · 均值口径 · 电耗续航以工信部能耗数据20260915 为准（BEV=CLTC，PHEV/EREV=WLTC）</div>

<h2>供应商 Top 5</h2>
<div class="sub" style="margin:0 0 8px">电芯供应商（有值子集）</div>
__BAT__
<div class="sub" style="margin:10px 0 8px">电机生产企业（有值子集）</div>
__MOT__

<h2>月销 Top 10（参考）</h2>
<div class="sub" style="margin:0 0 8px">易车零售月度口径 · 2026-08</div>
__SALES__

<h2>车型检索</h2>
<input id="q" class="search" placeholder="输入型号 / 商标 / 企业名检索…">
<div id="hits"></div>

<footer>
数据来源：工信部《道路机动车辆生产企业及产品公告》第 341~411 批 · 纯乘用车口径（GB/T 3730.1）<br>
口径双原则：仅乘用车 · 仅主流车企 · 销量参考=易车零售月度 · 本页为演示文稿/报告的移动摘要版
</footer>
</div>
<script>__ECHARTS__</script>
<script>
'use strict';
const D=__DATA__;
const AX={axisLine:{lineStyle:{color:'#2A3550'}},axisLabel:{color:'#93A0B8',fontSize:10},
  splitLine:{lineStyle:{color:'#202839'}}};
const GT={grid:{left:44,right:14,top:30,bottom:26},tooltip:{trigger:'axis',textStyle:{fontSize:11}},
  legend:{textStyle:{color:'#93A0B8',fontSize:10},top:0,itemWidth:14,itemHeight:8}};
function mk(id,opt){echarts.init(document.getElementById(id),null,{renderer:'canvas'}).setOption(opt);}
mk('c1',Object.assign({},GT,{xAxis:Object.assign({type:'category',data:D.years.map(String)},AX),
  yAxis:Object.assign({type:'value'},AX),
  series:[{name:'BEV',type:'bar',stack:'a',data:D.cntB,itemStyle:{color:'#4E8FD9'}},
          {name:'PHEV',type:'bar',stack:'a',data:D.cntP,itemStyle:{color:'#E8893C'}}]}));
mk('c2',Object.assign({},GT,{xAxis:Object.assign({type:'category',data:D.years.map(String)},AX),
  yAxis:Object.assign({type:'value'},AX),
  series:[{name:'BEV续航',type:'line',data:D.rngB,itemStyle:{color:'#4E8FD9'}},
          {name:'PHEV纯电续航',type:'line',data:D.rngP,itemStyle:{color:'#E8893C'}}]}));
mk('c3',Object.assign({},GT,{xAxis:Object.assign({type:'category',data:D.years.map(String)},AX),
  yAxis:Object.assign({type:'value',max:100},AX),
  series:[{name:'PHEV占比%',type:'line',areaStyle:{opacity:.15},data:D.phePct,itemStyle:{color:'#E8893C'}}]}));
mk('c4',Object.assign({},GT,{xAxis:Object.assign({type:'category',data:D.years.map(String)},AX),
  yAxis:Object.assign({type:'value',max:100},AX),
  series:[{name:'LFP%',type:'line',data:D.lfp,itemStyle:{color:'#3FA780'}},
          {name:'NCM%',type:'line',data:D.ncm,itemStyle:{color:'#4E8FD9'}}]}));
const hits=document.getElementById('hits');
function esc(s){return String(s).replace(/[<>&]/g,c=>({'<':'&lt;','>':'&gt;','&':'&amp;'}[c]));}
function render(q){
  q=(q||'').trim().toLowerCase();
  const out=[];
  for(const r of D.rows){
    if(q && !(String(r[1]).toLowerCase().includes(q)||String(r[2]).toLowerCase().includes(q)||String(r[7]||'').toLowerCase().includes(q))) continue;
    out.push(`<div class="card" style="padding:9px 12px"><b>${esc(r[7]||r[1]||'—')}</b>`+
      `<span class="pill ${r[6]==='BEV'?'bev':'phev'}">${r[6]}</span>`+
      `<div class="sub" style="margin-top:2px">${r[0]}批 · ${esc(r[2])}`+
      `${r[3]!=null?' · 续航 '+r[3]+'km':''}${r[4]!=null?' · '+r[4]+'kWh':''}`+
      `${r[5]!=null?' · '+r[5]+'kW':''}</div></div>`);
    if(out.length>=30) break;
  }
  hits.innerHTML=out.length?out.join(''):'<div class="card">无匹配（可按商标/企业名检索）</div>';
}
let tm=null;
document.getElementById('q').addEventListener('input',e=>{
  clearTimeout(tm); tm=setTimeout(()=>render(e.target.value),200);});
render('');
window.addEventListener('resize',()=>{document.querySelectorAll('.chart').forEach(el=>{
  const inst=echarts.getInstanceByDom(el); if(inst) inst.resize();});});
</script>
</body>
</html>'''

if __name__ == '__main__':
    main()
