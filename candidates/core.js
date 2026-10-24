/* 三款候选看板共用的图表核心：数据 → ECharts option
 * 依赖：window.DATA（底表 JSON）、window.THEME（各款自定义配色）
 * 全局筛选：state.power ∈ {ALL,BEV,PHEV}、state.years=[from,to]、state.segs=Set|null
 */
(function () {
  const D = window.DATA;
  const T = window.THEME;
  const state = { power: 'ALL', years: [2021, 2026], segs: null, metric: 'range', segMetric: 'range' };
  const charts = {};
  const YEARS = D.yearly.years;

  const C = {
    bev: T.bev, phev: T.phev, lfp: T.lfp || T.bev, ncm: T.ncm || T.phev, other: T.other || '#9aa0a6',
    text: T.text, text2: T.text2, grid: T.grid, bg: T.bg, palette: T.palette,
  };

  // ---------- 工具 ----------
  const yrMask = () => YEARS.map(y => y >= state.years[0] && y <= state.years[1]);
  const fy = arr => arr.filter((_, i) => yrMask()[i]);
  const years = () => fy(YEARS).map(String);
  const showBEV = () => state.power !== 'PHEV';
  const showPHEV = () => state.power !== 'BEV';
  const r1 = v => v == null ? '-' : (Math.round(v * 10) / 10).toLocaleString();
  const r0 = v => v == null ? '-' : Math.round(v).toLocaleString();

  const base = (extra) => Object.assign({
    color: C.palette,
    textStyle: { color: C.text, fontFamily: T.font },
    grid: { left: 8, right: 12, top: 44, bottom: 8, containLabel: true },
    tooltip: { trigger: 'axis', backgroundColor: T.tipBg, borderColor: C.grid, textStyle: { color: T.tipText, fontSize: 12 }, confine: true },
    legend: { top: 4, textStyle: { color: C.text2, fontSize: 11 }, itemWidth: 14, itemHeight: 8, icon: 'roundRect' },
  }, extra);
  const cat = (data, extra) => Object.assign({ type: 'category', data, axisLine: { lineStyle: { color: C.grid } }, axisTick: { show: false }, axisLabel: { color: C.text2, fontSize: 11 } }, extra);
  const val = (name, extra) => Object.assign({ type: 'value', name, nameTextStyle: { color: C.text2, fontSize: 10, align: 'left' }, splitLine: { lineStyle: { color: C.grid, type: 'dashed' } }, axisLabel: { color: C.text2, fontSize: 11 } }, extra);
  const line = (name, data, color, extra) => Object.assign({ name, type: 'line', data, smooth: T.smooth, symbol: 'circle', symbolSize: 6, lineStyle: { width: 2.5, color }, itemStyle: { color }, emphasis: { focus: 'series' } }, extra);
  const bar = (name, data, color, extra) => Object.assign({ name, type: 'bar', data, itemStyle: { color, borderRadius: T.barRadius }, barMaxWidth: 36, emphasis: { focus: 'series' } }, extra);
  const area = (color) => ({ areaStyle: { opacity: .12, color } });
  const seriesByPower = (bevArr, phevArr, unitFmt) => {
    const s = [];
    if (showBEV()) s.push(line('BEV', fy(bevArr), C.bev, unitFmt ? area(C.bev) : {}));
    if (showPHEV()) s.push(line('PHEV/EREV', fy(phevArr), C.phev, unitFmt ? area(C.phev) : {}));
    return s;
  };

  // ---------- 指标字典（指标探索器 / 细分市场） ----------
  const METRICS = {
    range: { label: '平均续航', unit: 'km', bev: D.annual.cols['BEV续航'], phev: D.annual.cols['PHEV续航'] },
    cap: { label: '平均电池容量', unit: 'kWh', bev: D.annual.cols['BEV容量'], phev: D.annual.cols['PHEV容量'] },
    power: { label: '平均电机功率', unit: 'kW', bev: D.annual.cols['BEV功率'], phev: D.annual.cols['PHEV功率'] },
    mass: { label: '平均整备质量', unit: 'kg', bev: D.annual.cols['BEV质量'], phev: D.annual.cols['PHEV质量'] },
    vdens: { label: '整车能量密度', unit: 'Wh/kg', bev: D.annual.cols['BEV整车密度'], phev: D.annual.cols['PHEV整车密度'] },
    cdens: { label: '电池能量密度', unit: 'Wh/kg', bev: D.cell_density.cols['BEV均值(Wh/kg)'], phev: D.cell_density.cols['PHEV均值(Wh/kg)'] },
    pw: { label: '功率质量比', unit: 'kW/kg×1000', bev: D.pw_ratio.cols['BEV(kW/kg×1000)'], phev: D.pw_ratio.cols['PHEV(kW/kg×1000)'] },
    count: { label: '公告车型数', unit: '款', bev: D.yearly.cols['BEV(款)'], phev: D.yearly.cols['PHEV(款)'] },
  };
  const SEG_METRICS = { n: ['车型数', '款'], range: ['平均续航', 'km'], cap: ['平均容量', 'kWh'], power: ['平均功率', 'kW'], mass: ['平均整备质量', 'kg'] };

  // ---------- 图表 option 构造器 ----------
  const O = {};

  O.yearlyCount = () => {
    const s = [];
    if (showBEV()) s.push(bar('BEV', fy(D.yearly.cols['BEV(款)']), C.bev, { stack: 'n' }));
    if (showPHEV()) s.push(bar('PHEV/EREV', fy(D.yearly.cols['PHEV(款)']), C.phev, { stack: 'n' }));
    if (state.power === 'ALL') s.push(line('PHEV占比%', fy(D.yearly.cols['PHEV占比%']), T.accent, { yAxisIndex: 1, areaStyle: null }));
    return base({ xAxis: cat(years()), yAxis: [val('款'), val('%', { splitLine: { show: false }, max: 100 })], series: s });
  };

  O.rangeTrend = () => {
    const s = [];
    const build = (src, name, color) => {
      s.push({ name: name + ' P25–P75', type: 'line', data: fy(src.cols['P25(km)']), lineStyle: { opacity: 0 }, stack: name, symbol: 'none', silent: true, tooltip: { show: false } });
      s.push({ name: name + ' P25–P75', type: 'line', data: fy(src.cols['P75(km)']).map((v, i) => v - fy(src.cols['P25(km)'])[i]), lineStyle: { opacity: 0 }, stack: name, symbol: 'none', areaStyle: { color, opacity: .12 }, tooltip: { show: false } });
      s.push(line(name + ' 均值', fy(src.cols['均值(km)']), color));
      s.push(line(name + ' 中位数', fy(src.cols['中位数(km)']), color, { lineStyle: { type: 'dashed', width: 1.5, color }, symbol: 'none' }));
    };
    if (showBEV()) build(D.range_bev, 'BEV', C.bev);
    if (showPHEV()) build(D.range_phev, 'PHEV/EREV', C.phev);
    return base({ xAxis: cat(years()), yAxis: val('km'), series: s });
  };

  O.capacityTrend = () => base({ xAxis: cat(years()), yAxis: val('kWh'), series: seriesByPower(D.capacity.cols['BEV均值(kWh)'], D.capacity.cols['PHEV均值(kWh)'], true) });
  O.cellDensity = () => base({ xAxis: cat(years()), yAxis: val('Wh/kg', { min: 100 }), series: seriesByPower(D.cell_density.cols['BEV均值(Wh/kg)'], D.cell_density.cols['PHEV均值(Wh/kg)']) });
  O.powerTrend = () => base({ xAxis: cat(years()), yAxis: val('kW'), series: seriesByPower(D.power.cols['BEV均值(kW)'], D.power.cols['PHEV均值(kW)'], true) });
  O.massTrend = () => base({ xAxis: cat(years()), yAxis: val('kg', { min: 1400 }), series: seriesByPower(D.mass.cols['BEV均值(kg)'], D.mass.cols['PHEV均值(kg)']) });
  O.pwRatio = () => base({ xAxis: cat(years()), yAxis: val('kW/kg×1000'), series: seriesByPower(D.pw_ratio.cols['BEV(kW/kg×1000)'], D.pw_ratio.cols['PHEV(kW/kg×1000)']) });
  O.vehDensity = () => base({ xAxis: cat(years()), yAxis: val('Wh/kg'), series: seriesByPower(D.veh_density.cols['BEV(Wh/kg)'], D.veh_density.cols['PHEV(Wh/kg)'], true) });

  O.batteryType = () => {
    const src = state.power === 'PHEV' ? D.bt_phev : D.bt_bev;
    const mk = (k, name, color) => ({ name, type: 'line', stack: 'p', data: fy(src.cols[k]), smooth: T.smooth, symbol: 'none', lineStyle: { width: 1, color }, itemStyle: { color }, areaStyle: { color, opacity: .75 }, emphasis: { focus: 'series' } });
    return base({
      title: { text: state.power === 'PHEV' ? 'PHEV/EREV' : (state.power === 'ALL' ? 'BEV（切换动力类型查看 PHEV）' : 'BEV'), left: 'right', top: 2, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } },
      tooltip: { trigger: 'axis', valueFormatter: v => r1(v) + '%' },
      xAxis: cat(years()), yAxis: val('%', { max: 100 }),
      series: [mk('LFP占比%', '磷酸铁锂 LFP', C.lfp), mk('NCM占比%', '三元锂 NCM', C.ncm), mk('其他占比%', '其他', C.other)],
    });
  };

  const distOption = (src, unit) => {
    const ys = fy(src.years.map(String));
    const tot = src.years.map((_, i) => src.buckets.reduce((a, b) => a + (src.data[b][i] || 0), 0));
    const s = src.buckets.map((b, bi) => ({
      name: b + unit, type: 'bar', stack: 'd', barMaxWidth: 48,
      data: fy(src.data[b].map((v, i) => tot[i] ? +(v / tot[i] * 100).toFixed(1) : 0)),
      itemStyle: { color: T.seq[bi % T.seq.length] }, emphasis: { focus: 'series' },
    }));
    return base({ tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => v + '%' }, legend: { top: 4, type: 'scroll', textStyle: { color: C.text2, fontSize: 10 } }, xAxis: cat(ys), yAxis: val('占比%', { max: 100 }), series: s });
  };
  O.rangeDist = () => distOption(state.power === 'PHEV' ? D.dist_range_phev : D.dist_range_bev, 'km');
  O.capDist = () => distOption(state.power === 'PHEV' ? D.dist_cap_phev : D.dist_cap_bev, 'kWh');

  O.efficiency = () => base({
    xAxis: cat(years()), yAxis: val('kWh/100km', { min: 10 }),
    series: [line('BEV 均值', fy(D.efficiency.cols['BEV均值']), C.bev, area(C.bev)), line('LFP 均值', fy(D.efficiency.cols['LFP均值']), C.lfp, { lineStyle: { type: 'dashed', width: 1.8, color: C.lfp } }), line('NCM 均值', fy(D.efficiency.cols['NCM均值']), C.ncm, { lineStyle: { type: 'dashed', width: 1.8, color: C.ncm } })],
  });
  O.phevFuel = () => base({
    xAxis: cat(years()), yAxis: [val('L/100km'), val('km', { splitLine: { show: false } })],
    series: [line('综合油耗', fy(D.phev_fuel.cols['综合油耗(L/100km)']), C.phev), line('B状态油耗', fy(D.phev_fuel.cols['B状态油耗(L/100km)']), T.accent, { lineStyle: { type: 'dashed', width: 1.8, color: T.accent } }), bar('纯电续航', fy(D.phev_fuel.cols['纯电续航(km)']), C.phev, { yAxisIndex: 1, itemStyle: { color: C.phev, opacity: .25, borderRadius: T.barRadius } })],
  });

  O.metricExplorer = () => {
    const m = METRICS[state.metric];
    return base({ xAxis: cat(years()), yAxis: val(m.unit, { scale: true }), series: seriesByPower(m.bev, m.phev, true) });
  };
  O.indexed = () => {
    // 2021=100 的指数化多指标折线（按动力类型）
    const keys = ['range', 'cap', 'power', 'mass', 'vdens'];
    const p = state.power === 'PHEV' ? 'phev' : 'bev';
    const s = keys.map((k, i) => { const arr = METRICS[k][p]; const b = arr[0]; return line(METRICS[k].label, fy(arr.map(v => +(v / b * 100).toFixed(1))), C.palette[i % C.palette.length], { symbolSize: 4 }); });
    return base({ title: { text: (p === 'bev' ? 'BEV' : 'PHEV/EREV') + ' · 2021=100', right: 0, top: 2, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, tooltip: { trigger: 'axis', valueFormatter: v => v }, xAxis: cat(years()), yAxis: val('指数', { scale: true }), series: s });
  };

  O.hhi = () => base({ xAxis: cat(years()), yAxis: [val('电池HHI'), val('电机HHI', { splitLine: { show: false } })], series: [line('电池(电芯)供应商 HHI', fy(D.hhi.cols['电池供应商HHI']), C.palette[0], area(C.palette[0])), line('电机供应商 HHI', fy(D.hhi.cols['电机供应商HHI']), C.palette[1], { yAxisIndex: 1 })] });
  const crOpt = src => base({ xAxis: cat(years()), yAxis: val('%', { max: 100 }), series: ['CR3%', 'CR5%', 'CR10%'].map((k, i) => line(k.replace('%', ''), fy(src.cols[k]), C.palette[i], i === 2 ? area(C.palette[2]) : {})) });
  O.crBatt = () => crOpt(D.cr_batt);
  O.crMotor = () => crOpt(D.cr_motor);
  const rankOpt = (arr, color) => {
    const a = arr.slice().reverse();
    return base({ grid: { left: 8, right: 40, top: 8, bottom: 8, containLabel: true }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, legend: { show: false }, xAxis: val('', { splitLine: { show: false } }), yAxis: cat(a.map(x => x.name.length > 12 ? x.name.slice(0, 12) + '…' : x.name), { axisLabel: { color: C.text, fontSize: 11 } }), series: [bar('配套车型数', a.map(x => x.n), color, { label: { show: true, position: 'right', color: C.text2, fontSize: 11 } })] });
  };
  O.rankBatt = () => rankOpt(D.rank_batt, C.palette[0]);
  O.rankMotor = () => rankOpt(D.rank_motor, C.palette[1]);
  O.combo = () => rankOpt(D.combo, T.accent);
  const heat = (m, color) => {
    const data = []; let max = 0;
    m.rows.forEach((r, i) => m.cols.forEach((c, j) => { data.push([j, i, m.data[i][j]]); max = Math.max(max, m.data[i][j]); }));
    return base({ grid: { left: 8, right: 8, top: 8, bottom: 40, containLabel: true }, tooltip: { position: 'top', formatter: p => `${m.rows[p.value[1]]} × ${m.cols[p.value[0]]}<br/>${p.value[2]} 款` }, xAxis: cat(m.cols.map(x => x.length > 8 ? x.slice(0, 8) + '…' : x), { splitArea: { show: true }, axisLabel: { interval: 0, rotate: 20, fontSize: 10, color: C.text2 } }), yAxis: cat(m.rows, { splitArea: { show: true } }), visualMap: { min: 0, max, calculable: false, orient: 'horizontal', left: 'center', bottom: 0, itemHeight: 90, textStyle: { color: C.text2, fontSize: 10 }, inRange: { color: [T.heatLow, color] } }, series: [{ type: 'heatmap', data, label: { show: true, fontSize: 10, color: C.text }, itemStyle: { borderColor: C.bg, borderWidth: 2 } }] });
  };
  O.brandBatt = () => heat(D.brand_batt, C.palette[0]);
  O.brandMotor = () => heat(D.brand_motor, C.palette[1]);

  const segRows = () => {
    const src = state.power === 'PHEV' ? D.seg_phev : D.seg_bev;
    return src.filter(r => !state.segs || state.segs.has(r.seg));
  };
  O.segBar = () => {
    const rows = segRows().slice().sort((a, b) => (b[state.segMetric] || 0) - (a[state.segMetric] || 0)).slice(0, 14);
    const [label, unit] = SEG_METRICS[state.segMetric];
    const color = state.power === 'PHEV' ? C.phev : C.bev;
    return base({ grid: { left: 8, right: 12, top: 30, bottom: 8, containLabel: true }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' }, valueFormatter: v => r1(v) + ' ' + unit }, legend: { show: false }, title: { text: `${state.power === 'PHEV' ? 'PHEV/EREV' : 'BEV'} · ${label}（${unit}）`, left: 0, top: 0, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, xAxis: cat(rows.map(r => r.seg), { axisLabel: { interval: 0, rotate: 35, fontSize: 10, color: C.text2 } }), yAxis: val(unit, { scale: true }), series: [bar(label, rows.map(r => r[state.segMetric]), color, { label: { show: true, position: 'top', fontSize: 9, color: C.text2, formatter: p => r0(p.value) } })] });
  };
  O.segScatter = () => {
    const rows = segRows().filter(r => r.range && r.cap);
    const maxN = Math.max(...rows.map(r => r.n));
    const color = state.power === 'PHEV' ? C.phev : C.bev;
    return base({ grid: { left: 8, right: 20, top: 30, bottom: 8, containLabel: true }, legend: { show: false }, tooltip: { trigger: 'item', formatter: p => `<b>${p.data[3]}</b><br/>续航 ${r0(p.data[0])} km · 容量 ${r1(p.data[1])} kWh<br/>功率 ${r0(p.data[4])} kW · 整备 ${r0(p.data[5])} kg<br/>${p.data[2]} 款` }, title: { text: '气泡大小 = 车型数', left: 0, top: 0, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, xAxis: val('续航 km', { scale: true, splitLine: { lineStyle: { color: C.grid, type: 'dashed' } } }), yAxis: val('容量 kWh', { scale: true }), series: [{ type: 'scatter', data: rows.map(r => [r.range, r.cap, r.n, r.seg, r.power, r.mass]), symbolSize: d => 8 + Math.sqrt(d[2] / maxN) * 46, itemStyle: { color, opacity: .55, borderColor: color, borderWidth: 1 }, label: { show: true, formatter: p => p.data[3], position: 'top', fontSize: 10, color: C.text }, labelLayout: { hideOverlap: true } }] });
  };
  O.segCmp = () => {
    const src = state.segMetric === 'cap' ? D.seg_cap_cmp : D.seg_range_cmp;
    const unit = state.segMetric === 'cap' ? 'kWh' : 'km';
    const rows = src.filter(r => r.y2023 != null && r.y2025 != null && (!state.segs || state.segs.has(r.seg)));
    return base({ tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, title: { text: `BEV ${unit === 'km' ? '续航' : '容量'} · 2023 vs 2025`, left: 0, top: 0, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, legend: { top: 4, right: 0 }, xAxis: cat(rows.map(r => r.seg), { axisLabel: { interval: 0, rotate: 30, fontSize: 10, color: C.text2 } }), yAxis: [val(unit, { scale: true }), val('变化%', { splitLine: { show: false } })], series: [bar('2023', rows.map(r => r.y2023), T.heatLow2 || C.grid), bar('2025', rows.map(r => r.y2025), C.bev), line('变化%', rows.map(r => r.chg), T.accent, { yAxisIndex: 1 })] });
  };
  O.segMass = () => base({ grid: { left: 8, right: 40, top: 8, bottom: 8, containLabel: true }, legend: { show: false }, tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } }, xAxis: val('kg', { splitLine: { show: false }, min: 1500 }), yAxis: cat(D.seg_mass.slice().reverse().map(r => r.seg)), series: [bar('平均整备质量', D.seg_mass.slice().reverse().map(r => r.mass), C.palette[3], { label: { show: true, position: 'right', fontSize: 10, color: C.text2, formatter: p => r0(p.value) } })] });

  O.corr = () => {
    const m = state.power === 'PHEV' ? D.corr_phev : D.corr_bev;
    const data = []; m.labels.forEach((_, i) => m.labels.forEach((_, j) => data.push([j, i, m.data[i][j]])));
    return base({ grid: { left: 8, right: 8, top: 8, bottom: 40, containLabel: true }, tooltip: { position: 'top', formatter: p => `${m.labels[p.value[1]]} × ${m.labels[p.value[0]]}：r = ${p.value[2]}` }, title: { text: state.power === 'PHEV' ? 'PHEV/EREV · Pearson' : 'BEV · Pearson', right: 0, top: 0, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, xAxis: cat(m.labels, { axisLabel: { interval: 0, fontSize: 10, color: C.text2 } }), yAxis: cat(m.labels), visualMap: { min: 0, max: 1, orient: 'horizontal', left: 'center', bottom: 0, itemHeight: 90, textStyle: { color: C.text2, fontSize: 10 }, inRange: { color: [T.heatLow, C.bev] } }, series: [{ type: 'heatmap', data, label: { show: true, fontSize: 10, color: C.text }, itemStyle: { borderColor: C.bg, borderWidth: 2 } }] });
  };
  O.yoy = () => {
    const src = state.power === 'PHEV' ? D.yoy_phev : D.yoy_bev;
    const ys = fy(src.years.map(String));
    return base({ tooltip: { trigger: 'axis', valueFormatter: v => v == null ? '-' : v + '%' }, title: { text: state.power === 'PHEV' ? 'PHEV/EREV 同比%' : 'BEV 同比%', right: 0, top: 2, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, legend: { top: 4, type: 'scroll', textStyle: { color: C.text2, fontSize: 10 } }, xAxis: cat(ys), yAxis: val('%'), series: src.rows.map((r, i) => line(r.name, fy(r.vals), C.palette[i % C.palette.length], { symbolSize: 4, lineStyle: { width: r.name.includes('数量') ? 1.5 : 2.2, type: r.name.includes('数量') ? 'dashed' : 'solid', color: C.palette[i % C.palette.length] } })) });
  };
  O.quadrant = () => base({ tooltip: { trigger: 'item', formatter: p => `${p.name}<br/>${r0(p.value)} 款 · ${p.percent}%` }, legend: { bottom: 0, top: null, textStyle: { color: C.text2, fontSize: 11 } }, title: { text: D.quadrant.note, left: 0, top: 0, textStyle: { fontSize: 11, color: C.text2, fontWeight: 400 } }, series: [{ type: 'pie', radius: ['42%', '68%'], center: ['50%', '50%'], data: D.quadrant.items.map(i => ({ name: i.name, value: i.n })), label: { color: C.text, fontSize: 11, formatter: '{b}\n{d}%' }, itemStyle: { borderColor: C.bg, borderWidth: 2, borderRadius: 4 } }] });
  O.kpiRadar = () => {
    const ind = ['续航', '电池容量', '电机功率', '整备质量', '能量密度'];
    const b = D.kpi_bev, p = D.kpi_phev;
    const mx = ind.map((_, i) => Math.max(b[i].mean, p[i].mean) * 1.1);
    return base({ tooltip: { trigger: 'item' }, legend: { bottom: 0, top: null, textStyle: { color: C.text2 } }, radar: { indicator: ind.map((n, i) => ({ name: n, max: mx[i] })), radius: '62%', axisName: { color: C.text2, fontSize: 11 }, splitLine: { lineStyle: { color: C.grid } }, splitArea: { show: false }, axisLine: { lineStyle: { color: C.grid } } }, series: [{ type: 'radar', data: [{ name: 'BEV 均值', value: b.map(x => x.mean), itemStyle: { color: C.bev }, areaStyle: { opacity: .15 } }, { name: 'PHEV/EREV 均值', value: p.map(x => x.mean), itemStyle: { color: C.phev }, areaStyle: { opacity: .15 } }] }] });
  };

  // ---------- KPI ----------
  function kpis() {
    const i = yrMask().lastIndexOf(true);
    const bevN = D.yearly.cols['BEV(款)'], phevN = D.yearly.cols['PHEV(款)'];
    const y = YEARS[i];
    const pick = (b, p) => state.power === 'BEV' ? b : state.power === 'PHEV' ? p : null;
    const total = fy(bevN).reduce((a, v) => a + v, 0) * (showBEV() ? 1 : 0) + fy(phevN).reduce((a, v) => a + v, 0) * (showPHEV() ? 1 : 0);
    return [
      { label: `公告车型总量（${state.years[0]}–${state.years[1]}）`, value: r0(total), unit: '款', sub: state.power === 'ALL' ? `BEV ${r0(fy(bevN).reduce((a, v) => a + v, 0))} / PHEV ${r0(fy(phevN).reduce((a, v) => a + v, 0))}` : (state.power === 'BEV' ? 'BEV 纯电' : 'PHEV/EREV 插混·增程') },
      { label: `${y} 年 PHEV 占比`, value: r1(D.yearly.cols['PHEV占比%'][i]), unit: '%', sub: '插混/增程持续放量' },
      { label: `${y} 年 ${state.power === 'PHEV' ? 'PHEV纯电' : 'BEV'}续航均值`, value: r0(pick(METRICS.range.bev[i], METRICS.range.phev[i]) ?? METRICS.range.bev[i]), unit: 'km', sub: `较 2021 年 ${(v=>v>=0?'+':'')(((pick(METRICS.range.bev[i], METRICS.range.phev[i]) ?? METRICS.range.bev[i]) / (pick(METRICS.range.bev[0], METRICS.range.phev[0]) ?? METRICS.range.bev[0]) - 1))}${r1(((pick(METRICS.range.bev[i], METRICS.range.phev[i]) ?? METRICS.range.bev[i]) / (pick(METRICS.range.bev[0], METRICS.range.phev[0]) ?? METRICS.range.bev[0]) - 1) * 100)}%` },
      { label: `${y} 年 ${state.power === 'PHEV' ? 'PHEV' : 'BEV'}电池容量均值`, value: r1(pick(METRICS.cap.bev[i], METRICS.cap.phev[i]) ?? METRICS.cap.bev[i]), unit: 'kWh', sub: '大电量化持续' },
      { label: `${y} 年 ${state.power === 'PHEV' ? 'PHEV' : 'BEV'} LFP 占比`, value: r1((state.power === 'PHEV' ? D.bt_phev : D.bt_bev).cols['LFP占比%'][i]), unit: '%', sub: '磷酸铁锂主导' },
      { label: `${y} 年电池供应商 HHI`, value: r0(D.hhi.cols['电池供应商HHI'][i]), unit: '', sub: `CR3 ${r1(D.cr_batt.cols['CR3%'][i])}% · 集中度${D.hhi.cols['电池供应商HHI'][i] < 1500 ? '低' : D.hhi.cols['电池供应商HHI'][i] < 2500 ? '中' : '高'}` },
    ];
  }

  // ---------- 渲染 ----------
  function render(id, key) {
    const el = document.getElementById(id); if (!el) return;
    if (!charts[id]) charts[id] = echarts.init(el, null, { renderer: 'canvas' });
    charts[id].setOption(O[key](), true);
  }
  const REG = [];
  function register(id, key) { REG.push([id, key]); }
  function renderAll() {
    REG.forEach(([id, key]) => { try { render(id, key); } catch (e) { console.error(key, e); } });
    document.dispatchEvent(new CustomEvent('kpis', { detail: kpis() }));
  }
  let rt; window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => Object.values(charts).forEach(c => c.resize()), 120); });
  document.addEventListener('visibilitychange', () => Object.values(charts).forEach(c => c.resize()));

  window.CORE = { state, register, renderAll, render, resizeAll: () => Object.values(charts).forEach(c => c.resize()), METRICS, SEG_METRICS, YEARS, kpis, D,
    segList: () => Array.from(new Set(D.seg_bev.concat(D.seg_phev).map(r => r.seg))) };
})();
