/**
 * dcd_extract.js — 懂车帝参数抽取(Tabbit 用户会话通道)
 *
 * 运行环境:Tabbit LocalAgent 的 nodejs 任务运行时(带用户已登录的浏览器会话)。
 * 调用方式(由 dcd_extract.ps1 包装):
 *   %LOCALAPPDATA%\Tabbit\LocalAgent\bin\tabbit-cli.exe nodejs --task dcd-params \
 *     --request-id <ID> --timeout-ms 180000 < dcd_extract.js
 *   stdin 喂本文件;参数经环境变量 DCD_MODE / DCD_INPUT / DCD_OUT 传入。
 *
 * ⚠ 实测状态说明(务必读):
 *   本文件是适配层,PageProxy 段的能力探测顺序( tabbit.* → page.* → browser.* )
 *   是按常见浏览器自动化形态写的推断,未经本机实测。首次运行必须先进 probe 模式:
 *     $env:DCD_MODE="probe"; ... < dcd_extract.js
 *   probe 会把运行时实际存在的全局对象、方法名、两个关键 URL 的页面结构 dump 到
 *   DCD_OUT,按实际形态修正 PageProxy 后再跑 search/fetch。
 *   若你本机已有跑通的原始抽取脚本(09-20 那批 raw/dcd_refill/ 的产出脚本),
 *   优先用它:dcd_extract.ps1 -JsPath <原脚本>。
 *
 * 模式:
 *   probe  — 环境与页面结构探测(首次校准用)
 *   search — 输入批次请求 JSON → 输出每 kw 的候选车系列表 [{sid, series_name}]
 *   fetch  — 输入定链 JSON(accept 的 sid 列表) → 输出 raw/dcd_refill/dcd_s<sid>.json
 *             + dcd_fuel_form.json(能源类型,用于 BEV/PHEV 动力守卫)
 *
 * 速率纪律:每页 1.5~3s 等待;批 5 车系/次(180s 限额内)。只读公开参数页,不绕任何验证。
 */
'use strict';
const fs = require('fs');

const MODE = process.env.DCD_MODE || 'probe';
const INPUT = process.env.DCD_INPUT || '';
const OUT = process.env.DCD_OUT || '';
const RAW_DIR = process.env.DCD_RAW || '';

// URL 形态(按仓库既有产出反推;probe 模式会验证,不准改这里)
const URL_SEARCH = kw => `https://www.dongchedi.com/auto/search?keyword=${encodeURIComponent(kw)}`;
const URL_PARAMS = sid => `https://www.dongchedi.com/auto/params/${sid}`;
const FIELDS = ['length', 'wheelbase', 'curb_weight', 'cltc_recharge_mileage', 'recharge_mileage',
  'power_consumption', 'battery_capacity', 'battery_type', 'battery_energy_density',
  'total_electric_power', 'total_electric_torque', 'front_electric_max_power',
  'front_electric_max_torque', 'rear_electric_max_power', 'rear_electric_max_torque'];
const sleep = ms => new Promise(r => setTimeout(r, ms));
const jitter = () => 1500 + Math.floor(Math.random() * 1500);

function dump(obj) {
  if (!OUT) { console.log(JSON.stringify(obj, null, 1)); return; }
  fs.writeFileSync(OUT, JSON.stringify(obj, null, 1));
}

/** 环境探测:运行时到底给了什么。 */
function probeEnv() {
  const names = ['tabbit', 'page', 'browser', 'context', 'agent', 'cdp', 'puppeteer', 'playwright'];
  const found = {};
  for (const n of names) {
    if (typeof globalThis[n] !== 'undefined') {
      const v = globalThis[n];
      found[n] = { type: typeof v, keys: safeKeys(v) };
    }
  }
  return { node: process.version, globals: found, env_mode: MODE };
}
function safeKeys(o) {
  try {
    return Object.getOwnPropertyNames(o).filter(k => !k.startsWith('_')).slice(0, 60);
  } catch (e) { return ['<无法枚举>']; }
}

/** 页面代理:按能力探测顺序拿到一个能导航+取 DOM/JSON 的对象。 */
async function getPage() {
  if (globalThis.tabbit && typeof globalThis.tabbit.getPage === 'function') {
    return { p: await globalThis.tabbit.getPage(), kind: 'tabbit.getPage' };
  }
  if (globalThis.page && typeof globalThis.page.goto === 'function') {
    return { p: globalThis.page, kind: 'global.page' };
  }
  if (globalThis.browser && typeof globalThis.browser.newPage === 'function') {
    return { p: await globalThis.browser.newPage(), kind: 'browser.newPage' };
  }
  if (globalThis.tabbit && typeof globalThis.tabbit.open === 'function') {
    return { p: globalThis.tabbit, kind: 'tabbit.open' };
  }
  return null;
}

async function goto(p, url) {
  if (typeof p.goto === 'function') return p.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });
  if (typeof p.navigate === 'function') return p.navigate(url);
  if (typeof p.open === 'function') return p.open(url);
  throw new Error('页面对象无 goto/navigate/open 方法 — 见 probe 输出修正 PageProxy');
}

async function readNextData(p) {
  // 优先直接取 window.__NEXT_DATA__;退化到 HTML 文本里抠
  if (typeof p.evaluate === 'function') {
    const nd = await p.evaluate(() => {
      try { return window.__NEXT_DATA__ || null; } catch (e) { return null; }
    });
    if (nd) return nd;
    const htmlTxt = await p.evaluate(() => document.documentElement.outerHTML);
    return parseNextFromHtml(htmlTxt);
  }
  if (typeof p.content === 'function') {
    const htmlTxt = await p.content();
    return parseNextFromHtml(htmlTxt);
  }
  throw new Error('页面对象无 evaluate/content 方法 — 见 probe 输出修正 PageProxy');
}

function parseNextFromHtml(htmlTxt) {
  const m = /<script id="__NEXT_DATA__" type="application\/json">([\s\S]*?)<\/script>/.exec(htmlTxt || '');
  if (!m) return null;
  try { return JSON.parse(m[1]); } catch (e) { return null; }
}

/** 搜索页 → 候选车系列表。选择器按 DCD 搜索页常见结构推断,probe 模式会验证。 */
async function searchCandidates(p, kw) {
  await goto(p, URL_SEARCH(kw));
  await sleep(jitter());
  const out = [];
  if (typeof p.evaluate === 'function') {
    const list = await p.evaluate(() => {
      const res = [];
      const anchors = Array.from(document.querySelectorAll('a[href*="/auto/series/"]'));
      for (const a of anchors) {
        const m = /\/auto\/series\/(\d+)/.exec(a.getAttribute('href') || '');
        if (m) res.push({ sid: m[1], series_name: (a.textContent || '').trim().slice(0, 60) });
      }
      return res;
    });
    out.push(...list);
  }
  // 去重(同 sid 保留第一个非空名)
  const seen = new Map();
  for (const c of out) {
    if (!c.sid) continue;
    if (!seen.has(c.sid) || (c.series_name && !seen.get(c.sid).series_name)) seen.set(c.sid, c);
  }
  return [...seen.values()].slice(0, 12);
}

/** 参数页 → rawData.cars[] 逐款字段抽取。 */
async function fetchSeries(p, sid) {
  await goto(p, URL_PARAMS(sid));
  await sleep(jitter());
  const nd = await readNextData(p);
  const raw = nd && nd.props && nd.props.pageProps && nd.props.pageProps.rawData;
  if (!raw) return { sid, error: '参数页未取到 __NEXT_DATA__.props.pageProps.rawData' };
  const cars = [];
  for (const c of (raw.car_info || raw.cars || [])) {
    const info = c.info || c.vals || {};
    const vals = {};
    for (const k of FIELDS) {
      const cell = info[k];
      vals[k] = cell && typeof cell === 'object' ? (cell.value != null ? cell.value : '') : (cell != null ? cell : '');
    }
    cars.push({ id: String(c.id || c.car_id || ''), name: String(c.name || c.car_name || ''), year: String(c.year || ''), vals });
  }
  // 能源类型(燃料形式):用于 BEV/PHEV 动力守卫
  const fuel = [];
  for (const c of (raw.car_info || raw.cars || [])) {
    const info = c.info || c.vals || {};
    const ff = (info.fuel_form && (info.fuel_form.value || info.fuel_form)) || '';
    if (c.id) fuel.push({ id: String(c.id), name: String(c.name || c.car_name || ''), ff: String(ff || '') });
  }
  return { sid, series: String(raw.series_name || raw.name || ''), cars, fuel };
}

async function main() {
  const env = probeEnv();
  if (MODE === 'probe') {
    const pg = await getPage();
    let pageInfo = { page_proxy: pg ? pg.kind : '未找到可用页面对象' };
    if (pg) {
      try {
        await goto(pg.p, URL_PARAMS('4499'));
        await sleep(2000);
        const nd = await readNextData(pg.p);
        pageInfo.params_4499_nextdata = nd ? 'OK(取到 __NEXT_DATA__)' : 'FAIL(无 __NEXT_DATA__)';
        pageInfo.rawData_keys = nd && nd.props && nd.props.pageProps && nd.props.pageProps.rawData
          ? safeKeys(nd.props.pageProps.rawData) : null;
      } catch (e) { pageInfo.params_4499_error = String(e && e.message || e); }
      try {
        await goto(pg.p, URL_SEARCH('宏光MINIEV'));
        await sleep(2000);
        const cand = await searchCandidates(pg.p, '宏光MINIEV');
        pageInfo.search_test = cand.slice(0, 5);
      } catch (e) { pageInfo.search_error = String(e && e.message || e); }
    }
    dump({ probe: env, page: pageInfo, at: new Date().toISOString() });
    console.log('probe 完成 →', OUT || '(stdout)');
    return;
  }

  const pg = await getPage();
  if (!pg) {
    dump({ error: '未找到可用页面对象,先跑 probe 模式并按输出修正 PageProxy', probe: env });
    process.exitCode = 2;
    return;
  }
  const input = JSON.parse(fs.readFileSync(INPUT, 'utf8'));

  if (MODE === 'search') {
    const candidates = {};
    for (const s of input.series || []) {
      try {
        candidates[s.qkey] = await searchCandidates(pg.p, s.kw);
        if ((!candidates[s.qkey] || !candidates[s.qkey].length) && s.kw_alt) {
          candidates[s.qkey + '#alt'] = await searchCandidates(pg.p, s.kw_alt);
        }
      } catch (e) {
        candidates[s.qkey] = [];
        candidates[s.qkey + '#error'] = String(e && e.message || e);
      }
      await sleep(jitter());
    }
    const doc = { batch_id: input.batch_id, mode: 'search', at: new Date().toISOString(), candidates };
    fs.writeFileSync(OUT, JSON.stringify(doc, null, 1));
    console.log(`search 完成:${Object.keys(candidates).length} kw →`, OUT);
    return;
  }

  if (MODE === 'fetch') {
    const sids = (input.series || []).filter(s => s.status === 'accept' && s.sid).map(s => String(s.sid));
    const uniq = [...new Set(sids)];
    const fuelAll = {};
    const results = [];
    for (const sid of uniq) {
      const fp = `${RAW_DIR}/dcd_s${sid}.json`;
      if (fs.existsSync(fp)) { results.push({ sid, skipped: '文件已存在' }); continue; }
      try {
        const r = await fetchSeries(pg.p, sid);
        if (r.error) { results.push({ sid, error: r.error }); continue; }
        fs.writeFileSync(fp, JSON.stringify({ series: r.series, cars: r.cars }, null, 1));
        for (const f of r.fuel) {
          (fuelAll[sid] = fuelAll[sid] || []).push({ id: f.id, name: f.name, ff: f.ff });
        }
        results.push({ sid, cars: r.cars.length });
      } catch (e) {
        results.push({ sid, error: String(e && e.message || e) });
      }
      await sleep(jitter());
    }
    // 合并进 fuel_form(保留已有条目)
    const ffp = `${RAW_DIR}/dcd_fuel_form.json`;
    let ffOld = {};
    if (fs.existsSync(ffp)) { try { ffOld = JSON.parse(fs.readFileSync(ffp, 'utf8')); } catch (e) {} }
    for (const [sid, lst] of Object.entries(fuelAll)) ffOld[sid] = lst;
    fs.writeFileSync(ffp, JSON.stringify(ffOld, null, 1));
    fs.writeFileSync(OUT, JSON.stringify({ batch_id: input.batch_id, mode: 'fetch', at: new Date().toISOString(), results }, null, 1));
    console.log(`fetch 完成:${uniq.length} sid →`, RAW_DIR);
    return;
  }

  dump({ error: `未知 DCD_MODE=${MODE}(应为 probe/search/fetch)` });
  process.exitCode = 2;
}

main().catch(e => {
  dump({ fatal: String(e && e.stack || e), mode: MODE });
  process.exitCode = 1;
});
