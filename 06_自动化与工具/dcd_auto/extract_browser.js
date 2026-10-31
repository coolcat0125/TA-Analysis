/**
 * extract_browser.js — 懂车帝官网车型参数直抓工具(零依赖 · 抽离 Tabbit)
 *
 * 原理:在懂车帝官网任一页面的控制台里运行,用【同源 fetch】抓三个页面:
 *   ① 搜索页  /auto/search?keyword=<kw>           → 候选车系(sid + 车系名)
 *   ② 车系页  /auto/series/<sid>                   → 该车系的车型清单(carId + 车型名)
 *   ③ 参数页  /auto/params-carIds-x-<id>-<id>...   → __NEXT_DATA__ 的 rawData.car_info[]
 * 抽 15 个参数字段 + 能源类型,累积后每 15 个车系打一个 bundle 下载。
 * 同源 fetch 不受 CORS 限制,也不需要任何第三方工具——抓取发生在真实浏览器会话内,
 * 但工具本身是完全独立的一段 JS,不绑定 Tabbit 或任何 agent 运行时。
 *
 * ⚠ URL 形态说明(2026-09-22):
 *   参数页的正确路径是 params-carIds-x-<车型id>(车型级 id,非车系 sid),由用户提供样本确认。
 *   车型 id 需从车系页取——所以是三步而非两步。多车型批量形态假定为
 *   params-carIds-x-<id1>-<id2>(连字符分隔);若实际是其他分隔符,probe 会抓到时替换。
 *   脚本对 /auto/params/<sid> 旧形态保留 fallback。
 *
 * 用法:
 *   1. 浏览器打开 https://www.dongchedi.com/ (停在该域名下,不要关)
 *   2. F12 → Console → 粘贴本文件全文 → 回车
 *   3. 首次运行弹出文件选择框 → 选 dcd_auto/state/targets.json
 *   4. 自动开始;每 15 个车系自动下载一个 bundle 到默认下载目录
 *   5. 中断后重跑会跳过已完成的(localStorage 断点)
 *
 * ⚠ 首次使用务必先跑探查模式:把下面第一行改成 const DCD_PROBE = true; 再粘贴。
 *    它会抓 搜索页+车系页+参数页 各一个,把真实结构 dump 成 probe.json。
 *    也可直接把参数页「另存为 dcd_probe_page.html」放到 dcd_auto/state/ 给维护者核对。
 *
 * 纪律:只读公开页面;每页间隔 1.5~3s;不绕任何验证码/登录墙。
 */
const DCD_PROBE = false;              // ← 首次改 true 跑探查
const BUNDLE_EVERY = 15;
const LS_KEY = 'dcd_extract_progress_v2';   // v2:三步流程,与旧版断点不兼容

const FIELDS = ['length', 'wheelbase', 'curb_weight', 'cltc_recharge_mileage', 'recharge_mileage',
  'power_consumption', 'battery_capacity', 'battery_type', 'battery_energy_density',
  'total_electric_power', 'total_electric_torque', 'front_electric_max_power',
  'front_electric_max_torque', 'rear_electric_max_power', 'rear_electric_max_torque'];

const sleep = ms => new Promise(r => setTimeout(r, ms));
const jitter = () => 1500 + Math.floor(Math.random() * 1500);
const log = (...a) => console.log('[DCD]', ...a);

/** 同源抓取 HTML 文本 */
async function getHtml(url) {
  const r = await fetch(url, { credentials: 'include', headers: { 'Accept': 'text/html' } });
  if (!r.ok) throw new Error(`HTTP ${r.status} ${url}`);
  return await r.text();
}

/** 从 HTML 里抠 __NEXT_DATA__ */
function parseNextData(htmlTxt) {
  const doc = new DOMParser().parseFromString(htmlTxt, 'text/html');
  const el = doc.getElementById('__NEXT_DATA__');
  if (!el) return null;
  try { return JSON.parse(el.textContent); } catch (e) { return null; }
}

/** ① 搜索页 → 候选车系 [{sid, series_name}] */
async function searchCandidates(kw) {
  const htmlTxt = await getHtml(`/auto/search?keyword=${encodeURIComponent(kw)}`);
  const doc = new DOMParser().parseFromString(htmlTxt, 'text/html');
  const out = [];
  for (const a of Array.from(doc.querySelectorAll('a[href*="/auto/series/"]'))) {
    const m = /\/auto\/series\/(\d+)/.exec(a.getAttribute('href') || '');
    if (m) out.push({ sid: m[1], series_name: (a.textContent || '').trim().slice(0, 60) });
  }
  const seen = new Map();
  for (const c of out) {
    if (!c.sid) continue;
    if (!seen.has(c.sid) || (c.series_name && !seen.get(c.sid).series_name)) seen.set(c.sid, c);
  }
  return [...seen.values()].slice(0, 12);
}

/** ② 车系页 → 车型清单 [{id, name}] */
async function fetchSeriesCars(sid) {
  const htmlTxt = await getHtml(`/auto/series/${sid}`);
  // 先试 __NEXT_DATA__( SSR 数据最全 )
  const nd = parseNextData(htmlTxt);
  let cars = [];
  const walk = (o, depth) => {
    if (!o || depth > 6 || typeof o !== 'object') return;
    if (Array.isArray(o)) { o.forEach(x => walk(x, depth + 1)); return; }
    if ((o.id != null) && (o.name != null || o.car_name != null) &&
        (/车|版|型|km|续航|Pro|Plus|Max|Li/i.test(String(o.name || o.car_name || '')) || o.year != null)) {
      const id = String(o.id);
      if (/^\d+$/.test(id) && id.length >= 4) {
        cars.push({ id, name: String(o.name || o.car_name || '').slice(0, 60) });
      }
    }
    for (const k in o) walk(o[k], depth + 1);
  };
  if (nd) walk(nd.props && nd.props.pageProps, 0);
  // 兜底:从 DOM 里的车型链接抠 /auto/params-carIds-x-<id> 或 /auto/series/<sid>/<id>
  if (!cars.length) {
    const doc = new DOMParser().parseFromString(htmlTxt, 'text/html');
    const seen = new Set();
    for (const a of Array.from(doc.querySelectorAll('a[href]'))) {
      const href = a.getAttribute('href') || '';
      let m = /params-carIds-x-(\d+)/.exec(href);
      if (m && !seen.has(m[1])) { seen.add(m[1]); cars.push({ id: m[1], name: (a.textContent || '').trim().slice(0, 60) }); continue; }
      m = /\/auto\/series\/\d+\/(\d+)/.exec(href);
      if (m && !seen.has(m[1])) { seen.add(m[1]); cars.push({ id: m[1], name: (a.textContent || '').trim().slice(0, 60) }); }
    }
  }
  // 去重
  const uniq = new Map();
  for (const c of cars) if (!uniq.has(c.id)) uniq.set(c.id, c);
  return [...uniq.values()].slice(0, 60);
}

/** ③ 参数页 → rawData;抽 cars[] 与 fuel[](支持多车型 id 一次抓取) */
async function fetchParams(carIds) {
  const ids = carIds.join('-');
  let htmlTxt = null, usedUrl = '';
  for (const u of [`/auto/params-carIds-x-${ids}`, `/auto/params/${carIds[0]}`]) {
    try { htmlTxt = await getHtml(u); usedUrl = u; break; } catch (e) { log('  参数页尝试失败:', u, e.message); }
  }
  if (!htmlTxt) return { error: '参数页两种 URL 形态均失败' };
  const nd = parseNextData(htmlTxt);
  const raw = nd && nd.props && nd.props.pageProps && nd.props.pageProps.rawData;
  if (!raw) {
    return { error: '未取到 __NEXT_DATA__.props.pageProps.rawData', usedUrl, html_head: htmlTxt.slice(0, 300) };
  }
  const list = raw.car_info || raw.cars || [];
  const cars = [], fuel = [];
  for (const c of list) {
    const info = (c && c.info) || {};
    const vals = {};
    for (const k of FIELDS) {
      const cell = info[k];
      vals[k] = cell && typeof cell === 'object' ? (cell.value != null ? String(cell.value) : '')
                                                 : (cell != null ? String(cell) : '');
    }
    cars.push({ id: String((c && c.id) || ''), name: String((c && c.name) || (c && c.car_name) || ''), year: String((c && c.year) || ''), vals });
    const ffCell = info.fuel_form || info.energy_type || info.fuel_type || info.fuel;
    const ff = ffCell && typeof ffCell === 'object' ? ffCell.value : ffCell;
    if (c && c.id) fuel.push({ id: String(c.id), name: String(c.name || c.car_name || ''), ff: String(ff || '') });
  }
  const infoKeys = list.length ? Object.keys((list[0] && list[0].info) || {}) : [];
  return { cars, fuel, infoKeys, usedUrl, series_name: String(raw.series_name || raw.name || '') };
}

/** 选最佳候选:车系名与搜索词归一后全等 > 包含 > 首个 */
function pickCandidate(cands, kw) {
  if (!cands.length) return null;
  const norm = s => String(s || '').toLowerCase().replace(/[\s\-_.·]/g, '');
  const nkw = norm(kw);
  let best = null, bestScore = -1;
  for (const c of cands) {
    const nm = norm(c.series_name);
    let s = 0;
    if (nm && nkw) {
      if (nm === nkw) s = 100;
      else if (nm.includes(nkw)) s = 60;
      else if (nkw.includes(nm) && nm.length >= 2) s = 40;
    }
    if (s > bestScore) { bestScore = s; best = c; }
  }
  return bestScore >= 0 ? best : cands[0];
}

function download(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 1)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function loadDone() {
  try { return new Set(JSON.parse(localStorage.getItem(LS_KEY) || '[]')); }
  catch (e) { return new Set(); }
}
function saveDone(set) {
  try { localStorage.setItem(LS_KEY, JSON.stringify([...set])); }
  catch (e) { log('localStorage 写入失败(不影响抓取,仅失去断点):', e.message); }
}

/** 探查模式:三个页面各抓一个,dump 结构 */
async function probe() {
  const out = { tool: 'extract_browser.js', mode: 'probe', at: new Date().toISOString(), href: location.href };
  const kw = '宏光MINIEV';
  try {
    const cands = await searchCandidates(kw);
    out.step1_search = { kw, count: cands.length, candidates: cands.slice(0, 8) };
  } catch (e) { out.step1_search_error = String(e.message || e); }
  await sleep(jitter());
  const sid = (out.step1_search && out.step1_search.candidates && out.step1_search.candidates[0] || {}).sid;
  if (sid) {
    try {
      const cars = await fetchSeriesCars(sid);
      out.step2_series = { sid, cars_count: cars.length, cars_sample: cars.slice(0, 10) };
    } catch (e) { out.step2_series_error = String(e.message || e); }
    await sleep(jitter());
    if ((out.step2_series || {}).cars_count) {
      try {
        const ids = out.step2_series.cars_sample.slice(0, 5).map(c => c.id);
        const r = await fetchParams(ids);
        out.step3_params = { tried_ids: ids, usedUrl: r.usedUrl, cars_count: (r.cars || []).length,
                             info_keys: r.infoKeys, sample_car: (r.cars || [])[0],
                             error: r.error, html_head: r.html_head };
      } catch (e) { out.step3_params_error = String(e.message || e); }
    }
  }
  download(out, 'dcd_probe.json');
  log('探查完成,已下载 dcd_probe.json。请检查:');
  log('  ① step1_search.candidates 是否有候选(搜索页选择器是否有效)');
  log('  ② step2_series.cars_count 是否 >0(车系页车型清单是否取到)');
  log('  ③ step3_params.info_keys 是否含 15 个目标字段、usedUrl 是哪种形态');
  log('  ④ 若 step2/3 失败,请把参数页另存为 dcd_probe_page.html 放到 dcd_auto/state/');
  console.table(out.step3_params && out.step3_params.info_keys || []);
}

async function main() {
  if (DCD_PROBE) return await probe();
  if (!/dongchedi\.com/.test(location.hostname)) {
    log('⚠ 当前页面不在懂车帝域名下,同源 fetch 会失败。请先打开 https://www.dongchedi.com/ 再粘贴运行');
    return;
  }
  const targets = await new Promise(resolve => {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.json,application/json';
    inp.onchange = async () => {
      try { resolve(JSON.parse(await inp.file.text())); }
      catch (e) { log('targets.json 解析失败:', e.message); resolve(null); }
    };
    inp.click();
  });
  if (!targets || !targets.targets) { log('未选择或无法解析 targets.json,退出'); return; }

  const done = loadDone();
  const todo = targets.targets.filter(t => !done.has(t.qkey));
  log(`目标 ${targets.targets.length} 个,已完成 ${done.size} 个,本次待抓 ${todo.length} 个`);
  if (!todo.length) { log('全部已完成。如需重抓,清 localStorage 键 ' + LS_KEY); return; }

  let buf = [], seq = 1, okCount = 0, failCount = 0;
  const flush = (final = false) => {
    if (!buf.length) return;
    download({ tool: 'extract_browser.js', version: 2, created_at: new Date().toISOString(),
               source: '懂车帝官网(用户浏览器同源 fetch)', count: buf.length, series: buf },
             `dcd_bundle_${String(seq).padStart(3, '0')}.json`);
    log(`已下载 dcd_bundle_${String(seq).padStart(3, '0')}.json(${buf.length} 个车系)`);
    seq++; buf = [];
  };

  for (let i = 0; i < todo.length; i++) {
    const t = todo[i];
    const tag = `[${i + 1}/${todo.length}] ${t.family}/${t.part}(kw=${t.kw})`;
    try {
      let cands = [];
      try { cands = await searchCandidates(t.kw); } catch (e) { log(tag, '搜索失败:', e.message); }
      await sleep(jitter());
      const pick = pickCandidate(cands, t.kw);
      if (!pick) {
        buf.push({ qkey: t.qkey, kw: t.kw, sid: null, series_name: null, candidates: cands,
                   cars: [], fuel: [], captured_at: new Date().toISOString(), error: '搜索无候选车系' });
        failCount++;
        done.add(t.qkey); saveDone(done);
      } else {
        // ② 车系页取车型清单
        let carList = [];
        try { carList = await fetchSeriesCars(pick.sid); }
        catch (e) { log(tag, '车系页失败:', e.message); }
        await sleep(jitter());
        if (!carList.length) {
          buf.push({ qkey: t.qkey, kw: t.kw, sid: pick.sid, series_name: pick.series_name,
                     candidates: cands, cars: [], fuel: [], captured_at: new Date().toISOString(),
                     error: '车系页未取到车型清单' });
          failCount++;
        } else {
          // ③ 参数页(一次最多 20 个车型 id,避免 URL 过长)
          let cars = [], fuel = [], err = null;
          for (let s = 0; s < carList.length; s += 20) {
            const chunk = carList.slice(s, s + 20).map(c => c.id);
            const r = await fetchParams(chunk);
            if (r.error) { err = r.error; break; }
            cars = cars.concat(r.cars || []); fuel = fuel.concat(r.fuel || []);
            if (s + 20 < carList.length) await sleep(jitter());
          }
          buf.push({ qkey: t.qkey, kw: t.kw, sid: pick.sid, series_name: pick.series_name,
                     candidates: cands, car_list: carList, cars, fuel,
                     captured_at: new Date().toISOString(), error: err });
          if (err) failCount++; else okCount++;
          log(tag, err ? `参数页异常:${err}` : `✓ sid=${pick.sid} ${carList.length} 车型 / ${cars.length} 款型参数`);
        }
        done.add(t.qkey); saveDone(done);
      }
    } catch (e) {
      log(tag, '未预期错误:', e.message || e);
      failCount++;
    }
    if (buf.length >= BUNDLE_EVERY) flush();
    await sleep(jitter());
  }
  flush(true);
  log(`汇总:成功 ${okCount},失败/无参 ${failCount}`);
  log('把下载目录里的 dcd_bundle_*.json 拷到 dcd_auto/state/bundles/ 后运行 python dcd_auto/import_bundle.py');
}

main().catch(e => console.error('[DCD] 致命错误:', e));
