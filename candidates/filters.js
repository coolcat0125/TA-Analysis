/* 共用筛选条：动力类型 / 年份区间 / 细分市场 / 指标选择 */
(function () {
  const S = CORE.state;
  function h(html) { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstChild; }

  window.mountFilters = function (container, opts = {}) {
    const segs = CORE.segList();
    const el = h(`
      <div class="filters">
        <div class="f-group">
          <span class="f-label">动力类型</span>
          <div class="seg" data-k="power">
            <button data-v="ALL" class="on">全部</button><button data-v="BEV">BEV</button><button data-v="PHEV">PHEV/EREV</button>
          </div>
        </div>
        <div class="f-group">
          <span class="f-label">年份</span>
          <select class="sel" data-k="y0">${CORE.YEARS.map(y => `<option ${y === 2021 ? 'selected' : ''}>${y}</option>`).join('')}</select>
          <span class="f-dash">–</span>
          <select class="sel" data-k="y1">${CORE.YEARS.map(y => `<option ${y === 2026 ? 'selected' : ''}>${y}</option>`).join('')}</select>
        </div>
        <div class="f-group f-segs">
          <span class="f-label">细分市场</span>
          <details class="dd">
            <summary class="sel"><span class="dd-txt">全部</span></summary>
            <div class="dd-menu">
              <div class="dd-actions"><a data-a="all">全选</a><a data-a="none">清空</a><a data-a="main">主流五档</a></div>
              ${segs.map(s => `<label><input type="checkbox" value="${s}" checked> ${s}</label>`).join('')}
            </div>
          </details>
        </div>
        <div class="f-group f-metric" ${opts.metric === false ? 'style="display:none"' : ''}>
          <span class="f-label">细分指标</span>
          <select class="sel" data-k="segMetric">${Object.entries(CORE.SEG_METRICS).map(([k, [l]]) => `<option value="${k}" ${k === 'range' ? 'selected' : ''}>${l}</option>`).join('')}</select>
        </div>
        <button class="f-reset" title="重置筛选">重置</button>
      </div>`);
    container.appendChild(el);

    const apply = () => { CORE.renderAll(); document.dispatchEvent(new CustomEvent('filters-changed')); };
    el.querySelector('.seg').addEventListener('click', e => {
      const b = e.target.closest('button'); if (!b) return;
      el.querySelectorAll('.seg button').forEach(x => x.classList.toggle('on', x === b));
      S.power = b.dataset.v; apply();
    });
    const y0 = el.querySelector('[data-k=y0]'), y1 = el.querySelector('[data-k=y1]');
    const yr = () => { let a = +y0.value, b = +y1.value; if (a > b) [a, b] = [b, a]; S.years = [a, b]; apply(); };
    y0.onchange = yr; y1.onchange = yr;
    el.querySelector('[data-k=segMetric]').onchange = e => { S.segMetric = e.target.value; apply(); };
    const boxes = Array.from(el.querySelectorAll('.dd-menu input'));
    const txt = el.querySelector('.dd-txt');
    const syncSeg = () => {
      const on = boxes.filter(b => b.checked).map(b => b.value);
      S.segs = on.length === boxes.length ? null : new Set(on);
      txt.textContent = S.segs ? (on.length === 0 ? '未选择' : on.length <= 2 ? on.join('、') : `已选 ${on.length} 项`) : '全部';
      apply();
    };
    boxes.forEach(b => b.onchange = syncSeg);
    el.querySelector('.dd-actions').addEventListener('click', e => {
      const a = e.target.dataset.a; if (!a) return;
      const main = ['Car-A', 'Car-B', 'Car-C', 'Car-D', 'SUV-B', 'SUV-C', 'SUV-D', 'MPV-D'];
      boxes.forEach(b => b.checked = a === 'all' ? true : a === 'none' ? false : main.includes(b.value));
      syncSeg();
    });
    document.addEventListener('click', e => { const d = el.querySelector('.dd'); if (d.open && !d.contains(e.target)) d.open = false; });
    el.querySelector('.f-reset').onclick = () => {
      S.power = 'ALL'; S.years = [2021, 2026]; S.segs = null; S.segMetric = 'range';
      el.querySelectorAll('.seg button').forEach(x => x.classList.toggle('on', x.dataset.v === 'ALL'));
      y0.value = '2021'; y1.value = '2026'; boxes.forEach(b => b.checked = true); txt.textContent = '全部';
      el.querySelector('[data-k=segMetric]').value = 'range'; apply();
    };
    return el;
  };

  // 指标探索器下拉
  window.mountMetricPicker = function (container) {
    const sel = h(`<select class="sel">${Object.entries(CORE.METRICS).map(([k, m]) => `<option value="${k}">${m.label}（${m.unit}）</option>`).join('')}</select>`);
    sel.onchange = e => { CORE.state.metric = e.target.value; CORE.render('c-explorer', 'metricExplorer'); };
    container.appendChild(sel);
  };
})();
