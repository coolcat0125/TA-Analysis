# -*- coding: utf-8 -*-
"""render_dashboard.py — 生成 DCD优化监控.html(实时监控版)

数据源:dcd_auto/state/pipeline_state.json(+ state/queue.json)。
形态:单文件 HTML,数据内嵌 + <meta refresh 15s> 自动刷新,无外部依赖,双击离线可开。
循环每跑完一批由 run_loop.ps1 调本脚本刷新;人工也可随时手动跑。

用法:
    python render_dashboard.py [--out 路径] [--no-coverage]
"""
import argparse
import datetime
import html
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dcd_common import (DASHBOARD, FIELD_MAP, PIPELINE_STATE, QUEUE_FILE, WB,
                        is_empty, load_rows)

REFRESH_SEC = 15
STALE_MIN = 10
AUTO_FIELDS = [f for f in FIELD_MAP if f != '百公里电耗(kWh/100km)']


def coverage():
    """当前 8 个自动字段的填充率(读底表;--no-coverage 或读失败时返回 None)。"""
    try:
        hdr, hi, rows = load_rows()
        out = {}
        for f in AUTO_FIELDS:
            n = sum(1 for r in rows if not is_empty(r[hi[f]]))
            out[f] = {'filled': n, 'total': len(rows), 'rate': round(n / len(rows) * 100, 1)}
        return out
    except Exception:
        return None


def esc(s):
    return html.escape(str(s if s is not None else ''))


def badge(text, kind):
    cls = {'ok': 'b-ok', 'warn': 'b-warn', 'bad': 'b-bad', 'info': 'b-info',
           'idle': 'b-idle'}.get(kind, 'b-info')
    return f'<span class="badge {cls}">{esc(text)}</span>'


def build(st, q, cov):
    now = datetime.datetime.now()
    loop = st.get('loop', {})
    status = loop.get('status', 'idle')
    status_map = {'running': ('运行中', 'ok'), 'paused': ('已暂停', 'bad'),
                  'idle': ('空闲', 'idle'), 'stopped': ('已停止', 'warn')}
    st_text, st_kind = status_map.get(status, (status, 'info'))

    updated = st.get('updated_at')
    stale = False
    if updated:
        try:
            dt = (now - datetime.datetime.fromisoformat(updated)).total_seconds() / 60
            stale = dt > STALE_MIN
        except Exception:
            pass

    base = st.get('baseline') or {}
    bv = base.get('verify') or {}
    t = st.get('totals', {})
    entries = q.get('entries', [])
    total_series = len(entries)
    done = t.get('series_done', 0)
    pct = (done / total_series * 100) if total_series else 0

    batches = st.get('batches', [])
    fills = st.get('recent_fills', [])
    blockers = st.get('blockers', [])

    # ---- KPI 卡
    kpis = [
        ('计划车系', f'{total_series}', f"缺口格 {q.get('total_gap_cells', 0):,}"),
        ('已完成', f'{done}', f"拒绝 {t.get('series_rejected', 0)} · 滞留 {t.get('series_held', 0)}"),
        ('累计绿填', f"{t.get('green_fill', 0):,}", f"红修 {t.get('red_fill', 0):,}(默认关)"),
        ('批次', f"{t.get('batches_applied', 0)}", f"回退 {t.get('batches_rolled_back', 0)}"),
        ('门禁', f"CRIT {bv.get('criticalTotal', '-')}", f"WARN {bv.get('warnTotal', '-')}"),
        ('护栏拦截', f"{t.get('cells_held', 0):,}", 'W5成对/污染锚点/歧义'),
    ]
    kpi_html = ''.join(
        f'<div class="kpi"><div class="kpi-l">{esc(k)}</div><div class="kpi-v">{esc(v)}</div>'
        f'<div class="kpi-s">{esc(s)}</div></div>' for k, v, s in kpis)

    # ---- 字段覆盖表
    if cov:
        rows_cov = ''.join(
            f"<tr><td>{esc(f)}</td><td class='num'>{cov[f]['filled']:,}/{cov[f]['total']:,}</td>"
            f"<td class='num'>{cov[f]['rate']}%</td>"
            f"<td><div class='bar'><div class='bar-in' style='width:{cov[f]['rate']}%'></div></div></td></tr>"
            for f in AUTO_FIELDS)
        cov_html = f"""<table><thead><tr><th>字段(自动填充集)</th><th>已填/总数</th><th>覆盖率</th><th></th></tr></thead>
        <tbody>{rows_cov}</tbody></table>"""
    else:
        cov_html = '<p class="muted">覆盖率未计算(--no-coverage 或底表不可读)</p>'

    # ---- 批次历史
    if batches:
        rows_b = ''
        for b in reversed(batches[-20:]):
            oc = b.get('outcome', '')
            ocb = {'applied': ('已应用', 'ok'), 'rolled-back': ('已回退', 'bad'),
                   'dry-run': ('预演', 'info'), 'no-op': ('无可填', 'idle')}.get(oc, (oc, 'info'))
            va = b.get('verify_after') or {}
            rows_b += (f"<tr><td>{esc(b.get('batch_id'))}</td><td>{esc((b.get('at') or '')[5:16])}</td>"
                       f"<td class='num'>{len(b.get('series', []))}</td>"
                       f"<td class='num'>{b.get('green_fill', 0)}</td>"
                       f"<td class='num'>{b.get('held', 0)}</td>"
                       f"<td class='num'>{b.get('ambiguous', 0)}</td>"
                       f"<td class='num'>{va.get('warnTotal', '-')}</td>"
                       f"<td>{badge(*ocb)}</td></tr>")
        batch_html = f"""<table><thead><tr><th>批次</th><th>时间</th><th>车系</th><th>绿填</th><th>护栏拦下</th>
        <th>歧义跳过</th><th>WARN(后)</th><th>结果</th></tr></thead><tbody>{rows_b}</tbody></table>"""
    else:
        batch_html = '<p class="muted">尚无批次记录</p>'

    # ---- 当前批次 / 队列预览
    cur = loop.get('current_batch')
    cur_html = f"<p>当前批次:<b>{esc(cur)}</b></p>" if cur else '<p class="muted">无活动批次</p>'
    if loop.get('paused_reason'):
        cur_html += f"<p class='alert'>暂停原因:{esc(loop['paused_reason'])}</p>"

    upcoming = ''.join(
        f"<tr><td class='num'>{e['priority_rank']}</td><td>{esc(e['family'])} / {esc(e['part'])}</td>"
        f"<td class='num'>{e['rows']}</td><td class='num'>{e['gap_cells']}</td>"
        f"<td>{'共享名' if e.get('shared_name') else ''}</td></tr>"
        for e in entries[:12])
    queue_html = f"""<table><thead><tr><th>#</th><th>车系(族/名)</th><th>行数</th><th>缺口格</th><th>标记</th></tr></thead>
    <tbody>{upcoming}</tbody></table>"""

    # ---- 最近填充
    if fills:
        rows_f = ''.join(
            f"<tr><td>{esc((f.get('batch') or ''))}</td><td>{esc(f.get('part'))}</td>"
            f"<td class='num'>{f.get('row')}</td><td>{esc(f.get('code'))}</td>"
            f"<td>{esc(f.get('field'))}</td><td>{esc(f.get('new'))}</td></tr>"
            for f in fills[:25])
        fills_html = f"""<table><thead><tr><th>批次</th><th>车系</th><th>行</th><th>产品型号</th><th>字段</th><th>新值</th></tr></thead>
        <tbody>{rows_f}</tbody></table>"""
    else:
        fills_html = '<p class="muted">尚无填充记录</p>'

    # ---- blocker
    if blockers:
        blk_html = ''.join(
            f"<div class='blocker'><b>{esc(b.get('batch'))}</b> {esc('; '.join(b.get('problems', [])))}"
            f"<span class='muted'> — {esc(b.get('action'))}</span></div>"
            for b in reversed(blockers[-5:]))
    else:
        blk_html = '<p class="muted">无回退/阻塞记录</p>'

    # ---- ETA 面板
    p = st.get('progress') or {}
    if p:
        eta_html = f"""<table>
        <tr><th>已完成车系</th><th>剩余车系</th><th>剩余批次</th><th>apply实测均耗时</th><th>抽取假设耗时</th><th>预计剩余</th><th>预计完成时刻</th></tr>
        <tr><td class="num">{p.get('series_done')} / {p.get('series_total')}({p.get('pct')}%)</td>
        <td class="num">{p.get('series_remaining')}</td><td class="num">{p.get('batches_remaining')}</td>
        <td class="num">{p.get('avg_apply_sec_per_batch') or '-'} s/批</td>
        <td class="num">{p.get('assumed_extract_sec_per_batch')} s/批</td>
        <td class="num"><b>{p.get('eta_human')}</b></td>
        <td class="num">{p.get('eta_finish_at')}</td></tr></table>
        <p class="muted">{esc(p.get('note'))} · 每批 {p.get('assumed_batch_size')} 车系</p>"""
    else:
        eta_html = '<p class="muted">尚无进度数据(跑过 apply 或执行 eta 子命令后显示)</p>'

    stale_html = ("<div class='stale'>⚠ 监控数据已停更超过 %d 分钟——循环可能已停止,"
                  "底表 SHA/门禁基线可能已变化,请回终端查看</div>" % STALE_MIN) if stale else ''

    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta http-equiv="refresh" content="{REFRESH_SEC}">
<title>DCD 自动优化监控</title>
<style>
:root{{color-scheme:light}}
*{{box-sizing:border-box}}
body{{font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;margin:0;background:#f5f6f8;color:#1f2329}}
.wrap{{max-width:1180px;margin:0 auto;padding:22px}}
h1{{font-size:21px;margin:0 0 4px}}
h2{{font-size:15px;margin:26px 0 10px;color:#333;border-left:4px solid #e8734a;padding-left:8px}}
.meta{{color:#666;font-size:12.5px;margin-bottom:14px}}
.kpis{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}}
.kpi{{background:#fff;border:1px solid #e4e6eb;border-radius:10px;padding:12px 14px}}
.kpi-l{{font-size:12px;color:#777}}
.kpi-v{{font-size:22px;font-weight:700;margin:2px 0}}
.kpi-s{{font-size:11.5px;color:#999}}
.panel{{background:#fff;border:1px solid #e4e6eb;border-radius:10px;padding:14px 16px;margin-top:12px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:6px}}
th{{text-align:left;color:#666;font-weight:600;border-bottom:2px solid #eee;padding:6px 8px}}
td{{border-bottom:1px solid #f0f0f0;padding:6px 8px}}
.num{{text-align:right;font-variant-numeric:tabular-nums}}
.bar{{background:#eee;border-radius:4px;height:9px;width:100%;min-width:90px}}
.bar-in{{background:#4a9e5c;height:9px;border-radius:4px}}
.badge{{display:inline-block;padding:1px 8px;border-radius:9px;font-size:11.5px}}
.b-ok{{background:#e6f4ea;color:#1e7e34}}.b-warn{{background:#fff4e0;color:#a15c00}}
.b-bad{{background:#fdecea;color:#c0392b}}.b-info{{background:#e8f0fe;color:#1a56b0}}
.b-idle{{background:#eee;color:#666}}
.alert{{background:#fdecea;border:1px solid #f5c6c0;border-radius:8px;padding:8px 12px;font-size:12.5px}}
.stale{{background:#fff4e0;border:1px solid #f0d9a8;border-radius:8px;padding:9px 13px;margin:10px 0;font-size:13px}}
.blocker{{background:#fdf3f2;border-left:3px solid #c0392b;padding:7px 11px;margin:6px 0;font-size:12.5px;border-radius:0 6px 6px 0}}
.muted{{color:#999;font-size:12.5px}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}
.foot{{margin-top:26px;color:#999;font-size:11.5px;line-height:1.7}}
@media(max-width:900px){{.kpis{{grid-template-columns:repeat(3,1fr)}}.grid2{{grid-template-columns:1fr}}}}
</style></head><body><div class="wrap">
<h1>DCD 懂车帝参数自动优化 · 监控看板</h1>
<div class="meta">数据截至 {esc(updated or '-')} · 页面 {now:%Y-%m-%d %H:%M:%S} 生成 · 每 {REFRESH_SEC}s 自动刷新 ·
循环状态 {badge(st_text, st_kind)} · 基线 HEAD {esc(base.get('git_head'))} · 底表 SHA {esc(base.get('wb_sha'))}</div>
{stale_html}
<div class="kpis">{kpi_html}</div>
<div class="panel"><b>总体进度</b>
<div class="bar" style="margin-top:8px;height:14px"><div class="bar-in" style="width:{pct:.1f}%;height:14px"></div></div>
<div class="meta" style="margin-top:6px">已完成车系 {done} / {total_series}({pct:.1f}%)</div></div>
<h2>进度与预计完成时间</h2>
<div class="panel">{eta_html}</div>
<h2>字段覆盖率(自动填充集 8 项)</h2>
<div class="panel">{cov_html}</div>
<h2>批次历史</h2>
<div class="panel">{batch_html}</div>
<div class="grid2">
<div><h2>当前批次</h2><div class="panel">{cur_html}</div></div>
<div><h2>回退与阻塞</h2><div class="panel">{blk_html}</div></div>
</div>
<h2>队列预览(缺口优先前 12)</h2>
<div class="panel">{queue_html}</div>
<h2>最近填充(25 条)</h2>
<div class="panel">{fills_html}</div>
<div class="foot">
纪律红线:只补空 · 唯一共识才写 · BEV 禁填油耗/发动机 · 车长/轴距成对(W5 铁律) · 22:40-06:00 禁写库 · 每批三件套自检(verify CRITICAL=0 / audit strict PASS / 双看板重生成) · 红修默认关闭待批准。<br>
数据来源:懂车帝参数页(Tabbit 用户会话通道)· 台账落底表《变更记录》· 状态文件 dcd_auto/state/pipeline_state.json · 本看板由 render_dashboard.py 生成。
</div>
</div></body></html>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=None)
    ap.add_argument('--no-coverage', action='store_true')
    a = ap.parse_args()
    st = json.load(open(PIPELINE_STATE, encoding='utf-8')) if os.path.exists(PIPELINE_STATE) else {}
    q = json.load(open(QUEUE_FILE, encoding='utf-8')) if os.path.exists(QUEUE_FILE) else {'entries': []}
    cov = None if a.no_coverage else coverage()
    out = a.out or DASHBOARD
    html_doc = build(st, q, cov)
    open(out, 'w', encoding='utf-8').write(html_doc)
    print(f'看板已生成: {out} ({len(html_doc) / 1024:.0f} KB)')


if __name__ == '__main__':
    main()
