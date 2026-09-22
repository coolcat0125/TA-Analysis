# -*- coding: utf-8 -*-
"""dcd_pipeline.py — DCD 自动优化批次编排器(护栏全自动)

子命令:
  baseline                     捕获当前门禁基线(verify/audit/底表SHA/git HEAD)
  next-batch --n 5             从队列取下一批车系,写 state/search/b<NNN>.json(抽取请求)
  decide --batch b001          读抽取结果 → decide_series 定链 → state/decided/b001.json
  apply --batch b001 [--dry-run] [--force]
                               备份底表 → 逐车系应用 → 三件套自检 → 失败自动回退并暂停
  status                       打印状态摘要

安全链(任一不过即回退整批并暂停循环):
  1) apply 前备份底表到 state/backups/pre_<batch>.xlsx
  2) apply 后跑 verify_consistency.py:CRITICAL 必须=0;W4/W5/W6/W7 不得增;W2 增幅须解释
  3) audit_data.py --strict 退出码必须=0
  4) 写库时段纪律:本地时间 22:40 后拒绝写库(--force 可越权,需人工理由)

红修默认关闭(--allow-t4-redfix 开启,且 T4 补丁本身待 Orchestrator/人工批准)。
"""
import argparse
import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dcd_common import (BATCH_DIR, BACKUP_DIR, DECIDED_DIR, FIELD_MAP, PIPELINE_STATE,
                        QUEUE_FILE, RAW_DCD, SEARCH_DIR, STATE, STOP_FLAG, WB,
                        ensure_dirs, today)
from dcd_refill_core import apply_entry, decide_series, ledger_max_seq

import openpyxl

# 自动填充字段集:默认 8 字段(排除 百公里电耗——DCD 为 CLTC 工况,与底表既有
# NEDC/官方口径混填会制造 W2 电耗矛盾;该字段覆盖率已 95.3%,性价比低)。
# --include-consumption 可纳入(自行承担 W2 增幅)。
AUTO_FIELDS = [f for f in FIELD_MAP if f != '百公里电耗(kWh/100km)']
CONSUMPTION_FIELD = '百公里电耗(kWh/100km)'
WRITE_CUTOFF = (22, 40)          # 22:40 后不写库(项目纪律)
WARN_HARD = ('W4', 'W5', 'W6', 'W7', 'W8')   # 这些类任何增幅都回退(W8=媒体车系品牌冲突)
WARN_SOFT = ('W2',)              # 允许增幅但必须入报告


# ---------------------------------------------------------------- state 读写
def load_state():
    if os.path.exists(PIPELINE_STATE):
        return json.load(open(PIPELINE_STATE, encoding='utf-8'))
    return {'updated_at': None, 'loop': {'status': 'idle', 'current_batch': None,
                                         'paused_reason': None},
            'baseline': None, 'totals': {}, 'series': {}, 'batches': [],
            'recent_fills': [], 'blockers': []}


def save_state(st):
    st['updated_at'] = datetime.datetime.now().isoformat(timespec='seconds')
    json.dump(st, open(PIPELINE_STATE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


def wb_sha():
    h = hashlib.sha256()
    with open(WB, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()[:16]


def git_head():
    try:
        return subprocess.run(['git', 'rev-parse', '--short', 'HEAD'], capture_output=True,
                              text=True, cwd=os.path.dirname(WB)).stdout.strip()
    except Exception:
        return None


# ---------------------------------------------------------------- 自检
def run_verify(wb_path=None):
    """跑 verify_consistency,返回 {criticalTotal, warnTotal, warn, pass}。"""
    out = os.path.join(STATE, '_verify_last.json')
    cmd = [sys.executable, os.path.join(os.path.dirname(WB), 'verify_consistency.py'),
           '--json', out]
    if wb_path:
        cmd += ['--input', wb_path]
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=os.path.dirname(WB))
    try:
        d = json.load(open(out, encoding='utf-8'))
        return {'criticalTotal': d.get('criticalTotal', 0), 'warnTotal': d.get('warnTotal', 0),
                'warn': d.get('warn', {}), 'pass': d.get('pass', False), 'raw': out}
    except Exception:
        return {'criticalTotal': -1, 'warnTotal': -1, 'warn': {}, 'pass': False,
                'error': r.stdout + r.stderr}


def run_audit():
    """跑 audit_data --strict,返回 (ok, detail)。"""
    r = subprocess.run([sys.executable, os.path.join(os.path.dirname(WB), 'audit_data.py'),
                        '--strict'], capture_output=True, text=True, cwd=os.path.dirname(WB))
    return r.returncode == 0, (r.stdout + r.stderr)[-400:]


def cmd_baseline(st):
    v = run_verify()
    a_ok, a_detail = run_audit()
    st['baseline'] = {'captured_at': datetime.datetime.now().isoformat(timespec='seconds'),
                      'git_head': git_head(), 'wb_sha': wb_sha(),
                      'verify': v, 'audit_strict_pass': a_ok}
    st['loop']['status'] = 'idle'
    st['loop']['paused_reason'] = None
    save_state(st)
    print(f"基线已捕获: HEAD={st['baseline']['git_head']} SHA={st['baseline']['wb_sha']} "
          f"CRITICAL={v['criticalTotal']} WARN={v['warnTotal']} audit={'PASS' if a_ok else 'FAIL'}")
    return 0


# ---------------------------------------------------------------- next-batch
def cmd_next_batch(st, n=5, from_mapping=False):
    """取下一批。from_mapping=True 时只挑『已有人审映射 + DCD 产物已存在』的车系,
    并直接按人审结论生成 search 结果文件——用于无 Tabbit 环境下跑通管线
    (人审即定链,decide_series 已实测可复现该结论,见 README 自测)。"""
    ensure_dirs()
    q = json.load(open(QUEUE_FILE, encoding='utf-8'))
    existing = [f for f in os.listdir(SEARCH_DIR) if f.startswith('b') and f.endswith('.json')]
    next_no = max([int(f[1:-5]) for f in existing if f[1:-5].isdigit()] + [0]) + 1
    batch_id = f'b{next_no:03d}'

    # 已人审接受的映射:kw 归一 → sid/series
    reviewed = {}
    mp = os.path.join(RAW_DCD, 'series_mapping.json')
    if os.path.exists(mp):
        for e in json.load(open(mp, encoding='utf-8')).get('reviewed', []):
            if e.get('ok'):
                from dcd_common import norm_name
                reviewed[norm_name(e.get('kw') or e.get('series') or '')] = e

    picked = []
    for e in q['entries']:
        s = st['series'].get(e['qkey'], {}).get('status')
        if s in DONE_STATUS:
            continue
        # 普通模式跳过已抽取车系(数据在位,不需重复搜索;其残余缺口由 from-mapping 模式消化)
        if not from_mapping and e.get('already_extracted'):
            continue
        if from_mapping:
            from dcd_common import norm_name
            npart = norm_name(e['part'])
            hit = None
            for nk, m in reviewed.items():
                # 前缀锚定匹配(防自由子串交叉误配,如 part='U5' ↔ kw='北汽EU5'):
                # 底表通用名称常是人审搜索词的省略('比亚迪唐Pro' vs kw '比亚迪唐')
                if not nk or not npart:
                    continue
                nser = norm_name(m.get('series') or '')
                anchored = npart.startswith(nk) or nk.startswith(npart)
                in_series = len(npart) >= 3 and nser and npart in nser   # 'EU5' ⊂ '北京EU5'
                if (anchored or in_series) and len(min(npart, nk, key=len)) >= 2:
                    hit = m
                    break
            if not hit:
                continue
            if not os.path.exists(os.path.join(RAW_DCD, f"dcd_s{hit['sid']}.json")):
                continue
            # 关键:沿用人审的搜索词作为 kw(底表裸 part 常不是有效搜索词,如 '01')
            e = dict(e, kw=hit.get('kw') or e['kw'],
                     _map_sid=str(hit['sid']), _map_series=hit.get('series'))
        if len(picked) >= n:
            break
        picked.append(e)

    if not picked:
        print('队列已无待处理车系(全部 done/rejected/held)。用 --retry 或重建队列。'
              if not from_mapping else
              '无『已有人审映射且 DCD 产物在位』的待处理车系;先跑 Tabbit 抽取或去掉 --from-mapping')
        return 1

    doc = {'batch_id': batch_id, 'created_at': datetime.datetime.now().isoformat(timespec='seconds'),
           'size': len(picked), 'from_mapping': bool(from_mapping),
           'series': [{k: v for k, v in p.items() if not k.startswith('_')} |
                      ({'map_sid': p['_map_sid'], 'map_series': p.get('_map_series')}
                       if p.get('_map_sid') else {})
                      for p in picked],
           'candidates': {}}
    json.dump(doc, open(os.path.join(SEARCH_DIR, f'{batch_id}.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    # from_mapping 模式:人审结论即定链,直接生成 search 结果(单候选=人审接受的车系)
    if from_mapping:
        cands = {}
        for p in picked:
            if p.get('_map_sid'):
                cands[p['qkey']] = [{'sid': p['_map_sid'],
                                     'series_name': p.get('_map_series') or p['part']}]
        json.dump({'batch_id': batch_id, 'from_mapping': True, 'candidates': cands},
                  open(os.path.join(SEARCH_DIR, f'{batch_id}.results.json'), 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=1)
        print(f"批次 {batch_id} 已生成({len(picked)} 车系,from-mapping 模式:人审结论已作为定链输入)")
    else:
        print(f"批次 {batch_id} 已生成({len(picked)} 车系): {os.path.join(SEARCH_DIR, batch_id + '.json')}")
    st['loop']['current_batch'] = batch_id
    save_state(st)
    for p in picked:
        print(f"  #{p['priority_rank']:<4} {p['family']} / {p['part']}  gap={p['gap_cells']}"
              + (f"  sid={p['_map_sid']}" if p.get('_map_sid') else ''))
    return 0


# ---------------------------------------------------------------- decide
def cmd_decide(st, batch_id):
    req_path = os.path.join(SEARCH_DIR, f'{batch_id}.json')
    res_path = os.path.join(SEARCH_DIR, f'{batch_id}.results.json')
    if not os.path.exists(req_path):
        print(f'缺批次请求文件: {req_path}'); return 1
    if not os.path.exists(res_path):
        print(f'缺抽取结果: {res_path}(先跑 Tabbit 搜索阶段)'); return 1
    req = json.load(open(req_path, encoding='utf-8'))
    res = json.load(open(res_path, encoding='utf-8'))
    cands = res.get('candidates') or {}
    decided = {'batch_id': batch_id, 'decided_at': datetime.datetime.now().isoformat(timespec='seconds'),
               'series': []}
    for s in req['series']:
        r = decide_series(s['kw'], cands.get(s['qkey']) or cands.get(s['kw']) or [],
                          s.get('expect_brand'), s.get('shared_name'))
        # kw_alt 兜底:主kw无候选时用 族+名 再搜一次的结果
        if r['status'] == 'reject' and s.get('kw_alt'):
            r2 = decide_series(s['kw_alt'], cands.get(s['qkey'] + '#alt') or [],
                               s.get('expect_brand'), True)
            if r2['status'] == 'accept':
                r = r2
        item = dict(s); item.update(r)
        decided['series'].append(item)
        st['series'][s['qkey']] = {'status': r['status'], 'batch': batch_id,
                                   'sid': r.get('sid'), 'reason': r.get('reason'),
                                   'updated_at': datetime.datetime.now().isoformat(timespec='seconds')}
    json.dump(decided, open(os.path.join(DECIDED_DIR, f'{batch_id}.json'), 'w', encoding='utf-8'),
              ensure_ascii=False, indent=1)
    save_state(st)
    acc = sum(1 for s in decided['series'] if s['status'] == 'accept')
    print(f"批次 {batch_id} 定链完成: accept={acc} / {len(decided['series'])}")
    for s in decided['series']:
        print(f"  {s['status']:<7} {s['family']} / {s['part']}  {s.get('reason', '')[:46]}")
    return 0


# ---------------------------------------------------------------- apply
def local_now():
    """项目本地时间。默认按用户所在时区 UTC+8(Asia/Singapore)折算,避免依赖
    tzdata(Windows Python 常缺 IANA 库);需要别的时区设 DCD_TZ_OFFSET(小时数)。"""
    import datetime as _dt
    off = 8
    try:
        off = float(os.environ.get('DCD_TZ_OFFSET', '8'))
    except Exception:
        pass
    return _dt.datetime.now(_dt.timezone(_dt.timedelta(hours=off)))


def in_write_window(force=False):
    if force:
        return True, '--force 越权'
    now = local_now()
    if (now.hour, now.minute) >= WRITE_CUTOFF or now.hour < 6:
        return False, (f'本地时间 {now:%H:%M}(按 UTC{local_now().utcoffset()} 折算)'
                       f' 在禁写窗(22:40-06:00),--force 可越权')
    return True, ''


def cmd_apply(st, batch_id, dry_run=False, force=False, allow_t4=False,
              include_consumption=False):
    ensure_dirs()
    dec_path = os.path.join(DECIDED_DIR, f'{batch_id}.json')
    if not os.path.exists(dec_path):
        print(f'缺定链文件: {dec_path}(先跑 decide)'); return 1
    dec = json.load(open(dec_path, encoding='utf-8'))
    accepted = [s for s in dec['series'] if s['status'] == 'accept' and s.get('sid')]
    if not accepted:
        print('本批无 accept 车系,无需应用'); return 0
    ok_win, why = in_write_window(force)
    if not ok_win and not dry_run:
        print(f'拒绝写库: {why}')
        st['loop']['status'] = 'paused'
        st['loop']['paused_reason'] = why
        save_state(st)
        return 2

    fields = list(AUTO_FIELDS) + ([CONSUMPTION_FIELD] if include_consumption else [])
    started_at = datetime.datetime.now().isoformat(timespec='seconds')
    fuel = {}
    fp = os.path.join(RAW_DCD, 'dcd_fuel_form.json')
    if os.path.exists(fp):
        fuel = json.load(open(fp, encoding='utf-8'))
    fuel_ff = {(str(s), str(sp['id'])): str(sp.get('ff') or '')
               for s, lst in fuel.items() for sp in (lst or [])
               if isinstance(sp, dict) and sp.get('id')}

    # 1) dry-run 预演(不备份不写)
    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']; led = wb['变更记录']
    hdr = [str(c.value or '').strip() for c in ws[1]]; hi = {h: i for i, h in enumerate(hdr)}
    seq = ledger_max_seq(led)
    plans, missing = [], []
    for s in accepted:
        dj = os.path.join(RAW_DCD, f"dcd_s{s['sid']}.json")
        if not os.path.exists(dj):
            missing.append(s); continue
        dcd = json.load(open(dj, encoding='utf-8'))
        rep, seq = apply_entry(wb, ws, led, hi, seq, s, dcd, fuel_ff, s['sid'], fields,
                               dry_run=True)
        plans.append((s, rep))
    wb.close()
    total_green = sum(r['green_fill'] for _, r in plans)
    total_red = sum(r['red_fix'] for _, r in plans)
    total_held = sum(len(r['held']) for _, r in plans)
    if missing:
        for s in missing:
            st['series'][s['qkey']]['status'] = 'fetch_failed'
        print(f"警告: {len(missing)} 个车系缺 DCD 数据文件(抽取未完成),跳过")
    print(f"dry-run: green={total_green} red={total_red} held={total_held} "
          f"(车系 {len(plans)}/{len(accepted)})")
    if dry_run or total_green + total_red == 0:
        if not dry_run:
            for s, r in plans:
                st['series'][s['qkey']] = {'status': 'done', 'batch': batch_id,
                                           'reason': '无可填格(歧义或已满)',
                                           'updated_at': datetime.datetime.now().isoformat(timespec='seconds')}
        _record_batch(st, batch_id, plans, outcome='dry-run' if dry_run else 'no-op',
                      started_at=started_at)
        update_progress(st)
        save_state(st)
        return 0

    # 2) 备份 + 实写
    backup = os.path.join(BACKUP_DIR, f'pre_{batch_id}.xlsx')
    shutil.copy2(WB, backup)
    wb = openpyxl.load_workbook(WB)
    ws = wb['NEV公告参数汇总']; led = wb['变更记录']
    hdr = [str(c.value or '').strip() for c in ws[1]]; hi = {h: i for i, h in enumerate(hdr)}
    seq = ledger_max_seq(led)
    real = []
    for s, _ in plans:
        dcd = json.load(open(os.path.join(RAW_DCD, f"dcd_s{s['sid']}.json"), encoding='utf-8'))
        rep, seq = apply_entry(wb, ws, led, hi, seq, s, dcd, fuel_ff, s['sid'], fields,
                               dry_run=False)
        real.append((s, rep))
        st['series'][s['qkey']] = {'status': 'applied', 'batch': batch_id, 'sid': s['sid'],
                                   'green': rep['green_fill'], 'held': len(rep['held']),
                                   'updated_at': datetime.datetime.now().isoformat(timespec='seconds')}
    wb.save(WB)
    print(f"已写入底表: green={sum(r['green_fill'] for _, r in real)}")

    # 3) 三件套自检
    v = run_verify()
    base = st.get('baseline') or {}
    bv = base.get('verify') or {}
    problems = []
    if v['criticalTotal'] != 0:
        problems.append(f"CRITICAL={v['criticalTotal']}(基线 {bv.get('criticalTotal')})")
    for k in WARN_HARD:
        if v['warn'].get(k, 0) > bv.get('warn', {}).get(k, 0):
            problems.append(f"{k} 增幅 {bv.get('warn', {}).get(k, 0)}→{v['warn'].get(k, 0)}")
    a_ok, a_detail = run_audit()
    if not a_ok:
        problems.append('audit --strict 未通过')
    if problems:
        shutil.copy2(backup, WB)                       # 回退
        v2 = run_verify()
        why = '; '.join(problems)
        st['loop']['status'] = 'paused'
        st['loop']['paused_reason'] = f'批次 {batch_id} 自检失败已回退: {why}'
        for s, _ in real:
            st['series'][s['qkey']]['status'] = 'rolled_back'
        _record_batch(st, batch_id, real, outcome='rolled-back',
                      verify_after=v2, problems=problems, started_at=started_at)
        update_progress(st)
        st.setdefault('blockers', []).append(
            {'at': datetime.datetime.now().isoformat(timespec='seconds'),
             'batch': batch_id, 'problems': problems, 'action': '底表已回退,循环暂停'})
        save_state(st)
        print('!! 自检失败,底表已回退,循环暂停:')
        for p in problems:
            print('   -', p)
        return 3
    # 4) 通过
    st['baseline'] = {'captured_at': datetime.datetime.now().isoformat(timespec='seconds'),
                      'git_head': git_head(), 'wb_sha': wb_sha(),
                      'verify': v, 'audit_strict_pass': a_ok}
    _record_batch(st, batch_id, real, outcome='applied', verify_after=v, started_at=started_at)
    update_progress(st)
    save_state(st)
    print(f"自检通过: CRITICAL=0 WARN={v['warnTotal']}(基线 {bv.get('warnTotal')}) audit=PASS")
    p = st.get('progress') or {}
    if p.get('series_remaining'):
        print(f"进度: {p['series_done']}/{p['series_total']} 车系({p['pct']}%),"
              f"剩余 {p['series_remaining']} 个 / {p['batches_remaining']} 批,"
              f"预计 {p['eta_human']}(约 {p['eta_finish_at']} 完成)")
    return 0


def _record_batch(st, batch_id, plans, outcome, verify_after=None, problems=None,
                  started_at=None):
    t = st.setdefault('totals', {})
    finished = datetime.datetime.now()
    dur = None
    if started_at:
        try:
            dur = (finished - datetime.datetime.fromisoformat(started_at)).total_seconds()
        except Exception:
            dur = None
    b = {'batch_id': batch_id, 'at': finished.isoformat(timespec='seconds'),
         'started_at': started_at, 'duration_sec': round(dur, 1) if dur else None,
         'outcome': outcome, 'green_fill': sum(r['green_fill'] for _, r in plans),
         'red_fix': sum(r['red_fix'] for _, r in plans),
         'held': sum(len(r['held']) for _, r in plans),
         'ambiguous': sum(r['ambiguous_skip'] for _, r in plans),
         'series': [{'qkey': s['qkey'], 'part': s['part'], 'family': s['family'],
                     'sid': s.get('sid'), 'green': r['green_fill'], 'held': len(r['held']),
                     'errors': len(r['errors'])} for s, r in plans],
         'verify_after': verify_after, 'problems': problems}
    st['batches'].append(b)
    st['batches'] = st['batches'][-100:]
    t['batches_applied'] = t.get('batches_applied', 0) + (1 if outcome == 'applied' else 0)
    t['batches_rolled_back'] = t.get('batches_rolled_back', 0) + (1 if outcome == 'rolled-back' else 0)
    if outcome == 'applied':
        # 只有真正落库的批计入累计;回退批的格数已撤销,单列备查
        t['green_fill'] = t.get('green_fill', 0) + b['green_fill']
        t['red_fix'] = t.get('red_fill', 0) + b['red_fix']
        t['cells_held'] = t.get('cells_held', 0) + b['held']
    elif outcome == 'rolled-back':
        t['fills_rolled_back'] = t.get('fills_rolled_back', 0) + b['green_fill'] + b['red_fix']
    # recent_fills 只记真正落库的批;dry-run/no-op 不记(防重复计数),
    # rolled-back 的格已撤销也不记(避免看板显示不存在的数据)
    if outcome == 'applied':
        for s, r in plans:
            for f in r['fills']:
                st['recent_fills'].insert(0, dict(f, batch=batch_id, part=s['part']))
        st['recent_fills'] = st['recent_fills'][:300]
    n_done = sum(1 for v in st['series'].values() if v.get('status') in ('applied', 'done'))
    t['series_done'] = n_done
    t['series_rejected'] = sum(1 for v in st['series'].values() if v.get('status') == 'rejected')
    t['series_held'] = sum(1 for v in st['series'].values() if v.get('status') == 'hold')


# ---------------------------------------------------------------- status
def cmd_status(st):
    print(f"state: {PIPELINE_STATE}")
    print(f"loop: {st['loop'].get('status')}  current={st['loop'].get('current_batch')}  "
          f"paused_reason={st['loop'].get('paused_reason')}")
    b = st.get('baseline')
    if b:
        print(f"baseline: HEAD={b.get('git_head')} SHA={b.get('wb_sha')} "
              f"CRITICAL={b['verify'].get('criticalTotal')} WARN={b['verify'].get('warnTotal')}")
    print('totals:', json.dumps(st.get('totals', {}), ensure_ascii=False))
    print(f"batches: {len(st.get('batches', []))}  recent_fills: {len(st.get('recent_fills', []))}")
    for blk in st.get('blockers', [])[-3:]:
        print('blocker:', blk.get('batch'), blk.get('problems'))


# ---------------------------------------------------------------- 进度与 ETA
DONE_STATUS = ('done', 'rejected', 'hold', 'applied', 'rolled_back')
# 注意:状态词表统一为 decide_series 的返回值(accept/reject/hold)+ 应用侧
# (applied/done/rolled_back)。2026-09-22 曾因跳过名单写成 'held'(decide 写 'hold')
# 导致被 hold 车系永远跳不掉、循环反复选同一批——已修正,勿再改词形。


def update_progress(st, batch_size=5):
    """计算实时进度与预计完成时间,写入 st['progress']。

    ETA 模型(假设明示,便于事后校正):
      - apply 侧耗时:取本会话已应用批次的实测平均值(含自检)
      - 抽取侧耗时:每批 EXTRACT_SEC_PER_BATCH 秒(仓库配方:5 车系/批,
        每页 1.5~3s 等待 + 两次导航 + 180s 限额内),可用环境变量覆盖
      - 剩余批次 = 剩余车系 / 每批车系数(向上取整)
    """
    import math
    q = json.load(open(QUEUE_FILE, encoding='utf-8')) if os.path.exists(QUEUE_FILE) else {'entries': []}
    total = len(q['entries'])
    done = sum(1 for e in q['entries'] if st['series'].get(e['qkey'], {}).get('status') in DONE_STATUS)
    remaining = total - done
    applied = [b for b in st.get('batches', []) if b.get('duration_sec')]
    avg_apply = (sum(b['duration_sec'] for b in applied) / len(applied)) if applied else None
    try:
        extract_sec = float(os.environ.get('EXTRACT_SEC_PER_BATCH', '90'))
    except Exception:
        extract_sec = 90.0
    per_batch = (avg_apply or 20.0) + extract_sec
    batches_left = math.ceil(remaining / batch_size) if batch_size else 0
    eta_sec = batches_left * per_batch
    st['progress'] = {
        'updated_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'series_total': total, 'series_done': done, 'series_remaining': remaining,
        'pct': round(done / total * 100, 1) if total else 0.0,
        'gap_cells_total': q.get('total_gap_cells', 0),
        'applied_batches': sum(1 for b in st.get('batches', []) if b.get('outcome') == 'applied'),
        'avg_apply_sec_per_batch': round(avg_apply, 1) if avg_apply else None,
        'assumed_extract_sec_per_batch': extract_sec,
        'assumed_batch_size': batch_size,
        'batches_remaining': batches_left,
        'eta_seconds': round(eta_sec),
        'eta_human': f'{eta_sec / 3600:.1f} 小时' if eta_sec >= 3600 else f'{eta_sec / 60:.0f} 分钟',
        'eta_finish_at': (datetime.datetime.now() + datetime.timedelta(seconds=eta_sec)).strftime('%Y-%m-%d %H:%M'),
        'note': ('apply 耗时为全部已计时批次实测均值(含无可填批);真正写库的批需另加底表备份+保存约 10s。'
                 '抽取耗时按仓库配方假设(5 车系/批、每页 1.5~3s、180s 限额),'
                 '可用 EXTRACT_SEC_PER_BATCH 覆盖;前 3 个真实批次后应以实测值校正'),
    }
    return st['progress']


def cmd_resume(st):
    """人工处理完回退原因后,解除暂停。"""
    if st['loop'].get('status') != 'paused':
        print(f"当前状态是 {st['loop'].get('status')},无需 resume")
        return 0
    reason = st['loop'].get('paused_reason')
    st['loop']['status'] = 'idle'
    st['loop']['paused_reason'] = None
    save_state(st)
    print(f"已解除暂停(原因曾是: {reason})")
    print('建议先跑一次 apply --batch <该批> --dry-run 确认,再继续循环')
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('cmd', choices=['baseline', 'next-batch', 'decide', 'apply', 'status', 'resume', 'eta'])
    ap.add_argument('--batch')
    ap.add_argument('--n', type=int, default=5)
    ap.add_argument('--from-mapping', action='store_true',
                    help='只取已有人审映射且 DCD 产物在场的车系,直接按人审结论定链(无 Tabbit 时跑通用)')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--force', action='store_true')
    ap.add_argument('--allow-t4-redfix', action='store_true')
    ap.add_argument('--include-consumption', action='store_true')
    a = ap.parse_args()
    st = load_state()
    if a.cmd == 'baseline':
        return cmd_baseline(st)
    if a.cmd == 'next-batch':
        return cmd_next_batch(st, a.n, a.from_mapping)
    if a.cmd == 'decide':
        if not a.batch:
            print('decide 需要 --batch'); return 1
        return cmd_decide(st, a.batch)
    if a.cmd == 'apply':
        if not a.batch:
            print('apply 需要 --batch'); return 1
        return cmd_apply(st, a.batch, a.dry_run, a.force, a.allow_t4_redfix,
                         a.include_consumption)
    if a.cmd == 'status':
        cmd_status(st); return 0
    if a.cmd == 'resume':
        return cmd_resume(st)
    if a.cmd == 'eta':
        p = update_progress(st, a.n)
        save_state(st)
        print(json.dumps(p, ensure_ascii=False, indent=1))
        return 0


if __name__ == '__main__':
    sys.exit(main())
