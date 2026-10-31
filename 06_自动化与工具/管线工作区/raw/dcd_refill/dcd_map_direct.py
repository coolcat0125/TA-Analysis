# -*- coding: utf-8 -*-
"""dcd_map_direct.py — 直连搜索映射 kw→sid（Wave 1-A）

对 batch_state.json 中未映射的批次车系，用直连搜索页（SSR searchData）取候选，
过品牌守卫（复用 media_fill.brand_conflict 词表）后回填 sid，并追加 series_mapping.json。
逐条落盘（断点续跑）；拒收留档 reason。

用法：
  python dcd_map_direct.py                      # 全部未映射批次车系
  python dcd_map_direct.py --batches r01,r02    # 指定批次
  python dcd_map_direct.py --limit 5            # 限速试跑
"""
import argparse
import io
import json
import os
import re
import sys
import time
import urllib.parse
from datetime import datetime

import openpyxl

import dcd_direct_fetch as d

HERE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(HERE, 'batch_state.json')
MAPPING = os.path.join(HERE, 'series_mapping.json')
_ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
sys.path.insert(0, _ROOT)
from media_fill import brand_conflict  # noqa: E402
_WB = [os.path.join(_ROOT, f) for f in os.listdir(_ROOT)
       if f.endswith('.xlsx') and '410' in f and '汇总' in f][0]

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def norm(s):
    return str(s or '').lower().replace(' ', '').replace('·', '')


def load_json(p, default):
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            return json.load(f)
    return default


def build_kw_brands():
    """通用名称(规范) → set(企业名称)，供品牌守卫。"""
    wb = openpyxl.load_workbook(_WB, read_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    hi = {h: i for i, h in enumerate(hdr) if h}
    out = {}
    for r in it:
        gen = str(r[hi['通用名称']] or '').strip()
        ent = str(r[hi['企业名称']] or '').strip()
        if not gen or not ent:
            continue
        for part in re.split(r'[,，、/]', gen):
            k = norm(part)
            if len(k) >= 2:
                out.setdefault(k, set()).add(ent)
    wb.close()
    return out


def _ala_series_cards(cell):
    """cell_type=26 ala 聚合卡（is_car_series_list=1）：display.serie[] 内嵌车系，
    sid 在 murl 的 cid= 参数，品牌在 display.url 的 brand_name= 参数。"""
    disp = cell.get('display') or {}
    out = []
    if not disp.get('is_car_series_list'):
        return out
    m = re.search(r'brand_name=([^&]+)', str(disp.get('url') or ''))
    brand = urllib.parse.unquote(m.group(1)) if m else ''
    for it in disp.get('serie') or []:
        if not isinstance(it, dict):
            continue
        mm = re.search(r'cid=(\d+)', str(it.get('murl') or ''))
        nm = str(it.get('name') or '')
        if mm and nm:
            out.append({'sid': mm.group(1), 'series_name': nm, 'brand_name': brand})
    return out


def search_candidates(ch, kw):
    from urllib.request import Request, urlopen
    q = urllib.parse.quote(kw)
    code, html = d.fetch(f'https://www.dongchedi.com/search?keyword={q}', ch)
    nd = d.parse_next_data(html) if code == 200 else None
    sd = ((nd or {}).get('props', {}).get('pageProps', {}) or {}).get('searchData') or {}
    cands, seen = [], set()
    for cell in sd.get('data') or []:
        if not isinstance(cell, dict):
            continue
        # ala 聚合卡：serie[] 内嵌车系列表（仰望U7 等查询的真实形态）
        if cell.get('cell_type') == 26:
            ala = _ala_series_cards(cell)
            if ala:
                for c in ala:
                    if c['sid'] not in seen:
                        seen.add(c['sid'])
                        cands.append(c)
                continue
        # 车系卡：cell_type=26，名字/品牌在 display 子对象
        disp = cell.get('display') or {}
        sid = cell.get('series_id')
        name = str(disp.get('series_name') or cell.get('series_name') or '')
        brand = str(disp.get('sub_brand_name') or disp.get('brand_name')
                    or cell.get('brand_name') or '')
        if sid is None or not name or str(sid) in seen:
            continue
        seen.add(str(sid))
        cands.append({'sid': str(sid), 'series_name': name, 'brand_name': brand})
    return cands


def pick(kw, brands, cands):
    """返回 (cand, score, note) 或 (None, None, reason)。"""
    def brand_pass(c):
        if not brands:
            return True
        series_key = c['brand_name'] + c['series_name']
        return not any(brand_conflict(b, series_key) for b in brands)

    scored = []
    for c in cands:
        nn = norm(c['series_name'])
        if not nn:
            continue
        nk = norm(kw)
        if nn == nk:
            s = 0
        elif nk in nn or nn in nk:
            s = 1
        else:
            s = 2
        scored.append((s, c))
    scored.sort(key=lambda x: (x[0], len(x[1]['series_name'])))
    # 唯一车系卡 + 品牌通过 → 直接收（拉丁名/中文名差异由 DCD 搜索背书）
    if len(scored) == 1 and brand_pass(scored[0][1]):
        return scored[0][1], scored[0][0], 'sole_candidate_brand_ok'
    for s, c in scored[:3]:
        if s >= 2:
            break
        if not brand_pass(c):
            continue
        return c, s, f'brand_ok({len(brands)}家)'
    if scored:
        return None, None, 'name_mismatch_or_brand_conflict'
    return None, None, 'no_candidate'


def variants_for(kw, brands):
    """生成消解变体：原名→去后缀(EV/Ultra/…)→去尾数字→品牌词+名。"""
    nk = norm(kw)
    outs = [nk]
    strip = re.sub(r'(ev|ultra|plus|pro|max|dmi|emi)+$', '', nk)
    if strip != nk and len(strip) >= 4:
        outs.append(strip)
    nodig = re.sub(r'\d+$', '', nk)
    if nodig != nk and len(nodig) >= 3:
        outs.append(nodig)
    btoks = []
    from media_fill import brand_tokens_of
    for b in brands:
        for bt in brand_tokens_of(b):
            if bt not in btoks:
                btoks.append(bt)
    for bt in btoks[:3]:
        outs.append(bt + nk)
        if nodig != nk and len(nodig) >= 3:
            outs.append(bt + nodig)
    seen, out = set(), []
    for v in outs:
        if v not in seen:
            seen.add(v)
            out.append(v)
    return out


def alias_pass(a, st, mapping, kw_brands, ch):
    """对拒收条目做变体消解：已映射变体→alias；否则逐变体直连搜索。"""
    wanted = set(a.batches.split(',')) if a.batches else None
    ok_sid = {}
    for m in mapping['reviewed']:
        if m.get('ok') and m.get('sid'):
            ok_sid[norm(m['kw'])] = str(m['sid'])
    jobs = []
    for bid, b in sorted(st['batches'].items()):
        if wanted and bid not in wanted:
            continue
        for s in b.get('series', []):
            if s.get('ok') and s.get('sid'):
                continue
            jobs.append((bid, s))
    if a.limit:
        jobs = jobs[:a.limit]
    print(f'alias 消解: {len(jobs)} 条')
    done = 0
    for i, (bid, s) in enumerate(jobs):
        kw = s['kw']
        brands = set()
        for k, vset in kw_brands.items():
            if k == norm(kw) or (len(k) > len(norm(kw)) and norm(kw) in k):
                brands |= vset
        got = None
        for v in variants_for(kw, brands):
            if v in ok_sid:
                got = ({'sid': ok_sid[v], 'series_name': v, 'brand_name': ''},
                       0, f'alias_of({v})')
                break
            try:
                cands = search_candidates(ch, v)
            except Exception as e:  # noqa: BLE001
                print(f'[{i+1}] {kw} 搜索失败({v}): {e}')
                continue
            cand, score, note = pick(kw, brands, cands)
            if cand:
                got = (cand, score, f'variant={v} {note}')
                break
            time.sleep(2)
        if got:
            cand, score, note = got
            s['sid'] = cand['sid']
            s['ok'] = True
            s.pop('reject_reason', None)
            mapping['reviewed'].append({
                'kw': kw, 'sid': cand['sid'], 'series': cand['series_name'],
                'ok': True, 'via': 'direct-alias', 'note': note,
                'time': datetime.now().strftime('%m-%d %H:%M')})
            ok_sid[norm(kw)] = str(cand['sid'])
            print(f'[{i+1}/{len(jobs)}] {bid} {kw} -> MAP({note}) '
                  f"{cand['series_name']} sid={cand['sid']}")
        else:
            print(f'[{i+1}/{len(jobs)}] {bid} {kw} -> 仍无源')
        with open(STATE, 'w', encoding='utf-8') as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
        with open(MAPPING, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=1)
        done += 1
        time.sleep(2)
    print(f'alias 完成 {done}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--batches', default='')
    ap.add_argument('--limit', type=int, default=0)
    ap.add_argument('--retry-rejected', action='store_true',
                    help='清除 wanted 批次内此前拒收记录并重试')
    ap.add_argument('--alias', action='store_true',
                    help='变体消解：后缀/数字剥离+品牌词组合+已映射别名')
    a = ap.parse_args()

    st = load_json(STATE, {'batches': {}})
    mapping = load_json(MAPPING, {'reviewed': [], 'field_map': {}})
    if a.alias:
        alias_pass(a, st, mapping, build_kw_brands(), d.load_cookie_header())
        return
    if a.retry_rejected:
        wanted_set = set(a.batches.split(',')) if a.batches else None
        rj_kw = set()
        for bid, b in st['batches'].items():
            if wanted_set and bid not in wanted_set:
                continue
            for s in b.get('series', []):
                if not s.get('ok') and not s.get('sid'):
                    rj_kw.add(norm(s['kw']))
        if rj_kw:
            mapping['reviewed'] = [m for m in mapping['reviewed']
                                   if m.get('ok') or norm(m['kw']) not in rj_kw]
    kw_brands = build_kw_brands()
    mapped_kw = {norm(m['kw']) for m in mapping['reviewed']}

    wanted = a.batches.split(',') if a.batches else sorted(st['batches'])
    jobs = []
    for bid in wanted:
        b = st['batches'].get(bid) or {}
        for s in b.get('series', []):
            if s.get('ok') and s.get('sid'):
                continue
            if norm(s['kw']) in mapped_kw:
                continue
            jobs.append((bid, s))
    if a.limit:
        jobs = jobs[:a.limit]
    print(f'待映射: {len(jobs)} 车系')

    ch = d.load_cookie_header()
    done = 0
    for i, (bid, s) in enumerate(jobs):
        kw = s['kw']
        try:
            cands = search_candidates(ch, kw)
        except Exception as e:  # noqa: BLE001
            print(f'[{i+1}] {kw} 搜索失败: {e}')
            time.sleep(5)
            continue
        brands = set()
        for k, vset in kw_brands.items():
            # 严格匹配：相等，或 kw 是更长通用名的一部分（防 'a5'⊂'ora5' 式污染）
            if k == norm(kw) or (len(k) > len(norm(kw)) and norm(kw) in k):
                brands |= vset
        cand, score, note = pick(kw, brands, cands)
        entry = {'kw': kw,
                 'sid': cand['sid'] if cand else '',
                 'series': cand['series_name'] if cand else
                           (cands[0]['series_name'] if cands else ''),
                 'ok': bool(cand),
                 'via': 'direct-search',
                 'time': datetime.now().strftime('%m-%d %H:%M')}
        if cand:
            entry['note'] = f"{note} score={score} brand={cand['brand_name']}"
            s['sid'] = cand['sid']
            s['ok'] = True
        else:
            entry['reason'] = note
            entry['candidates'] = cands[:3]
            s['ok'] = False
            s['reject_reason'] = note
        mapping['reviewed'].append(entry)
        with open(STATE, 'w', encoding='utf-8') as f:
            json.dump(st, f, ensure_ascii=False, indent=1)
        with open(MAPPING, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=1)
        flag = 'MAP' if cand else f'REJ({note})'
        print(f'[{i+1}/{len(jobs)}] {bid} {kw} -> {flag} '
              f"{cand['series_name'] if cand else ''} "
              f"候选={[c['series_name'] for c in cands[:3]]}")
        done += 1
        time.sleep(2)
    print(f'完成 {done}，映射落 series_mapping.json + batch_state.json')


if __name__ == '__main__':
    main()
