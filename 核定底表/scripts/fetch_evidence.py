# -*- coding: utf-8 -*-
"""Track B 自动取证：汽车之家参数通道（Phase 2b）。

通道（17 media_fill 已验证，AGENTS 2026-09-09/09-17/09-27 经验全部内化）：
- suggest → wordtype==3 车系卡，wordid=seriesid（int 强校验）
- listSpec → 在售(list)+停售(otherlist) 款型
- getParamConf → 参数矩阵（长度/宽度/高度/轴距）
- 品牌冲突守卫：候选车系品牌词与行板块品牌词冲突 → 拒绝（防泛化名撞名）
- 限流：1.2s/请求；paramConf 空返回 → 冷却 30s 重试一次（09-09 静默限流教训）

证据分级：汽车之家参数页 = B（权威媒体）；与公告众数一致 → cross_check=match（双源）。
断点续跑：evidence_cache/ 下逐系缓存。

用法：python fetch_evidence.py            # 全量（跳过已缓存）
      python fetch_evidence.py --group byd
      python fetch_evidence.py --retry-miss
"""
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import ssl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import DATA

CACHE = os.path.join(DATA, 'evidence_cache', 'autohome')
OUT = os.path.join(DATA, 'evidence_auto.json')
WORKLIST = os.path.join(DATA, 'coverage_worklist.json')

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
CTX = ssl.create_default_context()
SLEEP = 1.2
PARAM_KEYS = ('长度(mm)', '宽度(mm)', '高度(mm)', '轴距(mm)', '长*宽*高(mm)')

# 拉丁板块名 → 中文（suggest 查询词修复用）
LATIN_BOARD = {'STERRA': '星纪元', 'HYPER': '昊铂', 'ORA': '欧拉', 'ICAR': 'iCAR',
               'AION': 'AION', 'TIGGO': '瑞虎', 'ARRIZO': '艾瑞泽'}
# 个别车系 suggest 查询的专属提示（已逐一实测验证 2026-10-25）
QUERY_HINT = {
    '糯玉米': '长安Lumin', '小蚂蚁': '小蚂蚁', '熊猫MINI': '熊猫mini',
    '五菱': '五菱宏光MINIEV', '宏光': '宏光MINIEV', '仰望U9': '仰望U9',
    '奔奔E-STAR': '奔奔E-Star', '奔奔E-STAR国民版': '奔奔E-Star',
    '之星9': '长安之星9', '大拿': '五菱红大拿',
    '吉利M9': '银河M9', '启源A06': '长安启源A06', '长安G318': '深蓝G318',
    '长安V919': '长安凯程V919', '摩卡PHEV': '魏牌 摩卡', '山海L6C-DM': '捷途山海L6',
    '极星2': 'Polestar 2', '元EV': '比亚迪元',
    # 第二批（2026-10-25 suggest 逐条实测）
    '吉利E5': '银河E5', '吉利E8': '银河E8', '吉利A7': '银河A7 EM', '09PHEV': '领克09',
    'XC60T8': '沃尔沃XC60', '名爵MG07': 'MG7', '名爵MG4EV': 'MG4', '名爵MG4X': 'MG 4X',
    'IMAX8DMH': '荣威iMAX8新能源', '大通EV90': '新途EV90', '之光EV': '五菱之光',
    '埃安HYPERGT': '昊铂GT', 'HYCANZ03': '合创Z03', 'HYCANV09': '合创V09',
    'AIONSPIO': 'AION S Plus', '启源A05': '长安启源A05经典', '启源A05L': '长安启源A05L',
    '启源A07': '长安启源A07', '启源Q05': '长安启源Q05',
    'ARRIZO8C-DM': '艾瑞泽8', 'TIGGO8C-DM': '瑞虎8',
    '小鹏M03': '小鹏MONA M03', 'iCAR03': 'iCAR 03', 'STERRAET': '星纪元 ET',
    '星耀7': '银河星耀7', '埃安RT': 'AION RT', '熊猫MINI': '熊猫',
    '坦克500': '坦克500', '坦克400': '坦克400', '风云T9C-DM': '风云T9',
    '山海T2': '捷途山海T2', '山海L7PLUSC-DM': '捷途山海L7', '领克Z20': '领克Z20',
}
# 板块品牌词（冲突守卫参照）
BOARD_TOKENS = {
    'byd': ['比亚迪', '腾势', '方程豹', '仰望', '领汇'], 'nio': ['蔚来', '乐道', '萤火虫'],
    'xiaomi': ['小米'], 'tesla': ['特斯拉'], 'liauto': ['理想'], 'xpeng': ['小鹏'],
    'gwm': ['长城', '哈弗', '魏牌', '坦克', '欧拉', '沙龙'],
    'changan': ['长安', '深蓝', '启源', '阿维塔'],
    'gac': ['广汽', '埃安', '传祺', '昊铂', '合创', '启境'],
    'chery': ['奇瑞', '星途', '捷途', 'icar', '智界', '开瑞', '星纪元'],
    'geely': ['吉利', '几何', '帝豪', '领克', '极氪', '沃尔沃', 'smart', '极星', '远程',
              '雷达', '翼真', '英伦', '枫叶', '路特斯', '极越', '睿蓝', '银河'],
    'saic': ['上汽', '荣威', '名爵', '五菱', '宝骏', '大通', '飞凡', '智己'],
}
ALL_TOKENS = sorted({t for v in BOARD_TOKENS.values() for t in v}, key=len, reverse=True)


def get(u, timeout=25):
    req = urllib.request.Request(u, headers={"User-Agent": UA, "Accept": "application/json"})
    return json.loads(urllib.request.urlopen(req, timeout=timeout, context=CTX).read().decode("utf-8", "ignore"))


def suggest(q):
    d = get("https://sou.api.autohome.com.cn/sug/_suggest?plat=pc&uid=0&q="
            + urllib.parse.quote(q))
    return [(str(x.get("key") or ""), x.get("wordid"))
            for x in (d.get("result") or {}).get("data", []) or []
            if x.get("wordtype") == 3]


def brand_tokens_of(s):
    n = str(s or '').lower().replace(' ', '')
    return [b for b in ALL_TOKENS if b.lower() in n]


def conflict(group, cand_key):
    """候选车系名与行板块品牌词冲突判定（跨品牌=必错）。"""
    btoks = set(BOARD_TOKENS.get(group, []))
    stoks = set(brand_tokens_of(cand_key))
    if not btoks or not stoks:
        return False
    return not (btoks & stoks)


def queries_for(series, board):
    """suggest 查询词序列。"""
    qs = []
    if series in QUERY_HINT:
        qs.append(QUERY_HINT[series])
    qs.append(series)
    for lat, cn in LATIN_BOARD.items():
        if series.upper().startswith(lat):
            qs.append(series.upper().replace(lat, cn, 1))
    if board and board not in ('其他',):
        qs.append(f'{board}{series}')
    seen, out = set(), []
    for q in qs:
        if q and q.lower() not in seen:
            seen.add(q.lower())
            out.append(q)
    return out


def norm_key(s):
    return re.sub(r'[\s\u200b\u200c\u200d\ufeff\-·]', '', str(s or '')).upper()


def pick_candidate(cands, series, group):
    """挑与车系名重叠度最高且无品牌冲突的候选。"""
    sk = norm_key(series)
    best = None
    for key, wid in cands:
        try:
            wid = int(wid)
        except (TypeError, ValueError):
            continue
        if conflict(group, key):
            continue
        ck = norm_key(key)
        if ck == sk or (sk in ck) or (ck in sk and len(ck) >= 2):
            best = (key, wid)
            break
        if best is None:
            best = (key, wid)          # 无重叠时先记一个，最后再兜底
    return best


def list_specs(sid):
    r = (get(f"https://car-web-api.autohome.com.cn/car/spec/listSpec"
             f"?type=0x001f&from=1&pm=1&pluginversion=11.64.8&seriesid={sid}").get("result") or {})
    on, off = [], []
    for grp in (r.get("list") or []):
        for sp in (grp.get("speclist") or []):
            on.append((sp.get("id"), sp.get("name")))
    for grp in (r.get("otherlist") or []):
        for sp in (grp.get("speclist") or []):
            off.append((sp.get("id"), sp.get("name")))
    return (r.get("seriesname"), on, off)


def dims(sid, spid):
    """参数矩阵 → {长度, 宽度, 高度, 轴距}。空返回重试一次（限流冷却）。"""
    for attempt in (1, 2):
        try:
            r = (get(f"https://car-web-api.autohome.com.cn/car/param/getParamConf"
                     f"?mode=1&site=2&seriesid={sid}&specid={spid}").get("result") or {})
        except Exception:
            time.sleep(30)
            continue
        if not r.get("datalist"):
            time.sleep(30)
            continue
        tm = {}
        for grp in r.get("titlelist") or []:
            for it in grp.get("items") or []:
                tm[it.get("titleid")] = it.get("itemname")
        vals = {}
        for pc in (r.get("datalist") or [{}])[0].get("paramconflist") or []:
            n = tm.get(pc.get("titleid"))
            if n in PARAM_KEYS:
                v = pc["sublist"][0].get("name") if pc.get("sublist") else pc.get("itemname")
                vals[n] = v
        if vals:
            return vals
    return {}


def num_of(v):
    m = re.search(r'-?\d+(?:\.\d+)?', str(v or ''))
    return float(m.group(0)) if m else None


def fetch_one(series, board, group):
    """单系取证：返回缓存 dict。"""
    chosen = None
    for q in queries_for(series, board):
        cands = suggest(q)
        time.sleep(SLEEP)
        if cands:
            chosen = pick_candidate(cands, series, group)
            if chosen:
                break
    if not chosen:
        return {'status': 'no_candidate', 'queries': queries_for(series, board)}
    key, sid = chosen
    name, on, off = list_specs(sid)
    time.sleep(SLEEP)
    specs = (on[:3] + off[:2]) or []
    if on:
        specs = [on[0]] + [s for s in off[:2] if s] + on[1:3]
    rows = []
    for spid, spname in specs:
        v = dims(sid, spid)
        time.sleep(SLEEP)
        if v.get('长度(mm)'):
            rows.append({'spec': spname, 'len': num_of(v.get('长度(mm)')),
                         'wb': num_of(v.get('轴距(mm)')),
                         'width': num_of(v.get('宽度(mm)')),
                         'height': num_of(v.get('高度(mm)'))})
    return {'status': 'ok' if rows else 'no_dims', 'seriesid': sid, 'ah_key': key,
            'ah_name': name, 'n_on': len(on), 'n_off': len(off),
            'url': f'https://car.autohome.com.cn/config/series/{sid}.html',
            'dims': rows}


def main():
    only_group = sys.argv[sys.argv.index('--group') + 1] if '--group' in sys.argv else None
    retry_miss = '--retry-miss' in sys.argv
    only = None
    if '--only' in sys.argv:
        only = set(sys.argv[sys.argv.index('--only') + 1].split(','))
    os.makedirs(CACHE, exist_ok=True)
    work = json.load(open(WORKLIST, encoding='utf-8'))
    if only_group:
        work = [w for w in work if w['group'] == only_group]
    if only:
        work = [w for w in work if f"{w['group']}|{w['series']}" in only]
    n = {'ok': 0, 'no_candidate': 0, 'no_dims': 0, 'cached': 0}
    for i, w in enumerate(work):
        fp = os.path.join(CACHE, f"{w['group']}_{w['series']}.json".replace('/', '_'))
        if os.path.exists(fp) and not retry_miss:
            old = json.load(open(fp, encoding='utf-8'))
            if old.get('status') == 'ok':
                n['cached'] += 1
                continue
        try:
            r = fetch_one(w['series'], w['board'], w['group'])
        except Exception as e:
            r = {'status': f'err:{type(e).__name__}'}
            time.sleep(5)
        r['group'], r['series'], r['tier'] = w['group'], w['series'], w['tier']
        json.dump(r, open(fp, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        st = r['status'] if r['status'] == 'ok' else (
            'no_candidate' if r['status'] == 'no_candidate' else
            'no_dims' if r['status'] == 'no_dims' else 'err')
        n[st] = n.get(st, 0) + 1
        if (i + 1) % 20 == 0:
            print(f'  [{i + 1}/{len(work)}] {n}')
    # 汇总
    out = {}
    for fn in os.listdir(CACHE):
        d = json.load(open(os.path.join(CACHE, fn), encoding='utf-8'))
        out[f"{d.get('group')}|{d.get('series')}"] = d
    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('DONE', n, '→', OUT)


if __name__ == '__main__':
    main()
