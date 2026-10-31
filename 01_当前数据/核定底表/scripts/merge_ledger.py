# -*- coding: utf-8 -*-
"""车长核定台账合并器：官方锚点(A) + 汽车之家(B) + 15 迁移种子 → car_len_master.json。

优先级：411 正式官方 > 412 公示官方 > 汽车之家在售 > 种子。
人工仲裁通道：data/car_len_overrides.json（group|series → {len,wb,conf,src,url,note}），
最高优先级，本脚本只应用不生成——保证人工判断有留痕、可复跑。

分级：A=公告官方 art 页；B=汽车之家参数（权威媒体）；C=单一来源待复核；null=未核定不猜。
代际多值：汽车之家长度与公告众数不一致、但两者都出现在汽车之家款型矩阵 → 属换代真实差异，
记 代际多值，primary 取在售新款。
"""
import collections
import json
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import DATA, CAR_LEN_JSON
from fetch_official import CAR_LEN_SEED, OFFICIAL_JSON

EVID = os.path.join(DATA, 'evidence_auto.json')
OVERRIDES = os.path.join(DATA, 'car_len_overrides.json')
WORKLIST = os.path.join(DATA, 'coverage_worklist.json')
REVIEW_MD = os.path.join(DATA, 'car_len_review.md')


# 泛名守卫：通用名称只有品牌词的兜底车系（如 23 行「比亚迪」），挂任何单车系证据都是错的
GENERIC_SERIES = {
    '比亚迪', '长安', '吉利', '奇瑞', '五菱', '长城', '上汽', '广汽', '传祺', '埃安',
    '哈弗', '欧拉', '荣威', '名爵', '智己', '飞凡', '蔚来', '小鹏', '理想', '小米',
    '特斯拉', '腾势', '深蓝', '启源', '星途', '捷途', '智界', '仰望', '方程豹', '领克',
    '极氪', '魏牌', '坦克', '几何', '银河', '凯翼', '大通', '宝骏', '未命名',
}


def mode_of(vals):
    c = collections.Counter(vals)
    return c.most_common(1)[0][0] if c else None


def load_evidence():
    """evidence_auto.json 优先；未汇总时直接扫缓存目录（支持中途试跑）。"""
    if os.path.exists(EVID):
        return json.load(open(EVID, encoding='utf-8'))
    cache = os.path.join(DATA, 'evidence_cache', 'autohome')
    if not os.path.isdir(cache):
        return {}
    out = {}
    for fn in os.listdir(cache):
        try:
            d = json.load(open(os.path.join(cache, fn), encoding='utf-8'))
            out[f"{d.get('group')}|{d.get('series')}"] = d
        except Exception:
            continue
    return out


def main():
    anchors = json.load(open(OFFICIAL_JSON, encoding='utf-8')) if os.path.exists(OFFICIAL_JSON) else {}
    evid = load_evidence()
    work = json.load(open(WORKLIST, encoding='utf-8')) if os.path.exists(WORKLIST) else []
    ovr = json.load(open(OVERRIDES, encoding='utf-8')) if os.path.exists(OVERRIDES) else {}

    ledger, review = {}, []
    keys = {f"{w['group']}|{w['series']}" for w in work}

    def add(group, series):
        k = f'{group}|{series}'
        if k in ledger:
            return
        a = anchors.get(k) or []
        e = evid.get(k) or {} if series not in GENERIC_SERIES else {}
        if series in GENERIC_SERIES and not a:
            ledger[k] = {'group': group, 'series': series, 'len': None, 'wb': None,
                         'conf': None, 'src': '未核定(泛名兜底)', 'url': None,
                         'note': '通用名称仅为品牌词，无法归属具体车系', 'alts': [],
                         'cross': '泛名', 'verified': str(date.today())}
            return
        o411 = [x for x in a if x.get('status') == '正式' and x.get('len') is not None]
        o412 = [x for x in a if x.get('status') == '公示' and x.get('len') is not None]
        dims = e.get('dims') or []
        on_lens = []          # 在售款型长度（dims 顺序=在售优先）
        all_lens = [d['len'] for d in dims if d.get('len')]
        primary_wb = None
        seed = CAR_LEN_SEED.get(series)

        entry = {'group': group, 'series': series, 'len': None, 'wb': None,
                 'conf': None, 'src': None, 'url': None, 'note': '',
                 'alts': [], 'cross': None, 'verified': str(date.today())}

        # ── primary 选择 ──
        if o411:
            entry.update(len=mode_of([x['len'] for x in o411]),
                         wb=mode_of([x['wb'] for x in o411 if x.get('wb') is not None]),
                         conf='A', src='官方411公示', url=o411[0].get('url'))
        elif o412:
            entry.update(len=mode_of([x['len'] for x in o412]),
                         wb=mode_of([x['wb'] for x in o412 if x.get('wb') is not None]),
                         conf='A', src='官方412公示(公示)', url=o412[0].get('url'))
        elif dims:
            d0 = next((x for x in dims if x.get('len')), None)   # 首个带长度的款型
            if d0:
                entry.update(len=d0.get('len'), wb=d0.get('wb'), conf='B',
                             src='汽车之家在售', url=e.get('url'))
            primary_wb = d0.get('wb') if d0 else None
        elif seed:
            L, c, note = seed
            entry.update(len=L, conf=c, src='15迁移种子', note=note)
        if entry['len'] is None:
            entry['conf'] = None
            entry['src'] = '未核定'

        # ── 交叉核 ──
        srcs = {'官方411': sorted({x['len'] for x in o411 if x.get('len') is not None}),
                '官方412': sorted({x['len'] for x in o412 if x.get('len') is not None}),
                '汽车之家': sorted({x for x in all_lens if x is not None}),
                '种子': [seed[0]] if seed else []}
        uniq = sorted({v for vals in srcs.values() for v in vals if v is not None})
        if len(uniq) <= 1:
            entry['cross'] = '一致' if uniq else '无值'
        else:
            pub = [w for w in work if w['group'] == group and w['series'] == series]
            pub_mode = None
            if pub and pub[0].get('lens'):
                pub_mode = pub[0]['lens'][0][0]      # 公告众数（按行数降序第一）
            if (entry['len'] in all_lens and pub_mode in all_lens
                    and entry['len'] != pub_mode):
                entry['cross'] = f'代际多值(在售{entry["len"]:.0f}/公告众数{pub_mode:.0f})'
            else:
                entry['cross'] = '分歧'
                review.append({'group': group, 'series': series,
                               'chosen': entry['len'], 'src': entry['src'],
                               'pub_mode': pub_mode, **srcs})
        entry['alts'] = sorted({int(v) for v in uniq if v is not None and v != entry['len']})[:6]
        ledger[k] = entry

    # 全范围 + 官方锚点里有但不在范围的
    for w in work:
        add(w['group'], w['series'])
    for k in anchors:
        if k not in ledger:
            g, s = k.split('|', 1)
            add(g, s)

    # ── 人工仲裁覆盖 ──
    n_ovr = 0
    for k, v in ovr.items():
        if k in ledger:
            ledger[k].update(len=v.get('len'), wb=v.get('wb'), conf=v.get('conf', 'B'),
                             src=v.get('src', '人工仲裁'), url=v.get('url'),
                             note=v.get('note', ''), cross='人工仲裁')
            n_ovr += 1

    json.dump(ledger, open(CAR_LEN_JSON, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    cnt = collections.Counter((e['conf'], e['src']) for e in ledger.values())
    L = ['# 车长核定台账 · 分歧仲裁清单', f'', f'运行：{date.today()} · 台账 {len(ledger)} 系 · '
         f'人工仲裁 {n_ovr} 条', '',
         '## 覆盖统计', '']
    for (conf, src), n in sorted(cnt.items(), key=lambda kv: -kv[1]):
        L.append(f'- {conf} / {src}：{n} 系')
    L += ['', '## 分歧清单（需仲裁）', '']
    for r in review:
        if f"{r['group']}|{r['series']}" in ovr:
            continue
        L.append(f"- **[{r['group']}] {r['series']}** 选定 {r['chosen']} ({r['src']}) "
                 f"公告众数 {r['pub_mode']} | 411:{r['官方411']} 412:{r['官方412']} "
                 f"之家:{r['汽车之家']} 种子:{r['种子']}")
    open(REVIEW_MD, 'w', encoding='utf-8').write('\n'.join(L))
    print('台账', len(ledger), '系 | 分歧', len(review), '| 仲裁', n_ovr, '→', CAR_LEN_JSON)


if __name__ == '__main__':
    main()
