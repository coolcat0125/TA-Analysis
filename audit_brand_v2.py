# -*- coding: utf-8 -*-
"""audit_brand_v2.py — 媒体填充行品牌一致性审计（离线优先，v2 限流加固）

原理：
  1. 从底表 col34(媒体校验来源) 提取 autohome seriesId（480 个去重）
  2. listSpec 逐个取车系名（0.55s 限速 + ERR 退避重试 + 断点续跑缓存）
  3. 车系名含品牌词的：与行商标(产品商标+企业名称)品牌词比对，含别名表
  4. 车系名不含品牌词的：记 unresolvable（离线无法定性，不判错）
输出：audit-output/media_brand_audit_v2.json
"""
import json
import os
import re
import sys
import time

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import media_fill as mf  # noqa: E402

CACHE = os.path.join(HERE, 'audit-output', '_media_series_names.json')
OUT = os.path.join(HERE, 'audit-output', 'media_brand_audit_v2.json')

BRANDS = ['比亚迪', '特斯拉', '北汽', '北京汽车制造厂', '广汽埃安', '广汽', '埃安', 'aion',
          '长安', '深蓝', '阿维塔', '启源', '奇瑞', '吉利', '银河', '极氪', 'zeekr', '几何',
          '领克', '长城', '欧拉', '魏牌', '坦克', '红旗', '一汽', '奔腾', '上汽', '荣威',
          '名爵', '东风', '岚图', '猛士', '纳米', '奕派', '风神', '极狐', '蔚来', '乐道',
          '萤火虫', '小鹏', '理想', '零跑', '哪吒', '腾势', '方程豹', '仰望', '问界',
          '智界', '享界', '尊界', '小米', '五菱', '宝骏', '大众', '丰田', '本田', '日产',
          '别克', '奥迪', '奔驰', '宝马', '沃尔沃', '起亚', '现代', '福特', '大通', '依维柯',
          '赛力斯', '合众', 'smart', '路特斯', 'lotus', '极石', '远航', '大运', '创维',
          'skyworth', 'aion', '高合', 'hiphi', '摩登', '雷丁', '恒驰', '爱驰', '天际',
          '威马', '云度', '电咖', '新特', '国机智骏', '江铃', '江淮', '思皓', '钇为',
          '海马', '众泰', '江南', '大乘', '裕路', '国新', '中期', '东风日产', '东风本田',
          '广汽丰田', '广汽本田', '华晨', '金康', '赛力斯', '蓝电', '睿蓝', '枫叶', '曹操']
ALIAS = {('北京', '北汽'), ('北京汽车制造厂', '北汽'), ('广汽', '埃安'), ('广汽', 'aion'),
         ('广汽埃安', 'aion'), ('广汽埃安', '埃安'), ('上汽', '荣威'), ('上汽', '名爵'),
         ('上汽', '大通'), ('上汽', '五菱'), ('一汽', '奔腾'), ('吉利', '银河'), ('吉利', '极氪'),
         ('吉利', 'zeekr'), ('吉利', '几何'), ('吉利', '领克'), ('吉利', '睿蓝'), ('长城', '欧拉'),
         ('长城', '魏牌'), ('长城', '坦克'), ('东风', '岚图'), ('东风', '猛士'), ('东风', '纳米'),
         ('东风', '奕派'), ('东风', '风神'), ('东风', '日产'), ('东风', '本田'), ('东风', '启辰'),
         ('比亚迪', '腾势'), ('比亚迪', '方程豹'), ('比亚迪', '仰望'), ('赛力斯', '问界'),
         ('奇瑞', '智界'), ('北汽', '极狐'), ('奇瑞', '享界'), ('奇瑞', '星途'), ('五菱', '宝骏'),
         ('吉利', '曹操'), ('吉利', '枫叶'), ('江淮', '思皓'), ('江淮', '钇为'), ('江铃', '江淮')}


def brand_tokens(s):
    n = str(s or '').lower().replace(' ', '')
    hits = [b for b in BRANDS if b in n]
    return hits


def main():
    wb = openpyxl.load_workbook(os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx'),
                                read_only=True, data_only=True)
    cr = list(wb['NEV公告参数汇总'].iter_rows(min_row=2, values_only=True))
    wb.close()
    sid_rows = {}
    for r in cr:
        src = str(r[34] or '')
        for m in re.findall(r'series/(\d+)', src):
            sid_rows.setdefault(m, []).append(r)
    print(f'rows_with_media={sum(len(v) for v in sid_rows.values())} distinct={len(sid_rows)}')

    names = json.load(open(CACHE, encoding='utf-8')) if os.path.exists(CACHE) else {}
    errs = 0
    for i, (sid, rows) in enumerate(sorted(sid_rows.items())):
        if sid in names and names[sid] not in ('ERR',):
            continue
        ok = False
        for attempt in range(3):
            try:
                info = mf.list_spec(sid)
                nm = info.get('seriesname')
                if nm:
                    names[sid] = nm
                    ok = True
                break
            except Exception as e:
                errs += 1
                wait = 8 * (attempt + 1)
                print(f'  ERR sid={sid} attempt{attempt+1}: {str(e)[:60]} -> wait {wait}s', flush=True)
                time.sleep(wait)
        if not ok:
            names[sid] = 'ERR'
        if (i + 1) % 40 == 0:
            json.dump(names, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False)
            print(f'  progress {i+1}/{len(sid_rows)} cached', flush=True)
        time.sleep(0.55)
    json.dump(names, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False)
    err_n = sum(1 for v in names.values() if v == 'ERR')
    print(f'series names resolved: {len(names) - err_n}/{len(names)} (ERR={err_n})')

    flags = []
    unresolvable = 0
    checked = 0
    for sid, rows in sid_rows.items():
        sname = names.get(sid, '')
        stoks = brand_tokens(sname)
        if not stoks:
            unresolvable += len(rows)
            continue
        for r in rows:
            checked += 1
            btoks = brand_tokens(str(r[2] or '') + ' ' + str(r[3] or ''))
            if not btoks:
                continue
            conflict = all(
                bt != st and (bt, st) not in ALIAS and (st, bt) not in ALIAS
                for bt in btoks for st in stoks
            )
            if conflict:
                flags.append({'批次': r[0], '型号': r[1], '商标': str(r[2] or '')[:14],
                              '通用名称': str(r[6] or '')[:22], '车系': sname, 'sid': sid,
                              '行品牌词': btoks, '车系品牌词': stoks})
    print(f'checked={checked} unresolvable_rows={unresolvable} brand_conflicts={len(flags)}')
    json.dump({'checked': checked, 'unresolvableRows': unresolvable, 'conflicts': flags},
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('saved', OUT)
    for f in flags[:20]:
        print('  ', f['批次'], f['型号'], f['商标'], '|', f['通用名称'], '->', f['车系'])


if __name__ == '__main__':
    main()
