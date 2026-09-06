# -*- coding: utf-8 -*-
"""批次监控：检查工信部 410批正式公告 / 411批公示 / 购置税目录第三十四批 是否发布。

数据源（研究_批次跟踪_20260906.md §5.3，已验证可用）：
  1. 工信部主站站内搜索接口（索引有效，409批可命中作对照）
  2. 装备中心「公告发布」栏目实时分页接口（dataproxy，绕过静态缓存）

用法：
  python check_new_batch.py            # 检查并在 Update/批次监控日志.md 追加一行
  python check_new_batch.py --no-log   # 只打印不落日志

返回码：0=无新批次；10=发现410批正式公告；11=发现411批公示（可叠加，取最大）
注意：运行前须清除 HTTPS_PROXY/HTTP_PROXY（本机代理失效，见总控计划）。
"""
import json, os, re, sys, argparse
from datetime import datetime
from urllib.parse import quote
from urllib.request import Request, urlopen, build_opener, ProxyHandler

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, 'Update', '批次监控日志.md')

for k in ('HTTPS_PROXY', 'HTTP_PROXY', 'https_proxy', 'http_proxy'):
    os.environ.pop(k, None)
OPENER = build_opener(ProxyHandler({}))  # 强制直连，不走任何代理
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)', 'Accept': 'application/json'}

def fetch(url, data=None, timeout=25):
    req = Request(url, data=data, headers=UA)
    with OPENER.open(req, timeout=timeout) as r:
        return r.read().decode('utf-8', 'replace')

def miit_search(q):
    """主站搜索接口：返回命中条数与首条(标题,日期,URL)。
    响应结构：data.searchResult.dataResults[].data{title_text, cdate(ms), url}"""
    url = ('https://www.miit.gov.cn/search-front-server/api/search/info'
           '?websiteid=110000000000000&q=' + quote(q) + '&p=1&pg=10')
    try:
        d = json.loads(fetch(url))
    except Exception as e:
        return None, f'接口异常: {e}'
    try:
        recs = d['data']['searchResult']['dataResults']
        total = d['data']['searchResult'].get('totalCount', len(recs))
    except Exception:
        # 命中为 0 时接口可能省略 searchResult 节点
        if isinstance(d, dict) and d.get('code') == '200':
            return 0, []
        return None, f'响应结构未知: {str(d)[:200]}'
    out = []
    for r in recs[:3]:
        item = r.get('data', r)
        title = re.sub('<[^>]+>', '', str(item.get('title_text') or item.get('title') or ''))
        cd = str(item.get('cdate') or '')
        date = datetime.fromtimestamp(int(cd) / 1000).strftime('%Y-%m-%d') if cd.isdigit() else cd[:10]
        out.append((title[:70], date, str(item.get('url') or '')))
    return int(total) if isinstance(total, int) else len(recs), out

def eidc_latest():
    """装备中心「公告发布」栏目最新条目"""
    url = ('https://www.miit-eidc.org.cn/module/web/jpage/dataproxy.jsp'
           '?startrecord=1&endrecord=45&perpage=15')
    form = ('col=1&webid=12&path=https%3A%2F%2Fwww.miit-eidc.org.cn%2F&columnid=1691'
            '&sourceContentType=1&unitid=4638&webname=%E5%B7%A5%E4%B8%9A%E5%92%8C%E4%BF%A1%E6%81%AF%E5%8C%96'
            '%E9%83%A8%E8%A3%85%E5%A4%87%E5%B7%A5%E4%B8%9A%E5%8F%91%E5%B1%95%E4%B8%AD%E5%BF%83&permissiontype=0').encode()
    try:
        html = fetch(url, data=form)
    except Exception as e:
        return [f'接口异常: {e}']
    titles = re.findall(r"title=['\"]?([^'\"<>]+)['\"]?", html)
    dates = re.findall(r'(2026-\d{2}-\d{2})', html)
    if not titles:
        return ['响应无条目: ' + html[:150]]
    return [f'{t.strip()} {d}' for t, d in zip(titles[:3], (dates + [''])[:3])]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--no-log', action='store_true')
    args = ap.parse_args()
    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    n410, r410 = miit_search('第410批')
    n409, r409 = miit_search('第409批')   # 对照组：索引有效性
    n411, r411 = miit_search('第411批')
    eidc = eidc_latest()

    print(f'== 批次监控 {now} ==')
    print(f'主站搜索「第410批」: {n410} 条命中')
    for t in (r410 if isinstance(r410, list) else []): print('   -', t)
    print(f'主站搜索「第409批」(对照): {n409} 条命中')
    print(f'主站搜索「第411批」: {n411} 条命中')
    print('装备中心「公告发布」最新:')
    for t in eidc: print('   -', t)

    hit410 = isinstance(n410, int) and n410 > 0
    hit411 = isinstance(n411, int) and n411 > 0
    code = 0
    if hit410: code = max(code, 10)
    if hit411: code = max(code, 11)
    print(f'结论: {"发现410批正式公告线索!" if hit410 else "410批正式公告未发布"} '
          f'{"发现411批公示线索!" if hit411 else "411批公示未发布"} (exit={code})')

    if not args.no_log:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        if not os.path.exists(LOG):
            with open(LOG, 'w', encoding='utf-8') as f:
                f.write('# 批次监控日志\n\n| 时间 | 410批正式 | 411批公示 | 对照(409命中) | 装备中心最新 |\n|---|---|---|---|---|\n')
        with open(LOG, 'a', encoding='utf-8') as f:
            f.write(f'| {now} | {"命中("+str(n410)+")" if hit410 else "未发布"} '
                    f'| {"命中("+str(n411)+")" if hit411 else "未发布"} '
                    f'| {n409 if isinstance(n409,int) else "接口异常"} '
                    f'| {eidc[0][:60] if eidc else "-"} |\n')
    sys.exit(code)

if __name__ == '__main__':
    main()
