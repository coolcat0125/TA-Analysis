# -*- coding: utf-8 -*-
"""fetch_doc3_old_batches.py — 老批次（341~378）公告页发现+购置税目录下载转换

通道：MIIT 主站搜索（check_new_batch 同源，公开无滑块）→ 公告页 → 附件 .doc
识别：链接文本含「减免车辆购置税」的附件（多附件公告取购置税目录）；
下载限速 ≥1.5s；LibreOffice 显式 UTF8 过滤转换（09-08 教训）。
产出：../../02-数据底表/data/nev-announcements/raw/b{N}/taxfree.doc + taxfree.txt
断点：已存在 taxfree.txt 的批次跳过；逐批落盘 doc3_old_fetch_log.json。
用法：python fetch_doc3_old_batches.py [--batches 341,342,...] [--dry]
"""
import io
import json
import os
import re
import subprocess
import sys
import time
import urllib.parse
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join('..', '..', '02-数据底表', 'data', 'nev-announcements', 'raw')
LOG = os.path.join(HERE, 'doc3_old_fetch_log.json')
UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
BATCHES_DEFAULT = [341, 342, 343, 344, 345, 346, 348, 351, 354, 355, 356, 357,
                   359, 365, 366, 367, 368, 369, 370, 371, 372, 373, 374, 375, 376, 377, 378]


def get(url, timeout=25):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def search_page(batch):
    """MIIT 主站搜索（check_new_batch.miit_search 同源结构：searchResult.dataResults[].data）"""
    q = urllib.parse.quote(f'第{batch}批')
    u = ('https://www.miit.gov.cn/search-front-server/api/search/info'
         f'?websiteid=110000000000000&q={q}&p=1&pg=10')
    d = json.loads(get(u, 20))
    recs = (d.get('data', {}) or {}).get('searchResult', {}) or {}
    for r in recs.get('dataResults', []) or []:
        item = r.get('data', r)
        title = re.sub('<[^>]+>', '', str(item.get('title_text') or item.get('title') or ''))
        url = str(item.get('url') or '')
        if f'第{batch}批' in title and '生产企业及产品' in title:
            return urllib.parse.urljoin('https://www.miit.gov.cn/', url), title
    # 次选：公告类命中但标题变体
    for r in recs.get('dataResults', []) or []:
        item = r.get('data', r)
        title = re.sub('<[^>]+>', '', str(item.get('title_text') or item.get('title') or ''))
        url = str(item.get('url') or '')
        if f'第{batch}批' in title and '道路机动车辆' in title:
            return urllib.parse.urljoin('https://www.miit.gov.cn/', url), title
    return None, None


def find_tax_doc(page_url):
    """公告页内找购置税目录附件（.doc/.docx/.pdf）"""
    html = get(page_url, 25).decode('utf-8', 'ignore')
    links = re.findall(r'href="([^"]+\.(?:doc|docx|pdf))"[^>]*>([^<]*)<', html, re.I)
    for href, text in links:
        t = text.strip()
        if '购置税' in t or '购置税' in urllib.parse.unquote(href):
            url = href if href.startswith('http') else urllib.parse.urljoin(page_url, href)
            return url, re.sub(r'[\x00-\x1f]', '', t)
    # 兜底：附件区所有 doc
    for href, text in links:
        if '.pdf' not in href.lower():
            url = href if href.startswith('http') else urllib.parse.urljoin(page_url, href)
            return url, re.sub(r'[\x00-\x1f]', '', text.strip())
    return None, None


SOFFICE = r'C:\Program Files\LibreOffice\program\soffice.exe'


def convert_utf8(doc_path, txt_path):
    r = subprocess.run([SOFFICE, '--headless', '--convert-to',
                        'txt:Text (encoded):UTF8', '--outdir', os.path.dirname(doc_path), doc_path],
                       capture_output=True, timeout=120)
    base = os.path.splitext(os.path.basename(doc_path))[0] + '.txt'
    produced = os.path.join(os.path.dirname(doc_path), base)
    if os.path.exists(produced):
        if produced != txt_path:
            os.replace(produced, txt_path)
        return True
    return False


def main():
    batches = BATCHES_DEFAULT
    if '--batches' in sys.argv:
        batches = [int(x) for x in sys.argv[sys.argv.index('--batches') + 1].split(',')]
    log = {}
    if os.path.exists(LOG):
        log = json.load(open(LOG, encoding='utf-8'))
    for b in batches:
        d = os.path.join(RAW, f'b{b}')
        txt_path = os.path.join(d, 'taxfree.txt')
        if os.path.exists(txt_path):
            log[str(b)] = {'status': 'exists'}
            continue
        os.makedirs(d, exist_ok=True)
        entry = {'batch': b}
        try:
            page, title = search_page(b)
            entry['page'] = page
            if not page:
                entry['status'] = 'no_page'
            else:
                time.sleep(1.5)
                doc_url, label = find_tax_doc(page)
                entry['doc_url'] = doc_url
                entry['label'] = label
                if not doc_url:
                    entry['status'] = 'no_doc'
                else:
                    time.sleep(1.5)
                    ext = os.path.splitext(doc_url)[1] or '.doc'
                    doc_path = os.path.join(d, f'taxfree{ext}')
                    # 中文文件名需逐段编码（343 教训：api-gateway fileName 含中文→ascii 编码错）
                    safe_url = re.sub(r'([^:/?#]+)', lambda m: urllib.parse.quote(m.group(1)) if re.search(r'[\u4e00-\u9fff]', m.group(1)) else m.group(1), doc_url)
                    data = get(safe_url, 90)
                    open(doc_path, 'wb').write(data)
                    entry['bytes'] = len(data)
                    time.sleep(1.0)
                    if ext == '.doc' and convert_utf8(doc_path, txt_path):
                        entry['status'] = 'ok'
                    elif ext == '.docx':
                        entry['status'] = 'docx_need_convert'
                    else:
                        entry['status'] = 'convert_fail'
        except Exception as e:
            entry['status'] = f'err:{type(e).__name__}'
            entry['detail'] = str(e)[:100]
        log[str(b)] = entry
        json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print(f"b{b}: {entry.get('status')} {(entry.get('label') or '')[:28]} {entry.get('bytes','')}")
        time.sleep(1.5)
    ok = sum(1 for v in log.values() if v.get('status') in ('ok', 'exists'))
    print(f'完成: ok/exists {ok}/{len(batches)} → {LOG}')


if __name__ == '__main__':
    main()
