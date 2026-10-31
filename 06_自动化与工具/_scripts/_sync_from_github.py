#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 GitHub coolcat0125/TA-analysis 仓库下载最新核心文件到本地。
使用 GitHub Contents API，支持大文件（通过 download_url）。
"""
import urllib.request
import urllib.parse
import json
import base64
import os
import sys
import hashlib

DEST = r'C:\00 AI\01 Project\14 TA Analysis'
OWNER = 'coolcat0125'
REPO = 'TA-analysis'
BRANCH = 'main'
API_BASE = f'https://api.github.com/repos/{OWNER}/{REPO}/contents'

# 需要下载的核心文件清单
FILES_TO_DOWNLOAD = [
    'NEV公告参数汇总表_合并版（341~410批）.xlsx',
    'NEV公告数据看板.html',
    'NEV公告数据看板_发布版.html',
    'CHANGELOG.md',
    'generate_dashboard.py',
    'dashboard_template.html',
    'README.md',
    'fill_length_v373.py',
    'refresh_workbook_v37.py',
    '.gitignore',
    '.gitattributes',
    'AUDIT.md',
    'quality_gates.json',
    'audit_data.py',
]

def api_get(url):
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/vnd.github.v3+json',
    })
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read())

def download_file(path):
    encoded = urllib.parse.quote(path)
    url = f'{API_BASE}/{encoded}?ref={BRANCH}'
    print(f'  Fetching: {path}')
    data = api_get(url)

    if isinstance(data, list):
        print(f'  WARNING: {path} is a directory, skipping')
        return None

    if 'content' in data and data['content']:
        # 小文件：base64 直读
        content = base64.b64decode(data['content'])
        sha = data.get('sha', '')
        size = data.get('size', 0)
        print(f'    base64 decode: {len(content)} bytes (sha={sha[:12]})')
        return content, sha, size
    elif 'download_url' in data and data['download_url']:
        # 大文件：走 download_url
        print(f'    Downloading from: {data["download_url"][:80]}...')
        req = urllib.request.Request(data['download_url'], headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=120) as resp:
            content = resp.read()
        sha = data.get('sha', '')
        size = data.get('size', 0)
        print(f'    downloaded: {len(content)} bytes (expected {size}, sha={sha[:12]})')
        return content, sha, size
    else:
        print(f'  ERROR: unexpected response for {path}: {list(data.keys())}')
        return None

def main():
    os.makedirs(DEST, exist_ok=True)
    results = []

    print(f'=== GitHub Sync: {OWNER}/{REPO}@{BRANCH} ===')
    print(f'Destination: {DEST}')
    print(f'Files to download: {len(FILES_TO_DOWNLOAD)}')
    print()

    for path in FILES_TO_DOWNLOAD:
        try:
            result = download_file(path)
            if result is None:
                results.append((path, 'FAILED', 0, ''))
                continue

            content, sha, size = result
            out_path = os.path.join(DEST, path)

            # 备份已有文件（如果存在且不同）
            if os.path.exists(out_path):
                old_size = os.path.getsize(out_path)
                with open(out_path, 'rb') as f:
                    old_hash = hashlib.sha256(f.read()).hexdigest()
                new_hash = hashlib.sha256(content).hexdigest()
                if old_hash == new_hash:
                    print(f'    SKIP: identical to local (size={old_size})')
                    results.append((path, 'UNCHANGED', old_size, sha[:12]))
                    continue
                else:
                    print(f'    UPDATE: local {old_size}B -> remote {len(content)}B')

            with open(out_path, 'wb') as f:
                f.write(content)
            print(f'    SAVED: {out_path}')
            results.append((path, 'UPDATED', len(content), sha[:12]))

        except Exception as e:
            print(f'    ERROR: {e}')
            results.append((path, f'ERROR: {e}', 0, ''))

    print()
    print('=== Summary ===')
    for path, status, size, sha in results:
        print(f'  [{status:10s}] {size:>10,} B  {path}  (sha={sha})')

    updated = sum(1 for _, s, _, _ in results if s == 'UPDATED')
    unchanged = sum(1 for _, s, _, _ in results if s == 'UNCHANGED')
    failed = sum(1 for _, s, _, _ in results if s.startswith('FAILED') or s.startswith('ERROR'))
    print(f'\nTotal: {len(results)} | Updated: {updated} | Unchanged: {unchanged} | Failed: {failed}')

if __name__ == '__main__':
    main()
