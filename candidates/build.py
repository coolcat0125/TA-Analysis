#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把三款候选看板打包为自包含 HTML（内嵌 ECharts / 数据 / 脚本），双击离线可用。
用法：python3 candidates/build.py   → 输出 candidates/dist/看板A_*.html 等
"""
import os, re, json, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(HERE, 'dist')
os.makedirs(OUT, exist_ok=True)

subprocess.check_call([sys.executable, os.path.join(HERE, 'extract_base_data.py')])
def rd(p): return open(p, encoding='utf-8').read()
ech = rd(os.path.join(ROOT, 'echarts.min.js'))
data = rd(os.path.join(HERE, 'base_data.json'))
core = rd(os.path.join(HERE, 'core.js'))
filt = rd(os.path.join(HERE, 'filters.js'))

def inline(html):
    html = html.replace('<script src="../../echarts.min.js"></script>', '<script>' + ech + '</script>')
    html = html.replace('<script src="../base_data.js"></script>', '<script>window.DATA=' + data + ';</script>')
    html = html.replace('<script src="../core.js"></script>', '<script>' + core + '</script>')
    html = html.replace('<script src="../filters.js"></script>', '<script>' + filt + '</script>')
    return html

# 开发态 base_data.js
open(os.path.join(HERE, 'base_data.js'), 'w', encoding='utf-8').write('window.DATA=' + data + ';')

for name, out in [('A_editorial', '看板A_行业简报式.html'), ('B_console', '看板B_分析控制台.html'), ('C_mobile', '看板C_移动卡片式.html')]:
    src = os.path.join(HERE, name, 'index.html')
    html = inline(rd(src))
    open(os.path.join(OUT, out), 'w', encoding='utf-8').write(html)
    print('built', out, round(os.path.getsize(os.path.join(OUT, out)) / 1024), 'KB')
