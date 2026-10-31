# -*- coding: utf-8 -*-
"""serve_monitor.py — DCD 监控看板本地实时服务

用法：python serve_monitor.py [端口，默认 8797]
然后浏览器打开 http://localhost:8797/DCD监控看板.html
看板每 3 秒自动拉取 raw/dcd_refill/status.json（由 dcd_batch_run.py status 刷新）。
Ctrl+C 停止。
"""
import http.server
import os
import socketserver
import sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8797
HERE = os.path.dirname(os.path.abspath(__file__))


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=HERE, **kw)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

    def log_message(self, fmt, *args):
        pass  # 静默


with socketserver.TCPServer(('', PORT), Handler) as httpd:
    print(f'DCD 监控看板实时服务已启动：')
    print(f'  → http://localhost:{PORT}/DCD监控看板.html')
    print(f'  数据源：raw/dcd_refill/status.json（每 3s 刷新）')
    print(f'  停止：Ctrl+C')
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print('\n已停止')
