#!/usr/bin/env python3
"""本地预览：python3 candidates/serve.py → http://localhost:8080/ 打开三选一入口。"""
import http.server, os, sys
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **k): super().__init__(*a, directory=ROOT, **k)
    def do_GET(self):
        if self.path in ('/', ''): self.send_response(302); self.send_header('Location', '/candidates/index.html'); self.end_headers(); return
        return super().do_GET()
    def log_message(self, *a): pass
port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
http.server.ThreadingHTTPServer(('0.0.0.0', port), H).serve_forever()
