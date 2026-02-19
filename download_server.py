#!/usr/bin/env python3
"""강제 다운로드 서버 - Content-Disposition: attachment 헤더 추가"""
from http.server import HTTPServer, BaseHTTPRequestHandler
import os
import urllib.parse

SERVE_DIR = "/home/user/webapp"

class ForceDownloadHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        # URL 디코딩
        path = urllib.parse.unquote(self.path.lstrip('/'))

        # 루트 접속 시 파일 목록 표시
        if not path:
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            files = [f for f in os.listdir(SERVE_DIR) if os.path.isfile(os.path.join(SERVE_DIR, f))]
            py_files = [f for f in files if f.endswith('.py')]
            html = '<html><body><h2>다운로드 가능한 파일</h2><ul>'
            for f in sorted(py_files):
                encoded = urllib.parse.quote(f)
                html += f'<li><a href="/{encoded}" download="{f}">{f}</a></li>'
            html += '</ul></body></html>'
            self.wfile.write(html.encode('utf-8'))
            return

        filepath = os.path.join(SERVE_DIR, path)
        if not os.path.isfile(filepath):
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b'File not found')
            return

        filesize = os.path.getsize(filepath)
        filename = os.path.basename(filepath)

        self.send_response(200)
        # 강제 다운로드 헤더
        self.send_header('Content-Type', 'application/octet-stream')
        self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.send_header('Content-Length', str(filesize))
        self.end_headers()

        with open(filepath, 'rb') as f:
            self.wfile.write(f.read())

    def log_message(self, format, *args):
        pass  # 로그 출력 억제

if __name__ == '__main__':
    server = HTTPServer(('0.0.0.0', 7778), ForceDownloadHandler)
    print('Download server running on port 7778')
    server.serve_forever()
