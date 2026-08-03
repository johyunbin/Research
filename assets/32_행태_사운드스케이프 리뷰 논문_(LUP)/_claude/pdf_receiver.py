# -*- coding: utf-8 -*-
"""
Paper32 — 로컬 PDF 수신 서버 (127.0.0.1:18923)
브라우저 페이지 JS가 same-origin fetch로 받은 PDF blob을 POST하면 저장.
POST /save?id=ID_0123 (body=PDF bytes) → fulltext/pdf/ID_0123.pdf
GET /status → 저장 카운트. CORS 전면 허용(로컬 수신 전용).
"""
import os, re
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

DEST = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude\fulltext\pdf"
os.makedirs(DEST, exist_ok=True)


class H(BaseHTTPRequestHandler):
    timeout = 30
    protocol_version = "HTTP/1.1"

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Private-Network", "true")

    def do_OPTIONS(self):
        self.send_response(204); self._cors(); self.end_headers()

    def do_GET(self):
        n = len([f for f in os.listdir(DEST) if f.endswith(".pdf")])
        body = f"saved={n}".encode()
        self.send_response(200); self._cors()
        self.send_header("Content-Type", "text/plain")
        self.send_header("Content-Length", str(len(body))); self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        m = re.search(r"id=(ID_\d{4})", self.path)
        length = int(self.headers.get("Content-Length", 0))
        data = self.rfile.read(length) if length else b""
        if not m or data[:5] != b"%PDF-" or length < 10000:
            body = b"reject"
            self.send_response(400); self._cors()
            self.send_header("Content-Length", str(len(body))); self.end_headers()
            self.wfile.write(body); return
        with open(os.path.join(DEST, m.group(1) + ".pdf"), "wb") as f:
            f.write(data)
        body = f"ok:{m.group(1)}:{length}".encode()
        self.send_response(200); self._cors()
        self.send_header("Content-Type", "text/plain")
        self.send_header("Content-Length", str(len(body))); self.end_headers()
        self.wfile.write(body)

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    print("receiver on 127.0.0.1:18923 (threading)")
    ThreadingHTTPServer(("127.0.0.1", 18923), H).serve_forever()
