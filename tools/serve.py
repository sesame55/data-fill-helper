"""本機測試用伺服器：不讓瀏覽器快取，改完程式重新整理就是新版。
用法（在專案資料夾）：python tools/serve.py   → http://127.0.0.1:8000/
"""
import functools
import http.server
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class NoCacheHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


if __name__ == "__main__":
    handler = functools.partial(NoCacheHandler, directory=ROOT)
    with http.server.ThreadingHTTPServer(("127.0.0.1", 8000), handler) as httpd:
        print("Serving http://127.0.0.1:8000/", flush=True)
        httpd.serve_forever()
