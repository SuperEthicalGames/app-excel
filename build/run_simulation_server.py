"""Run the local app HTTP server without opening an unrelated browser window."""

from __future__ import annotations

import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))

import server  # noqa: E402

if len(sys.argv) > 1:
    server.PORT = int(sys.argv[1])

httpd = server.ThreadingHTTPServer((server.HOST, server.PORT), server.AppHandler)
try:
    httpd.serve_forever()
finally:
    httpd.server_close()
