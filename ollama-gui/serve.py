#!/usr/bin/env python3
"""Tiny zero-dependency web server for the Ollama Chat GUI.

Serves index.html and reverse-proxies Ollama's API paths (/api/*, /v1/*) to an
Ollama instance, so the browser talks to a single origin and no CORS setup is
needed. Responses are streamed through chunk by chunk, so tokens appear live.

Usage:
    python3 serve.py                      # http://localhost:8080 -> http://localhost:11434
    OLLAMA_URL=http://10.0.0.5:11434 PORT=9000 python3 serve.py
"""
import http.client
import os
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit

OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434").rstrip("/")
HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8080"))
PROXY_PREFIXES = ("/api/", "/v1/")
HERE = os.path.dirname(os.path.abspath(__file__))

# Hop-by-hop headers must not be forwarded by a proxy (RFC 9110 section 7.6.1).
HOP_BY_HOP = {
    "connection", "keep-alive", "proxy-authenticate", "proxy-authorization",
    "te", "trailer", "transfer-encoding", "upgrade", "host", "content-length",
}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=HERE, **kwargs)

    def do_GET(self):
        if self.path.startswith(PROXY_PREFIXES):
            self.proxy()
        else:
            super().do_GET()

    def do_HEAD(self):
        if self.path.startswith(PROXY_PREFIXES):
            self.proxy()
        else:
            super().do_HEAD()

    def do_POST(self):
        if self.path.startswith(PROXY_PREFIXES):
            self.proxy()
        else:
            self.send_error(405)

    do_DELETE = do_POST

    def proxy(self):
        target = urlsplit(OLLAMA_URL)
        conn_cls = http.client.HTTPSConnection if target.scheme == "https" else http.client.HTTPConnection
        length = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(length) if length else None
        headers = {k: v for k, v in self.headers.items() if k.lower() not in HOP_BY_HOP}
        # Ollama checks Origin against OLLAMA_ORIGINS; drop it since this is now a
        # same-origin, server-to-server call.
        headers.pop("Origin", None)
        try:
            conn = conn_cls(target.hostname, target.port, timeout=600)
            conn.request(self.command, target.path + self.path, body=body, headers=headers)
            resp = conn.getresponse()
        except OSError as exc:
            self.send_error(502, f"Cannot reach Ollama at {OLLAMA_URL}: {exc}")
            return

        self.send_response(resp.status, resp.reason)
        for k, v in resp.getheaders():
            if k.lower() not in HOP_BY_HOP | {"server", "date"}:
                self.send_header(k, v)
        # HTTP/1.0 response without Content-Length: the body ends when the
        # connection closes, which lets us stream tokens as they arrive.
        self.send_header("Connection", "close")
        self.end_headers()
        try:
            while True:
                chunk = resp.read1(8192)
                if not chunk:
                    break
                self.wfile.write(chunk)
                self.wfile.flush()
        except (BrokenPipeError, ConnectionResetError):
            pass  # browser pressed Stop / closed the tab
        finally:
            conn.close()


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Ollama Chat GUI on http://{HOST}:{PORT}  (proxying /api, /v1 -> {OLLAMA_URL})")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        sys.exit(0)


if __name__ == "__main__":
    main()
