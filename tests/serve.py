#!/usr/bin/env python3
"""Serve the app and its tests locally with caching off, so edits always load.

    python3 tests/serve.py        then open http://localhost:8766/tests/
"""
import http.server, os, sys

os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def log_message(self, *args):
        pass


port = int(sys.argv[1]) if len(sys.argv) > 1 else 8766
print(f'http://localhost:{port}/tests/')
http.server.ThreadingHTTPServer(('127.0.0.1', port), Handler).serve_forever()
