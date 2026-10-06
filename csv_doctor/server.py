"""Local-only HTTP UI. File bytes are held in memory, not sent to external services."""
import base64
import binascii
from http.server import BaseHTTPRequestHandler, HTTPServer
from io import BytesIO
import json
from pathlib import Path
import secrets
import tempfile
import zipfile
from .core import load_table, diagnose, apply, export


def make_handler():
    sessions = {}
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass  # Avoid recording uploaded filenames/content in access logs.

        def reply(self, status, data, mime='application/json'):
            if mime == 'application/json':
                data = json.dumps(data, ensure_ascii=False).encode('utf-8')
            self.send_response(status)
            self.send_header('Content-Type', mime)
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'self'; img-src 'self' blob:; frame-ancestors 'none'; base-uri 'none'")
            if mime == 'application/zip':
                self.send_header('Content-Disposition', 'attachment; filename="csv-doctor-results.zip"')
            self.end_headers()
            self.wfile.write(data)

        def local_request(self):
            port = self.server.server_port
            hosts = {f'127.0.0.1:{port}', f'localhost:{port}'}
            origin = self.headers.get('Origin')
            return self.headers.get('Host') in hosts and (origin is None or origin in {f'http://{x}' for x in hosts})

        def do_GET(self):
            if not self.local_request():
                return self.reply(403, {'error': 'Local same-origin access only'})
            if self.path != '/':
                return self.reply(404, {'error': 'Not found'})
            self.reply(200, Path(__file__).with_name('ui.html').read_bytes(), 'text/html; charset=utf-8')

        def do_POST(self):
            if not self.local_request():
                return self.reply(403, {'error': 'Local same-origin access only'})
            try:
                if self.headers.get('Content-Type', '').split(';')[0] != 'application/json':
                    return self.reply(415, {'error': 'JSON request required'})
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 < length <= 15 * 1024 * 1024:
                    return self.reply(413, {'error': 'Request exceeds 15 MiB'})
                request = json.loads(self.rfile.read(length))
                if not isinstance(request, dict):
                    raise ValueError('Request must be a JSON object')
                if self.path == '/api/import':
                    raw = base64.b64decode(request['data'], validate=True)
                    table = load_table(raw, request.get('filename', 'input.csv'),
                                       request.get('encoding', 'auto'), request.get('delimiter', 'auto'), request.get('sheet'))
                    token = secrets.token_urlsafe(24)
                    if len(sessions) >= 5:
                        sessions.pop(next(iter(sessions)))
                    sessions[token] = table
                    result = {'token': token, 'header': table['header'], 'rows': table['rows'][:20],
                              'encoding': table['encoding'], 'delimiter': table['delimiter'], 'diagnosis': diagnose(table)}
                    result['diagnosis']['issues'] = result['diagnosis']['issues'][:200]
                    return self.reply(200, result)
                if self.path not in ('/api/preview', '/api/export'):
                    return self.reply(404, {'error': 'Not found'})
                table = sessions.get(request.get('token'))
                if table is None:
                    return self.reply(400, {'error': 'Session expired; import the file again'})
                result = apply(table, request['recipe'])
                if self.path == '/api/preview':
                    return self.reply(200, {'header': result['header'], 'rows': result['rows'][:20],
                                           'changes': result['changes'][:200], 'unresolved': result['unresolved'][:200],
                                           'output_rows': len(result['rows']), 'quarantine_count': len(result['quarantine']),
                                           'change_count': len(result['changes']), 'unresolved_count': len(result['unresolved'])})
                with tempfile.TemporaryDirectory() as tmp:
                    export(tmp, table, result)
                    content = BytesIO()
                    with zipfile.ZipFile(content, 'w', zipfile.ZIP_DEFLATED) as archive:
                        for path in sorted(Path(tmp).iterdir()):
                            archive.writestr(path.name, path.read_bytes())
                self.reply(200, content.getvalue(), 'application/zip')
            except (ValueError, KeyError, TypeError, binascii.Error, OSError, zipfile.BadZipFile) as exc:
                self.reply(400, {'error': str(exc)})
    return Handler


def serve(port=8765):
    server = HTTPServer(('127.0.0.1', port), make_handler())
    print(f'CSV Doctor: http://127.0.0.1:{server.server_port} — Ctrl+C to stop')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
