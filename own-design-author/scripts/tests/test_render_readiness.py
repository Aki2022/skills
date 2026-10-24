"""Visible resource readiness must not wait forever for background fetch streams."""
import json
import os
import subprocess
from functools import partial
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from threading import Event, Thread

import pytest
from test_core import SKILL, project, check
from test_ship_integration import browser_project, browser_env


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')
def test_background_stream_does_not_replace_visible_resource_readiness(project):
    env = browser_env(); root = browser_project(project, env)
    page = root / 'index.html'
    page.write_text(page.read_text().replace('</body>', '<script>fetch("/background").then(r=>r.text()).catch(()=>{})</script></body>'))
    assert check(project)['exit_code'] != 2
    stop = Event(); requested = Event()
    class Handler(SimpleHTTPRequestHandler):
        def log_message(self, *args): pass
        def do_GET(self):
            if self.path == '/background':
                requested.set(); self.send_response(200); self.send_header('Content-Length','999999'); self.end_headers()
                self.wfile.write(b'x'); self.wfile.flush(); stop.wait(15)
            else: super().do_GET()
    server = ThreadingHTTPServer(('127.0.0.1',0),partial(Handler,directory=str(root)))
    thread = Thread(target=server.serve_forever,daemon=True); thread.start()
    output = root / 'design/.resolved/readiness.json'
    argv = [os.environ.get('DESIGN_TEST_NODE','node'),str(SKILL.parent/'own-website-audit/scripts/measure_lp.mjs'),
        'http://127.0.0.1:'+str(server.server_port),'--manifest',str(root/'design/.resolved/web.json'),
        '--environment',str(root/'design/render-environment.json'),'--viewports','375','--schemes','light',
        '--profile','product-ui','--json',str(output)]
    try:
        run = subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True,timeout=12)
        assert run.returncode == 0, run.stdout+run.stderr
        result = json.loads(output.read_text())
        assert requested.is_set() and len(result['runs']) == 2
        assert result['environment']['repeat_equal'] is True and result['failures'] == []
        page.write_text(page.read_text().replace('<p>Stable text.</p>', '<img src="/missing.png" alt="Required image"><p>Stable text.</p>'))
        run = subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True,timeout=12)
        assert run.returncode != 0
    finally:
        stop.set(); server.shutdown(); server.server_close(); thread.join()
