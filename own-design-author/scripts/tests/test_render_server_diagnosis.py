"""A render server that never answers names itself: which row, what it last answered, and where its output is.

Without the name, a pre-push that starts four servers and probes sixteen paths stops with one sentence that fits
every one of them, and the output that would explain it was discarded.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ship  # noqa: E402

SERVE = ('import sys,http.server;'
         'print("server says hello",flush=True);'
         'http.server.ThreadingHTTPServer(("127.0.0.1",int(sys.argv[1])),http.server.SimpleHTTPRequestHandler).serve_forever()')


@pytest.fixture(autouse=True)
def short_wait(monkeypatch):
    monkeypatch.setattr(ship, 'READY_SECONDS', 2)


def test_a_path_that_never_answers_names_its_row_status_and_log(tmp_path):
    servers = {'DESIGN_WEB_URL': {'command': [sys.executable, '-c', SERVE, '{port}'], 'path': '/'},
               'DESIGN_WEB_MISSING_URL': {'same_as': 'DESIGN_WEB_URL', 'path': '/missing'}}
    with pytest.raises(ValueError) as caught:
        with ship.render_servers(tmp_path, servers, ship.hook_free_env()):
            pass
    message = str(caught.value)
    assert 'DESIGN_WEB_MISSING_URL' in message and 'HTTP 404' in message and '/missing' in message
    log = Path(message.split('log: ')[1].strip())
    assert 'server says hello' in log.read_text()  # the shared server's output is kept, not discarded


def test_a_server_that_exits_names_its_row_exit_code_and_log(tmp_path):
    crash = 'import sys;print("boom before listening",file=sys.stderr);sys.exit(7)'
    servers = {'DESIGN_LP_URL': {'command': [sys.executable, '-c', crash, '{port}'], 'path': '/'}}
    with pytest.raises(ValueError) as caught:
        with ship.render_servers(tmp_path, servers, ship.hook_free_env()):
            pass
    message = str(caught.value)
    assert 'DESIGN_LP_URL' in message and 'exit 7' in message
    assert 'boom before listening' in Path(message.split('log: ')[1].strip()).read_text()


def test_a_server_that_never_listens_says_no_connection(tmp_path):
    idle = 'import time;time.sleep(30)'
    servers = {'DESIGN_REPORT_URL': {'command': [sys.executable, '-c', idle, '{port}'], 'path': '/'}}
    with pytest.raises(ValueError) as caught:
        with ship.render_servers(tmp_path, servers, ship.hook_free_env()):
            pass
    message = str(caught.value)
    assert 'DESIGN_REPORT_URL' in message and 'no connection' in message and 'did not become ready' in message


def test_logs_are_removed_when_every_server_answers(tmp_path):
    servers = {'DESIGN_WEB_URL': {'command': [sys.executable, '-c', SERVE, '{port}'], 'path': '/'}}
    with ship.render_servers(tmp_path, servers, ship.hook_free_env()) as (_env, _urls, processes):
        log_dir = Path(processes[0].log_dir)
        assert log_dir.is_dir()
    assert not log_dir.exists()
