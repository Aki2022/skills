"""Fonts served from another origin (a CDN such as Google Fonts) are accepted only when pinned by URL and SHA-256.

The fixed environment used to accept only fonts the repository serves from its own origin, so a site that keeps
Google Fonts could not reach L2 at all. A pinned remote font is fetched at render time, and its bytes must match
the pin: a changed, unlisted or undeclared-origin font is rejected and the glyphs it would draw count as fallback.
"""
import hashlib
import json
import os
import subprocess
from functools import partial
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from threading import Thread

import pytest

from test_core import SKILL, project, check, write  # noqa: F401
from test_fixed_render import NODE, HELPER, pin

BROWSER = pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')


def load(tmp_path, config):
    path = tmp_path / 'design/render-environment.json'
    write(path, config)
    js = ('import {loadEnvironment} from ' + json.dumps(HELPER.as_uri()) + ';'
          'try {loadEnvironment(process.argv.at(-1),process.cwd(),[375],["light"]);console.log("ok")}catch(e){console.log("rejected")}')
    return subprocess.run([NODE, '--input-type=module', '-e', js, str(path)], cwd=tmp_path,
                          capture_output=True, text=True).stdout.strip()


BASE = {'schema_version': 1, 'runtime': {'browser_version': '1', 'browser_sha256': 'a' * 64, 'playwright_version': '1',
        'node_version': 'v1', 'platform': 'p', 'arch': 'a', 'os_release': 'r', 'launch_args': '--x'},
        'clock': '2020-01-01T00:00:00.000Z', 'locale': 'en-US', 'timezone': 'UTC', 'device_scale_factor': 1,
        'reduced_motion': 'reduce', 'viewports': [{'width': 375, 'height': 812}], 'schemes': ['light'], 'fonts': []}
REMOTE = {'origins': ['https://fonts.gstatic.com'], 'families': ['Noto Sans JP'],
          'files': {'https://fonts.gstatic.com/s/notosansjp/v1/a.woff2': 'b' * 64}}


def test_environment_accepts_pinned_remote_fonts_without_local_fonts(tmp_path):
    assert load(tmp_path, {**BASE, 'remote_fonts': REMOTE}) == 'ok'


@pytest.mark.parametrize('remote', [
    {**REMOTE, 'files': {}},                                                   # nothing pinned
    {**REMOTE, 'origins': []},
    {**REMOTE, 'families': []},
    {**REMOTE, 'files': {'https://other.example/a.woff2': 'b' * 64}},         # file outside the declared origins
    {**REMOTE, 'files': {'https://fonts.gstatic.com/a.woff2': 'short'}},      # not a sha256
    {**REMOTE, 'origins': ['https://fonts.gstatic.com/path']},                # an origin has no path
    {**REMOTE, 'origins': ['file://x']},
    {**REMOTE, 'extra': 1},
])
def test_environment_rejects_malformed_remote_font_pins(tmp_path, remote):
    assert load(tmp_path, {**BASE, 'remote_fonts': remote}) == 'rejected'


def test_environment_without_any_font_is_still_rejected(tmp_path):
    assert load(tmp_path, BASE) == 'rejected'


# --- real browser --------------------------------------------------------------------------------

class Quiet(SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')  # what a font CDN sends
        super().end_headers()


def serve(directory):
    server = ThreadingHTTPServer(('127.0.0.1', 0), partial(Quiet, directory=str(directory)))
    Thread(target=server.serve_forever, daemon=True).start()
    return server, 'http://127.0.0.1:' + str(server.server_port)


@pytest.fixture
def cdn_site(project, tmp_path):
    """The page on one origin, its stylesheet and font on another (the CDN)."""
    root, _ = project
    env = dict(os.environ)
    if env.get('DESIGN_TEST_NODE_MODULES'):
        env['NODE_PATH'] = env['DESIGN_TEST_NODE_MODULES']
    config = pin(root, env)
    font_bytes = (root / 'design/assets/fixture.woff2').read_bytes()
    cdn = tmp_path / 'cdn'
    (cdn / 's').mkdir(parents=True)
    (cdn / 's/inter.woff2').write_bytes(font_bytes)
    site, cdn_origin = None, None
    cdn_server, cdn_origin = serve(cdn)
    (cdn / 'style.css').write_text('@font-face{font-family:Inter;src:url(' + cdn_origin + '/s/inter.woff2) format("woff2")}')
    tokens = {'color': {'$type': 'color'}, 'font': {'$type': 'fontFamily', '$value': ['Inter', 'sans-serif'],
              '$extensions': {'org.own.design': {'origin': 'proposal'}}},
              'radius': {'$type': 'dimension', '$value': {'value': 0, 'unit': 'px'}, '$extensions': {'org.own.design': {'origin': 'proposal'}}}}
    for name, components, alpha in [('primary', [0, 0, 0], 1), ('white', [1, 1, 1], 1), ('clear', [0, 0, 0], 0)]:
        tokens['color'][name] = {'$value': {'colorSpace': 'srgb', 'components': components,
                                 'hex': '#ffffff' if name == 'white' else '#000000', 'alpha': alpha},
                                 '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    write(root / 'design/tokens/brand.tokens.json', tokens)
    (root / 'index.html').write_text(
        '<html lang="en"><head><title>Fixture</title><link rel="stylesheet" href="' + cdn_origin + '/style.css"><style>'
        'body{font:16px/1.5 Inter,sans-serif;background:white;color:black}'
        '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
        '</style></head><body><main><h1>Fixture</h1><p>Stable text.</p></main></body></html>')
    site_server, site = serve(root)
    sha = lambda data: hashlib.sha256(data).hexdigest()  # noqa: E731
    remote = {'origins': [cdn_origin], 'families': ['Inter'],
              'files': {cdn_origin + '/style.css': sha((cdn / 'style.css').read_bytes()), cdn_origin + '/s/inter.woff2': sha(font_bytes)}}

    def measure(remote_fonts):
        write(root / 'design/render-environment.json', {**config, **({'fonts': [], 'remote_fonts': remote_fonts} if remote_fonts else {})})
        l0 = check(project)
        assert l0['exit_code'] != 2, l0['errors']
        output = root / 'design/.resolved/remote-test.json'
        output.unlink(missing_ok=True)
        run = subprocess.run([NODE, str(SKILL.parent / 'own-website-audit/scripts/measure_lp.mjs'), site + '/index.html',
                              '--manifest', str(root / 'design/.resolved/web.json'),
                              '--environment', str(root / 'design/render-environment.json'),
                              '--viewports', '375', '--schemes', 'light', '--json', str(output)],
                             cwd=root, env=env, capture_output=True, text=True)
        return run, json.loads(output.read_text()) if output.exists() else None
    yield measure, remote, cdn, root, env, site
    site_server.shutdown()
    cdn_server.shutdown()


def font_failure(data):
    return next((f for f in data['failures'] if f['id'] == 'font-fallback'), None)


@BROWSER
def test_pinned_cdn_font_draws_the_glyphs(cdn_site):
    measure, remote, *_ = cdn_site
    run, data = measure(remote)
    assert run.returncode == 0, run.stdout + run.stderr
    evidence = data['runs'][0]['render_evidence']
    assert evidence['font_glyphs'] > 0 and evidence['fallbacks'] == [] and evidence['rejected_font_requests'] == 0
    font_url = next(u for u in remote['files'] if u.endswith('.woff2'))
    assert evidence['font_sources'] == [{'family': 'Inter', 'sha256s': [remote['files'][font_url]]}]


@BROWSER
def test_cdn_font_without_a_pin_is_rejected_as_before(cdn_site):
    measure, *_ = cdn_site
    run, data = measure(None)
    assert run.returncode != 0 and font_failure(data), run.stdout


@BROWSER
def test_cdn_font_whose_bytes_changed_is_rejected(cdn_site):
    measure, remote, cdn, *_ = cdn_site
    path = cdn / 's/inter.woff2'
    path.write_bytes(path.read_bytes() + b'\0')  # the CDN now serves different bytes under the pinned URL
    run, data = measure(remote)
    assert run.returncode != 0 and font_failure(data), run.stdout
    assert data['runs'][0]['render_evidence']['rejected_font_requests'] >= 1


@BROWSER
def test_cdn_font_url_that_is_not_pinned_is_rejected(cdn_site):
    measure, remote, *_ = cdn_site
    files = {u: h for u, h in remote['files'].items() if not u.endswith('.woff2')}
    files[next(iter(remote['files'])) + '?unused'] = 'c' * 64  # keep the pin non-empty, but not for the font used
    run, data = measure({**remote, 'files': files})
    assert run.returncode != 0 and font_failure(data), run.stdout


@BROWSER
def test_pin_tool_records_what_the_page_loads_from_the_declared_origins(cdn_site):
    measure, remote, cdn, root, env, site = cdn_site
    config = pin(root, env)  # the pin is written before any remote font is known
    write(root / 'design/render-environment.json', {**config, 'fonts': []})
    run = subprocess.run([NODE, str(SKILL / 'scripts/pin_remote_fonts.mjs'), site + '/index.html',
                          '--environment', str(root / 'design/render-environment.json'),
                          '--origins', remote['origins'][0]], cwd=root, env=env, capture_output=True, text=True)
    assert run.returncode == 0, run.stderr
    assert json.loads(run.stdout) == remote
