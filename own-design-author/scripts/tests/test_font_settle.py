"""Glyph evidence is read only after every font the page is loading has settled.

On a real site (officeohana_company, 2026-11-15) the same page and pinned fonts fell back to a system font in
1 of 6 renders, only at 375px: a font face started loading after `document.fonts.ready` had resolved, and the
glyph check ran while it was still loading, so `font-display: swap` text was drawn in a system font. These
tests make that race deterministic with a font that starts loading late and arrives slowly.
"""
import hashlib
import json
import os
import subprocess
import time
from functools import partial
from http.server import ThreadingHTTPServer
from threading import Thread

import pytest

from test_core import SKILL, project, check, write  # noqa: F401
import layer_paths as lp
from test_fixed_render import NODE, pin
from test_remote_fonts import Quiet

pytestmark = pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')

LATE = ('<script>setTimeout(()=>{const p=document.createElement("p");p.style.fontFamily="Late,sans-serif";'
        'p.textContent="Late text";document.querySelector("main").append(p)},400)</script>')


class Slow(Quiet):
    delay = 0.0
    css_delay = 0.0

    def do_GET(self):
        if self.path.endswith('late.woff2'):
            time.sleep(self.delay)
        if self.path.endswith('style.css'):
            time.sleep(self.css_delay)
        super().do_GET()


@pytest.fixture
def late_font(project, tmp_path):
    root, _ = project
    env = dict(os.environ)
    if env.get('DESIGN_TEST_NODE_MODULES'):
        env['NODE_PATH'] = env['DESIGN_TEST_NODE_MODULES']
    config = pin(root, env)
    font = (lp.path(root, 'assets', 'fixture.woff2')).read_bytes()
    cdn = tmp_path / 'cdn'
    (cdn / 's').mkdir(parents=True)
    (cdn / 's/inter.woff2').write_bytes(font)
    (cdn / 's/late.woff2').write_bytes(font)
    handler = type('SlowCdn', (Slow,), {})
    cdn_server = ThreadingHTTPServer(('127.0.0.1', 0), partial(handler, directory=str(cdn)))
    Thread(target=cdn_server.serve_forever, daemon=True).start()
    origin = 'http://127.0.0.1:' + str(cdn_server.server_port)
    (cdn / 'style.css').write_text(
        '@font-face{font-family:Inter;src:url(' + origin + '/s/inter.woff2) format("woff2")}'
        '@font-face{font-family:Late;font-display:swap;src:url(' + origin + '/s/late.woff2) format("woff2")}')
    tokens = {'color': {'$type': 'color'}, 'font': {'$type': 'fontFamily', '$value': ['Inter', 'sans-serif'],
              '$extensions': {'org.own.design': {'origin': 'proposal'}}},
              'radius': {'$type': 'dimension', '$value': {'value': 0, 'unit': 'px'}, '$extensions': {'org.own.design': {'origin': 'proposal'}}}}
    tokens['late'] = {'$type': 'fontFamily', '$value': ['Late', 'sans-serif'], '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    for name, components, alpha in [('primary', [0, 0, 0], 1), ('white', [1, 1, 1], 1), ('clear', [0, 0, 0], 0)]:
        tokens['color'][name] = {'$value': {'colorSpace': 'srgb', 'components': components,
                                 'hex': '#ffffff' if name == 'white' else '#000000', 'alpha': alpha},
                                 '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    write(lp.path(root, 'design_dir', 'tokens', 'brand.tokens.json'), tokens)
    (root / 'index.html').write_text(
        '<html lang="en"><head><title>Fixture</title><link rel="stylesheet" href="' + origin + '/style.css"><style>'
        'body{font:16px/1.5 Inter,sans-serif;background:white;color:black}p{margin:0 0 16px}'
        '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
        '</style></head><body><main><h1 style="margin:0 0 16px">Fixture</h1><p>Stable text.</p></main>' + LATE + '</body></html>')
    from test_remote_fonts import serve
    site_server, site = serve(root)
    sha = lambda data: hashlib.sha256(data).hexdigest()  # noqa: E731
    remote = {'origins': [origin], 'families': ['Inter'],
              'files': {origin + '/style.css': sha((cdn / 'style.css').read_bytes()),
                        origin + '/s/inter.woff2': sha(font), origin + '/s/late.woff2': sha(font)}}
    write(lp.path(root, 'render_environment'), {**config, 'fonts': [], 'remote_fonts': remote})

    page = (root / 'index.html').read_text()

    def measure(delay, css_delay=0.0, script=True):
        handler.delay = delay
        handler.css_delay = css_delay
        (root / 'index.html').write_text(page if script else page.replace(LATE, ''))
        l0 = check(project)
        assert l0['exit_code'] != 2, l0['errors']
        output = lp.path(root, 'resolved', 'settle-test.json')
        output.unlink(missing_ok=True)
        run = subprocess.run([NODE, str(SKILL.parent / 'own-web-build/scripts/measure_render.mjs'), site + '/index.html',
                              '--manifest', str(lp.path(root, 'resolved', 'web.json')),
                              '--environment', str(lp.path(root, 'render_environment')),
                              '--profile', 'lp', '--viewports', '375', '--schemes', 'light', '--json', str(output)],
                             cwd=root, env=env, capture_output=True, text=True, timeout=120)
        return run, json.loads(output.read_text()) if output.exists() else None
    yield measure
    site_server.shutdown()
    cdn_server.shutdown()


def fallback(data):
    return next((f for f in data['failures'] if f['id'] == 'font-fallback'), None)


def test_a_font_that_starts_loading_late_is_waited_for_before_the_glyph_check(late_font):
    run, data = late_font(4)
    assert run.returncode == 0, run.stdout + run.stderr
    assert fallback(data) is None and data['environment']['repeat_equal'] is True
    assert all(r['render_evidence']['fallbacks'] == [] for r in data['runs'])


def test_a_font_that_never_settles_fails_the_render_instead_of_passing(late_font):
    run, data = late_font(40)  # longer than the settle limit
    assert run.returncode == 2, run.stdout + run.stderr
    assert data['execution_error'] == 'font-unsettled', data.get('failures')


def test_a_stylesheet_that_arrives_late_is_waited_for_before_fonts_count_as_settled(late_font):
    # Seen once in 24 real renders: the whole page drawn in system fonts with no pinned glyph. A page without
    # scripts fires DOMContentLoaded before its stylesheets load, so no font face existed yet and the font set
    # looked idle. The render waits for the page's load event (stylesheets included) before fonts can settle.
    run, data = late_font(0, css_delay=2, script=False)
    assert run.returncode == 0, run.stdout + run.stderr
    assert all(r['render_evidence']['fallbacks'] == [] for r in data['runs'])
