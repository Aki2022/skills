"""Each code check the item map relies on fires on a page that breaks it and stays silent on a clean page."""
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))
from item_checks import COLLECTED  # noqa: E402

DOM_CHECKS = sorted(c for c in COLLECTED if not c.startswith('axe:'))

BAD = '''<!doctype html><html lang="ja"><head><title>Page</title></head><body>
<h1>Other</h1><h1>Second</h1>
<input placeholder="x"><input readonly><input disabled><input maxlength="3"><input pattern="[0-9]+">
<ol><li>a</li></ol><div tabindex="2">t</div><input type="date"><button disabled>off</button>
<table><tr><td colspan="2">12</td></tr></table>
<div aria-expanded="false">menu</div>
<a href="#a">こちら</a> <a href="#b" style="cursor:default">plain</a>
<button style="width:20px;overflow:hidden;white-space:nowrap">a very long label</button>
<input style="height:20px">
<label for="i1">Name</label><input id="i1">
<input id="p" onpaste="return false">
<form><p role="alert">error</p></form>
<p>・one<br>・two</p>
<button><svg width="10" height="10"><rect width="10" height="10"></rect></svg></button>
<a href="#c"><svg width="10" height="10"><rect width="10" height="10"></rect></svg>Text</a>
<a href="https://example.org/" target="_blank">external</a>
<p>text <a href="#d" style="text-decoration:none;color:inherit">link</a> more</p>
<p>氏　名</p>
<p style="font-style:italic">''' + 'long italic text ' * 10 + '''</p>
<div role="dialog" style="border:none">dialog</div>
<input type="file" style="display:none">
<div aria-roledescription="carousel">''' + '<div aria-roledescription="slide">s</div>' * 7 + '''</div>
<div style="display:flex;flex-direction:column-reverse"><button>first</button><button>second</button></div>
<input style="border:1px solid #eeeeee;background:#ffffff">
<svg width="10" height="10" fill="#dddddd"><rect width="10" height="10"></rect></svg>
<p style="font-size:13px">small</p><p style="font-weight:300">thin</p>
<div style="padding:13px">pad</div><div style="box-shadow:0 0 3px #000">shadow</div>
<p style="text-align:justify">justified</p>
</body></html>'''

GOOD = '''<!doctype html><html lang="ja"><head><title>Clean page | Site</title>
<style>body{background:#ffffff;color:#000000;font:400 16px/1.5 sans-serif;margin:0}
input,button,select{font:inherit;padding:8px;border:1px solid #333333;background:#ffffff;color:#000000;height:40px}
label{display:block}td,th{padding:8px;vertical-align:top}h1{font-size:32px;font-weight:700;margin:0}
a{text-decoration:underline;color:#000000}p{margin:0}</style></head><body>
<h1>Clean page</h1>
<label for="n">Name</label><input id="n" type="text">
<table><tr><th>Item</th><th>Count</th></tr><tr><td>Apples</td><td style="text-align:right">12</td></tr></table>
<button aria-expanded="false" aria-controls="m">Menu</button><div id="m" hidden>menu</div>
<p>Read the <a href="#guide">application guide</a> first.</p>
<p><a href="https://example.org/" target="_blank">Example（新規タブで開きます）</a></p>
<button><svg aria-hidden="true" width="10" height="10" fill="#000000"><rect width="10" height="10"></rect></svg>Save</button>
<p>Name and address</p>
</body></html>'''

RUNNER = '''
import {createRequire} from "node:module";
import {collectItemChecks} from %s;
const require = createRequire(process.cwd() + "/");
const {chromium} = require("playwright");
const [html, manifest] = JSON.parse(process.argv[2]);
const browser = await chromium.launch();
const page = await browser.newPage({viewport:{width:800, height:2000}});
await page.setContent(html);
console.log(JSON.stringify(await page.evaluate(collectItemChecks, manifest)));
await browser.close();
'''


def collect(html, tmp_path):
    if os.environ.get('DESIGN_BROWSER_TEST') != '1':
        pytest.skip('explicit browser fixture run required')
    script = tmp_path / 'run.mjs'
    script.write_text(RUNNER % json.dumps((SCRIPTS / 'resolved_contract.mjs').as_uri()))
    manifest = {'item_checks': {'collect': DOM_CHECKS, 'applicability': {}},
                'palette': {'dimension': ['0px', '8px', '16px', '32px', '40px'], 'fontWeight': ['400', '700'], 'shadow': []}}
    env = dict(os.environ)
    if os.environ.get('DESIGN_TEST_NODE_MODULES'):
        env['NODE_PATH'] = os.environ['DESIGN_TEST_NODE_MODULES']
    run = subprocess.run([os.environ.get('DESIGN_TEST_NODE', 'node'), str(script), json.dumps([html, manifest])],
                         cwd=Path(os.environ['DESIGN_TEST_NODE_MODULES']).parent if os.environ.get('DESIGN_TEST_NODE_MODULES') else tmp_path,
                         env=env, capture_output=True, text=True)
    assert run.returncode == 0, run.stderr
    return json.loads(run.stdout)['results']


def test_every_dom_check_fires_on_a_page_that_breaks_it(tmp_path):
    results = collect(BAD, tmp_path)
    silent = [c for c in DOM_CHECKS if results[c]['count'] == 0]
    assert not silent, silent


def test_no_dom_check_fires_on_a_clean_page(tmp_path):
    results = collect(GOOD, tmp_path)
    noisy = {c: results[c] for c in DOM_CHECKS if results[c]['count']}
    assert not noisy, noisy


SR_ONLY = 'position:absolute;width:1px;height:1px;overflow:hidden;clip-path:inset(50%);white-space:nowrap'


@pytest.mark.parametrize('body,fires', [
    # The catalog asks for no disabled control at all, or one with readable text right by it saying why.
    ('<button disabled>送信する</button>', True),
    ('<input type="submit" disabled value="送信する">', True),
    ('<button aria-disabled="true">送信する</button>', True),
    ('<fieldset disabled><button>申請する</button></fieldset>', True),             # disabled by its fieldset
    ('<button disabled aria-describedby="r">送信</button><p id="r" hidden>本人確認後に送信できます</p>', True),
    ('<button disabled aria-describedby="r">送信</button><p id="r" style="' + SR_ONLY + '">本人確認後に送信できます</p>', True),
    ('<button disabled aria-describedby="missing">送信</button>', True),
    ('<button disabled aria-describedby="r">送信</button><p id="r"> </p>', True),
    ('<button disabled aria-describedby="r">送信</button><p id="r">本人確認が完了すると送信できます</p>', False),
    ('<fieldset disabled><p id="r">送信中です</p><button aria-describedby="r">申請する</button></fieldset>', False),
    ('<button aria-disabled="true" aria-describedby="r">送信</button><p id="r">送信中です</p>', False),
    ('<button aria-disabled="false">確認する</button>', False),
    ('<button>確認する</button>', False),
])
def test_disabled_button_needs_a_readable_reason_tied_to_it(tmp_path, body, fires):
    page = '<!doctype html><html lang="ja"><head><title>t</title></head><body>' + body + '</body></html>'
    assert (collect(page, tmp_path)['dom:disabled-button']['count'] > 0) is fires
