"""A page must measure the same way twice. These guard two sources of repeat mismatch that made
real pages fail with non-deterministic-render: a focus rule that left the page scrolled before
the capture, and region captures taken without waiting for the pixels to settle."""
import json
import os
from pathlib import Path
import subprocess

import pytest

SCRIPTS = Path(__file__).resolve().parents[1]


def run_node(source, payload, tmp_path):
    if os.environ.get('DESIGN_BROWSER_TEST') != '1':
        pytest.skip('explicit browser fixture run required')
    script = tmp_path / 'run.mjs'
    script.write_text(source)
    env = dict(os.environ)
    modules = os.environ.get('DESIGN_TEST_NODE_MODULES')
    if modules:
        env['NODE_PATH'] = modules
    run = subprocess.run([os.environ.get('DESIGN_TEST_NODE', 'node'), str(script), json.dumps(payload)],
                         cwd=Path(modules).parent if modules else tmp_path, env=env, capture_output=True, text=True)
    assert run.returncode == 0, run.stderr
    return json.loads(run.stdout)


FOCUS_RUNNER = '''
import {createRequire} from "node:module";
import {collectItemChecks} from %s;
const require = createRequire(process.cwd() + "/");
const {chromium} = require("playwright");
const [html, manifest] = JSON.parse(process.argv[2]);
const browser = await chromium.launch();
const page = await browser.newPage({viewport:{width:800, height:600}});
await page.setContent(html);
const results = (await page.evaluate(collectItemChecks, manifest)).results;
const scrollY = await page.evaluate(() => window.scrollY);
console.log(JSON.stringify({results, scrollY}));
await browser.close();
'''


def test_focus_rule_leaves_the_page_where_it_was(tmp_path):
    html = ('<!doctype html><html><body style="margin:0"><div style="height:5000px">tall</div>'
            '<a href="#x" style="outline-style:solid">last link</a></body></html>')
    rule = {'id': 'focus-ring', 'kind': 'style', 'targets': ['a[href]'], 'property': 'outline-style',
            'state': 'focus', 'allowlist': ['solid', 'auto']}
    manifest = {'item_checks': {'collect': ['rule:focus-ring'], 'applicability': {}},
                'palette': {'dimension': ['0px'], 'fontWeight': ['400'], 'shadow': []}, 'rules': [rule]}
    out = run_node(FOCUS_RUNNER % json.dumps((SCRIPTS / 'resolved_contract.mjs').as_uri()), [html, manifest], tmp_path)
    assert out['results']['rule:focus-ring'].get('checked') is not False
    assert out['scrollY'] == 0, out


CAPTURE_RUNNER = '''
import {createRequire} from "node:module";
import {stableCapture} from %s;
import {createHash} from "node:crypto";
const require = createRequire(process.cwd() + "/");
const {chromium} = require("playwright");
const [html] = JSON.parse(process.argv[2]);
const browser = await chromium.launch();
const page = await browser.newPage({viewport:{width:400, height:300}});
await page.setContent(html);
const clip = {x:0, y:0, width:100, height:100};
let out;
try {
  const a = await stableCapture(page, {fullPage:true, clip});
  const b = await stableCapture(page, {fullPage:true, clip});
  const h = (p) => createHash("sha256").update(p).digest("hex");
  out = {ok:true, same:h(a)===h(b)};
} catch (e) { out = {ok:false, error:String(e.message)}; }
console.log(JSON.stringify(out));
await browser.close();
'''


def capture(html, tmp_path):
    return run_node(CAPTURE_RUNNER % json.dumps((SCRIPTS / 'fixed_render.mjs').as_uri()), [html], tmp_path)


def test_region_capture_of_a_settled_page_repeats_exactly(tmp_path):
    out = capture('<div style="width:100px;height:100px;background:#336699"></div>', tmp_path)
    assert out == {'ok': True, 'same': True}, out


def test_region_capture_refuses_pixels_that_never_settle(tmp_path):
    html = ('<div id="b" style="width:100px;height:100px;background:#000"></div><script>'
            'let n=0;setInterval(()=>{n=(n+37)%256;document.getElementById("b").style.background='
            '`rgb(${n},${255-n},${n})`},5)</script>')
    out = capture(html, tmp_path)
    assert out['ok'] is False and 'stabilize' in out['error'], out
