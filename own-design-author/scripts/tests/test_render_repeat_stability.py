"""A page must measure the same way twice. These guard two sources of repeat mismatch that made
real pages fail with non-deterministic-render: a focus rule that left the page scrolled before
the capture, and region captures taken without waiting for the pixels to settle."""
import json
import os
from pathlib import Path
import subprocess

import pytest
from test_core import project  # noqa: F401  (fixture used by the measure_lp tests below)

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


def focus_check(link_style, tmp_path):
    html = ('<!doctype html><html><head><style>a{outline:none}' + link_style + '</style></head>'
            '<body style="margin:0"><div style="height:5000px">tall</div><a href="#x">last link</a></body></html>')
    rule = {'id': 'focus-ring', 'kind': 'style', 'targets': ['a[href]'], 'property': 'outline-style',
            'state': 'focus', 'allowlist': ['solid']}
    manifest = {'item_checks': {'collect': ['rule:focus-ring'], 'applicability': {}},
                'palette': {'dimension': ['0px'], 'fontWeight': ['400'], 'shadow': []}, 'rules': [rule]}
    return run_node(FOCUS_RUNNER % json.dumps((SCRIPTS / 'resolved_contract.mjs').as_uri()), [html, manifest], tmp_path)


def test_focus_rule_measures_the_focused_style_without_scrolling(tmp_path):
    # The ring exists only while focused, so a pass proves the check read the focused state.
    out = focus_check('a:focus{outline-style:solid}', tmp_path)
    row = out['results']['rule:focus-ring']
    assert row.get('checked') is not False and row['count'] == 0, out
    assert out['scrollY'] == 0, out


def test_focus_rule_still_fails_a_missing_focus_ring_without_scrolling(tmp_path):
    out = focus_check('', tmp_path)
    assert out['results']['rule:focus-ring']['count'] == 1, out
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


# Through measure_lp itself, so reverting measureRegion to a single capture turns these red.
WIDE = ("data:image/svg+xml," "%3Csvg xmlns='http://www.w3.org/2000/svg' width='600' height='200'%3E"
        "%3Crect width='600' height='200' fill='%23336699'/%3E%3C/svg%3E")
TALL = ("data:image/svg+xml," "%3Csvg xmlns='http://www.w3.org/2000/svg' width='100' height='400'%3E"
        "%3Crect width='100' height='400' fill='%23336699'/%3E%3C/svg%3E")
FLICKER = ('<div style="position:relative;width:300px;height:100px">'
           '<img src="SRC" alt="chart" style="width:300px;height:100px;object-fit:FIT;display:block">'
           '<div id="f" style="position:absolute;left:0;top:0;width:60px;height:60px;background:#000"></div></div>'
           '<script>let n=0;setInterval(()=>{n=(n+53)%256;document.getElementById("f").style.background='
           '`rgb(${n},${255-n},${n})`},5)</script>')


def measure_flicker(project, src, fit):
    from test_measure_profiles import BROWSER, NORMAL, prepare, measure
    if not BROWSER:
        pytest.skip('explicit real browser run required')
    root = prepare(project)
    return measure(root, NORMAL.replace('CONTENT', FLICKER.replace('SRC', src).replace('FIT', fit)), 'product-ui')


def test_an_image_region_that_never_settles_fails_instead_of_passing(project):
    run, report = measure_flicker(project, WIDE, 'contain')
    ids = [row['id'] for row in report['failures']]
    assert run.returncode == 1 and 'image-unstable' in ids, report


def test_over_crop_is_still_judged_when_the_region_pixels_never_settle(project):
    run, report = measure_flicker(project, TALL, 'cover')
    ids = [row['id'] for row in report['failures']]
    assert 'image-over-crop' in ids and 'image-unstable' in ids, report
