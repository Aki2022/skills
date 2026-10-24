"""Computed token membership is checked by a real Chromium renderer."""
import json
import os
import subprocess
from pathlib import Path

import pytest

NODE = os.environ.get('DESIGN_TEST_NODE', 'node')
MODULE = (Path(__file__).resolve().parents[1] / 'resolved_contract.mjs').as_uri()


def collect(html, palette):
    script = ('import {createRequire} from "node:module";'
              'import {collectDesign} from ' + json.dumps(MODULE) + ';'
              'const require=createRequire(process.cwd()+"/");'
              'const {chromium}=require("playwright");'
              'const browser=await chromium.launch({headless:true});'
              'try{const page=await browser.newPage();'
              'await page.setContent(JSON.parse(process.argv[1]));'
              'console.log(JSON.stringify(await page.evaluate(collectDesign, '
              '{palette:JSON.parse(process.argv[2])})));'
              '}finally{await browser.close()}')
    run = subprocess.run([NODE, '--input-type=module', '-e', script, json.dumps(html), json.dumps(palette)],
                         env=os.environ.copy(), capture_output=True, text=True)
    return run, json.loads(run.stdout) if run.returncode == 0 else None


PALETTE = {'color': ['#000000', '#ffffff', 'transparent'],
           'fontFamily': ['Arial'], 'dimension': ['-20px', '0px', '8px']}
NORMAL = ('<style>body{font-family:Arial;color:black;background:white}'
          '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
          '</style><main><p>Visible text</p></main>')


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')
def test_negative_transform_dimension_does_not_break_radius_palette():
    run, result = collect(NORMAL, PALETTE)
    assert run.returncode == 0, run.stderr
    assert result['targetCount'] > 0 and result['failures'] == []
    run, result = collect(NORMAL.replace('border-radius:0', 'border-radius:7px'), PALETTE)
    assert run.returncode == 0, run.stderr
    assert any(row['property'] == 'borderTopLeftRadius' for row in result['failures'])
    run, result = collect(NORMAL.replace('color:black;background', 'color:rebeccapurple;background'), PALETTE)
    assert run.returncode == 0, run.stderr
    assert any(row['property'] == 'color' for row in result['failures'])


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')
@pytest.mark.parametrize('kind,value', [('dimension', 'not-a-length'), ('color', 'not-a-color'), ('fontFamily', '!!!')])
def test_invalid_palette_values_still_fail(kind, value):
    palette = {key: list(values) for key, values in PALETTE.items()}
    palette[kind].append(value)
    run, _ = collect(NORMAL, palette)
    assert run.returncode != 0 and 'invalid palette CSS value' in run.stderr


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')
def test_only_painted_properties_count_but_visible_mutations_fail():
    html = NORMAL + ('<style>.unused{border:0 solid rebeccapurple;outline:0 solid rebeccapurple;'
                     'text-decoration:none rebeccapurple;background:rgba(99,88,77,0)}'
                     '.hidden{display:none;color:rebeccapurple;font-family:serif;border-radius:7px}'
                     '.faded{opacity:0;color:rebeccapurple;font-family:serif}'
                     '.wrapper{color:rebeccapurple;font-family:serif}</style>'
                     '<div class="unused">Allowed visible text</div>'
                     '<div class="hidden"><p>Hidden text</p></div>'
                     '<div class="faded"><p>Faded descendant</p></div>'
                     '<div class="wrapper"><span style="font-family:Arial;color:black">Allowed child</span></div>')
    run, result = collect(html, PALETTE)
    assert run.returncode == 0, run.stderr
    assert result['failures'] == [], result['failures']
    for before, after, property_name in [
        ('border:0 solid', 'border:1px solid', 'borderTopColor'),
        ('outline:0 solid', 'outline:1px solid', 'outlineColor'),
        ('text-decoration:none', 'text-decoration:underline', 'textDecorationColor'),
        ('rgba(99,88,77,0)', 'rgba(99,88,77,.5)', 'backgroundColor'),
        ('display:none', 'display:block', 'color'),
        ('opacity:0', 'opacity:1', 'fontFamily'),
    ]:
        run, result = collect(html.replace(before, after), PALETTE)
        assert run.returncode == 0, run.stderr
        assert any(row['property'] == property_name for row in result['failures']), (after, result)


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')
def test_hidden_only_dom_is_not_reported_as_measured_targets():
    run, result = collect('<style>body{font-family:Arial}</style><div hidden>Unavailable content</div>', PALETTE)
    assert run.returncode == 0, run.stderr
    assert result['targetCount'] == 0


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit real browser run required')
def test_palette_probe_is_not_animated_by_application_css():
    html = NORMAL + '<style>*{transition-duration:10s!important;animation-duration:10s!important}</style>'
    run, result = collect(html, PALETTE)
    assert run.returncode == 0, run.stderr
    assert result['failures'] == [], result['failures']
