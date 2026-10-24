"""Endpoint profiles and native label hit areas use real Chromium measurements."""
import json
import os
import subprocess

import pytest
from test_core import SKILL, project, check, write

NODE = os.environ.get('DESIGN_TEST_NODE', 'node')
MEASURE = SKILL.parent / 'own-website-audit/scripts/measure_lp.mjs'
BROWSER = os.environ.get('DESIGN_BROWSER_TEST') == '1'
NORMAL = ('<!doctype html><html lang="en"><head><title>Product fixture</title><style>'
          'body{font:16px/1.5 Arial;background:white;color:black;margin:0}'
          '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
          'button{font:inherit;color:black;background:white;width:100px;height:48px}'
          'input{appearance:none;width:16px;height:16px;margin:0;border:1px solid black;background:white}'
          'label{display:flex;align-items:center;gap:8px;min-height:48px;width:200px}'
          '</style></head><body><main><h1>Product fixture</h1>CONTENT</main></body></html>')


def prepare(project):
    root, _ = project
    tokens = {'color': {'$type': 'color'},
              'font': {'$type': 'fontFamily', '$value': ['Arial']},
              'radius': {'$type': 'dimension', '$value': {'value': 0, 'unit': 'px'}}}
    for name, values, alpha in [('primary', [0, 0, 0], 1), ('white', [1, 1, 1], 1), ('clear', [0, 0, 0], 0)]:
        tokens['color'][name] = {'$value': {'colorSpace': 'srgb', 'components': values, 'alpha': alpha,
                                           'hex': '#ffffff' if name == 'white' else '#000000'}}
    for key, value in tokens.items():
        for token in ([value] if '$value' in value else [v for k, v in value.items() if not k.startswith('$')]):
            token['$extensions'] = {'org.own.design': {'origin': 'proposal'}}
    write(root / 'design/tokens/brand.tokens.json', tokens)
    assert check(project)['exit_code'] != 2
    return root


def measure(root, html, profile=None):
    page = root / 'index.html'; page.write_text(html)
    output = root / 'measurement.json'
    argv = [NODE, str(MEASURE), page.as_uri(), '--manifest', str(root / 'design/.resolved/web.json'),
            '--viewports', '375', '--schemes', 'light', '--json', str(output)]
    if profile is not None: argv += ['--profile', profile]
    run = subprocess.run(argv, cwd=root, env=os.environ.copy(), capture_output=True, text=True)
    return run, json.loads(output.read_text()) if output.is_file() else None


@pytest.mark.skipif(not BROWSER, reason='explicit real browser run required')
def test_product_ui_skips_only_marketing_fold_cta(project):
    root = prepare(project)
    html = NORMAL.replace('CONTENT', '<button data-cta style="margin-top:900px">Submit</button>')
    run, lp = measure(root, html)
    assert run.returncode == 1 and any(row['id'] == 'fold-cta' for row in lp['failures'])
    assert lp['profile'] == 'lp'
    run, ui = measure(root, html, 'product-ui')
    assert run.returncode == 0, run.stdout + run.stderr
    assert ui['profile'] == 'product-ui'
    assert ui['skipped'] == [{'id': 'fold-cta', 'reason': 'marketing first-fold CTA is not applicable to product-ui'}]
    run, broken = measure(root, html.replace('color:black;background:white', 'color:#777777;background:white'), 'product-ui')
    assert run.returncode == 1
    assert any(row['id'] == 'design-palette' for row in broken['failures'])
    assert any(row['id'] in ('contrast', 'axe-contrast') for row in broken['failures'])
    run, _ = measure(root, html, 'unknown')
    assert run.returncode == 2


@pytest.mark.skipif(not BROWSER, reason='explicit real browser run required')
def test_associated_contiguous_label_is_measured_and_real_click_toggles(project):
    root = prepare(project)
    label = '<label><input id="consent" type="checkbox">Consent</label>'
    run, result = measure(root, NORMAL.replace('CONTENT', label), 'product-ui')
    assert run.returncode == 0, run.stdout + run.stderr
    assert not result['runs'][0]['tapFails']
    script = ('import {createRequire} from "node:module";const require=createRequire(process.cwd()+"/");'
              'const {chromium}=require("playwright");const browser=await chromium.launch();'
              'try{const page=await browser.newPage();await page.goto(process.argv[1]);'
              'await page.locator("label").click({position:{x:150,y:24}});'
              'if(!await page.locator("#consent").isChecked())process.exitCode=1;'
              '}finally{await browser.close()}')
    click = subprocess.run([NODE, '--input-type=module', '-e', script, (root / 'index.html').as_uri()],
                           cwd=root, capture_output=True, text=True)
    assert click.returncode == 0, click.stderr
    # Offscreen label geometry is tested after scrolling and restored for fold evidence.
    run, result = measure(root, NORMAL.replace('CONTENT', '<div style="height:900px"></div>' + label), 'product-ui')
    assert run.returncode == 0, run.stdout + run.stderr
    assert not result['runs'][0]['tapFails'] and result['runs'][0]['fold']['h1Visible'] is True
    for bad in [
        '<input id="consent" type="checkbox"><label for="consent" style="margin-top:48px">Consent</label>',
        '<input id="consent" type="checkbox"><label for="consent" hidden style="display:none">Consent</label>',
        '<label><input id="consent" type="checkbox" style="pointer-events:none">Consent</label>'.replace('<label>', '<label style="pointer-events:none">'),
        '<input id="consent" type="checkbox"><label for="other">Consent</label>',
        '<label style="position:relative"><input id="consent" type="checkbox">Consent<span style="position:absolute;inset:0;pointer-events:auto" role="button"></span></label>',
    ]:
        run, result = measure(root, NORMAL.replace('CONTENT', bad), 'product-ui')
        assert run.returncode == 1
        assert any(row['id'] == 'tap-target' for row in result['failures']), bad


def test_ship_rejects_measurement_using_another_profile(tmp_path):
    from ship_contract import validate_web_result
    from test_ship_contract import result_fixture
    result = result_fixture(tmp_path)
    path = tmp_path / 'design/.resolved/l2-0.json'
    measured = json.loads(path.read_text())
    measured['profile'] = 'product-ui'
    write(path, measured)
    with pytest.raises(ValueError, match='profile'):
        validate_web_result(tmp_path, result, {'DESIGN_WEB_URL': 'http://127.0.0.1:12345/'})


def test_validation_endpoint_profiles_are_explicit_and_strict():
    from web_check import validation_endpoint
    assert validation_endpoint({'env':'DESIGN_WEB_URL'}) == ('DESIGN_WEB_URL', 'lp')
    assert validation_endpoint({'env':'DESIGN_WEB_URL','profile':'product-ui'}) == ('DESIGN_WEB_URL','product-ui')
    for invalid in [None, {}, {'env':'DESIGN_WEB_URL','profile':'unknown'}, {'env':'DESIGN_WEB_URL','profile':'lp','skip':True}]:
        with pytest.raises(ValueError): validation_endpoint(invalid)
