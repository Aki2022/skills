"""Every adopted (defined) item is verified by a machine check or an AFM judgment, must or not."""
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

from test_core import project, write, check, approve, item, SKILL  # noqa: F401  (fixture reuse)


def set_items(root, items):
    config = json.loads((root / 'design/design.yaml').read_text())
    config['items'] = items
    write(root / 'design/design.yaml', config)


def doc_section(root, anchor, body):
    path = root / 'design/DESIGN.md'
    path.write_text(path.read_text() + f'\n\n## {anchor}\n\n{body}\n')


def test_color_token_item_is_machine_checked(project):
    result = check(project)
    assert result['manifest']['web']['coverage']['items'] == {'machine': ['dads-000001'], 'afm': []}
    assert result['afm_candidates'] == []


def test_doc_only_item_is_routed_to_afm_with_rule_and_evidence(project):
    root, _ = project
    doc_section(root, 'dads-000001', 'Group related controls with tight spacing.')
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['doc:DESIGN.md#dads-000001']}})
    result = check(project)
    assert result['manifest']['web']['coverage']['items'] == {'machine': [], 'afm': ['dads-000001']}
    [candidate] = result['afm_candidates']
    assert candidate['id'] == 'dads-000001'
    assert 'A synthetic fixture policy.' in candidate['rule']
    assert 'Group related controls with tight spacing.' in candidate['evidence']
    assert result['counts']['afm'] == 1
    assert not any('AFM' in w for w in result['warnings'])


def test_dimension_token_alone_is_not_machine_checked(project):
    root, _ = project
    write(root / 'design/tokens/space.tokens.json', {'space': {'$type': 'dimension', 'md': {
        '$value': {'value': 16, 'unit': 'px'}, '$extensions': {'org.own.design': {'origin': 'proposal'}}}}})
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['token:space.md']}})
    assert check(project)['manifest']['web']['coverage']['items']['afm'] == ['dads-000001']


def test_nice_to_have_item_with_uncovered_rule_fails_ship(project):
    root, catalogs = project
    write(catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': [item(grade='nice-to-have')]})
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['rule:contrast-text']}})
    write(root / 'design/checks/web.yaml', {'rules': [{'id': 'contrast-text', 'kind': 'contrast',
        'targets': ['body'], 'minimum': 4.5}]})
    result = check(project, ship=True)
    assert result['exit_code'] == 2
    assert any('uncovered required evidence 1' in e for e in result['errors'])


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1', reason='explicit browser fixture run required')
def test_design_check_runs_afm_on_the_l2_capture_and_html(project, tmp_path):
    from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
    from functools import partial
    from threading import Thread
    from test_fixed_render import pin
    from test_afm_judge import STUB
    root, catalogs = project
    env = dict(os.environ, DESIGN_AUDIT_NODE=os.environ['DESIGN_TEST_NODE'])
    if os.environ.get('DESIGN_TEST_NODE_MODULES'):
        env['NODE_PATH'] = os.environ['DESIGN_TEST_NODE_MODULES']
    write(root / 'design/render-environment.json', pin(root, env))
    tokens = json.loads((root / 'design/tokens/brand.tokens.json').read_text())
    for name, components, hex_value, alpha in [('white', [1, 1, 1], '#ffffff', 1), ('clear', [0, 0, 0], '#000000', 0)]:
        tokens['color'][name] = {'$value': {'colorSpace': 'srgb', 'components': components, 'hex': hex_value, 'alpha': alpha},
                                 '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    tokens['font'] = {'$type': 'fontFamily', '$value': ['Inter', 'sans-serif'], '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    tokens['radius'] = {'$type': 'dimension', '$value': {'value': 0, 'unit': 'px'}, '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    write(root / 'design/tokens/brand.tokens.json', tokens)
    doc_section(root, 'dads-000001', 'Group related controls with tight spacing.')
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['doc:DESIGN.md#dads-000001']}})
    approve(project)
    write(root / 'design/validation.json', {'schema_version': 1, 'html': ['index.html'],
          'urls': [{'env': 'DESIGN_FIXTURE_URL'}], 'viewports': [375], 'schemes': ['light']})
    (root / 'index.html').write_text(
        '<html lang="en"><head><title>Fixture</title><meta name="viewport" content="width=device-width">'
        '<style>@font-face{font-family:Inter;src:url(/fixture.woff2)}body{font:16px/1.5 Inter,sans-serif;background:white;color:black}'
        '*{border-radius:0;border-color:black;outline-color:black;text-decoration-color:black}'
        ':focus-visible{outline:2px solid black}@media(min-width:1px){}</style></head><body><main><h1>Fixture</h1></main></body></html>')
    stub = tmp_path / 'fm'
    stub.write_text(STUB.replace("if mode == 'drop'", "open(os.environ['STUB_ARGS'], 'a').write(json.dumps(sys.argv) + '\\n')\nif mode == 'drop'"))
    stub.chmod(0o755)
    env.update(DESIGN_AFM_COMMAND=str(stub), DESIGN_AFM_CACHE=str(tmp_path / 'cache'),
               STUB_CALLS=str(tmp_path / 'calls'), STUB_ARGS=str(tmp_path / 'args.json'))

    class Handler(SimpleHTTPRequestHandler):
        def log_message(self, *args):
            pass
    server = ThreadingHTTPServer(('127.0.0.1', 0), partial(Handler, directory=str(root)))
    Thread(target=server.serve_forever, daemon=True).start()
    env['DESIGN_FIXTURE_URL'] = 'http://127.0.0.1:' + str(server.server_port)
    argv = [sys.executable, str(SKILL / 'scripts/web_check.py'), '--repo', str(root), '--catalog-dir', str(catalogs)]
    try:
        run = subprocess.run(argv, env=env, capture_output=True, text=True)
        result = json.loads(run.stdout)
        assert run.returncode == 1, result
        assert [s['stage'] for s in result['stages']] == ['L0', 'L1', 'L2', 'AFM'], (result['errors'], run.stderr[-2000:])
        assert result['stages'][-1]['judged'] == 2
        record = json.loads((root / 'design/.resolved/afm.json').read_text())
        assert record['artifacts'] == ['L2:DESIGN_FIXTURE_URL:375:light', 'html:index.html']
        calls = [json.loads(line) for line in (tmp_path / 'args.json').read_text().splitlines()]
        images = [a for argv_ in calls for a in argv_ if a.endswith('.png')]
        assert len(images) == 1 and Path(images[0]).is_file()  # the stable L2 capture is what AFM judged
        env['STUB_VIOLATE'] = 'dads-000001'
        (tmp_path / 'cache').rename(tmp_path / 'cache-old')
        result = json.loads(subprocess.run(argv, env=env, capture_output=True, text=True).stdout)
        assert any(w.startswith('AFM advisory: dads-000001') for w in result['warnings']), result
        env['STUB_MODE'] = 'drop'
        (tmp_path / 'cache').rename(tmp_path / 'cache-old2')
        run = subprocess.run(argv, env=env, capture_output=True, text=True)
        assert run.returncode == 2 and 'AFM left' in ' '.join(json.loads(run.stdout)['errors'])
        env['DESIGN_AFM_COMMAND'] = str(tmp_path / 'absent')
        run = subprocess.run(argv, env=env, capture_output=True, text=True)
        assert run.returncode == 2 and 'AFM' in ' '.join(json.loads(run.stdout)['errors'])
    finally:
        server.shutdown()
