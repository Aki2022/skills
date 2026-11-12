"""Every adopted (defined) item is verified: by code where the skill's map says code can, by AFM only for the rest."""
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

from test_core import project, write, check, approve, item, SKILL  # noqa: F401  (fixture reuse)
from classify import load_policy, classify_response
from item_checks import load_map, document_key, heading_path, L2_GATED, COLLECTED


def set_items(root, items):
    config = json.loads((root / 'design/design.yaml').read_text())
    config['items'] = items
    write(root / 'design/design.yaml', config)


def doc_section(root, anchor, body):
    path = root / 'design/DESIGN.md'
    path.write_text(path.read_text() + f'\n\n## {anchor}\n\n{body}\n')


def catalog_item(document, heading, grade='must'):
    """A fixture item that sits in a real catalog document/heading so the skill map applies to it."""
    row = item(grade=grade)
    row.update({'document': document + '/index.md', 'title': heading.split(' > ')[-1], 'headings': heading.split(' > ')})
    policy = load_policy()
    choice = 'global1-lp0-product0-presentation0' if grade == 'must' else 'global0-lp0-product0-presentation0'
    answer = {'type': 'choice', 'choice': choice, 'confidence': 1.0,
              'probabilities': {k: float(k == choice) for k in policy['outcomes']}}
    row['classification'] = classify_response(row, answer, policy['model'], policy)
    return row


def use_catalog_item(project, row):
    root, catalogs = project
    write(catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': [row]})
    doc_section(root, 'dads-000001', 'How this repository applies the rule.')
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['doc:DESIGN.md#dads-000001']}})


def add_tokens(root, extra):
    tokens = json.loads((root / 'design/tokens/brand.tokens.json').read_text())
    tokens.update(extra)
    write(root / 'design/tokens/brand.tokens.json', tokens)


def dimension(value, unit='px'):
    return {'$type': 'dimension', '$value': {'value': value, 'unit': unit},
            '$extensions': {'org.own.design': {'origin': 'proposal'}}}


def test_skill_map_names_real_catalog_headings_and_known_checkers():
    mapping = load_map()
    catalog = json.loads((SKILL / 'catalog/2026-10-14.json').read_text())
    paths = {(document_key(i), heading_path(i)) for i in catalog['items']}
    documents = {document_key(i) for i in catalog['items']}
    seen = set()
    for row in mapping['checks']:
        assert set(row['checkers']) <= L2_GATED | COLLECTED
        for path in row['paths']:
            assert (row['document'], path) in paths, (row['document'], path)
            assert (row['document'], path) not in seen, ('mapped twice', path)
            seen.add((row['document'], path))
    assert set(mapping['applicability']) <= documents


def test_color_token_item_is_machine_checked(project):
    result = check(project)
    assert result['manifest']['web']['coverage']['items'] == {'machine': ['dads-000001'], 'afm': []}
    assert result['afm_candidates'] == []


def test_doc_only_item_outside_the_map_goes_to_afm(project):
    root, _ = project
    doc_section(root, 'dads-000001', 'Group related controls with tight spacing.')
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['doc:DESIGN.md#dads-000001']}})
    result = check(project)
    assert result['manifest']['web']['coverage']['items'] == {'machine': [], 'afm': ['dads-000001']}
    [candidate] = result['afm_candidates']
    assert 'A synthetic fixture policy.' in candidate['rule']
    assert 'Group related controls with tight spacing.' in candidate['evidence']
    assert result['counts']['afm'] == 1 and not any('AFM' in w for w in result['warnings'])


def test_mapped_item_is_machine_when_its_tokens_exist_and_afm_when_not(project):
    root, _ = project
    use_catalog_item(project, catalog_item('foundations/spacing', '余白ルールの例'))
    assert check(project)['manifest']['web']['coverage']['items']['afm'] == ['dads-000001']
    add_tokens(root, {'space': {'md': dimension(16)}})
    manifest = check(project)['manifest']['web']
    assert manifest['coverage']['item_routes']['dads-000001'] == {
        'route': 'machine', 'checkers': ['palette:spacing'], 'document': 'foundations/spacing'}
    assert manifest['item_checks']['collect'] == ['palette:spacing']


@pytest.mark.parametrize('path,route', [
    ('余白ルールの例', 'machine'),
    ('余白の種類 > パディング', 'machine'),
    ('余白の種類', 'afm'),                                # parent mixes padding and margin
    ('余白の種類 > マージン', 'afm'),                      # margin is not measured
    ('余白ルールの例 > 余白による関係性の明示', 'afm'),     # meaning, not a value
])
def test_only_the_exact_heading_path_is_mapped(project, path, route):
    root, _ = project
    use_catalog_item(project, catalog_item('foundations/spacing', path))
    add_tokens(root, {'space': {'md': dimension(16)}})
    assert check(project)['manifest']['web']['coverage']['item_routes']['dads-000001']['route'] == route


def test_component_item_carries_its_presence_selectors(project):
    use_catalog_item(project, catalog_item('components/table', 'ユースケース'))
    manifest = check(project)['manifest']['web']
    assert manifest['coverage']['items']['afm'] == ['dads-000001']
    assert manifest['item_checks']['applicability'] == {'components/table': ['table']}


def test_nice_to_have_item_with_uncovered_rule_fails_ship(project):
    root, catalogs = project
    write(catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': [item(grade='nice-to-have')]})
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['rule:contrast-text']}})
    write(root / 'design/checks/web.yaml', {'rules': [{'id': 'contrast-text', 'kind': 'contrast',
        'targets': ['body'], 'minimum': 4.5}]})
    result = check(project, ship=True)
    assert result['exit_code'] == 2
    assert any('uncovered required evidence 1' in e for e in result['errors'])


PAGE = ('<html lang="en"><head><title>Fixture</title><meta name="viewport" content="width=device-width">'
        '<style>@font-face{font-family:Inter;src:url(/fixture.woff2)}body{font:16px/1.5 Inter,sans-serif;background:white;color:black}'
        '*{border-radius:0;border-color:black;outline-color:black;text-decoration-color:black}'
        ':focus-visible{outline:2px solid black}@media(min-width:1px){}</style></head><body><main><h1>Fixture</h1>BODY</main></body></html>')


@pytest.fixture
def rendered(project, tmp_path):
    """A served fixture page plus a stub AFM model; yields a function that runs design-check."""
    if os.environ.get('DESIGN_BROWSER_TEST') != '1':
        pytest.skip('explicit browser fixture run required')
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
    tokens['radius'] = dimension(0)
    tokens['size'] = dimension(32)
    tokens['body'] = dimension(16)
    write(root / 'design/tokens/brand.tokens.json', tokens)
    write(root / 'design/validation.json', {'schema_version': 1, 'html': ['index.html'],
          'urls': [{'env': 'DESIGN_FIXTURE_URL'}], 'viewports': [375], 'schemes': ['light']})
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

    def run(body='', **extra):
        (root / 'index.html').write_text(PAGE.replace('BODY', body))
        approve(project)
        done = subprocess.run(argv, env={**env, **extra}, capture_output=True, text=True)
        return done.returncode, json.loads(done.stdout)
    yield run, root, tmp_path
    server.shutdown()


def stages(result):
    return [s['stage'] for s in result['stages']]


def test_afm_judges_only_leftover_items_on_the_l2_capture(rendered, project):
    run, root, tmp_path = rendered
    doc_section(root, 'dads-000001', 'Group related controls with tight spacing.')
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['doc:DESIGN.md#dads-000001']}})
    code, result = run()
    assert code == 1, result
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS', 'AFM']
    assert result['stages'][-1]['judged'] == 1
    record = json.loads((root / 'design/.resolved/afm.json').read_text())
    assert record['pairs'] == [['dads-000001', 'DESIGN_FIXTURE_URL']]
    calls = [json.loads(line) for line in (tmp_path / 'args.json').read_text().splitlines()]
    images = [a for argv_ in calls for a in argv_ if a.endswith('.png')]
    assert len(images) == 1 and Path(images[0]).is_file()  # the stable L2 capture is what AFM judged
    (tmp_path / 'cache').rename(tmp_path / 'cache-old')
    code, result = run(STUB_VIOLATE='dads-000001')
    assert any(w.startswith('AFM advisory: dads-000001') for w in result['warnings']), result
    (tmp_path / 'cache').rename(tmp_path / 'cache-old2')
    code, result = run(STUB_MODE='drop')
    assert code == 2 and 'AFM left' in ' '.join(result['errors'])
    code, result = run(DESIGN_AFM_COMMAND=str(tmp_path / 'absent'))
    assert code == 2 and 'AFM' in ' '.join(result['errors'])


def test_absent_component_skips_afm_entirely(rendered, project):
    run, root, tmp_path = rendered
    use_catalog_item(project, catalog_item('components/table', 'ユースケース'))
    code, result = run()
    assert code in (0, 1), result
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS']
    assert result['stages'][-1]['not_applicable'] == 1 and result['stages'][-1]['afm_pairs'] == 0
    assert not (tmp_path / 'args.json').exists()  # the model was never asked
    code, result = run('<table><tr><td>1</td></tr></table>')
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS', 'AFM'], result


def test_mapped_item_is_enforced_by_code_and_names_the_item(rendered, project):
    run, root, tmp_path = rendered
    use_catalog_item(project, catalog_item('foundations/spacing', '余白ルールの例', grade='nice-to-have'))
    code, result = run('<p style="padding:16px">ok</p>')
    assert code in (0, 1), result
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS']
    items = json.loads((root / 'design/.resolved/items.json').read_text())
    assert items['machine'] == {'dads-000001': {'DESIGN_FIXTURE_URL': 'pass'}}
    code, result = run('<p style="padding:13px">off scale</p>')
    assert code == 2, result
    assert any(e.startswith('DESIGN_FIXTURE_URL: palette:spacing') and '13px' in e and e.endswith('breaks adopted items dads-000001')
               for e in result['errors']), result['errors']
    items = json.loads((root / 'design/.resolved/items.json').read_text())
    assert items['machine'] == {'dads-000001': {'DESIGN_FIXTURE_URL': 'fail'}}
