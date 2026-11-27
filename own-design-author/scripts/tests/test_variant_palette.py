"""L2 checks colours against the variants the HTML declares, not against the union of every token.

A repository with two brands (variants) tags each brand token with `$extensions['org.own.design'].variant`; tokens
without the tag are shared. L0 writes a palette per variant into the manifest. The page declares its variants on
`<html data-design-variant="a">` (several are space-separated). With two or more variants and no declaration L2 stops.
"""
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import layer_paths as lp
from test_core import SKILL, check, project, write  # noqa: F401  (project is a fixture)

BROWSER = pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST') != '1',
                             reason='explicit browser fixture run required')


def color(hex_, **extra):
    n = [int(hex_[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    return {'$value': {'colorSpace': 'srgb', 'components': n, 'hex': hex_},
            '$extensions': {'org.own.design': {'origin': 'proposal', **extra}}}


def two_variant_tokens():
    return {
        'color': {'$type': 'color',
                  'primary': color('#000000'), 'white': color('#ffffff'),
                  'clear': {**color('#000000'), '$value': {**color('#000000')['$value'], 'alpha': 0}},
                  'brand-a': {'accent': color('#cc0000', variant='a')},
                  'brand-b': {'accent': color('#0000cc', variant='b')}},
        'font': {'$type': 'fontFamily', '$value': ['sans-serif'],
                 '$extensions': {'org.own.design': {'origin': 'proposal'}}},
        'radius': {'$type': 'dimension', '$value': {'value': 0, 'unit': 'px'},
                   '$extensions': {'org.own.design': {'origin': 'proposal'}}}}


def test_manifest_has_one_palette_per_variant_and_shared_tokens_are_in_all(project):
    root, _ = project
    write(lp.path(root, 'design_dir', 'tokens/brand.tokens.json'), two_variant_tokens())
    manifest = check(project)['manifest']['web']
    variants = manifest['variants']
    assert sorted(variants) == ['a', 'b']
    assert '#cc0000' in variants['a']['color'] and '#cc0000' not in variants['b']['color']
    assert '#0000cc' in variants['b']['color'] and '#0000cc' not in variants['a']['color']
    for name in variants:  # shared colours belong to every variant
        assert {'#000000', '#ffffff'} <= set(variants[name]['color'])
    assert {'#cc0000', '#0000cc'} <= set(manifest['palette']['color'])  # the union stays for single-variant readers


def test_a_repository_without_variants_has_no_variant_map(project):
    manifest = check(project)['manifest']['web']
    assert manifest['variants'] == {}


PAGE = ('<!doctype html><html lang="en"{decl}><head><title>Fixture</title>'
        '<style>body{{background:white;font:16px/1.5 sans-serif}}'
        '*{{color:black;border-color:black;outline-color:black;text-decoration-color:black;'
        'background-color:transparent;border-radius:0}}body{{background-color:white}}'
        '</style></head><body><main><h1>Fixture</h1><p>Readable <span style="color:{ink}">accent</span> text.</p>'
        '</main></body></html>')


@BROWSER
def test_l2_checks_against_the_declared_variant(project, tmp_path):
    root, _ = project
    node = os.environ['DESIGN_TEST_NODE']
    write(lp.path(root, 'design_dir', 'tokens/brand.tokens.json'), two_variant_tokens())
    assert check(project)['exit_code'] == 1
    manifest = lp.path(root, 'resolved', 'web.json')
    script = SKILL.parent / 'own-web-build/scripts/measure_render.mjs'
    page = root / 'index.html'

    def measure(html, name):
        page.write_text(html)
        output = tmp_path / (name + '.json')
        run = subprocess.run([node, str(script), page.as_uri(), '--manifest', str(manifest), '--profile', 'lp',
                              '--viewports', '375', '--schemes', 'light', '--json', str(output)],
                             cwd=SKILL.parent, capture_output=True, text=True)
        assert run.returncode in (0, 1), run.stderr
        return json.loads(output.read_text())

    def palette_failures(report):
        return [f for f in report['failures'] if f['id'] == 'design-palette']

    # Correct colours only: green.
    assert palette_failures(measure(PAGE.format(decl=' data-design-variant="b"', ink='#0000cc'), 'green')) == []
    # B declared, A's colour used: red on design-palette.
    wrong = palette_failures(measure(PAGE.format(decl=' data-design-variant="b"', ink='#cc0000'), 'wrong'))
    assert wrong and wrong[0]['sample'][0]['property'] == 'color'
    # Two variants, no declaration: stops, even though every colour is in the union.
    silent = palette_failures(measure(PAGE.format(decl='', ink='#cc0000'), 'silent'))
    assert silent and 'variant' in json.dumps(silent[0])
    # An undeclared-variant name is not a variant of this repository: stops.
    unknown = palette_failures(measure(PAGE.format(decl=' data-design-variant="c"', ink='#000000'), 'unknown'))
    assert unknown and 'variant' in json.dumps(unknown[0])
    # Both declared: a page that shows both brands may use both.
    both = measure(PAGE.format(decl=' data-design-variant="a b"', ink='#cc0000'), 'both')
    assert palette_failures(both) == []


@BROWSER
def test_single_variant_repository_needs_no_declaration(project, tmp_path):
    root, _ = project
    node = os.environ['DESIGN_TEST_NODE']
    tokens = two_variant_tokens()
    del tokens['color']['brand-b']
    write(lp.path(root, 'design_dir', 'tokens/brand.tokens.json'), tokens)
    assert check(project)['exit_code'] == 1
    page = root / 'index.html'
    page.write_text(PAGE.format(decl='', ink='#cc0000'))
    output = tmp_path / 'one.json'
    run = subprocess.run([node, str(SKILL.parent / 'own-web-build/scripts/measure_render.mjs'), page.as_uri(),
                          '--manifest', str(lp.path(root, 'resolved', 'web.json')), '--profile', 'lp',
                          '--viewports', '375', '--schemes', 'light', '--json', str(output)],
                         cwd=SKILL.parent, capture_output=True, text=True)
    assert run.returncode in (0, 1), run.stderr
    assert not [f for f in json.loads(output.read_text())['failures'] if f['id'] == 'design-palette']
