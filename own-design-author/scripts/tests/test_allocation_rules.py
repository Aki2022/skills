"""The allocation rules cover what the infoarch layer can say: every shape, and the order that does not change with the width."""
import json
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
RULES = json.loads((SKILLS / 'own-design-author/references/allocation-rules.json').read_text(encoding='utf-8'))


def test_every_shape_the_blueprint_schema_allows_has_a_width_and_a_part():
    schema = json.loads((SKILLS / 'own-infoarch-author/references/schema/blueprint-screen.v1.schema.json').read_text(encoding='utf-8'))
    shapes = set(schema['properties']['regions']['additionalProperties']['properties']['shape']['enum'])
    assert set(RULES['width_by_shape']) == shapes
    names = {b['name'] for b in RULES['breakpoints']}
    for shape, rule in RULES['width_by_shape'].items():
        assert names <= set(rule), shape
        assert all(1 <= rule[n] <= RULES['grid']['columns'] for n in names), shape
        assert rule['part']


def test_the_order_rule_does_not_depend_on_the_width():
    assert RULES['order']['vertical'] == 'priority-ascending' and 'width' not in json.dumps(RULES['order'], ensure_ascii=False).replace('画面幅で変わらない', '')
    mobile = next(b for b in RULES['breakpoints'] if b['name'] == 'mobile')
    assert mobile['columns_used'] == 1 and all(r['mobile'] == 1 for r in RULES['width_by_shape'].values())


def test_the_override_notation_requires_a_reason_and_is_written_down():
    layout = json.loads((SKILLS / 'own-design-author/references/layer-paths.json').read_text(encoding='utf-8'))['entries']
    assert 'file' not in RULES['override'] and layout[RULES['override']['file_entry']] == 'design/layout-overrides.yaml'  # a key into layer-paths.json, not a path
    assert 'reason' in RULES['override']['entry']
    text = (SKILLS / 'own-design-author/references/allocation-rules.md').read_text(encoding='utf-8')
    assert 'design/layout-overrides.yaml' in text and '理由' in text and all(t in text for t in ('dashboard', 'detail', 'list', 'form'))
