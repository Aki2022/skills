"""Screen conformance (SPEC-information-architecture-standardization v6, requirement 10; Quality 1).

The rendered page must show the outline's sections in order, each section's copy word for word and in order, each
copy inside its own section, and nothing outside the marks. Every condition has a red fixture (the table), checked on
DOM snapshots here and with a real browser in test_dom.
"""
import copy
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import dom  # noqa: E402
from test_check import OUTLINE, SKILL, SSOT  # noqa: E402
from test_dom import anchor, green_snapshot  # noqa: E402

TABLE = json.loads((SKILL / 'fixtures/screen-red-cases.json').read_text(encoding='utf-8'))
CONDITIONS = {c['id']: c for c in TABLE['conditions']}
CASES = TABLE['cases']


def apply(ops):
    snap = green_snapshot()
    for op in ops:
        if op['op'] == 'set':
            snap['anchors'][op['index']][op['field']] = op['value']
        elif op['op'] == 'append':
            a = op['anchor']
            snap['anchors'].append({**anchor(a['kind'], a['id'], a['top'], a['text']), 'section': a.get('section')})
        elif op['op'] == 'remove':
            snap['anchors'].pop(op['index'])
        elif op['op'] == 'unanchored':
            snap['unanchored'] = op['value']
        else:
            raise KeyError(op['op'])
    return snap


def fails(snap):
    return [f[0] for f in dom.compare(OUTLINE, SSOT, snap) if f[1] == 'fail']


def test_every_condition_has_a_red_case():
    covered = {c['condition'] for c in CASES}
    assert sorted(set(CONDITIONS) - covered) == [] and sorted(covered - set(CONDITIONS)) == []
    source = (SKILL / 'scripts/dom.py').read_text(encoding='utf-8')
    assert all(f"'{c['code']}'" in source for c in CONDITIONS.values())


def test_green_snapshot_conforms():
    assert dom.compare(OUTLINE, SSOT, green_snapshot()) == []


@pytest.mark.parametrize('case', CASES, ids=[c['id'] for c in CASES])
def test_red_case_fails_with_its_condition_code(case):
    assert CONDITIONS[case['condition']]['code'] in fails(apply(case['ops'])), dom.compare(OUTLINE, SSOT, apply(case['ops']))


def test_text_outside_the_marks_is_never_echoed():
    found = dom.compare(OUTLINE, SSOT, apply([{'op': 'unanchored', 'value': '業界No.1 の実績'}]))
    assert not any('業界' in f[2] for f in found)


def test_unknown_operation_is_rejected():
    with pytest.raises(KeyError):
        apply([{'op': 'nope'}])
