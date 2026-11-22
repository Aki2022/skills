"""audit_screen: how an approved layout-override row is matched against what was drawn (no browser needed).

SPEC-deliverable-production-layers requirement 4: a row allows the screen, the pair of regions and the widths it names and nothing
else, and the width cannot be left out (leaving it out used to mean every width). These cases feed `audit` hand-made snapshots, so
they run everywhere; the real-browser and CLI cases are in test_override_approvals.py.
"""
import copy
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / 'own-web-build/scripts'))
for p in (ROOT / 'own-infoarch-author/scripts/tests', ROOT / 'own-positioning-author/scripts/tests', ROOT / 'own-infoarch-author/scripts',
          ROOT / 'own-positioning-author/scripts', ROOT / 'own-doc-update/scripts'):
    sys.path.insert(0, str(p))

import audit_screen  # noqa: E402
from blueprint_fixtures import OVERVIEW  # noqa: E402


def snapshot(name, tops):
    """What measure() returns for one width: the three regions of OVERVIEW drawn at the given tops, the primary action painted, every state shown."""
    width = {'desktop': 1280, 'mobile': 375}[name]
    return {'viewport': {'name': name, 'width': width, 'height': 800}, 'screens': ['scr-overview'],
            'regions': [{'id': rid, 'visible': True, 'width': 300, 'height': 100, 'top': top} for rid, top in tops.items()],
            'actions': [{'id': 'act-0001', 'visible': True, 'painted': True, 'marked_primary': True}],
            'states': [{'id': sid, 'text': decl['text']} for sid, decl in OVERVIEW['states'].items()]}


IN_ORDER = {'rgn-0001': 0, 'rgn-0002': 200, 'rgn-0003': 400}
INVERTED = {'rgn-0001': 200, 'rgn-0002': 0, 'rgn-0003': 400}  # one inversion: priority 2 is drawn above priority 1
ROW = {'id': 'ovr-0001', 'screen': 'scr-overview', 'above': 'rgn-0002', 'below': 'rgn-0001', 'reason': '部署の比較を先に見たい'}


def run(rows, tops_by_width):
    findings, _ = audit_screen.audit(copy.deepcopy(OVERVIEW), [snapshot(n, t) for n, t in tops_by_width.items()], rows)
    return findings


def order_failures(findings):
    return sorted(m.split(':')[0] for lvl, c, m in findings if lvl == 'fail' and c == 'SCREEN_ORDER')


def test_a_conforming_page_needs_no_row():
    assert run([], {'desktop': IN_ORDER, 'mobile': IN_ORDER}) == []


def test_an_inversion_without_a_row_fails_at_every_width():
    findings = run([], {'desktop': INVERTED, 'mobile': INVERTED})
    assert {'desktop', 'mobile'} <= set(order_failures(findings))


def test_a_row_without_a_width_does_not_cover_any_width():
    """The first red case: `viewports` left out must not widen the row to every width (it used to: `o.get('viewports', [vp])`)."""
    row = {k: v for k, v in ROW.items()}
    assert 'viewports' not in row
    findings = run([row], {'desktop': INVERTED, 'mobile': INVERTED})
    assert {'desktop', 'mobile'} <= set(order_failures(findings))
    assert not [c for lvl, c, _ in findings if c == 'SCREEN_ORDER_OVERRIDDEN']


def test_a_row_with_an_empty_width_list_does_not_cover_any_width():
    findings = run([{**ROW, 'viewports': []}], {'desktop': INVERTED, 'mobile': INVERTED})
    assert {'desktop', 'mobile'} <= set(order_failures(findings))


def test_a_row_covers_only_the_widths_it_names():
    findings = run([{**ROW, 'viewports': ['desktop']}], {'desktop': INVERTED, 'mobile': INVERTED})
    assert order_failures(findings) == ['mobile']
    assert [(lvl, c) for lvl, c, _ in findings if c == 'SCREEN_ORDER_OVERRIDDEN'] == [('info', 'SCREEN_ORDER_OVERRIDDEN')]


def test_a_row_covers_only_the_pair_of_regions_it_names():
    other = {**ROW, 'viewports': ['desktop', 'mobile'], 'above': 'rgn-0003', 'below': 'rgn-0002'}  # not the pair that is inverted
    findings = run([other], {'desktop': INVERTED, 'mobile': INVERTED})
    assert order_failures(findings) == ['desktop', 'mobile']


def test_a_row_names_which_row_allowed_the_inversion():
    findings = run([{**ROW, 'viewports': ['desktop', 'mobile']}], {'desktop': INVERTED, 'mobile': INVERTED})
    assert order_failures(findings) == []
    assert all('ovr-0001' in m for lvl, c, m in findings if c == 'SCREEN_ORDER_OVERRIDDEN')
