"""audit_screen through its real CLI, in a tmp repository, against the owner's approval of each layout-override row.

SPEC-deliverable-production-layers requirement 4. The page is drawn in a real browser at two widths; an inversion passes only through a
row that holds the owner's current approval, for the screen, the pair of regions and the width it names. The approval is recorded the real
way (design-author approve, with the owner's words and date); nothing here writes the ledger by hand. When no browser can run, these tests
are SKIPPED and reported as such (never passed).

The matching alone (no browser) is test_override_matching.py; the rows' own red and green cases are own-design-author's
test_layout_override_approvals.py.
"""
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

ROOT = Path(__file__).resolve().parents[3]
for p in (ROOT / 'own-design-author/scripts/tests', ROOT / 'own-infoarch-author/scripts/tests', ROOT / 'own-positioning-author/scripts/tests',
          ROOT / 'own-infoarch-author/scripts', ROOT / 'own-positioning-author/scripts', ROOT / 'own-doc-update/scripts', ROOT / 'own-design-author/scripts'):
    sys.path.insert(0, str(p))

import layer_paths  # noqa: E402
import layout_overrides as lo  # noqa: E402
from test_core import approve as approve_tokens, project  # noqa: E402,F401  (project is a fixture: a design repository with tmp catalogs)
from test_ia_bridge import attach_blueprint  # noqa: E402  (the real blueprint fixture, the IA module on)
from test_layout_override_approvals import WORDS, DAY, dump, record  # noqa: E402  (the real CLI recording)
from test_screen_conformance import RED, browser  # noqa: E402,F401  (browser is a fixture: skips when no browser can measure)


SCREEN_ID = 'scr-overview'


def audit_cli(repo, html, overrides=True):
    cmd = [sys.executable, str(ROOT / 'own-web-build/scripts/audit_screen.py'), '--screen',
           str(layer_paths.path(repo, 'infoarch', 'blueprint', 'screens', f'{SCREEN_ID}.yaml')), '--html', str(html), '--repo', str(repo)]
    if overrides:
        cmd += ['--overrides', str(layer_paths.path(repo, 'layout_overrides'))]
    done = subprocess.run(cmd, capture_output=True, text=True, timeout=180, env={**os.environ, 'DESIGN_AUTHOR_PYTHON': sys.executable})
    out = json.loads(done.stdout)
    assert done.returncode == out['exit_code'], done.stdout + done.stderr
    return out


def findings(out, code, level='fail'):
    return [f['message'] for f in out.get('findings', []) if f['code'] == code and f['level'] == level]


@pytest.fixture
def repo(browser, project, tmp_path):
    """A design repository with the blueprint fixture and a built page of scr-overview; the browser pinned like the render environment."""
    root, catalogs = project
    approve_tokens(project)
    attach_blueprint(project)
    built = tmp_path / 'overview.html'
    made = subprocess.run([sys.executable, str(ROOT / 'own-web-build/scripts/build_screen.py'), '--repo', str(root), '--screen', SCREEN_ID, '--out', str(built)],
                          capture_output=True, text=True)
    assert made.returncode == 0, made.stdout + made.stderr
    target = layer_paths.path(root, 'render_environment')
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(layer_paths.path(browser, 'render_environment'), target)
    return root, catalogs, built.read_text(encoding='utf-8'), tmp_path


def page_with(repo, name):
    root, _, page, tmp = repo
    broken = RED[name][0](page)
    assert broken != page
    path = tmp / f'{name}.html'
    path.write_text(broken, encoding='utf-8')
    return path


def inverted_pairs(out):
    """[(above, below)] the rows must name, from the failures of an audit with no row: the lower-priority region is above the higher one."""
    return sorted({tuple(reversed(re.findall(r'rgn-\d+', m)[:2])) for m in findings(out, 'SCREEN_ORDER')})


def write_rows(root, pairs, viewports, first=1):
    rows = [{'id': f'ovr-{n:04d}', 'screen': SCREEN_ID, 'above': a, 'below': b, 'viewports': viewports, 'reason': f'{a} を先に見たいと、オーナーが指定した'}
            for n, (a, b) in enumerate(pairs, first)]
    dump(layer_paths.path(root, 'layout_overrides'), rows)
    return rows


def test_the_whole_path_request_page_the_owners_words_and_a_passing_audit(repo):
    root, catalogs, _, tmp = repo
    html = page_with(repo, 'priority-inverted')
    bare = audit_cli(root, html, overrides=False)
    pairs = inverted_pairs(bare)
    assert bare['exit_code'] == 2 and pairs, 'the page was not inverted'

    rows = write_rows(root, pairs, ['desktop', 'mobile'])
    unapproved = audit_cli(root, html)
    assert unapproved['exit_code'] == 2
    assert len(findings(unapproved, 'OVERRIDE_UNAPPROVED')) == len(rows)
    assert findings(unapproved, 'SCREEN_ORDER') and not findings(unapproved, 'SCREEN_ORDER_OVERRIDDEN', 'info')  # a row nobody approved allows nothing

    out = tmp / 'request'
    made = subprocess.run([sys.executable, str(ROOT / 'own-design-author/scripts/design_author.py'), 'override-request', '--repo', str(root), '--out', str(out)],
                          capture_output=True, text=True)
    assert made.returncode == 0, made.stdout + made.stderr
    page = (out / 'approval-request.html').read_text(encoding='utf-8')
    assert all(lo.key_for(r['id']) in page for r in rows)

    record(project_of(repo), tmp, *[r['id'] for r in rows])
    passed = audit_cli(root, html)
    assert passed['exit_code'] == 0, passed
    assert len(findings(passed, 'SCREEN_ORDER_OVERRIDDEN', 'info')) >= len(rows)
    assert all('ovr-' in m for m in findings(passed, 'SCREEN_ORDER_OVERRIDDEN', 'info'))


def project_of(repo):
    root, catalogs, _, _ = repo
    return root, catalogs


def approved_repo(repo, name='priority-inverted', viewports=('desktop', 'mobile')):
    root, _, _, tmp = repo
    html = page_with(repo, name)
    pairs = inverted_pairs(audit_cli(root, html, overrides=False))
    rows = write_rows(root, pairs, list(viewports))
    record(project_of(repo), tmp, *[r['id'] for r in rows])
    assert audit_cli(root, html)['exit_code'] == 0
    return root, html, rows


def test_changing_one_thing_in_an_approved_row_fails_the_audit_for_that_row_only(repo):
    root, html, rows = approved_repo(repo)
    assert len(rows) >= 1
    edited = yaml.safe_load(layer_paths.path(root, 'layout_overrides').read_text(encoding='utf-8'))
    edited[0]['reason'] = '別の理由に書き換えた'
    dump(layer_paths.path(root, 'layout_overrides'), edited)
    out = audit_cli(root, html)
    assert out['exit_code'] == 2
    assert [m.split(':')[0] for m in findings(out, 'OVERRIDE_UNAPPROVED')] == [edited[0]['id']]
    pair = (edited[0]['above'], edited[0]['below'])
    assert any(pair[0] in m and pair[1] in m for m in findings(out, 'SCREEN_ORDER'))  # its own pair is a failure again
    for other in edited[1:]:  # the other rows still allow theirs
        assert not any(other['above'] in m and other['below'] in m for m in findings(out, 'SCREEN_ORDER'))


def test_a_row_without_a_width_fails_and_allows_no_width(repo):
    root, html, rows = approved_repo(repo)
    bare = {m.split(':')[0] for m in findings(audit_cli(root, html, overrides=False), 'SCREEN_ORDER')}
    assert bare  # the widths at which the page is inverted
    edited = yaml.safe_load(layer_paths.path(root, 'layout_overrides').read_text(encoding='utf-8'))
    for r in edited:
        r.pop('viewports')
    dump(layer_paths.path(root, 'layout_overrides'), edited)
    out = audit_cli(root, html)
    assert out['exit_code'] == 2
    assert findings(out, 'OVERRIDE_NO_WIDTH') and not findings(out, 'SCREEN_ORDER_OVERRIDDEN', 'info')
    assert {m.split(':')[0] for m in findings(out, 'SCREEN_ORDER')} == bare  # a row without a width used to cover every width; it covers none


def test_a_row_for_the_desktop_does_not_allow_an_inversion_that_only_the_phone_has(repo):
    root, _, _, tmp = repo
    html = page_with(repo, 'phone-order-broken-only-on-the-phone')
    pairs = inverted_pairs(audit_cli(root, html, overrides=False))
    assert pairs
    rows = write_rows(root, pairs, ['desktop'])
    record(project_of(repo), tmp, *[r['id'] for r in rows])
    out = audit_cli(root, html)
    assert out['exit_code'] == 2 and {m.split(':')[0] for m in findings(out, 'SCREEN_ORDER')} == {'mobile'}
    # the same rows, widened to the phone and approved again, pass
    rows = write_rows(root, pairs, ['mobile'], first=len(rows) + 1)
    record(project_of(repo), tmp, *[r['id'] for r in rows])
    assert audit_cli(root, html)['exit_code'] == 0


def test_a_region_changing_under_an_approved_row_fails_the_audit(repo):
    root, html, rows = approved_repo(repo)
    path = layer_paths.path(root, 'infoarch', 'blueprint', 'screens', f'{SCREEN_ID}.yaml')
    screen = yaml.safe_load(path.read_text(encoding='utf-8'))
    screen['regions'][rows[0]['above']]['text'] = '領域の文言を変えた'
    dump(path, screen)
    out = audit_cli(root, html)
    assert out['exit_code'] == 2 and findings(out, 'OVERRIDE_UNAPPROVED')


def test_an_override_file_that_the_audit_is_not_given_allows_nothing(repo):
    root, html, rows = approved_repo(repo)
    out = audit_cli(root, html, overrides=False)
    assert out['exit_code'] == 2 and findings(out, 'SCREEN_ORDER')


def test_an_approved_row_does_not_cover_another_screens_audit(repo):
    """The rows are about one screen: a row of scr-person approved or not, the audit of scr-overview does not read it as its own."""
    root, html, rows = approved_repo(repo)
    edited = yaml.safe_load(layer_paths.path(root, 'layout_overrides').read_text(encoding='utf-8'))
    edited.append({'id': 'ovr-0090', 'screen': 'scr-person', 'above': 'rgn-0002', 'below': 'rgn-0001', 'viewports': ['desktop'], 'reason': '本人の推移を先に見たい'})
    dump(layer_paths.path(root, 'layout_overrides'), edited)
    out = audit_cli(root, html)
    assert out['exit_code'] == 0 and not findings(out, 'OVERRIDE_UNAPPROVED')  # scr-person's unapproved row is the design check's concern here
