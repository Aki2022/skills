"""Product UI screens: build from a blueprint, then measure the drawn page (real browser, two widths).

Every red case is a change to a conforming page that the check must stop; the conforming page must pass. A page with no marks
is not checked (never a pass). The measuring tests need node and Playwright (own-design-author's playwright_runtime); when they
cannot run, the tests are SKIPPED and reported as such, never passed.
"""
import copy
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / 'own-web-build/scripts'))
for p in (ROOT / 'own-infoarch-author/scripts/tests', ROOT / 'own-positioning-author/scripts/tests', ROOT / 'own-infoarch-author/scripts',
          ROOT / 'own-positioning-author/scripts', ROOT / 'own-doc-update/scripts'):
    sys.path.insert(0, str(p))

import audit_screen  # noqa: E402
import build_screen  # noqa: E402
from blueprint_fixtures import OVERVIEW, PERSON, PRODUCT, write_blueprint  # noqa: E402


PROJECT = None  # set by the `browser` fixture: the project whose render-environment.json pins the browser
PW_HOME = Path(os.environ.get('OWN_PLAYWRIGHT_HOME') or Path.home() / '.cache' / 'own-playwright')


def pinned_project(tmp_path_factory):
    """A project directory whose design/render-environment.json pins the Playwright version installed on this machine
    (the fixed environment the measuring runs in). None when there is none."""
    versions = sorted(p.name for p in PW_HOME.iterdir() if (p / 'node_modules/playwright').is_dir()) if PW_HOME.is_dir() else []
    if not versions:
        return None
    project = tmp_path_factory.mktemp('pw-project')
    (project / 'design').mkdir()
    (project / 'design/render-environment.json').write_text(json.dumps({'runtime': {'playwright_version': versions[-1]}}), encoding='utf-8')
    return project


def can_measure(tmp_path_factory):
    if shutil.which('node') is None:
        return 'node is not installed'
    project = pinned_project(tmp_path_factory)
    if project is None:
        return f'no Playwright is installed under {PW_HOME}'
    page = tmp_path_factory.mktemp('probe') / 'probe.html'
    page.write_text('<html><body><div data-ia-region="rgn-0001">x</div></body></html>', encoding='utf-8')
    try:
        audit_screen.measure(page, repo=project)
    except Exception as exc:  # noqa: BLE001
        return f'real-browser measuring is unavailable: {str(exc)[:120]}'
    return None


@pytest.fixture(scope='session')
def browser(tmp_path_factory):
    reason = can_measure(tmp_path_factory)
    if reason:
        pytest.skip(reason)
    global PROJECT
    PROJECT = pinned_project(tmp_path_factory)
    return PROJECT


@pytest.fixture
def built(tmp_path):
    repo = tmp_path / 'repo'
    repo.mkdir()
    write_blueprint(repo)
    out = tmp_path / 'overview.html'
    proc = subprocess.run([sys.executable, str(ROOT / 'own-web-build/scripts/build_screen.py'), '--repo', str(repo), '--screen', 'scr-overview',
                           '--out', str(out)], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    return out.read_text(encoding='utf-8'), tmp_path


def run(page, tmp_path, screen=OVERVIEW, overrides=None, project=None):
    html = tmp_path / 'case.html'
    html.write_text(page, encoding='utf-8')
    return audit_screen.audit(screen, audit_screen.measure(html, repo=project or PROJECT), overrides or [])


def codes(findings, level='fail'):
    return sorted({c for lvl, c, _ in findings if lvl == level})


# ---- the build ---------------------------------------------------------------------------------------------------

def test_the_build_keeps_the_priority_order_and_marks_everything(built):
    page, _ = built
    regions = re.findall(r'data-ia-region="(rgn-\d+)"', page)
    assert regions == ['rgn-0001', 'rgn-0002', 'rgn-0003']
    assert page.count('data-ia-primary="true"') == 1 and 'class="btn primary"' in page
    assert all(f'data-ia-state="{s}"' in page for s in OVERVIEW['states'])
    assert '<style>' in page and 'grid-column:span' in page


def test_the_build_refuses_a_failing_blueprint_and_an_existing_file(tmp_path):
    repo = tmp_path / 'repo'
    repo.mkdir()
    bad = copy.deepcopy(OVERVIEW)
    bad['actions']['act-0002'] = {'text': '部署を見る', 'region': 'rgn-0002', 'primary': True}
    write_blueprint(repo, screens=(bad, PERSON))
    out = tmp_path / 'x.html'
    proc = subprocess.run([sys.executable, str(ROOT / 'own-web-build/scripts/build_screen.py'), '--repo', str(repo), '--screen', 'scr-overview',
                           '--out', str(out)], capture_output=True, text=True)
    assert proc.returncode == 2 and not out.exists()
    out.write_text('keep', encoding='utf-8')
    proc = subprocess.run([sys.executable, str(ROOT / 'own-web-build/scripts/build_screen.py'), '--repo', str(repo), '--screen', 'scr-person',
                           '--out', str(out)], capture_output=True, text=True)
    assert proc.returncode == 2 and out.read_text(encoding='utf-8') == 'keep'


def test_the_allocation_rules_cover_every_shape_the_schema_allows():
    schema = json.loads((ROOT / 'own-infoarch-author/references/schema/blueprint-screen.v1.schema.json').read_text(encoding='utf-8'))
    shapes = set(schema['properties']['regions']['additionalProperties']['properties']['shape']['enum'])
    rules = build_screen.load_rules()
    assert set(rules['width_by_shape']) == shapes  # a shape without a width rule would be placed by accident
    for rule in rules['width_by_shape'].values():
        assert {'desktop', 'tablet', 'mobile', 'part'} <= set(rule)


# ---- the measuring (real browser) ----------------------------------------------------------------------------------

def test_a_conforming_page_passes_at_both_widths(browser, built):
    page, tmp = built
    findings, stats = run(page, tmp)
    assert codes(findings) == [] and stats['viewports'] == ['desktop', 'mobile'], findings


def test_the_second_screen_passes_too(browser, tmp_path):
    repo = tmp_path / 'repo'
    repo.mkdir()
    write_blueprint(repo)
    out = tmp_path / 'person.html'
    subprocess.run([sys.executable, str(ROOT / 'own-web-build/scripts/build_screen.py'), '--repo', str(repo), '--screen', 'scr-person', '--out', str(out)], check=True, capture_output=True)
    findings, _ = run(out.read_text(encoding='utf-8'), tmp_path, PERSON)
    assert codes(findings) == [], findings


RED = {
    # the lower-priority region is drawn above the higher-priority one
    'priority-inverted': (lambda p: p.replace('</style>', '[data-ia-region="rgn-0003"]{order:-1;grid-column:1/-1}</style>'), 'SCREEN_ORDER'),
    'second-primary-painted': (lambda p: p.replace('<button type="button" class="btn primary" data-ia-action="act-0001"', '<button type="button" class="btn primary" data-ia-action="act-0009" data-ia-primary="true">もう 1 つ</button><button type="button" class="btn primary" data-ia-action="act-0001"'), 'SCREEN_PRIMARY_MANY'),
    'declared-state-absent': (lambda p: re.sub(r'<div data-ia-state="error"[^>]*>.*?</div>', '', p), 'SCREEN_STATE_MISSING'),
    'region-mark-removed': (lambda p: p.replace('data-ia-region="rgn-0003"', 'data-x="rgn-0003"'), 'SCREEN_REGION_MISSING'),
    'region-not-drawn': (lambda p: p.replace('</style>', '[data-ia-region="rgn-0003"]{display:none}</style>'), 'SCREEN_REGION_HIDDEN'),
    'primary-not-painted': (lambda p: p.replace('</style>', '.btn.primary{background:transparent}</style>'), 'SCREEN_PRIMARY_WRONG'),
    'phone-order-broken-only-on-the-phone': (lambda p: p.replace('</style>', '@media (max-width:639px){.grp>.region:last-of-type{order:-1}.grp{display:flex;flex-direction:column}}</style>'), 'SCREEN_ORDER'),
}


@pytest.mark.parametrize('name', sorted(RED))
def test_every_red_case_is_stopped(name, browser, built):
    page, tmp = built
    mutate, code = RED[name]
    broken = mutate(page)
    assert broken != page, 'the case did not change the page'
    findings, _ = run(broken, tmp)
    assert code in codes(findings), (name, findings)


def test_the_priority_order_is_kept_at_the_desktop_and_the_phone_width(browser, built):
    page, tmp = built
    html = tmp / 'p.html'
    html.write_text(page, encoding='utf-8')
    for snap in audit_screen.measure(html, repo=PROJECT):
        order = [r['id'] for r in sorted(snap['regions'], key=lambda r: (r['top'], r['left']))]
        assert order == sorted(order), (snap['viewport']['name'], order)
    desktop, phone = audit_screen.measure(html, repo=PROJECT)
    width = lambda snap, rid: next(r['width'] for r in snap['regions'] if r['id'] == rid)
    assert width(desktop, 'rgn-0001') < width(phone, 'rgn-0001')  # the width changed with the screen; the order did not


def test_a_page_with_no_marks_is_not_checked(browser, tmp_path):
    findings, _ = run('<html><body><main><p>画面</p></main></body></html>', tmp_path)
    assert codes(findings) == ['SCREEN_NOT_CHECKED']


def test_an_inversion_the_owner_decided_on_is_reported_not_failed(browser, built):
    page, tmp = built
    broken = RED['priority-inverted'][0](page)
    findings, _ = run(broken, tmp)
    inversions = [m for lvl, c, m in findings if c == 'SCREEN_ORDER']
    assert inversions
    pairs = set()
    for m in inversions:
        a, b = re.findall(r'rgn-\d+', m)[:2]
        pairs.add((a, b))
    declared = [{'screen': 'scr-overview', 'above': b, 'below': a, 'reason': 'オーナーが指定した'} for a, b in pairs]
    findings, _ = run(broken, tmp, overrides=declared)
    assert codes(findings) == [] and 'SCREEN_ORDER_OVERRIDDEN' in codes(findings, 'info')


def test_an_override_without_a_reason_is_refused(tmp_path):
    f = tmp_path / 'o.yaml'
    f.write_text('- {screen: scr-overview, above: rgn-0006, below: rgn-0002, reason: ""}\n', encoding='utf-8')
    with pytest.raises(ValueError):
        audit_screen.overrides_for(f, 'scr-overview')
