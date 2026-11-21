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
import audit_html  # noqa: E402
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


def test_the_built_page_passes_every_structural_constraint_of_the_html_media(built):
    """own-web-build must satisfy media/html `html-*` before shipping: a built screen that fails them would make the gate red on a
    page the skill itself produced (focus visibility was missing, found when a fresh session ran the whole table)."""
    page, _ = built
    results = audit_html.check_structure(page)
    assert results, 'the structural check ran on nothing'
    assert [label for ok, label in results if not ok] == []


def test_the_built_page_meets_the_text_contrast_in_the_light_and_the_dark_theme(built):
    """What the build draws in the light theme is pinned twice: by the page's own `:root` values (body text, the outlined action and
    the filled primary action, in both themes) and by audit_html's L1 contrast check, which judges the base and the dark
    `prefers-color-scheme` theme separately (it used to resolve `:root` last-wins and judged only the dark pairs; found when the L1
    advisories of built screens were read)."""
    page, _ = built
    css = re.search(r'<style>(.*?)</style>', page, re.S).group(1)
    base = re.search(r'(?<![\w-])(?<!\{):root\{([^}]*)\}', css).group(1)
    dark = re.search(r'prefers-color-scheme:dark\)\{:root\{([^}]*)\}', css).group(1)
    minimum = 4.5
    for name, block in (('light', base), ('dark', dark)):
        v = dict(re.findall(r'--([\w-]+):([^;]+)', block))
        pairs = {'body': (v['fg'], v['bg']), 'outlined action': (v['accent'], v['bg']), 'primary action': (v['bg'], v['accent'])}
        assert len(pairs) == 3
        for what, (fg, bg) in pairs.items():
            ratio = audit_html.contrast_ratio(audit_html.parse_color(fg), audit_html.parse_color(bg))
            assert ratio >= minimum, f'{name} theme, {what}: {fg} on {bg} = {ratio:.2f}:1 (min {minimum})'
    results = audit_html.check_contrast(page, minimum)
    assert [label for ok, label in results if not ok] == []
    judged = ' '.join(label for _, label in results)
    # L1 judges a pair only when one rule sets both colours: body text and the filled primary action (the outlined action has no background)
    for fg, bg in (('fg', 'bg'), ('bg', 'accent')):
        for block in (base, dark):
            v = dict(re.findall(r'--([\w-]+):([^;]+)', block))
            assert f'{v[fg]} on {v[bg]}' in judged, f'L1 did not judge {v[fg]} on {v[bg]}'


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


# ---- the interpreter ---------------------------------------------------------------------------------------------
# `python3 audit_screen.py ...` is how the SKILL tells people to run it, and the python3 on PATH may not have PyYAML. The L0
# dependency interpreter (DESIGN_AUTHOR_PYTHON) has it, and build_screen already uses it; audit_screen must too, or it stops
# (exit 2) on a command that is written exactly as the SKILL says.

def _run_without_yaml(tmp_path, env_extra=None, drop=()):
    """Run audit_screen.py as a script in an interpreter where `import yaml` fails, like a python3 without PyYAML."""
    script = ROOT / 'own-web-build/scripts/audit_screen.py'
    screen = tmp_path / 'screen.yaml'
    screen.write_text('id: scr-x\n', encoding='utf-8')
    hide = ("import sys, runpy; sys.modules['yaml'] = None; sys.argv = ['audit_screen.py'] + sys.argv[1:]; "
            f"runpy.run_path({str(script)!r}, run_name='__main__')")
    env = {k: v for k, v in os.environ.items() if k not in ('DESIGN_AUTHOR_PYTHON', 'AUDIT_SCREEN_REEXEC', *drop)}
    env.update(env_extra or {})
    return subprocess.run([sys.executable, '-c', hide, '--screen', str(screen), '--html', str(tmp_path / 'absent.html')],
                          capture_output=True, text=True, env=env, timeout=120)


def test_audit_screen_starts_itself_again_under_the_dependency_interpreter(tmp_path):
    proc = _run_without_yaml(tmp_path, {'DESIGN_AUTHOR_PYTHON': sys.executable})
    assert 'PyYAML' not in proc.stdout + proc.stderr, proc.stdout + proc.stderr


def test_audit_screen_without_the_dependency_interpreter_stops_and_says_which_variable(tmp_path):
    proc = _run_without_yaml(tmp_path)
    assert proc.returncode == 2 and 'PyYAML' in proc.stdout and 'DESIGN_AUTHOR_PYTHON' in proc.stdout, proc.stdout + proc.stderr


def test_audit_screen_does_not_start_itself_again_in_a_loop(tmp_path):
    proc = _run_without_yaml(tmp_path, {'DESIGN_AUTHOR_PYTHON': sys.executable, 'AUDIT_SCREEN_REEXEC': '1'})
    assert proc.returncode == 2 and 'PyYAML' in proc.stdout, proc.stdout + proc.stderr


# ---- the dependency interpreter is broken ------------------------------------------------------------------------
# DESIGN_AUTHOR_PYTHON may point at something that cannot run the script: a path that does not exist, a directory, a file that
# is not executable, or a python without PyYAML. Every one of them must stop with exit 2 and name the variable, and must never
# loop. Stopping is the safe side; a traceback (exit 1) or a hang is not.

def _yamlless_interpreter(tmp_path):
    """A stand-in for a python without PyYAML: it runs the script it is given with `import yaml` made to fail, and logs each start."""
    log = tmp_path / 'starts.log'
    wrapper = tmp_path / 'python-without-yaml'
    wrapper.write_text(
        '#!/bin/sh\n'
        f'echo started >> "{log}"\n'
        f'exec "{sys.executable}" -c "import sys, runpy; sys.modules[\'yaml\'] = None; sys.argv = sys.argv[1:]; '
        'runpy.run_path(sys.argv[0], run_name=\'__main__\')" "$@"\n', encoding='utf-8')
    wrapper.chmod(0o755)
    return wrapper, log


def _broken_targets(tmp_path):
    not_executable = tmp_path / 'not-executable'
    not_executable.write_text('#!/bin/sh\n', encoding='utf-8')
    not_executable.chmod(0o644)
    return {'absent': tmp_path / 'no-such-python', 'directory': tmp_path, 'not-executable': not_executable}


@pytest.mark.parametrize('kind', ['absent', 'directory', 'not-executable'])
def test_audit_screen_stops_when_the_dependency_interpreter_cannot_be_started(kind, tmp_path):
    target = _broken_targets(tmp_path)[kind]
    proc = _run_without_yaml(tmp_path, {'DESIGN_AUTHOR_PYTHON': str(target)})
    assert proc.returncode == 2, proc.stdout + proc.stderr
    assert 'PyYAML' in proc.stdout and 'DESIGN_AUTHOR_PYTHON' in proc.stdout, proc.stdout + proc.stderr
    assert 'Traceback' not in proc.stderr, proc.stderr


def test_audit_screen_stops_once_when_the_dependency_interpreter_has_no_pyyaml(tmp_path):
    wrapper, log = _yamlless_interpreter(tmp_path)
    proc = _run_without_yaml(tmp_path, {'DESIGN_AUTHOR_PYTHON': str(wrapper)})
    assert proc.returncode == 2, proc.stdout + proc.stderr
    assert 'PyYAML' in proc.stdout and 'DESIGN_AUTHOR_PYTHON' in proc.stdout, proc.stdout + proc.stderr
    # started exactly once: the interpreter was really tried (the branch exists), and it was not tried again (no loop)
    assert log.read_text(encoding='utf-8').splitlines() == ['started']
