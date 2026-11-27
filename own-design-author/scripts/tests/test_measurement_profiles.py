"""The measurement after rendering is one tool (own-web-build) whose thresholds and per-profile rules are design's
(own-design-author/references/measurement-profiles.json). A profile says which checks run, which do not and why, and
the widths and colour schemes it measures. Nothing here names a kind of deliverable except the profile ids themselves.

The browser cases need DESIGN_BROWSER_TEST=1 (and a Playwright runtime); the rest run everywhere."""
import copy
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from layer_testkit import copy_skills, git_init, run  # noqa: E402
from test_core import SKILL, project, check, write  # noqa: E402,F401
import layer_paths as lp  # noqa: E402

SKILLS = SKILL.parent
TOOL_REL = 'own-web-build/scripts/measure_render.mjs'
TOOL = SKILLS / TOOL_REL
REGISTRY = SKILL / 'references/measurement-profiles.json'
NODE = os.environ.get('DESIGN_TEST_NODE', 'node')
BROWSER = os.environ.get('DESIGN_BROWSER_TEST') == '1'
NO_RUNTIME = {'OWN_PLAYWRIGHT_DIR': '/nonexistent-playwright', 'OWN_PLAYWRIGHT_HOME': '/nonexistent-playwright-home'}
KIND_WORDS = re.compile(r'landing|\blp\b|product|slide|pptx|print|poster|flyer', re.I)


def registry():
    return json.loads(REGISTRY.read_text())


def emitted_ids(source):
    """Every check id the tool can put into a report: fail(...) / warn(...) calls and a failure literal."""
    ids = set(re.findall(r'\b(?:fail|warn)\(\s*["\']([a-z][a-z0-9-]*)["\']', source))
    ids |= set(re.findall(r'failures\.push\(\{\s*id:\s*["\']([a-z][a-z0-9-]*)["\']', source))
    return ids


# ---- the files and the shape of the registry ----

def test_the_tool_lives_in_the_build_skill_and_the_thresholds_in_the_design_skill():
    assert TOOL.is_file(), 'the measurement tool is own-web-build/scripts/measure_render.mjs'
    assert REGISTRY.is_file(), 'the profiles are own-design-author/references/measurement-profiles.json'
    assert (SKILL / 'references/measurement-thresholds.md').is_file(), 'the threshold table moved to the design skill'


def test_the_tool_and_its_checks_carry_no_kind_of_deliverable_in_their_names():
    data = registry()
    names = [TOOL.name, *data['checks'], *data['thresholds']]
    assert not [n for n in names if KIND_WORDS.search(n)], names


def test_every_profile_lists_what_it_applies_and_skips_with_a_reason_and_covers_every_check_exactly():
    data = registry()
    checks = set(data['checks'])
    assert checks and set(data['profiles']), 'an empty registry would pass every comparison below'
    for name, profile in data['profiles'].items():
        applies = profile['applies']
        skips = {row['id']: row['reason'] for row in profile['skips']}
        assert len(applies) == len(set(applies)), name
        assert not set(applies) & set(skips), f'{name}: a check is both applied and skipped'
        assert set(applies) | set(skips) == checks, f'{name}: every check id is applied or skipped, none left unsaid'
        assert all(isinstance(r, str) and len(r.strip()) >= 10 for r in skips.values()), f'{name}: a skip needs its reason'
        assert profile['viewports'] and all(type(w) is int and w > 0 for w in profile['viewports']), name
        assert profile['schemes'] and set(profile['schemes']) <= {'light', 'dark'}, name


def test_the_registry_knows_exactly_the_ids_the_tool_can_emit():
    emitted = emitted_ids(TOOL.read_text())
    assert len(emitted) > 20, f'the scan found only {len(emitted)} ids; the pattern no longer matches the tool'
    known = set(registry()['checks'])
    assert emitted - known == set(), 'the tool emits ids no profile has decided on'
    assert known - emitted == set(), 'the registry names ids the tool never emits'


def test_each_check_has_a_lane_and_the_one_profile_specific_check_is_skipped_where_it_does_not_apply():
    data = registry()
    assert {row['lane'] for row in data['checks'].values()} == {'fail', 'warn'}
    assert {'lp', 'product-ui'} <= set(data['profiles'])
    skipped = {n: [r['id'] for r in p['skips']] for n, p in data['profiles'].items()}
    assert skipped['lp'] == []
    assert skipped['product-ui'] == ['fold-cta']


def test_no_profile_may_skip_a_check_that_guards_against_measuring_nothing():
    """A profile that skipped these would let an empty, broken or unverified measurement read as green."""
    guards = {'zero-targets', 'measurement-broken', 'render-execution-failed', 'axe-unavailable', 'render-environment-mismatch'}
    for name, profile in registry()['profiles'].items():
        assert guards <= set(profile['applies']), f'{name} skips a guard: {guards - set(profile["applies"])}'
    assert registry()['endpoint_default'] in registry()['profiles']


def test_thresholds_are_numbers_and_the_tool_holds_no_copy_of_them():
    data = registry()
    assert data['thresholds'] and all(type(v) in (int, float) for v in data['thresholds'].values())
    source = TOOL.read_text()
    block = re.search(r'const\s+T\s*=\s*\{(.*?)\n\};', source, re.S)
    assert block, 'the tool builds its thresholds in one block'
    assert not re.search(r':\s*\d', block.group(1)), 'a number is written in the tool instead of being read from the registry'


# ---- the loader (python side: ship and the L2 launcher use it instead of spelling the ids) ----

def test_the_loader_knows_the_profile_ids_and_refuses_anything_else():
    import measurement_profiles as mp
    assert set(mp.profile_ids()) == set(registry()['profiles'])
    assert mp.profile('product-ui')['skips'][0]['id'] == 'fold-cta'
    for bad in (None, '', 'unknown', 'LP', 'print'):
        with pytest.raises(ValueError):
            mp.profile(bad)


def test_the_ship_contract_and_the_l2_launcher_take_their_profile_ids_from_the_registry(monkeypatch):
    import measurement_profiles as mp
    from web_check import validation_endpoint
    assert validation_endpoint({'env': 'DESIGN_WEB_URL'}) == ('DESIGN_WEB_URL', 'lp')
    narrowed = {k: v for k, v in registry().items()}
    narrowed['profiles'] = {'product-ui': registry()['profiles']['product-ui']}
    monkeypatch.setattr(mp, 'load', lambda: narrowed)
    with pytest.raises(ValueError):
        validation_endpoint({'env': 'DESIGN_WEB_URL', 'profile': 'lp'})
    assert validation_endpoint({'env': 'DESIGN_WEB_URL', 'profile': 'product-ui'})[1] == 'product-ui'


# ---- the tool refuses to run without a known profile (no browser needed) ----

@pytest.mark.parametrize('extra, text', [([], 'profile'), (['--profile', 'unknown'], 'profile'),
                                         (['--profile', 'print'], 'profile')])
def test_the_tool_exits_2_without_a_known_profile_before_any_browser(tmp_path, extra, text):
    repo = git_init(tmp_path / 'plain')
    result = run([NODE, TOOL, 'https://example.invalid/', '--manifest', repo / 'web.json', *extra],
                 cwd=tmp_path, env=NO_RUNTIME)
    assert result.returncode == 2, result.stdout + result.stderr
    assert text in result.stderr and 'resolved manifest unavailable' not in result.stderr, result.stderr


# ---- real browser: the same tool, both profiles, red and green ----

NORMAL = ('<!doctype html><html lang="en"><head><title>Fixture</title><style>'
          'body{font:16px/1.5 Arial;background:white;color:black;margin:0}'
          '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
          'button{font:inherit;color:black;background:white;width:100px;height:48px}'
          '</style></head><body><main><h1>Fixture</h1>CONTENT</main></body></html>')


def prepare(project):
    root, _ = project
    tokens = {'color': {'$type': 'color'},
              'font': {'$type': 'fontFamily', '$value': ['Arial']},
              'radius': {'$type': 'dimension', '$value': {'value': 0, 'unit': 'px'}}}
    for name, values, alpha in [('primary', [0, 0, 0], 1), ('white', [1, 1, 1], 1), ('clear', [0, 0, 0], 0)]:
        tokens['color'][name] = {'$value': {'colorSpace': 'srgb', 'components': values, 'alpha': alpha,
                                           'hex': '#ffffff' if name == 'white' else '#000000'}}
    for value in [tokens['font'], tokens['radius'], *[v for k, v in tokens['color'].items() if not k.startswith('$')]]:
        value['$extensions'] = {'org.own.design': {'origin': 'proposal'}}
    write(lp.path(root, 'design_dir', 'tokens', 'brand.tokens.json'), tokens)
    assert check(project)['exit_code'] != 2
    return root


def measure(root, html, profile, tool=TOOL, extra=()):
    page = root / 'index.html'
    page.write_text(html)
    output = root / 'measurement.json'
    output.unlink(missing_ok=True)
    argv = [NODE, str(tool), page.as_uri(), '--manifest', str(lp.path(root, 'resolved', 'web.json')),
            '--viewports', '375', '--schemes', 'light', '--json', str(output), '--profile', profile, *extra]
    result = subprocess.run(argv, cwd=root, env=os.environ.copy(), capture_output=True, text=True)
    return result, json.loads(output.read_text()) if output.is_file() else None


BODY = '<p>Readable fixture copy that is long enough to count as body text on the page.</p>'
LOW_CONTRAST = NORMAL.replace('background:white;color:black', 'background:white;color:#999999')
assert LOW_CONTRAST != NORMAL, 'the low-contrast fixture must differ from the clean one'
BELOW_FOLD_CTA = NORMAL.replace('CONTENT', '<button data-cta style="margin-top:900px">Submit</button>')
browser = pytest.mark.skipif(not BROWSER, reason='explicit real browser run required')


@browser
@pytest.mark.parametrize('profile', ['lp', 'product-ui'])
def test_a_clean_page_is_green_and_a_low_contrast_page_is_red_under_every_profile(project, profile):
    root = prepare(project)
    ok, report = measure(root, NORMAL.replace('CONTENT', BODY), profile)
    assert ok.returncode == 0, ok.stdout + ok.stderr
    assert report['profile'] == profile and report['failures'] == []
    assert report['measured']['runs'] == 1 and report['measured']['targets'] > 0
    bad, broken = measure(root, LOW_CONTRAST.replace('CONTENT', BODY), profile)
    assert bad.returncode == 1, bad.stdout + bad.stderr
    assert any(row['id'] in ('contrast', 'axe-contrast') for row in broken['failures'])


@browser
def test_the_first_fold_cta_is_judged_only_where_the_profile_applies_it(project):
    root = prepare(project)
    lp_run, lp_report = measure(root, BELOW_FOLD_CTA, 'lp')
    assert lp_run.returncode == 1 and any(row['id'] == 'fold-cta' for row in lp_report['failures'])
    assert 'fold-cta' in lp_report['applied'] and lp_report['skipped'] == []
    ui_run, ui_report = measure(root, BELOW_FOLD_CTA, 'product-ui')
    assert ui_run.returncode == 0, ui_run.stdout + ui_run.stderr
    assert 'fold-cta' not in ui_report['applied']
    assert [row['id'] for row in ui_report['skipped']] == ['fold-cta']
    assert ui_report['skipped'][0]['reason'] == registry()['profiles']['product-ui']['skips'][0]['reason']


@browser
def test_a_threshold_comes_from_the_registry_not_from_the_tool(project, tmp_path):
    """Raise the body-size floor in a copy of the skills: the same page that passed now fails (the tool reads it)."""
    root = prepare(project)
    html = NORMAL.replace('CONTENT', BODY)
    ok, _ = measure(root, html, 'lp')
    assert ok.returncode == 0, ok.stdout + ok.stderr

    def raise_floor(data):
        pass
    skills = copy_skills(tmp_path / 'skills', ('own-design-author', 'own-web-build'), raise_floor)
    file = skills / 'own-design-author/references/measurement-profiles.json'
    data = json.loads(file.read_text())
    data['thresholds']['bodyFontSizeMin'] = 40
    file.write_text(json.dumps(data))
    stricter, report = measure(root, html, 'lp', tool=skills / TOOL_REL)
    assert stricter.returncode == 1 and any(row['id'] == 'body-font-size' for row in report['failures'])


@browser
def test_a_page_with_nothing_to_measure_is_not_a_pass(project):
    root = prepare(project)
    empty = '<!doctype html><html lang="en"><head><title>Empty</title></head><body></body></html>'
    result, report = measure(root, empty, 'lp')
    assert result.returncode != 0, result.stdout + result.stderr
    assert report['measured']['targets'] == 0


# ---- the floor under every composition: a text column that crowds characters is red (spec Requirements 15) ----

CROWD = SKILL / 'fixtures/text-crowding'


def crowd_page(name):
    # the fixture's own spacing is on an 8px grid; the harness heading's default margin would be the off-grid one
    return NORMAL.replace('<h1>', '<h1 style="margin:0 0 8px">').replace('CONTENT', (CROWD / name).read_text(encoding='utf-8'))


def crowd_ids(report):
    return [row['id'] for row in report['failures']]


def test_the_crowding_fixtures_exist_and_carry_no_proper_name():
    names = sorted(p.name for p in CROWD.glob('*.html'))
    assert names == ['one-column.html', 'three-flex-columns-crowded.html', 'two-columns-crowded.html'], names


def test_the_crowding_check_is_a_failing_check_that_both_profiles_apply_and_its_floor_is_a_registry_number():
    data = registry()
    assert data['checks']['text-crowding']['lane'] == 'fail'
    assert type(data['thresholds']['crowdedMinChars']) in (int, float) and data['thresholds']['crowdedMinChars'] > 0
    assert all('text-crowding' in p['applies'] for p in data['profiles'].values())


@browser
@pytest.mark.parametrize('profile', ['lp', 'product-ui'])
def test_two_columns_that_squeeze_text_to_a_few_characters_are_red_and_one_column_is_green(project, profile):
    root = prepare(project)
    red, report = measure(root, crowd_page('two-columns-crowded.html'), profile)
    assert red.returncode == 1, red.stdout + red.stderr
    row = next(r for r in report['failures'] if r['id'] == 'text-crowding')
    assert row['at'].endswith('/375px') and row['count'] >= 4 and row['min'] == registry()['thresholds']['crowdedMinChars']
    assert row['sample'] and all(s['chars'] < row['min'] and s['lines'] >= 2 for s in row['sample']), row
    green, clean = measure(root, crowd_page('one-column.html'), profile)
    assert green.returncode == 0, green.stdout + green.stderr
    assert 'text-crowding' not in crowd_ids(clean)


@browser
def test_the_floor_does_not_depend_on_how_the_columns_were_made(project):
    root = prepare(project)
    red, report = measure(root, crowd_page('three-flex-columns-crowded.html'), 'lp')
    assert red.returncode == 1 and 'text-crowding' in crowd_ids(report), red.stdout + red.stderr


@browser
def test_short_labels_and_one_line_blocks_are_not_crowding(project):
    """A narrow box that holds one short line (a number, a tag) is not a crowded column of text."""
    root = prepare(project)
    html = NORMAL.replace('<h1>', '<h1 style="margin:0 0 8px">').replace('CONTENT', '<ul style="display:flex;gap:8px;margin:0;padding:0 24px;list-style:none">'
                          + ''.join(f'<li style="width:48px;border:1px solid black">{t}</li>' for t in ('01', '02', '03', '04'))
                          + '</ul>' + BODY)
    ok, report = measure(root, html, 'lp')
    assert ok.returncode == 0 and 'text-crowding' not in crowd_ids(report), ok.stdout + ok.stderr


@browser
def test_the_floor_is_read_from_the_registry_not_from_the_tool(project, tmp_path):
    """Lower the floor in a copy of the skills: the page that was red is green (and the reverse for one column)."""
    root = prepare(project)
    skills = copy_skills(tmp_path / 'skills', ('own-design-author', 'own-web-build'), lambda data: None)
    file = skills / 'own-design-author/references/measurement-profiles.json'
    data = json.loads(file.read_text())
    data['thresholds']['crowdedMinChars'] = 1
    file.write_text(json.dumps(data))
    loose, report = measure(root, crowd_page('two-columns-crowded.html'), 'lp', tool=skills / TOOL_REL)
    assert 'text-crowding' not in crowd_ids(report), 'the tool kept its own copy of the floor'
    data['thresholds']['crowdedMinChars'] = 30
    file.write_text(json.dumps(data))
    strict, report = measure(root, crowd_page('one-column.html'), 'lp', tool=skills / TOOL_REL)
    assert strict.returncode == 1 and 'text-crowding' in crowd_ids(report)
