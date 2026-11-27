"""The plan's emphasis (計画の強調) and its check on the drawn page.

A block of the plan may say, per emphasis, which characters of one copy or one dictionary value are drawn larger than the
rest of it: `emphasis: [{target: cpy-…|value-…, text: …}]`. Before the HTML, plan_check.py reads that the target belongs
to the block and that the text is in it (and leaves something to be larger than). With `--html <file>` it draws the page
at each width and fails when the emphasized characters are not drawn larger than the rest of their target. Before this,
the plan said "値段を大きく" in free text, nothing compared it with the page, and the price was drawn at the size of the
course name (2026-11-26)."""
import copy
import json
import shutil
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from plan_testkit import PLAN, SCRIPT, Repo, codes, run  # noqa: E402
from test_screen_conformance import can_measure, pinned_project  # noqa: E402


@pytest.fixture
def repo(tmp_path):
    r = Repo(tmp_path)
    r.l0()
    return r


def with_emphasis(repo, block, emphasis, script=None):
    data = copy.deepcopy(PLAN)
    data['blocks'][block]['emphasis'] = emphasis
    repo.plan(data)
    if script is not None:
        repo.script(script)
        repo.l0()  # the script is an input of the manifest


def check(repo, *extra):
    rc, out, raw = run('plan_check.py', '--repo', repo.root, '--deliverable', 'lp-sample', *extra)
    assert out is not None, raw.stdout + raw.stderr
    assert out['exit_code'] == rc
    return rc, out


def script_with_value():
    # blk-0003's copy shows a dictionary value, the way a course card shows its name and price
    data = copy.deepcopy(SCRIPT)
    data['copies']['cpy-0004'] = {'block': 'blk-0003', 'order': 1, 'text': '導入は{{value-customer-count}}'}
    return data


def test_an_emphasis_on_a_copy_of_the_block_passes(repo):
    with_emphasis(repo, 'blk-0003', [{'target': 'cpy-0004', 'text': '8 割'}])
    rc, out = check(repo)
    assert rc == 0, out['findings']
    assert out['emphasis'] == 1


def test_an_emphasis_on_a_value_shown_in_the_block_passes(repo):
    with_emphasis(repo, 'blk-0003', [{'target': 'value-customer-count', 'text': '100'}], script=script_with_value())
    rc, out = check(repo)
    assert rc == 0, out['findings']


@pytest.mark.parametrize('emphasis,code', [
    ({'target': 'cpy-0099', 'text': '8 割'}, 'PLAN_EMPHASIS_TARGET_UNKNOWN'),           # no such copy
    ({'target': 'cpy-0002', 'text': '手間'}, 'PLAN_EMPHASIS_TARGET_UNKNOWN'),           # a copy of another block
    ({'target': 'value-phone', 'text': '03'}, 'PLAN_EMPHASIS_TARGET_UNKNOWN'),          # a value the block does not show
    ({'target': 'cpy-0004', 'text': '9 割'}, 'PLAN_EMPHASIS_TEXT_ABSENT'),              # not in the target
    ({'target': 'cpy-0004', 'text': '利用者の 8 割が続けている'}, 'PLAN_EMPHASIS_TEXT_ABSENT'),  # nothing left to be larger than
])
def test_an_emphasis_that_cannot_be_drawn_fails(repo, emphasis, code):
    with_emphasis(repo, 'blk-0003', [emphasis])
    rc, out = check(repo)
    assert rc == 2 and code in codes(out, 'fail'), out['findings']


def test_an_emphasis_outside_the_schema_fails(repo):
    with_emphasis(repo, 'blk-0003', [{'target': 'cpy-0004'}])
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_SCHEMA' in codes(out, 'fail'), out['findings']


# --- on the drawn page ------------------------------------------------------------------------------------------------

PAGE = '''<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>body{{font:16px sans-serif;margin:0}}{css}</style></head>
<body><section data-ia-section="blk-0003"><p data-ia-copy="cpy-0004">導入は<span data-dict="value-customer-count">{value}</span></p></section></body></html>'''


@pytest.fixture(scope='module')
def browser(tmp_path_factory):
    reason = can_measure(tmp_path_factory)
    if reason:
        pytest.skip(f'{reason}: the drawn page is not checked here (reported as skipped, not passed)')
    return pinned_project(tmp_path_factory)


def draw(repo, browser, value, css=''):
    shutil.copyfile(Path(browser) / 'ux/design/render-environment.json', repo.root / 'ux/design/render-environment.json')
    repo.l0()  # the render environment is an input of the manifest
    page = repo.root / 'site/sample.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(PAGE.format(value=value, css=css), encoding='utf-8')
    return page


@pytest.mark.parametrize('value,css,expect', [
    ('<span class="n">100</span>社以上', '.n{font-size:40px}', 0),       # the emphasized characters are larger
    ('<span class="n">100</span>社以上', '', 'PLAN_EMPHASIS_NOT_RENDERED'),  # split, but drawn at the same size
    ('100社以上', '', 'PLAN_EMPHASIS_NOT_RENDERED'),                         # not split at all
    ('<span class="n">100</span>社以上', 'body{font-size:40px}.n{font-size:16px}', 'PLAN_EMPHASIS_NOT_RENDERED'),  # smaller
])
def test_the_emphasis_must_be_drawn_larger(repo, browser, value, css, expect):
    with_emphasis(repo, 'blk-0003', [{'target': 'value-customer-count', 'text': '100'}], script=script_with_value())
    page = draw(repo, browser, value, css)
    rc, out = check(repo, '--html', page)
    if expect == 0:
        assert rc == 0, out['findings']
        assert out['emphasis_drawn'] == {'checked': 2, 'failed': 0}  # one emphasis at two widths
    else:
        assert rc == 2 and expect in codes(out, 'fail'), out['findings']


def test_an_emphasis_whose_target_is_not_on_the_page_fails(repo, browser):
    with_emphasis(repo, 'blk-0003', [{'target': 'value-customer-count', 'text': '100'}], script=script_with_value())
    page = draw(repo, browser, '100社以上')
    page.write_text(page.read_text(encoding='utf-8').replace(' data-dict="value-customer-count"', ''), encoding='utf-8')
    rc, out = check(repo, '--html', page)
    assert rc == 2 and 'PLAN_EMPHASIS_TARGET_NOT_DRAWN' in codes(out, 'fail'), out['findings']


def test_without_a_browser_the_drawn_check_is_not_a_pass(repo, tmp_path):
    with_emphasis(repo, 'blk-0003', [{'target': 'cpy-0004', 'text': '8 割'}])
    page = repo.root / 'site/sample.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text('<p data-ia-copy="cpy-0004">利用者の<b>8 割</b>が続けている</p>', encoding='utf-8')
    rc, out, raw = run('plan_check.py', '--repo', repo.root, '--deliverable', 'lp-sample', '--html', page,
                       env={'OWN_PLAYWRIGHT_DIR': '/nonexistent-playwright', 'OWN_PLAYWRIGHT_HOME': str(tmp_path / 'none')})
    assert out is not None, raw.stdout + raw.stderr
    assert rc == 2 and 'PLAN_EMPHASIS_NOT_MEASURED' in codes(out, 'fail'), out['findings']


def test_a_plan_without_emphasis_has_nothing_to_draw(repo, tmp_path):
    page = repo.root / 'site/sample.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text('<p>x</p>', encoding='utf-8')
    rc, out = check(repo, '--html', page)
    assert rc == 0, out['findings']
    assert out['emphasis'] == 0 and 'PLAN_EMPHASIS_NONE' in codes(out, 'info')
