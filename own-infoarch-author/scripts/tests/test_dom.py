"""ISSUE-03: rendered DOM vs structured outline.

Unit tests compare DOM snapshots (the JSON dom_snapshot.mjs writes) with the outline; they always run.
The integration test drives a real Chromium through Playwright and runs only when
INFOARCH_TEST_PLAYWRIGHT_DIR names a directory whose node_modules has playwright: it is reported as a skip
with that reason, never as a pass.
"""
import copy
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import dom  # noqa: E402
from test_check import OUTLINE, SSOT  # noqa: E402

SCRIPTS = Path(__file__).resolve().parents[1]
VP = {'name': 'mobile', 'width': 375, 'height': 812}


def anchor(kind, aid, top, text='x', visible=True, left=0, height=40):
    return {'kind': kind, 'id': aid, 'top': top, 'left': left, 'width': 300, 'height': height, 'visible': visible, 'text': text}


def green_snapshot():
    return {'viewport': VP, 'text_chars': 400, 'first_screen': '既存システムの運用、重くありませんか 運用の手間が半分になります', 'anchors': [
        anchor('problem', 'prb-0001', 10, '既存システムの運用、重くありませんか'),
        anchor('msg', 'msg-0001', 120, '運用の手間が半分になります'),
        anchor('msg', 'msg-0002', 900, '導入は100社以上'),
        anchor('fact', 'fct-0001', 905, '100社以上', height=20),
        anchor('msg', 'msg-0003', 1500, '資料請求はこちら（03-1234-5678）'),
        anchor('fact', 'fct-0002', 1505, '03-1234-5678', height=20),
    ]}


def codes(findings, level='fail'):
    return [f[0] for f in findings if f[1] == level]


def run(snapshot, outline=OUTLINE, ssot=SSOT):
    return dom.compare(outline, ssot, snapshot)


def test_green_snapshot_passes():
    assert codes(run(green_snapshot())) == []


@pytest.mark.parametrize('mutation,code', [
    (lambda s: s['anchors'].pop(1), 'IA_DOM_ANCHOR_MISSING'),
    (lambda s: s['anchors'][1].update(visible=False), 'IA_DOM_ANCHOR_HIDDEN'),
    (lambda s: s['anchors'][1].update(top=1600), 'IA_DOM_ORDER'),
    (lambda s: s['anchors'][3].update(text='120社以上'), 'IA_DOM_FACT_VALUE'),
    (lambda s: s['anchors'].pop(3), 'IA_DOM_FACT_MISSING'),
    (lambda s: s.update(text_chars=0, anchors=[], first_screen=''), 'IA_DOM_NO_TARGETS'),
    (lambda s: s['anchors'].append(anchor('msg', 'msg-0099', 2000)), 'IA_DOM_ANCHOR_UNKNOWN'),
])
def test_red_snapshots_fail(mutation, code):
    snap = green_snapshot()
    mutation(snap)
    assert code in codes(run(snap)), run(snap)


def test_cta_may_repeat_without_breaking_order():
    snap = green_snapshot()
    snap['anchors'].insert(2, anchor('msg', 'msg-0003', 300, '資料請求'))
    assert codes(run(snap)) == []


def test_fact_whitespace_is_normalized_but_value_must_match():
    snap = green_snapshot()
    snap['anchors'][3]['text'] = ' 100社以上\n'
    assert codes(run(snap)) == []


def test_two_runs_must_be_identical():
    a, b = green_snapshot(), green_snapshot()
    b['anchors'][1]['top'] = 121
    assert 'IA_DOM_UNSTABLE' in codes(dom.compare_runs([a, b]))
    assert codes(dom.compare_runs([a, copy.deepcopy(a)])) == []


PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')

HTML = '''<!doctype html><html><head><meta charset="utf-8"><style>body{{font:16px sans-serif;margin:0}} section{{min-height:200px}}</style></head><body>
<section data-ia-problem="prb-0001"><h1>既存システムの運用、重くありませんか</h1></section>
<section data-ia-msg="msg-0001"><h2>運用の手間が半分になります</h2></section>
<section data-ia-msg="msg-0002" style="margin-top:600px"><p>導入は<span data-ia-fact="fct-0001">100社以上</span></p></section>
<section data-ia-msg="msg-0003"{hidden}><a href="#">資料請求はこちら（<span data-ia-fact="fct-0002">03-1234-5678</span>）</a></section>
</body></html>'''


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('hidden,expect', [('', 0), (' style="opacity:0"', 2)])
def test_real_browser_end_to_end(tmp_path, hidden, expect):
    from test_check import write_repo
    write_repo(tmp_path)
    page = tmp_path / 'site/index.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(HTML.format(hidden=hidden), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    assert proc.returncode == expect, report
    if expect == 0:
        assert report['targets']['viewports'] == 2 and report['first_screens']


def test_missing_playwright_is_exit_2_not_a_pass(tmp_path):
    from test_check import write_repo
    write_repo(tmp_path)
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text('<p>x</p>', encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=str(tmp_path / 'nowhere'))
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, cwd=str(tmp_path), check=False)
    assert proc.returncode == 2, proc.stdout + proc.stderr
    assert 'IA_DOM_PLAYWRIGHT' in [f['code'] for f in json.loads(proc.stdout)['findings']], proc.stdout
