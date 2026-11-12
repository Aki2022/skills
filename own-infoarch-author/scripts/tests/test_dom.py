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
    # The repeat does not carry the outline's copy: reported, but the section itself still matches.
    assert 'IA_DOM_COPY_EXTRA' in codes(run(snap), 'warn'), run(snap)


# --- rendered copy must be the outline's copy (2026-11-12) ---

def test_rewritten_copy_fails():
    snap = green_snapshot()
    snap['anchors'][1]['text'] = '運用の手間がゼロになります'
    found = run(snap)
    assert 'IA_DOM_COPY_DRIFT' in codes(found), found
    assert not any('ゼロ' in f[2] for f in found), 'rendered copy is not echoed into findings'


def test_copy_with_an_added_claim_fails():
    snap = green_snapshot()
    snap['anchors'][2]['text'] = '導入は100社以上。業界No.1'
    assert 'IA_DOM_COPY_DRIFT' in codes(run(snap))


def test_copy_comparison_ignores_whitespace_and_width():
    snap = green_snapshot()
    snap['anchors'][1]['text'] = '運用の手間が\n半分に なります'
    snap['anchors'][2]['text'] = '導入は１００社以上'
    assert codes(run(snap)) == [] and 'IA_DOM_COPY_EXTRA' not in codes(run(snap), 'warn')


def test_hidden_copy_does_not_count_as_the_section():
    snap = green_snapshot()
    snap['anchors'][1]['text'] = '別の文'
    snap['anchors'].append(anchor('msg', 'msg-0001', 130, '運用の手間が半分になります', visible=False))
    assert 'IA_DOM_COPY_DRIFT' in codes(run(snap))


def test_text_outside_every_anchor_is_a_warning_without_echo():
    snap = green_snapshot()
    snap['unanchored'] = '業界No.1 の実績'
    found = run(snap)
    assert codes(found) == [] and 'IA_DOM_UNANCHORED_TEXT' in codes(found, 'warn'), found
    assert not any('業界' in f[2] for f in found)


def test_no_unanchored_text_is_quiet():
    snap = green_snapshot()
    snap['unanchored'] = ''
    assert 'IA_DOM_UNANCHORED_TEXT' not in codes(run(snap), 'warn')


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
    (tmp_path / 'site/index.html').write_text(HTML.format(hidden=''), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=str(tmp_path / 'nowhere'))
    if PW_DIR:
        # A valid inherited runtime must not override the explicitly selected missing root.
        env['NODE_PATH'] = str(Path(PW_DIR) / 'node_modules')
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, cwd=str(tmp_path), check=False)
    assert proc.returncode == 2, proc.stdout + proc.stderr
    assert 'IA_DOM_PLAYWRIGHT' in [f['code'] for f in json.loads(proc.stdout)['findings']], proc.stdout


ANIMATED = '''<!doctype html><html><head><meta charset="utf-8"><style>
body{{margin:0;font:16px sans-serif}} section{{min-height:200px}}
@keyframes up{{from{{opacity:0}}to{{opacity:1}}}} .anim{{animation:up .6s ease-out forwards}}
</style></head><body>
<section data-ia-problem="prb-0001" class="hero" style="opacity:0">既存システムの運用、重くありませんか</section>
<section data-ia-msg="msg-0001" class="anim">運用の手間が半分になります</section>
<section data-ia-msg="msg-0002" class="anim">導入は<span data-ia-fact="fct-0001">100社以上</span></section>
<section data-ia-msg="msg-0003" class="anim">資料請求はこちら（<span data-ia-fact="fct-0002">03-1234-5678</span>）</section>
<script>document.addEventListener('DOMContentLoaded',function(){{setTimeout(function(){{var h=document.querySelector('.hero');h.style.transition='opacity .8s ease-out .2s';h.style.opacity='1';}},100);}});</script>
</body></html>'''


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_page_that_fades_in_is_read_after_it_settles(tmp_path):
    """A hero shown by a timer + transition and sections with CSS animations must not read as hidden or unstable."""
    from test_check import write_repo
    write_repo(tmp_path)
    page = tmp_path / 'site/index.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(ANIMATED.format(), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    assert proc.returncode == 0, json.loads(proc.stdout)['findings']


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real monorepo runtime integration not run')
def test_default_runtime_resolves_from_monorepo_ancestor(tmp_path):
    """Without an explicit override, a nested app can use its repository's existing runtime."""
    project = tmp_path / 'project'
    nested = project / 'apps/web'
    nested.mkdir(parents=True)
    (project / 'node_modules').symlink_to(Path(PW_DIR) / 'node_modules', target_is_directory=True)
    page = project / 'index.html'
    page.write_text(HTML.format(hidden=''), encoding='utf-8')
    env = dict(os.environ)
    env.pop('INFOARCH_PLAYWRIGHT_DIR', None)
    env.pop('NODE_PATH', None)
    proc = subprocess.run(['node', str(SCRIPTS / 'dom_snapshot.mjs'), str(page), json.dumps([VP])],
                          cwd=nested, env=env, capture_output=True, text=True, check=False)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    snapshots = json.loads(proc.stdout)
    assert len(snapshots) == 1 and snapshots[0]['viewport'] == VP
    assert snapshots[0]['anchors'] and snapshots[0]['text_chars'] > 0


DRIFT = HTML.replace('運用の手間が半分になります', '運用の手間がゼロになります')
EXTRA = HTML.replace('</body>', '<footer><p>業界No.1 の実績</p></footer></body>')


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('page,expect,code', [(DRIFT, 2, 'IA_DOM_COPY_DRIFT'), (EXTRA, 1, 'IA_DOM_UNANCHORED_TEXT')])
def test_real_browser_reads_copy_and_unanchored_text(tmp_path, page, expect, code):
    from test_check import write_repo
    write_repo(tmp_path)
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text(page.format(hidden=''), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    assert proc.returncode == expect and code in [f['code'] for f in report['findings']], report
    if code == 'IA_DOM_UNANCHORED_TEXT':
        assert {r['viewport'] for r in report['unanchored']} == {'mobile', 'desktop'}
        assert all('業界No.1' in r['text'] for r in report['unanchored'])


# --- security review of #98 (2026-11-12) ---

def test_fact_anchor_with_an_unknown_id_fails():
    snap = green_snapshot()
    snap['anchors'].append(anchor('fact', 'fct-0099', 2000, '業界No.1'))
    assert 'IA_DOM_ANCHOR_UNKNOWN' in codes(run(snap)), run(snap)


def test_fact_anchor_for_a_fact_this_page_does_not_use_must_show_its_value():
    outline = copy.deepcopy(OUTLINE)
    outline['pages'][0]['sections'][3]['copy'] = '資料請求はこちら'  # the page no longer refers to fct-0002
    snap = green_snapshot()
    snap['anchors'][4]['text'] = '資料請求はこちら'
    snap['anchors'][5]['text'] = '業界No.1'  # a registered fact id wrapped around other text
    assert 'IA_DOM_FACT_VALUE' in codes(run(snap, outline=outline)), run(snap, outline=outline)
    snap['anchors'][5]['text'] = '03-1234-5678'  # showing the registered value is fine
    assert codes(run(snap, outline=outline)) == []


def test_drifted_repeats_are_sent_to_judgment(monkeypatch, tmp_path):
    from test_check import write_repo
    write_repo(tmp_path)
    (tmp_path / 'site').mkdir()
    (tmp_path / 'site/index.html').write_text('<p>x</p>', encoding='utf-8')

    def snap(root, html, vps):
        out = []
        for v in vps:
            s = green_snapshot()
            s['viewport'] = v
            s['anchors'].insert(2, anchor('msg', 'msg-0003', 300, '業界No.1 の資料を請求'))
            out.append(s)
        return out
    monkeypatch.setattr(dom, 'snapshot', snap)
    report = dom.run(tmp_path, 'lp-municipal', '2026-10-23')
    assert report.unanchored and all('業界No.1 の資料を請求' in r['text'] for r in report.unanchored), report.unanchored


ZERO_SIZE = HTML.replace('</body>', '<section data-ia-msg="msg-0001" style="height:0;min-height:0;overflow:visible">業界No.1 の実績</section></body>')


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_real_browser_zero_size_anchor_that_paints_text_is_checked(tmp_path):
    from test_check import write_repo
    write_repo(tmp_path)
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text(ZERO_SIZE.format(hidden=''), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    # The second msg-0001 has no box but its text is painted: it is a repeat with other copy, and its text is judged.
    assert 'IA_DOM_COPY_EXTRA' in [f['code'] for f in report['findings']], report['findings']
    assert report['unanchored'] and all('業界No.1' in r['text'] for r in report['unanchored'])


ZERO_SIZE_LOOSE = HTML.replace('</body>', '<div style="height:0;overflow:visible">業界No.1 の実績</div></body>')


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_real_browser_text_painted_from_a_zero_size_box_is_unanchored(tmp_path):
    from test_check import write_repo
    write_repo(tmp_path)
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text(ZERO_SIZE_LOOSE.format(hidden=''), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    assert 'IA_DOM_UNANCHORED_TEXT' in [f['code'] for f in report['findings']], report['findings']
    assert report['unanchored'] and all('業界No.1' in r['text'] for r in report['unanchored'])


# --- a section may carry both a problem and a message (security review of #100) ---

def _both():
    outline = copy.deepcopy(OUTLINE)
    outline['pages'][0]['sections'][0]['message'] = 'msg-0001'  # problem section that also carries the benefit
    outline['pages'][0]['sections'].pop(1)
    snap = green_snapshot()
    snap['anchors'].pop(1)
    # skeleton puts both attributes on one element, so the snapshot holds two anchors with the same text
    snap['anchors'].insert(1, anchor('msg', 'msg-0001', 10, '既存システムの運用、重くありませんか'))
    return outline, snap


def test_section_with_problem_and_message_checks_both_anchors():
    outline, snap = _both()
    assert codes(run(snap, outline=outline)) == [], run(snap, outline=outline)
    snap['anchors'][1]['text'] = '業界No.1'
    assert 'IA_DOM_COPY_DRIFT' in codes(run(snap, outline=outline)), run(snap, outline=outline)


def test_drifted_message_of_a_problem_section_is_sent_to_judgment():
    outline, snap = _both()
    snap['anchors'].insert(2, anchor('msg', 'msg-0001', 20, '業界No.1 の実績'))
    assert '業界No.1 の実績' in dom.unjudged_text(outline, SSOT, snap)
