"""Rendered DOM vs outline (v6 anchors: data-ia-section and data-ia-copy).

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
from test_check import OUTLINE, SSOT_VIEW as SSOT  # noqa: E402

SCRIPTS = Path(__file__).resolve().parents[1]
VP = {'name': 'mobile', 'width': 375, 'height': 812}


def anchor(kind, aid, top, text='x', visible=True, left=0, height=40, nested=False, section=None, dom_index=None, reorder=(),
           unverified=()):
    # dom_index is the anchor's place in document order; a fixture written top to bottom uses its top unless told otherwise.
    return {'kind': kind, 'id': aid, 'top': top, 'left': left, 'width': 300, 'height': height, 'visible': visible, 'text': text,
            'nested': nested, 'section': section, 'dom_index': top if dom_index is None else dom_index,
            'reorder': list(reorder), 'unverified': list(unverified)}


C1, C2 = '既存システムの運用、重くありませんか', '運用の手間が半分になります'
C3, C4, C5 = '導入は100社以上、止まらずに動き続けます', '保守の要らない運用へ', '資料請求はこちら（03-1234-5678）'


def green_snapshot():
    # index:        0 blk-0001, 1 cpy-0001, 2 cpy-0002, 3 blk-0002, 4 cpy-0003, 5 value-customer-count, 6 blk-0003, 7 cpy-0004,
    #               8 cpy-0005, 9 value-phone
    return {'viewport': VP, 'text_chars': 400, 'first_screen': f'{C1} {C2}', 'anchors': [
        anchor('section', 'blk-0001', 0, f'{C1} {C2}', height=200),
        anchor('copy', 'cpy-0001', 10, C1, section='blk-0001'),
        anchor('copy', 'cpy-0002', 120, C2, section='blk-0001'),
        anchor('section', 'blk-0002', 880, C3, height=200),
        anchor('copy', 'cpy-0003', 900, C3, section='blk-0002'),
        anchor('fact', 'value-customer-count', 905, '100社以上', height=20, section='blk-0002'),
        anchor('section', 'blk-0003', 1480, f'{C4} {C5}', height=200),
        anchor('copy', 'cpy-0004', 1490, C4, section='blk-0003'),
        anchor('copy', 'cpy-0005', 1500, C5, section='blk-0003'),
        anchor('fact', 'value-phone', 1505, '03-1234-5678', height=20, section='blk-0003'),
    ]}


def codes(findings, level='fail'):
    return [f[0] for f in findings if f[1] == level]


def run(snapshot, outline=OUTLINE, ssot=SSOT):
    return dom.compare(outline, ssot, snapshot)


def test_green_snapshot_passes():
    assert codes(run(green_snapshot())) == []


@pytest.mark.parametrize('mutation,code', [
    (lambda s: s['anchors'].pop(2), 'IA_DOM_ANCHOR_MISSING'),
    (lambda s: s['anchors'].pop(3), 'IA_DOM_ANCHOR_MISSING'),
    (lambda s: s['anchors'][2].update(visible=False), 'IA_DOM_ANCHOR_HIDDEN'),
    (lambda s: s['anchors'][2].update(dom_index=5), 'IA_DOM_COPY_ORDER'),
    (lambda s: s['anchors'][3].update(dom_index=1700), 'IA_DOM_ORDER'),
    (lambda s: s['anchors'][2].update(reorder=['div.cards: flex-direction: row-reverse']), 'IA_DOM_CSS_REORDER'),
    (lambda s: s['anchors'][3].update(reorder=['section: order: 2']), 'IA_DOM_CSS_REORDER'),
    (lambda s: s['anchors'][5].update(text='120社以上'), 'IA_DOM_FACT_VALUE'),
    (lambda s: s['anchors'].pop(5), 'IA_DOM_FACT_MISSING'),
    (lambda s: s.update(text_chars=0, anchors=[], first_screen=''), 'IA_DOM_NO_TARGETS'),
    (lambda s: s['anchors'].append(anchor('copy', 'cpy-0099', 2000)), 'IA_DOM_ANCHOR_UNKNOWN'),
    (lambda s: s['anchors'].append(anchor('section', 'sec-0099', 2000)), 'IA_DOM_ANCHOR_UNKNOWN'),
    # v4/v5 attributes left on a page must not pass silently
    (lambda s: s['anchors'].append(anchor('legacy-problem', 'prb-0001', 5)), 'IA_DOM_LEGACY_ATTRIBUTE'),
    (lambda s: s['anchors'].append(anchor('legacy-msg', 'msg-0001', 5)), 'IA_DOM_LEGACY_ATTRIBUTE'),
    (lambda s: s['anchors'].append(anchor('legacy-pain', 'pan-0001', 5)), 'IA_DOM_LEGACY_ATTRIBUTE'),
    (lambda s: s['anchors'].append(anchor('legacy-frame', 'hero', 5)), 'IA_DOM_LEGACY_ATTRIBUTE'),
    (lambda s: s['anchors'][2].update(nested=True), 'IA_DOM_ANCHOR_NESTED'),
])
def test_red_snapshots_fail(mutation, code):
    snap = green_snapshot()
    mutation(snap)
    assert code in codes(run(snap)), run(snap)


def test_order_is_document_order_not_the_drawn_position():
    """Side by side (2026-11-26): copies and blocks drawn left to right or a few px apart keep the DOM order."""
    snap = green_snapshot()
    snap['anchors'][1].update(top=14, left=0)     # cpy-0001 drawn lower than cpy-0002 ...
    snap['anchors'][2].update(top=10, left=200)   # ... and to its right: still first in the DOM
    snap['anchors'][3].update(top=0, left=900)    # blk-0002 beside blk-0001 at the top
    assert codes(run(snap)) == [], run(snap)


def test_a_reorder_on_a_hidden_or_unknown_anchor_is_not_reported_twice():
    snap = green_snapshot()
    snap['anchors'].append(anchor('copy', 'cpy-0002', 130, C2, visible=False, reorder=['order: 1']))
    assert 'IA_DOM_CSS_REORDER' not in codes(run(snap)), run(snap)


def test_order_that_cannot_be_judged_is_reported_as_unverified_not_passed():
    snap = green_snapshot()
    snap['anchors'][2]['unverified'] = ['p: position: absolute']
    found = run(snap)
    assert codes(found) == [] and codes(found, 'warn') == ['IA_DOM_ORDER_UNVERIFIED'], found


def test_a_repeated_call_to_action_is_another_copy_in_the_outline():
    """v6: a repeat on the page is a repeat in the outline (another copy of the CTA message in another section)."""
    snap = green_snapshot()
    snap['anchors'].insert(3, anchor('copy', 'cpy-0005', 300, '資料請求', section='blk-0001'))
    assert {'IA_DOM_COPY_REPEATED', 'IA_DOM_COPY_SECTION', 'IA_DOM_COPY_DRIFT'} <= set(codes(run(snap))), run(snap)


# --- rendered copy must be the outline's copy (2026-11-12) ---

def test_rewritten_copy_fails():
    snap = green_snapshot()
    snap['anchors'][2]['text'] = '運用の手間がゼロになります'
    found = run(snap)
    assert 'IA_DOM_COPY_DRIFT' in codes(found), found
    assert not any('ゼロ' in f[2] for f in found), 'rendered copy is not echoed into findings'


def test_copy_with_an_added_claim_fails():
    snap = green_snapshot()
    snap['anchors'][4]['text'] = '導入は100社以上、止まらずに動き続けます。業界No.1'
    assert 'IA_DOM_COPY_DRIFT' in codes(run(snap))


def test_copy_comparison_ignores_whitespace_and_width():
    snap = green_snapshot()
    snap['anchors'][2]['text'] = '運用の手間が\n半分に なります'
    snap['anchors'][4]['text'] = '導入は１００社以上、止まらずに 動き続けます'
    assert codes(run(snap)) == []


def test_hidden_copy_does_not_count_as_the_section():
    snap = green_snapshot()
    snap['anchors'][2]['text'] = '別の文'
    snap['anchors'].append(anchor('copy', 'cpy-0002', 130, C2, visible=False))
    assert 'IA_DOM_COPY_DRIFT' in codes(run(snap))


def test_text_outside_every_anchor_fails_without_echo():
    snap = green_snapshot()
    snap['unanchored'] = '業界No.1 の実績'
    found = run(snap)
    assert codes(found) == ['IA_DOM_UNANCHORED_TEXT'], found
    assert not any('業界' in f[2] for f in found)


def test_no_unanchored_text_is_quiet():
    snap = green_snapshot()
    snap['unanchored'] = ''
    assert 'IA_DOM_UNANCHORED_TEXT' not in codes(run(snap))


def test_fact_whitespace_is_normalized_but_value_must_match():
    snap = green_snapshot()
    snap['anchors'][5]['text'] = ' 100社以上\n'
    assert codes(run(snap)) == []


def test_fact_split_into_child_elements_still_matches():
    # A value may be split into child elements to style part of it (a larger price, a smaller unit). The snapshot joins
    # the children's text with a space, so a value with no space at the split must still match (2026-11-26).
    snap = green_snapshot()
    snap['anchors'][5]['text'] = '100 社以上'
    assert codes(run(snap)) == []


def test_fact_split_does_not_hide_a_changed_value():
    snap = green_snapshot()
    snap['anchors'][5]['text'] = '120 社以上'
    assert 'IA_DOM_FACT_VALUE' in codes(run(snap))


def test_two_runs_must_be_identical():
    a, b = green_snapshot(), green_snapshot()
    b['anchors'][2]['top'] = 121
    assert 'IA_DOM_UNSTABLE' in codes(dom.compare_runs([a, b]))
    assert codes(dom.compare_runs([a, copy.deepcopy(a)])) == []


PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')

HTML = '''<!doctype html><html><head><meta charset="utf-8"><style>body{{font:16px sans-serif;margin:0}} section{{min-height:200px}}</style></head><body>
<section data-ia-section="blk-0001" data-ia-function="claim"><h1 data-ia-copy="cpy-0001">既存システムの運用、重くありませんか</h1><h2 data-ia-copy="cpy-0002">運用の手間が半分になります</h2></section>
<section data-ia-section="blk-0002" data-ia-function="claim" style="margin-top:600px"><p data-ia-copy="cpy-0003">導入は<span data-dict="value-customer-count">100社以上</span>、止まらずに動き続けます</p></section>
<section data-ia-section="blk-0003" data-ia-function="action"{hidden}><p data-ia-copy="cpy-0004">保守の要らない運用へ</p><a href="#" data-ia-copy="cpy-0005">資料請求はこちら（<span data-dict="value-phone">03-1234-5678</span>）</a></section>
</body></html>'''


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('hidden,expect', [('', 0), (' style="opacity:0"', 2), (' data-ia-problem="prb-0001"', 2)])
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
<section data-ia-section="blk-0001" class="hero" style="opacity:0"><p data-ia-copy="cpy-0001">既存システムの運用、重くありませんか</p><p data-ia-copy="cpy-0002" class="anim">運用の手間が半分になります</p></section>
<section data-ia-section="blk-0002" class="anim"><p data-ia-copy="cpy-0003">導入は<span data-dict="value-customer-count">100社以上</span>、止まらずに動き続けます</p></section>
<section data-ia-section="blk-0003" class="anim"><p data-ia-copy="cpy-0004">保守の要らない運用へ</p><p data-ia-copy="cpy-0005">資料請求はこちら（<span data-dict="value-phone">03-1234-5678</span>）</p></section>
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
@pytest.mark.parametrize('page,expect,code', [(DRIFT, 2, 'IA_DOM_COPY_DRIFT'), (EXTRA, 2, 'IA_DOM_UNANCHORED_TEXT')])
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
    snap['anchors'].append(anchor('fact', 'value-unknown', 2000, '業界No.1'))
    assert 'IA_DOM_ANCHOR_UNKNOWN' in codes(run(snap)), run(snap)


def test_fact_anchor_for_a_fact_this_page_does_not_use_must_show_its_value():
    outline = copy.deepcopy(OUTLINE)
    outline['copies']['cpy-0005']['text'] = '資料請求はこちら'  # the page no longer refers to value-phone
    snap = green_snapshot()
    snap['anchors'][8]['text'] = '資料請求はこちら'
    snap['anchors'][9]['text'] = '業界No.1'  # a registered fact id wrapped around other text
    assert 'IA_DOM_FACT_VALUE' in codes(run(snap, outline=outline)), run(snap, outline=outline)
    snap['anchors'][9]['text'] = '03-1234-5678'  # showing the registered value is fine
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
            s['anchors'].insert(3, anchor('copy', 'cpy-0005', 300, '業界No.1 の資料を請求'))
            out.append(s)
        return out
    monkeypatch.setattr(dom, 'snapshot', snap)
    report = dom.run(tmp_path, 'lp-municipal', '2026-10-23')
    assert report.unanchored and all('業界No.1 の資料を請求' in r['text'] for r in report.unanchored), report.unanchored


ZERO_SIZE = HTML.replace('</body>', '<p data-ia-copy="cpy-0002" style="height:0;min-height:0;overflow:visible">業界No.1 の実績</p></body>')


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
    # The second cpy-0002 has no box but its text is painted: it is a repeated mark, and its text is still collected.
    assert 'IA_DOM_COPY_REPEATED' in [f['code'] for f in report['findings']], report['findings']
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


# --- declared text and official names (the unanchored warning means "text nobody manages", 2026-11-15) ---

def declared_outline():
    """OUTLINE plus a navigation bar, a section label and the product's official name on the page."""
    o = copy.deepcopy(OUTLINE)
    o['texts'] = {'nav': {'kind': 'navigation', 'text': 'サービス 使い方'},
                  'label-proof': {'kind': 'label', 'text': 'EVIDENCE', 'block': 'blk-0002'}}
    return o


def declared_snapshot():
    s = green_snapshot()
    s['anchors'] += [anchor('text', 'nav', 0, 'サービス 使い方', height=20),
                     anchor('text', 'label-proof', 880, 'EVIDENCE', height=20),
                     anchor('entity', 'name-web-server', 1700, 'Webサーバー', height=20)]
    s['unanchored'] = ''
    return s


def drun(snap, outline=None):
    return dom.compare(outline or declared_outline(), SSOT, snap)


def test_declared_text_and_official_name_pass_with_no_warning():
    found = drun(declared_snapshot())
    assert codes(found) == [] and codes(found, 'warn') == [], found


@pytest.mark.parametrize('mutation,code', [
    # the official name must be the whole text of its anchor, everywhere
    (lambda s: s['anchors'][-1].update(text='Web サーバー Pro'), 'IA_DOM_ENTITY_NAME'),
    (lambda s: s['anchors'][-1].update(text='Webサーバー 業界No.1'), 'IA_DOM_ENTITY_NAME'),
    (lambda s: s['anchors'].append(anchor('entity', 'name-web-server', 1800, 'Webサーバー2')), 'IA_DOM_ENTITY_NAME'),
    (lambda s: s['anchors'].append(anchor('entity', 'ent-0099', 1800, 'Webサーバー')), 'IA_DOM_ANCHOR_UNKNOWN'),
    # declared text must match on every visible occurrence (no COPY_EXTRA leniency)
    (lambda s: s['anchors'][-2].update(text='EVIDENCE 業界No.1'), 'IA_DOM_COPY_DRIFT'),
    (lambda s: s['anchors'].append(anchor('text', 'nav', 1900, 'サービス 料金')), 'IA_DOM_COPY_DRIFT'),
    (lambda s: s['anchors'].append(anchor('text', 'undeclared', 1900, '業界No.1')), 'IA_DOM_ANCHOR_UNKNOWN'),
    (lambda s: s['anchors'].pop(-3), 'IA_DOM_ANCHOR_MISSING'),   # navigation is still required on its page
    (lambda s: s['anchors'].pop(-2), 'IA_DOM_ANCHOR_MISSING'),
    (lambda s: s['anchors'][-2].update(visible=False), 'IA_DOM_ANCHOR_HIDDEN'),
    # declared text and names cannot carry other IA anchors inside them
    (lambda s: s['anchors'][-2].update(nested=True), 'IA_DOM_ANCHOR_NESTED'),
    (lambda s: s['anchors'][-1].update(nested=True), 'IA_DOM_ANCHOR_NESTED'),
])
def test_declared_text_and_name_red(mutation, code):
    snap = declared_snapshot()
    mutation(snap)
    assert code in codes(drun(snap)), drun(snap)


def test_navigation_may_move_without_breaking_order():
    snap = declared_snapshot()
    snap['anchors'][-3]['top'] = 3000  # e.g. a footer navigation
    assert codes(drun(snap)) == [], drun(snap)


def test_navigation_may_be_collapsed_at_a_viewport_but_must_exist():
    snap = declared_snapshot()
    snap['anchors'][-3]['visible'] = False  # e.g. a header menu hidden on mobile
    assert codes(drun(snap)) == [], drun(snap)
    snap['anchors'].pop(-3)
    assert 'IA_DOM_ANCHOR_MISSING' in codes(drun(snap))


def test_declared_text_with_a_fact_shows_the_value():
    o = declared_outline()
    o['texts']['nav']['text'] = 'サービス {{value-phone}}'
    snap = declared_snapshot()
    snap['anchors'][-3]['text'] = 'サービス 03-1234-5678'
    assert codes(drun(snap, o)) == [], drun(snap, o)


def test_unanchored_warning_points_at_declaring_the_text():
    snap = declared_snapshot()
    snap['unanchored'] = '業界No.1'
    found = [f for f in drun(snap) if f[0] == 'IA_DOM_UNANCHORED_TEXT']
    assert found and found[0][1] == 'fail' and '宣言' in found[0][2], found


DECLARED = '''<!doctype html><html><head><meta charset="utf-8"><style>body{{font:16px sans-serif;margin:0}} section{{min-height:200px}}</style></head><body>
<nav data-ia-text="nav">サービス 使い方</nav>
<section data-ia-section="blk-0001"><h1 data-ia-copy="cpy-0001">既存システムの運用、重くありませんか</h1><h2 data-ia-copy="cpy-0002">運用の手間が半分になります</h2></section>
<section data-ia-section="blk-0002" style="margin-top:600px"><p data-ia-text="label-proof">{label}</p><p data-ia-copy="cpy-0003">導入は<span data-dict="value-customer-count">100社以上</span>、止まらずに動き続けます</p></section>
<section data-ia-section="blk-0003"><p data-ia-copy="cpy-0004">保守の要らない運用へ</p><a href="#" data-ia-copy="cpy-0005">資料請求はこちら（<span data-dict="value-phone">03-1234-5678</span>）</a></section>
<footer><span data-dict="name-web-server">{name}</span></footer>
</body></html>'''


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('label,name,expect,code', [
    ('EVIDENCE', 'Webサーバー', 0, None),
    ('EVIDENCE 業界No.1', 'Webサーバー', 2, 'IA_DOM_COPY_DRIFT'),
    ('EVIDENCE', 'Webサーバー Pro', 2, 'IA_DOM_ENTITY_NAME'),
    # reads as the right label, but nothing is painted: the anchor text is the visible text, not innerText
    ('<span style="opacity:0">EVIDENCE</span>', 'Webサーバー', 2, 'IA_DOM_COPY_DRIFT'),
    ('EVIDENCE<span data-ia-copy="cpy-0003">導入は100社以上</span>', 'Webサーバー', 2, 'IA_DOM_ANCHOR_NESTED'),
])
def test_real_browser_declared_text_and_names(tmp_path, label, name, expect, code):
    from test_check import write_repo, CHANNEL, DRAFT, DECK
    write_repo(tmp_path, outlines=(declared_outline(), DECK))
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text(DECLARED.format(label=label, name=name), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    found = [f['code'] for f in report['findings']]
    assert proc.returncode == expect, report['findings']
    if code:
        assert code in found, report['findings']
    else:
        assert 'IA_DOM_UNANCHORED_TEXT' not in found and not report['unanchored'], report


# --- screen conformance with a real browser (requirement 10; fixtures/screen-red-cases.json) ---

SEC1 = '<section data-ia-section="blk-0001" data-ia-function="claim"><h1 data-ia-copy="cpy-0001">既存システムの運用、重くありませんか</h1><h2 data-ia-copy="cpy-0002">運用の手間が半分になります</h2></section>'
SEC2 = '<section data-ia-section="blk-0002" data-ia-function="claim"><p data-ia-copy="cpy-0003">導入は<span data-dict="value-customer-count">100社以上</span>、止まらずに動き続けます</p></section>'
SEC3 = '<section data-ia-section="blk-0003" data-ia-function="action"><p data-ia-copy="cpy-0004">保守の要らない運用へ</p><a href="#" data-ia-copy="cpy-0005">資料請求はこちら（<span data-dict="value-phone">03-1234-5678</span>）</a></section>'
PAGE = ('<!doctype html><html><head><meta charset="utf-8"><style>body{{font:16px sans-serif;margin:0}} section{{min-height:200px}}</style>'
        '</head><body>{body}</body></html>')
SCREEN_RED = [
    ('sections-swapped', SEC1 + SEC3 + SEC2, 'IA_DOM_ORDER'),
    ('copy-one-character', (SEC1 + SEC2 + SEC3).replace('半分になります', '半分になりました'), 'IA_DOM_COPY_DRIFT'),
    ('copies-swapped', SEC1.replace('<h1 data-ia-copy="cpy-0001">既存システムの運用、重くありませんか</h1><h2 data-ia-copy="cpy-0002">運用の手間が半分になります</h2>',
                                    '<h2 data-ia-copy="cpy-0002">運用の手間が半分になります</h2><h1 data-ia-copy="cpy-0001">既存システムの運用、重くありませんか</h1>') + SEC2 + SEC3,
     'IA_DOM_COPY_ORDER'),
    ('copy-in-other-section', SEC1.replace('<h2 data-ia-copy="cpy-0002">運用の手間が半分になります</h2>', '')
     + SEC2.replace('</section>', '<p data-ia-copy="cpy-0002">運用の手間が半分になります</p></section>') + SEC3, 'IA_DOM_COPY_SECTION'),
    ('text-added-outside', SEC1 + SEC2 + SEC3 + '<footer><p>いまだけ無料</p></footer>', 'IA_DOM_UNANCHORED_TEXT'),
]


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('name,body,code', [('conforming', SEC1 + SEC2 + SEC3, None)] + SCREEN_RED, ids=lambda v: v if isinstance(v, str) and len(v) < 30 else None)
def test_real_browser_screen_conformance(tmp_path, name, body, code):
    from test_check import write_repo
    write_repo(tmp_path)
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text(PAGE.format(body=body), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    found = [f['code'] for f in report['findings'] if f['level'] == 'fail']
    if code is None:
        assert proc.returncode == 0 and found == [], report['findings']
    else:
        assert proc.returncode == 2 and code in found, report['findings']


def test_every_screen_condition_has_a_real_browser_fixture():
    import json as _json
    table = _json.loads((SCRIPTS.parent / 'fixtures/screen-red-cases.json').read_text(encoding='utf-8'))
    required = {'IA_DOM_ORDER', 'IA_DOM_COPY_DRIFT', 'IA_DOM_COPY_ORDER', 'IA_DOM_COPY_SECTION', 'IA_DOM_UNANCHORED_TEXT',
                'IA_DOM_CSS_REORDER'}
    assert required <= {c['code'] for c in table['conditions']}
    assert {code for _, _, code in SCREEN_RED + CARD_RED} == required


# --- order is the DOM's reading order; CSS that redraws it in another order fails (2026-11-26) ---
# Three course cards (name, description, button) in one block, side by side. The old rule sorted anchors by the drawn
# (top, left) and failed any row; the order is now the document order, and CSS that changes the visual order fails.

CARD_NAMES = ('ライト', 'スタンダード', 'プレミアム')


def cards_outline():
    o = copy.deepcopy(OUTLINE)
    n = 6
    for card in CARD_NAMES:
        for text in (f'{card}コース', f'{card}の説明です', f'{card}を選ぶ'):
            o['copies'][f'cpy-{n:04d}'] = {'block': 'blk-0002', 'order': n - 4, 'message': 'msg-0002', 'text': text}
            n += 1
    return o


def card(i, style=''):
    name, base = CARD_NAMES[i], 6 + 3 * i
    attr = f' style="{style}"' if style else ''
    return (f'<div class="card"{attr}><h3 data-ia-copy="cpy-{base:04d}">{name}コース</h3>'
            f'<p data-ia-copy="cpy-{base + 1:04d}">{name}の説明です</p><a href="#" data-ia-copy="cpy-{base + 2:04d}">{name}を選ぶ</a></div>')


# the tops differ by a few px on purpose: a row drawn 1-3 px apart must not change the order
OFFSETS = ('margin-top:3px', '', 'margin-top:1px')


def cards_page(cards=None, row='display:flex;gap:8px', blocks=None, wrap=''):
    cards = cards or ''.join(card(i, OFFSETS[i]) for i in range(3))
    sec2 = SEC2.replace('</section>', f'<div class="cards" style="{row}">{cards}</div></section>')
    body = ''.join(blocks or (SEC1, sec2, SEC3))
    if wrap:
        body = f'<main style="{wrap}">{body}</main>'
    return PAGE.format(body=body)


def run_cards(tmp_path, html):
    from test_check import write_repo, DECK
    write_repo(tmp_path, outlines=(cards_outline(), DECK))
    (tmp_path / 'site').mkdir(exist_ok=True)
    (tmp_path / 'site/index.html').write_text(html, encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    return proc.returncode, report, {f['code'] for f in report['findings'] if f['level'] == 'fail'}


def _sec2():
    return cards_page().split('<body>')[1].split('</body>')[0].replace(SEC1, '').replace(SEC3, '')


CARD_GREEN = [
    ('cards-in-a-row', cards_page()),
    ('cards-stacked', cards_page(row='display:block')),
    ('grid-placed-in-dom-order', cards_page(cards=''.join(card(i, f'grid-column-start:{i + 1};grid-row-start:1') for i in range(3)),
                                            row='display:grid;grid-template-columns:repeat(3,1fr)')),
    ('blocks-in-a-row', cards_page(wrap='display:flex;gap:8px;align-items:flex-start')),
]
CARD_RED = [
    ('cards-swapped-in-dom', cards_page(cards=card(1) + card(0) + card(2)), 'IA_DOM_COPY_ORDER'),
    ('blocks-swapped-in-dom', cards_page(blocks=(SEC1, SEC3, _sec2())), 'IA_DOM_ORDER'),
    ('card-moved-by-order', cards_page(cards=card(0, 'order:2') + card(1) + card(2)), 'IA_DOM_CSS_REORDER'),
    ('row-reverse', cards_page(row='display:flex;flex-direction:row-reverse'), 'IA_DOM_CSS_REORDER'),
    ('wrap-reverse', cards_page(row='display:flex;flex-wrap:wrap-reverse'), 'IA_DOM_CSS_REORDER'),
    ('grid-placement', cards_page(cards=card(0, 'grid-column-start:3;grid-row-start:1') + card(1) + card(2),
                                  row='display:grid;grid-template-columns:repeat(3,1fr)'), 'IA_DOM_CSS_REORDER'),
    ('grid-dense', cards_page(row='display:grid;grid-template-columns:repeat(3,1fr);grid-auto-flow:row dense'), 'IA_DOM_CSS_REORDER'),
    ('blocks-column-reverse', cards_page(wrap='display:flex;flex-direction:column-reverse'), 'IA_DOM_CSS_REORDER'),
]


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('name,html', CARD_GREEN, ids=[n for n, _ in CARD_GREEN])
def test_real_browser_side_by_side_in_dom_order_passes(tmp_path, name, html):
    rc, report, fails = run_cards(tmp_path, html)
    assert rc == 0 and fails == set(), report['findings']


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('name,html,code', CARD_RED, ids=[n for n, _, _ in CARD_RED])
def test_real_browser_order_red(tmp_path, name, html, code):
    rc, report, fails = run_cards(tmp_path, html)
    assert rc == 2 and code in fails, report['findings']
    if code == 'IA_DOM_CSS_REORDER':
        assert not fails & {'IA_DOM_ORDER', 'IA_DOM_COPY_ORDER'}, report['findings']  # the DOM order itself is right


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_real_browser_order_moved_by_absolute_position_is_unverified_not_passed(tmp_path):
    html = cards_page(row='position:relative;height:300px', cards=card(0, 'position:absolute;left:400px;top:0') + card(1) + card(2))
    rc, report, fails = run_cards(tmp_path, html)
    assert rc == 1 and fails == set(), report['findings']
    assert 'IA_DOM_ORDER_UNVERIFIED' in {f['code'] for f in report['findings'] if f['level'] == 'warn'}
    assert report['uncovered'] and all(u['path'] == 'site/index.html' for u in report['uncovered']), report['uncovered']
