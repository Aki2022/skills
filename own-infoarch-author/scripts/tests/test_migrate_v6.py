"""One batch migration of a v5 SSOT to v6 (SPEC-information-architecture-standardization v6, requirement 17).

Promises are dropped, v5 messages become outline atoms, audiences become the outline's constant `target`, v5
sections and frames become section and copy ids, proofs become positioning evidence, the rendered v5 marks become
data-ia-section / data-ia-copy and modules.ia moves to the v6 contract. The tool prints the old -> new id map, the
items it failed to migrate (non-zero = exit 2, nothing written), the approvals and judgments that stop applying, and
the rollback steps. Dry run by default; it never writes design/approvals.yaml or design/judgments.json.
"""
import copy
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import ia_module  # noqa: E402
from test_check import check, cli, codes, dump  # noqa: E402

TABLE = json.loads((SKILL / 'fixtures/migrate-red-cases.json').read_text(encoding='utf-8'))
CONDITIONS = {c['id']: c for c in TABLE['conditions']}
CASES = TABLE['cases']
PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')


def _a(**kw):
    return {'state': 'proposed', 'provenance': 'owner', **kw}


# The v5 SSOT (info_arch.yaml schema version 3): audiences and messages in the SSOT, the strategy layer with
# promises and concepts, proofs that support messages.
V5_SSOT = {
    'schema_version': 3,
    'ledger': {'next': {'aud': 3, 'msg': 5, 'prf': 3, 'fct': 3, 'ent': 2, 'vce': 2, 'seg': 3, 'job': 3, 'pan': 4,
                        'gan': 2, 'val': 3, 'alt': 3, 'str': 3, 'cpt': 2, 'prm': 3},
               'tombstones': {'msg-0004': 'msg-0003', 'pan-0003': None}},
    'audiences': {
        'aud-0001': _a(label='自治体の情報部門', situation='既存システムの運用が重い', tone='丁寧で落ち着いた調子',
                       desired_action='資料請求', priorities=['msg-0001', 'msg-0002', 'msg-0003'], segment='seg-0001'),
        'aud-0002': _a(label='中小企業の経営者', situation='人手が足りない', tone='端的で親しみやすい調子',
                       priorities=['msg-0002', 'msg-0001', 'msg-0003'], segment='seg-0002')},
    'segments': {
        'seg-0001': _a(qualifier='自治体の情報部門', situation='既存システムの運用が重い'),
        'seg-0002': _a(qualifier='中小企業の経営者', situation='人手が足りない')},
    'jobs': {
        'job-0001': _a(text='既存システムを無理なく運用する', segment='seg-0001', push=['保守の手間が増え続ける']),
        'job-0002': _a(text='少ない人数で業務を回す', segment='seg-0002')},
    'pains': {
        'pan-0001': _a(text='運用が重い', job='job-0001'),
        'pan-0002': _a(text='人手が足りない', job='job-0002')},
    'gains': {'gan-0001': _a(text='職員が本来の業務に時間を使える', job='job-0001')},
    'values': {
        'val-0001': _a(text='運用の手間が減る', relieves=['pan-0001'], creates=['gan-0001']),
        'val-0002': _a(text='少人数でも回る', relieves=['pan-0002'])},
    'alternatives': {
        'alt-0001': _a(text='今のシステムを使い続ける', segment='seg-0001'),
        'alt-0002': _a(text='表計算で管理する', segment='seg-0002')},
    'strengths': {
        'str-0001': _a(text='自動更新で保守が要らない', versus=['alt-0001'], produces=['val-0001']),
        'str-0002': _a(text='設定なしで使い始められる', versus=['alt-0002'], produces=['val-0002'])},
    'concepts': {'cpt-0001': _a(category='自治体向け Web サーバー', concept='保守のいらない Web サーバー')},
    'promises': {
        'prm-0001': _a(text='運用の手間を半分にする', values=['val-0001']),
        'prm-0002': _a(text='少人数でも止まらない', values=['val-0002'])},
    'messages': {
        'msg-0001': _a(text='運用の手間が半分になる', role='benefit', addresses=['pan-0001'], serves=['val-0001', 'prm-0001']),
        'msg-0002': _a(text='多くの自治体で使われている', role='proof', serves=['val-0001', 'val-0002']),
        'msg-0003': _a(text='資料を請求する', role='cta', serves=['prm-0001', 'prm-0002'])},
    'proofs': {
        'prf-0001': _a(text='導入自治体の数', source='導入台帳', supports=['msg-0002'], as_of='2026-10-01'),
        'prf-0002': _a(text='導入前後の運用時間の比較', source='導入事例の計測', supports=['msg-0001'], as_of='2026-10-01')},
    'facts': {
        'fct-0001': _a(label='導入自治体数', value='100自治体以上', as_of='2026-10-01'),
        'fct-0002': _a(label='問い合わせ電話', value='03-1234-5678', evergreen=True)},
    'entities': {'ent-0001': _a(type='product', name='Webサーバー')},
    'voices': {'vce-0001': _a(traits=[
        {'trait': 'わかりやすい', 'is': '専門的な中身も噛み砕いて言う', 'is_not': '子ども扱いはしない'},
        {'trait': '誠実', 'is': 'できることを具体的に言う', 'is_not': '大げさな約束はしない'},
        {'trait': '前向き', 'is': '解決の道筋を示す', 'is_not': '浮ついた煽りはしない'}])},
}

# v5 outlines (schema version 2): pages of sections, each with one copy; frames declared per deliverable.
V5_OUTLINES = [
    {'schema_version': 2, 'id': 'lp-municipal', 'channel': 'web',
     'brief': {'purpose': '自治体の情報部門からの資料請求を増やす', 'action': '資料請求', 'medium': 'LP'},
     'audience': 'aud-0001', 'outputs': ['site/index.html'],
     'strategy': {'choose': ['val-0001', 'prm-0001'], 'rationale': '手間が減る価値を先に示す'},
     'frames': {'media': 'lp', 'instances': [{'id': 'fv', 'type': 'first-view-message'}, {'id': 'cta-main', 'type': 'cta'},
                                              {'id': 'order', 'type': 'section-order'}]},
     'pages': [{'id': 'top', 'label': 'トップ', 'sections': [
         {'role': 'pain', 'pain': 'pan-0001', 'copy': '既存システムの運用、重くありませんか'},
         {'role': 'benefit', 'message': 'msg-0001', 'frame': 'fv', 'copy': '運用の手間が半分になります'},
         {'role': 'proof', 'message': 'msg-0002', 'proofs': ['prf-0001'], 'copy': '導入は{{fct-0001}}'},
         {'role': 'cta', 'message': 'msg-0003', 'frame': 'cta-main', 'copy': '資料請求はこちら（{{fct-0002}}）'},
         {'role': 'navigation', 'key': 'nav-footer', 'copy': 'お問い合わせ窓口'}]}]},
    {'schema_version': 2, 'id': 'deck-smb', 'channel': 'deck',
     'brief': {'purpose': '中小企業向けの説明', 'action': '問い合わせ', 'medium': '資料', 'cut': ['msg-0001']},
     'audience': 'aud-0002', 'outputs': ['decks/smb.pptx'],
     'strategy': {'choose': ['val-0002'], 'rationale': '少人数でも回る価値を実績で支える'},
     'pages': [{'id': 's1', 'sections': [
         {'role': 'proof', 'message': 'msg-0002', 'proofs': ['prf-0001'], 'copy': '{{fct-0001}}が導入'},
         {'role': 'cta', 'message': 'msg-0003', 'copy': 'お問い合わせ {{fct-0002}}'}]}]},
]
CHANNELS = [{'schema_version': 2, 'media': 'web', 'outputs': ['site/**/*.html'], 'ship': False},
            {'schema_version': 2, 'media': 'deck', 'outputs': ['decks/*.pptx'], 'ship': False}]

# What the v5 factory shipped: the v5 skeleton, styled (classes, a wrapper). One section holds its copy in a single
# child element, one holds it as bare text.
V5_HTML = '''<!doctype html>
<html lang="ja">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>lp</title>
<style>body{font:16px sans-serif;margin:0} section{padding:24px}</style></head>
<body data-ia-outline="lp-municipal">
 <main data-ia-page="top" aria-label="トップ" class="page">
  <section class="hero" data-ia-role="pain" data-ia-pain="pan-0001">
    <p class="lead">既存システムの運用、重くありませんか</p>
  </section>
  <div class="wrap">
  <section data-ia-role="benefit" data-ia-msg="msg-0001" data-ia-frame="fv" class="benefit">
    <h2>運用の手間が半分になります</h2>
  </section>
  </div>
  <section data-ia-role="proof" data-ia-msg="msg-0002">導入は<span data-ia-fact="fct-0001">100自治体以上</span></section>
  <section data-ia-role="cta" data-ia-msg="msg-0003" data-ia-frame="cta-main">
    <a href="#" class="button">資料請求はこちら（<span data-ia-fact="fct-0002">03-1234-5678</span>）</a>
  </section>
  <section data-ia-role="navigation" data-ia-text="nav-footer"><nav>お問い合わせ窓口</nav></section>
 </main>
</body>
</html>
'''
DESIGN = {'schema_version': 1, 'modules': {'ia': '0.3.0'}}
APPROVALS = [
    {'key': 'ds:brand:token:color.primary', 'value_hash': 'a' * 64, 'approver': 'fixture-owner', 'date': '2026-11-14'},
    {'key': 'ia:strategy:val-0001', 'value_hash': 'b' * 64, 'approver': 'fixture-owner', 'date': '2026-11-14'},
    {'key': 'ia:strategy:prm-0001', 'value_hash': 'c' * 64, 'approver': 'fixture-owner', 'date': '2026-11-14'},
]
JUDGMENTS = {'rows': [
    {'candidate': {'id': 'ia-strategy-fit:lp-municipal', 'input': {'kind': 'ia-strategy-fit', 'catalog_id': 'lp-municipal'}}},
    {'candidate': {'id': 'ds-legibility:hero', 'input': {'kind': 'ds-legibility', 'catalog_id': 'hero'}}},
]}


def write_v5(root, ssot=V5_SSOT, outlines=V5_OUTLINES, html=V5_HTML, design=DESIGN):
    base = root / 'design/info_arch'
    dump(base / 'info_arch.yaml', ssot)
    for ch in CHANNELS:
        dump(base / 'channels' / f'{ch["media"]}.yaml', ch)
    for o in outlines:
        dump(base / 'outlines' / f'{o["id"]}.yaml', o)
    page = root / 'site/index.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(html, encoding='utf-8')
    (root / 'design/design.yaml').write_text(json.dumps(design, indent=2) + '\n', encoding='utf-8')
    dump(root / 'design/approvals.yaml', APPROVALS)
    (root / 'design/judgments.json').write_text(json.dumps(JUDGMENTS, indent=2) + '\n', encoding='utf-8')
    return base


def snapshot(root):
    return {p.relative_to(root).as_posix(): p.read_bytes() for p in sorted(root.rglob('*')) if p.is_file()}


def migrate(root, *extra):
    return cli('migrate-v6', '--root', str(root), *extra)


def load(path):
    return yaml.safe_load(path.read_text(encoding='utf-8'))


# ---- dry run, report ---------------------------------------------------------------------------------------------

def test_dry_run_reports_everything_and_writes_nothing(tmp_path):
    write_v5(tmp_path)
    before = snapshot(tmp_path)
    rc, out = migrate(tmp_path)
    assert rc == 0, out
    assert out['exit_code'] == 0 and out['status'] == 'preview' and out['written'] is False
    assert snapshot(tmp_path) == before  # dry run: not one byte changed
    assert out['unmigrated'] == [] and out['unmigrated_count'] == 0
    ids = out['id_map']
    assert ids['aud-0001'] == ['lp-municipal:target']
    assert ids['msg-0001'] == ['lp-municipal:msg-0001']
    assert sorted(ids['msg-0002']) == ['deck-smb:msg-0002', 'lp-municipal:msg-0002']
    assert ids['lp-municipal:top[0]'] == ['lp-municipal:sec-0001', 'lp-municipal:cpy-0001']
    assert ids['lp-municipal#fv'] == ['lp-municipal:sec-0002']          # a v5 frame becomes a section id
    assert ids['lp-municipal#order'] == ['lp-municipal:sections.order']  # the sequence frame is the order
    assert ids['prm-0001'] == [] and ids['cpt-0001'] == []               # promises and concepts are dropped
    assert sorted(ids['prf-0001']) == ['prf-0001', 'prf-0003']           # evidence of two segments: a copy each
    dropped = {d['item'] for d in out['dropped']}
    assert {'prm-0001', 'prm-0002', 'cpt-0001', 'ledger.tombstones.msg-0004'} <= dropped
    stops = out['stops_applying']
    assert [r['key'] for r in stops['approvals']] == ['ia:strategy:prm-0001', 'ia:strategy:val-0001']
    assert [r['id'] for r in stops['judgments']] == ['ia-strategy-fit:lp-municipal']  # design rows stay
    assert out['module_pin'] == {'from': '0.3.0', 'to': ia_module.VERSION}
    assert out['rollback'] and any('git' in step for step in out['rollback'])
    assert set(out['files']['write']) >= {'design/info_arch/info_arch.yaml', 'design/design.yaml', 'site/index.html',
                                          'design/info_arch/outlines/lp-municipal.yaml'}
    assert 'design/approvals.yaml' not in out['files']['write'] and 'design/judgments.json' not in out['files']['write']
    assert out['drafting_required']  # value kinds, categories and fulfills arrows are judgment, never invented


# ---- --write -----------------------------------------------------------------------------------------------------

def test_write_moves_the_ssot_outlines_marks_and_pin(tmp_path):
    write_v5(tmp_path)
    keep = {p: (tmp_path / p).read_bytes() for p in ('design/approvals.yaml', 'design/judgments.json')}
    rc, out = migrate(tmp_path, '--write')
    assert rc == 0, out
    assert out['status'] == 'migrated' and out['written'] is True
    assert {p: (tmp_path / p).read_bytes() for p in keep} == keep  # never writes approvals or judgments
    ssot = load(tmp_path / 'design/info_arch/info_arch.yaml')
    assert ssot['schema_version'] == 4
    assert not {'audiences', 'messages', 'promises', 'concepts'} & set(ssot)
    assert set(ssot['ledger']['next']) == {'prf', 'fct', 'ent', 'vce', 'seg', 'job', 'pan', 'gan', 'val', 'alt', 'str', 'cat'}
    assert ssot['ledger']['tombstones'] == {'pan-0003': None}
    assert ssot['pains']['pan-0001'] == _a(segment='seg-0001', text='運用が重い')
    assert ssot['values']['val-0001']['segment'] == 'seg-0001' and 'kind' not in ssot['values']['val-0001']
    assert 'push' not in ssot['jobs']['job-0001']
    keys = {(e['kind'], e['from'], e['to']) for e in ssot['edges']}
    assert {('has_job', 'seg-0001', 'job-0001'), ('need', 'job-0001', 'pan-0001'), ('need', 'job-0001', 'gan-0001'),
            ('satisfies', 'val-0001', 'pan-0001'), ('satisfies', 'val-0001', 'gan-0001'),
            ('lacks', 'str-0001', 'alt-0001'), ('produces', 'str-0001', 'val-0001'),
            ('supports', 'prf-0001', 'val-0001'), ('supports', 'prf-0003', 'val-0002'),
            ('supports', 'prf-0002', 'val-0001')} <= keys
    assert ssot['proofs']['prf-0003']['segment'] == 'seg-0002' and 'copy of prf-0001' in ssot['proofs']['prf-0003']['provenance']
    assert ssot['ledger']['next']['prf'] == 4
    lp = load(tmp_path / 'design/info_arch/outlines/lp-municipal.yaml')
    assert lp['schema_version'] == 3
    assert lp['target'] == {'segment': 'seg-0001', 'condition': '自治体の情報部門'}
    assert lp['tone'] == '丁寧で落ち着いた調子' and lp['objective'] == {'id': 'obj-0001', 'action': '資料請求'}
    assert lp['messages'] == {
        'msg-0001': {'value': 'val-0001', 'rank': 'lead', 'type': 'normal'},
        'msg-0002': {'value': 'val-0001', 'rank': 'support', 'type': 'normal'},
        'msg-0003': {'value': 'val-0001', 'rank': 'support', 'type': 'cta', 'achieves': 'obj-0001'}}
    assert [(s['order'], s['role'], s['places']) for s in lp['sections'].values()] == [
        (1, 'pain', ['msg-0001']), (2, 'aware', ['msg-0001']), (3, 'trust', ['msg-0002']), (4, 'act', ['msg-0003'])]
    assert lp['copies']['cpy-0003'] == {'section': 'sec-0003', 'order': 1, 'message': 'msg-0002', 'text': '導入は{{fct-0001}}'}
    assert lp['texts'] == {'nav-footer': {'kind': 'navigation', 'text': 'お問い合わせ窓口'}}
    deck = load(tmp_path / 'design/info_arch/outlines/deck-smb.yaml')
    assert deck['target']['segment'] == 'seg-0002'
    assert deck['messages']['msg-0002'] == {'value': 'val-0002', 'rank': 'lead', 'type': 'normal'}
    html = (tmp_path / 'site/index.html').read_text(encoding='utf-8')
    for legacy in ('data-ia-msg', 'data-ia-pain', 'data-ia-frame', 'data-ia-problem'):
        assert legacy not in html
    assert '<p class="lead" data-ia-copy="cpy-0001">' in html                    # a single child carries the copy
    assert 'data-ia-section="sec-0003"' in html and 'data-ia-copy="cpy-0003"' in html
    assert html.count('data-ia-section=') == 4 and html.count('data-ia-copy=') == 4
    assert json.loads((tmp_path / 'design/design.yaml').read_text(encoding='utf-8'))['modules']['ia'] == ia_module.VERSION
    rc, again = migrate(tmp_path)
    assert rc == 0 and again['status'] == 'already-v6'  # rerunnable: a v6 tree is a no-op


def draft_ssot(base):
    """The positioning judgment the migration lists and never invents: value kinds, member categories, recalls,
    includes and fulfills degrees (as the AI drafts them from the owner's material after the migration)."""
    ssot = load(base / 'info_arch.yaml')
    for v in ssot['values'].values():
        v['kind'] = 'difference'
    ssot['categories'] = {'cat-0001': _a(segment='seg-0001', text='自治体向け Web サーバー'),
                          'cat-0002': _a(segment='seg-0002', text='中小企業向けの業務システム')}
    ssot['ledger']['next']['cat'] = 3

    def edge(kind, frm, to, **kw):
        return {'kind': kind, 'from': frm, 'to': to, 'state': 'proposed', 'provenance': 'owner', **kw}
    ssot['edges'] += [edge('member_of', 'seg-0001', 'cat-0001'), edge('recalls', 'job-0001', 'cat-0001'),
                      edge('includes', 'cat-0001', 'alt-0001'), edge('fulfills', 'alt-0001', 'pan-0001', degree='none'),
                      edge('fulfills', 'alt-0001', 'gan-0001', degree='none'),
                      edge('member_of', 'seg-0002', 'cat-0002'), edge('recalls', 'job-0002', 'cat-0002'),
                      edge('includes', 'cat-0002', 'alt-0002'), edge('fulfills', 'alt-0002', 'pan-0002', degree='partial')]
    dump(base / 'info_arch.yaml', ssot)


def test_migrated_tree_passes_check_once_the_listed_drafting_is_done(tmp_path):
    """The migration's output is v6-shaped: the only work left is the judgment it lists (value kinds, categories,
    fulfills degrees, a normal message beside each CTA). With that drafted, `check` has no failure."""
    write_v5(tmp_path)
    rc, out = migrate(tmp_path, '--write')
    assert rc == 0, out
    rc, report = check(tmp_path)
    assert rc == 2  # not done yet: drafting_required is real work
    base = tmp_path / 'design/info_arch'
    draft_ssot(base)
    for oid, (into, cta_sec, cta_copy) in {'lp-municipal': ('sec-0003', 'sec-0004', 'cpy-0004'),
                                           'deck-smb': ('sec-0001', 'sec-0002', 'cpy-0002')}.items():
        o = load(base / f'outlines/{oid}.yaml')
        o['sections'][into]['places'] += o['sections'].pop(cta_sec)['places']
        o['copies'][cta_copy].update(section=into, order=2)
        dump(base / f'outlines/{oid}.yaml', o)
    rc, report = check(tmp_path)
    assert codes(report) == [], report['findings']
    assert rc == 1  # warnings only: the owner has not approved the new value paths yet
    assert 'IA_OUTLINE_VALUE_UNAPPROVED' in codes(report, 'warn')


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_rewritten_marks_pass_the_real_browser_screen_check(tmp_path):
    """dom fails on legacy marks (IA_DOM_LEGACY_ATTRIBUTE); after the rewrite the real browser finds every section and
    copy anchor in order, with the copy word for word. The outline's CTA-only section is left as migrated (an outline
    finding, not a screen one), so the page is the one the migration wrote."""
    write_v5(tmp_path)
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)

    def dom():
        proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                               '--outline', 'lp-municipal', '--date', '2026-10-23'], capture_output=True, text=True,
                              env=env, check=False)
        report = json.loads(proc.stdout)
        assert report['targets'].get('viewports') == 2, report  # the browser really ran (not stopped at check)
        return [c for c in codes(report) if c.startswith('IA_DOM_')]
    rc, out = migrate(tmp_path, '--write')
    assert rc == 0, out
    draft_ssot(tmp_path / 'design/info_arch')
    assert dom() == []
    page = tmp_path / 'site/index.html'
    page.write_text(page.read_text(encoding='utf-8').replace('data-ia-copy="cpy-0002"', 'data-ia-msg="msg-0001"'),
                    encoding='utf-8')
    assert 'IA_DOM_LEGACY_ATTRIBUTE' in dom()  # the same check is red on a mark the migration left behind


# ---- refusals and no-ops -----------------------------------------------------------------------------------------

def test_v6_tree_is_a_noop_and_unknown_version_fails(tmp_path):
    write_v5(tmp_path)
    ssot = copy.deepcopy(V5_SSOT)
    ssot['schema_version'] = 4
    dump(tmp_path / 'design/info_arch/info_arch.yaml', ssot)
    before = snapshot(tmp_path)
    rc, out = migrate(tmp_path, '--write')
    assert rc == 0 and out['status'] == 'already-v6' and snapshot(tmp_path) == before
    ssot['schema_version'] = 2
    dump(tmp_path / 'design/info_arch/info_arch.yaml', ssot)
    rc, out = migrate(tmp_path)
    assert rc == 2 and out['exit_code'] == 2


# ---- red cases: every unmigrated condition --------------------------------------------------------------------------

def _apply(root, ops):
    ssot, outlines = copy.deepcopy(V5_SSOT), {o['id']: copy.deepcopy(o) for o in V5_OUTLINES}
    design, html = copy.deepcopy(DESIGN), V5_HTML
    for op in ops:
        f = op['file']
        if f.startswith('html:'):
            assert op['op'] == 'replace' and op['old'] in html, op
            html = html.replace(op['old'], op['new'])
            continue
        doc = ssot if f == 'ssot' else design if f == 'design' else outlines[f.split(':', 1)[1]]
        *parents, last = op['path']
        for k in parents:
            doc = doc[k]
        if op['op'] == 'set':
            doc[last] = op['value']
        elif op['op'] == 'delete':
            del doc[last]
        elif op['op'] == 'append':
            doc[last].append(op['value'])
        else:
            raise AssertionError(op)
    write_v5(root, ssot, list(outlines.values()), html, design)


def test_every_condition_has_a_red_case_and_every_case_names_a_condition():
    covered = {c['condition'] for c in CASES}
    assert sorted(set(CONDITIONS) - covered) == []  # 対応の無い移しそこねの条件 0 件
    assert sorted(covered - set(CONDITIONS)) == []
    assert len({c['id'] for c in CASES}) == len(CASES)
    source = (SCRIPTS / 'migrate_v6.py').read_text(encoding='utf-8')
    for c in CONDITIONS.values():
        assert f"'{c['reason']}'" in source, c  # a reason the tool never raises would stay green forever


@pytest.mark.parametrize('case', CASES, ids=[c['id'] for c in CASES])
def test_red_case_fails_and_writes_nothing(case, tmp_path):
    _apply(tmp_path, case['ops'])
    before = snapshot(tmp_path)
    rc, out = migrate(tmp_path, '--write')
    assert rc == 2 and out['exit_code'] == 2, out
    assert out['unmigrated_count'] == len(out['unmigrated']) > 0
    assert CONDITIONS[case['condition']]['reason'] in {u['reason'] for u in out['unmigrated']}, out['unmigrated']
    assert out['written'] is False and snapshot(tmp_path) == before  # one batch: all or nothing
