"""Messaging SSOT + structured outline: green fixture, every required red case, impact lookup."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
POSITIONING_TESTS = SKILL.parent / 'own-positioning-author/scripts/tests'
sys.path.insert(0, str(POSITIONING_TESTS))
from ia_fixtures import EDGES, SSOT, _a, _e, confirm, dump  # noqa: E402,F401  (the positioning the outlines stand on)
DATE = '2026-10-23'
RED = json.loads((SKILL / 'fixtures/red-cases.json').read_text(encoding='utf-8'))['cases']

CHANNEL = {'schema_version': 2, 'media': 'web', 'outputs': ['site/**/*.html'], 'ship': True, 'ship_command': 'mise run deploy'}
DRAFT = {'schema_version': 2, 'media': 'deck', 'outputs': ['decks/*.pptx'], 'ship': False}
# outline v3 (requirement 6): stands on the target segment's positioning; constants, objective, then the three
# independent variables. Order lives in `order`, never in how the file is written.
OUTLINE = {
    'schema_version': 3, 'id': 'lp-municipal', 'channel': 'web', 'outputs': ['site/index.html'],
    'target': {'segment': 'seg-0001', 'condition': '既存システムの運用が重いと感じている情報部門の担当者'},
    'touchpoint': 'lp', 'tone': '丁寧で落ち着いた調子',
    'objective': {'id': 'obj-0001', 'action': '資料請求'},
    'messages': {
        'msg-0001': {'value': 'val-0001', 'rank': 'lead', 'type': 'normal', 'compares': ['alt-0001']},
        'msg-0002': {'value': 'val-0003', 'rank': 'support', 'type': 'normal'},
        'msg-0003': {'value': 'val-0001', 'rank': 'support', 'type': 'cta', 'achieves': 'obj-0001'}},
    'sections': {
        'sec-0001': {'order': 1, 'role': 'pain', 'places': ['msg-0001']},
        'sec-0002': {'order': 2, 'role': 'trust', 'places': ['msg-0002']},
        'sec-0003': {'order': 3, 'role': 'act', 'places': ['msg-0001', 'msg-0003']}},
    'copies': {
        'cpy-0001': {'section': 'sec-0001', 'order': 1, 'message': 'msg-0001', 'text': '既存システムの運用、重くありませんか'},
        'cpy-0002': {'section': 'sec-0001', 'order': 2, 'message': 'msg-0001', 'text': '運用の手間が半分になります'},
        'cpy-0003': {'section': 'sec-0002', 'order': 1, 'message': 'msg-0002', 'text': '導入は{{value-customer-count}}、止まらずに動き続けます'},
        'cpy-0004': {'section': 'sec-0003', 'order': 1, 'message': 'msg-0001', 'text': '保守の要らない運用へ'},
        'cpy-0005': {'section': 'sec-0003', 'order': 2, 'message': 'msg-0003', 'text': '資料請求はこちら（{{value-phone}}）'}},
}
DECK = {
    'schema_version': 3, 'id': 'deck-smb', 'channel': 'deck', 'outputs': ['decks/smb.pptx'],
    'target': {'segment': 'seg-0002', 'condition': '人手が足りない中小企業の経営者'},
    'touchpoint': 'deck', 'tone': '端的で親しみやすい調子',
    'objective': {'id': 'obj-0001', 'action': '問い合わせ'},
    'messages': {
        'msg-0001': {'value': 'val-0002', 'rank': 'lead', 'type': 'normal'},
        'msg-0002': {'value': 'val-0002', 'rank': 'support', 'type': 'cta', 'achieves': 'obj-0001'}},
    'sections': {
        'sec-0001': {'order': 1, 'role': 'aware', 'places': ['msg-0001']},
        'sec-0002': {'order': 2, 'role': 'act', 'places': ['msg-0001', 'msg-0002']}},
    'copies': {
        'cpy-0001': {'section': 'sec-0001', 'order': 1, 'message': 'msg-0001', 'text': '{{value-customer-count}}が導入、少人数でも回ります'},
        'cpy-0002': {'section': 'sec-0002', 'order': 1, 'message': 'msg-0001', 'text': '少人数で回す仕組みを'},
        'cpy-0003': {'section': 'sec-0002', 'order': 2, 'message': 'msg-0002', 'text': 'お問い合わせ {{value-phone}}'}},
}

# The docs dictionary the outlines copy from (docs/dictionary/<domain>.json): fixed values and the official name.
DICTIONARY = {'schema_version': 1, 'domain': 'product', 'summary': 'テスト用の辞書', 'entries': [
    {'id': 'value-customer-count', 'kind': 'value', 'text': '100社以上', 'meaning': '導入社数', 'expires': '2027-10-01'},
    {'id': 'value-phone', 'kind': 'value', 'text': '03-1234-5678', 'meaning': '問い合わせ電話', 'evergreen': True},
    {'id': 'name-web-server', 'kind': 'name', 'text': 'Webサーバー', 'meaning': '製品名'}]}


def view_of(ssot, dictionary=DICTIONARY):
    """An SSOT with the dictionary's values and names laid in the way the loader does (what dom, judge and skeleton read)."""
    import check as ia_check
    facts, entities = ia_check.terms_of(dictionary['entries'])
    return {**ssot, 'facts': facts, 'entities': entities}


SSOT_VIEW = view_of(SSOT)


def write_dictionary(root, dictionary):
    path = root / 'docs/dictionary' / f'{dictionary["domain"]}.json'
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(dictionary, ensure_ascii=False, indent=2), encoding='utf-8')


def write_repo(root, ssot=SSOT, channels=(CHANNEL, DRAFT), outlines=(OUTLINE, DECK), confirmed=True, dictionary=DICTIONARY):
    base = root / 'design/info_arch'
    dump(base / 'info_arch.yaml', ssot)
    if dictionary is not None:
        write_dictionary(root, dictionary)
    if confirmed:
        confirm(root, ssot)
    for ch in channels:
        dump(base / 'channels' / f'{ch["media"]}.yaml', ch)
    for o in outlines:
        dump(base / 'outlines' / f'{o["id"]}.yaml', o)
    return base


def record_red_case(command, rc, report):
    """RED_CASE_LOG=<file>: append one line per CLI run (test id, command, exit code, fail codes). The before/after
    table of the positioning split (WS-20261120 ISSUE-04) compares two such logs; unset, this does nothing."""
    import os
    path = os.environ.get('RED_CASE_LOG')
    if not path:
        return
    fails = sorted({f['code'] for f in report.get('findings', []) if f.get('level') == 'fail'})
    with open(path, 'a', encoding='utf-8') as fh:
        fh.write(json.dumps({'test': os.environ.get('PYTEST_CURRENT_TEST', '').split(' ')[0], 'command': command,
                             'exit_code': rc, 'fail_codes': fails}, ensure_ascii=False) + '\n')


POSITIONING_COMMANDS = {'positioning', 'migrate-approvals', 'review'}  # the positioning layer's CLI, not this skill's
POSITIONING_SCRIPTS = SKILL.parent / 'own-positioning-author/scripts'


def cli(*args):
    script = (POSITIONING_SCRIPTS / 'positioning_author.py') if args[0] in POSITIONING_COMMANDS else (SCRIPTS / 'infoarch_author.py')
    proc = subprocess.run([sys.executable, str(script), *args], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    report = json.loads(proc.stdout)
    record_red_case(args[0], proc.returncode, report)
    return proc.returncode, report


def check(root):
    rc, report = cli('check', '--root', str(root), '--date', DATE)
    assert report['exit_code'] == rc
    return rc, report


def codes(report, level='fail'):
    return [f['code'] for f in report['findings'] if f['level'] == level]

def test_green_fixture_passes_with_real_targets(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
    assert report['targets'] == {'channels': 2, 'outlines': 2, 'sections': 5}


def mutate(case, root):
    ssot, outline, deck = copy.deepcopy(SSOT), copy.deepcopy(OUTLINE), copy.deepcopy(DECK)
    dic = copy.deepcopy(DICTIONARY)
    channels = [copy.deepcopy(CHANNEL), copy.deepcopy(DRAFT)]
    cp = outline['copies']
    if case == 'outline-ref-broken':
        outline['sections']['sec-0002']['places'] = ['msg-0099']
    elif case == 'outline-ref-tombstoned':
        ssot['ledger']['next']['val'] = 5
        ssot['ledger']['tombstones'] = {'val-0004': 'val-0001'}
        outline['messages']['msg-0002']['value'] = 'val-0004'
    elif case == 'outline-ref-empty':
        ssot['values']['val-0003'] = {'state': 'empty', 'provenance': 'proposal'}
        ssot['edges'] = [e for e in ssot['edges'] if 'val-0003' not in (e['from'], e['to'])]
    elif case == 'outline-placeholder-unknown':
        cp['cpy-0003']['text'] = '導入は{{value-unknown}}'
    elif case == 'outline-placeholder-malformed':
        cp['cpy-0003']['text'] = '導入は{{fct-1}}と{{ msg-0001 }}'
    elif case == 'outline-fact-literal':
        cp['cpy-0003']['text'] = '導入は100社以上'
    elif case == 'outline-channel-unknown':
        outline['channel'] = 'poster'
    elif case == 'outline-output-undeclared':
        outline['outputs'] = ['other/index.html']
    elif case == 'outline-id-mismatch':
        base = write_repo(root, ssot, channels, [outline, deck], dictionary=dic)
        (base / 'outlines/lp-municipal.yaml').rename(base / 'outlines/lp-renamed.yaml')
        return
    elif case == 'expired-proof-shipped':
        ssot['proofs']['prf-0001']['as_of'] = '2024-01-01'
    elif case == 'expired-fact-shipped':
        dic['entries'][0]['expires'] = '2026-10-22'
    elif case == 'no-outlines':
        outline = deck = None
    elif case == 'no-channels':
        channels = []
    elif case == 'ssot-state-approved':
        ssot['values']['val-0001']['state'] = 'approved'
    elif case == 'ssot-ref-to-empty':
        ssot['ledger']['next']['seg'] = 4
        ssot['segments']['seg-0003'] = {'state': 'empty', 'provenance': 'proposal'}
        ssot['proofs']['prf-0003']['segment'] = 'seg-0003'
    elif case == 'ssot-id-not-issued':
        ssot['ledger']['next']['alt'] = 3
    elif case == 'ssot-voices-block':
        ssot['voices'] = {'vce-0001': {'state': 'proposed', 'provenance': 'owner', 'traits': []}}
    elif case == 'channel-ship-without-command':
        del channels[0]['ship_command']
    elif case == 'ssot-duplicate-key':
        write_repo(root, ssot, channels, [o for o in (outline, deck) if o], dictionary=dic)
        p = root / 'design/info_arch/info_arch.yaml'
        p.write_text(p.read_text(encoding='utf-8') + 'schema_version: 2\n', encoding='utf-8')
        return
    elif case == 'outline-placeholder-multiline':
        cp['cpy-0003']['text'] = '導入は{{fct-001\n}}と{{ msg-0001\n}}'
    elif case == 'ssot-id-noncanonical':
        ssot['ledger']['next']['str'] = 4
        ssot['strengths']['str-00003'] = copy.deepcopy(ssot['strengths']['str-0001'])
    elif case == 'ssot-evergreen-with-date':
        dic['entries'][1]['expires'] = '2020-01-01'
    elif case == 'outline-dot-directory':
        base = write_repo(root, ssot, channels, [outline, deck], dictionary=dic)
        (base / 'outlines/.drafts').mkdir()
        (base / 'outlines/.drafts/broken.yaml').write_text('a: [', encoding='utf-8')
        return
    elif case == 'ssot-id-zero-serial':
        ssot['strengths']['str-0000'] = copy.deepcopy(ssot['strengths']['str-0001'])
    elif case == 'ssot-id-trailing-newline':
        ssot['strengths']['str-0001\n'] = copy.deepcopy(ssot['strengths']['str-0001'])
    elif case == 'outline-text-section-broken':
        outline['texts'] = {'label-problem': {'kind': 'label', 'text': 'PROBLEM', 'section': 'sec-0099'}}
    elif case == 'outline-placeholder-padded':
        deck['copies']['cpy-0001']['text'] = '{{\nvalue-customer-count\n}}が導入'
    elif case == 'outline-v2':
        outline['schema_version'] = 2
    else:
        raise KeyError(case)
    write_repo(root, ssot, channels, [o for o in (outline, deck) if o], dictionary=dic)


@pytest.mark.parametrize('case', RED, ids=[c['id'] for c in RED])
def test_red_case_fails_with_exit_2(case, tmp_path):
    mutate(case['id'], tmp_path)
    rc, report = check(tmp_path)
    assert rc == 2, report['findings']
    assert case['code'] in codes(report), codes(report)


def test_red_case_table_and_mutations_agree(tmp_path):
    for c in RED:
        mutate(c['id'], tmp_path / c['id'])
    with pytest.raises(KeyError):
        mutate('not-a-case', tmp_path / 'x')


def test_expired_but_not_shipped_is_reported_not_failed(tmp_path):
    ssot = copy.deepcopy(SSOT)
    dic = copy.deepcopy(DICTIONARY)
    dic['entries'].append({'id': 'value-old-fee', 'kind': 'value', 'text': '月額 1万円', 'meaning': '旧料金', 'expires': '2025-01-01'})
    write_repo(tmp_path, ssot, dictionary=dic)
    rc, report = check(tmp_path)
    assert rc == 0 and 'IA_EXPIRED_UNSHIPPED' in codes(report, 'info'), report['findings']


def test_expiry_days_policy_overrides_default(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['policy'] = {'expiry_days': {'proof': 10}}
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_EXPIRED_SHIPPED' in codes(report), report['findings']


def test_evidence_off_every_shipped_path_expires_quietly(tmp_path):
    """prf-0002 backs val-0002, which only the draft deck stands on: expired, reported, not failed."""
    ssot = copy.deepcopy(SSOT)
    ssot['proofs']['prf-0002']['as_of'] = '2024-01-01'
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 0 and 'IA_EXPIRED_UNSHIPPED' in codes(report, 'info'), report['findings']


def test_impact_lists_outlines_and_outputs(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'value-customer-count')
    assert rc == 0 and out['outlines'] == ['deck-smb', 'lp-municipal']
    assert out['outputs'] == ['decks/smb.pptx', 'site/index.html']
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'seg-0002')
    assert out['outlines'] == ['deck-smb']
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'prf-0001')
    assert out['outlines'] == ['lp-municipal']  # evidence on the path of the LP's lead value


def test_crash_is_exit_2(tmp_path):
    write_repo(tmp_path)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'check', '--root', str(tmp_path), '--date', '2026-13-45'],
                          capture_output=True, text=True, check=False)
    assert proc.returncode == 2 and json.loads(proc.stdout)['exit_code'] == 2


def test_template_is_schema_valid(tmp_path):
    import shutil
    shutil.copytree(SKILL / 'templates/design', tmp_path / 'design')  # channels and the guide: this skill's
    shutil.copytree(SKILL.parent / 'own-positioning-author/templates/design', tmp_path / 'design', dirs_exist_ok=True)  # the empty SSOT
    rc, report = check(tmp_path)
    assert 'IA_SCHEMA' not in codes(report), report['findings']
    assert rc == 2 and 'IA_NO_OUTLINES' in codes(report)


def test_wrapper_uses_design_author_python(tmp_path):
    import os
    write_repo(tmp_path)
    env = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable)
    proc = subprocess.run([str(SCRIPTS / 'infoarch-author'), 'check', '--root', str(tmp_path), '--date', DATE],
                          capture_output=True, text=True, env=env, check=False)
    assert proc.returncode == 0, proc.stdout + proc.stderr


def test_impact_on_unknown_id_is_exit_2(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'val-9999')
    assert rc == 2 and 'IA_IMPACT_UNKNOWN_ID' in codes(out)


def test_dotfiles_in_channels_and_outlines_are_ignored(tmp_path):
    base = write_repo(tmp_path)
    (base / 'channels/.DS_Store').write_bytes(b'x')
    (base / 'outlines/.gitkeep').write_text('', encoding='utf-8')
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']


def test_wrapper_exits_2_when_interpreter_is_missing(tmp_path):
    import os
    write_repo(tmp_path)
    env = dict(os.environ, DESIGN_AUTHOR_PYTHON=str(tmp_path / 'no-such-python'))
    proc = subprocess.run([str(SCRIPTS / 'infoarch-author'), 'check', '--root', str(tmp_path)],
                          capture_output=True, text=True, env=env, check=False)
    assert proc.returncode == 2, (proc.returncode, proc.stdout, proc.stderr)


def test_lead_value_without_evidence_is_a_warning(tmp_path):
    """The lead message's value has nothing on its positioning path that evidence supports."""
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if e['kind'] != 'supports' or e['from'] != 'prf-0002']
    del ssot['proofs']['prf-0002']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    hits = [f for f in report['findings'] if f['code'] == 'IA_LEAD_UNSUPPORTED']
    assert rc == 1 and [f['level'] for f in hits] == ['warn'], report['findings']
    assert 'deck-smb' in hits[0]['message'] and 'val-0002' in hits[0]['message']


def test_declared_navigation_and_label_texts_pass(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['texts'] = {'nav': {'kind': 'navigation', 'text': 'サービス 使い方'},
                        'label-problem': {'kind': 'label', 'text': 'PROBLEM', 'section': 'sec-0001'}}
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
