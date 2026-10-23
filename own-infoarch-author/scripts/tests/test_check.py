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
DATE = '2026-10-23'
RED = json.loads((SKILL / 'fixtures/red-cases.json').read_text(encoding='utf-8'))['cases']

SSOT = {
    'schema_version': 2,
    'ledger': {'next': {'aud': 3, 'prb': 2, 'msg': 5, 'prf': 2, 'fct': 3, 'ent': 2, 'vce': 2},
               'tombstones': {'msg-0004': 'msg-0003'}},
    'audiences': {
        'aud-0001': {'state': 'proposed', 'provenance': 'owner', 'label': '自治体の情報部門', 'situation': '既存システムの運用が重い',
                     'pains': ['運用負荷'], 'desired_action': '資料請求', 'tone': '丁寧で落ち着いた調子',
                     'priorities': ['msg-0001', 'msg-0002', 'msg-0003']},
        'aud-0002': {'state': 'proposed', 'provenance': 'owner', 'label': '中小企業の経営者', 'situation': '人手が足りない',
                     'desired_action': '問い合わせ', 'priorities': ['msg-0002', 'msg-0001', 'msg-0003']},
    },
    'problems': {'prb-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '運用が重い', 'audience': 'aud-0001'}},
    'messages': {
        'msg-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '運用の手間が半分になる', 'role': 'benefit', 'addresses': ['prb-0001']},
        'msg-0002': {'state': 'proposed', 'provenance': 'owner', 'text': '多くの自治体で使われている', 'role': 'proof'},
        'msg-0003': {'state': 'proposed', 'provenance': 'owner', 'text': '資料を請求する', 'role': 'cta'},
    },
    'proofs': {'prf-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '導入自治体の数', 'source': '導入台帳',
                            'supports': ['msg-0002'], 'as_of': '2026-10-01'}},
    'facts': {
        'fct-0001': {'state': 'proposed', 'provenance': 'owner', 'label': '導入社数', 'value': '100社以上', 'as_of': '2026-10-01'},
        'fct-0002': {'state': 'proposed', 'provenance': 'owner', 'label': '問い合わせ電話', 'value': '03-1234-5678', 'evergreen': True},
    },
    'entities': {'ent-0001': {'state': 'proposed', 'provenance': 'owner', 'type': 'product', 'name': 'Webサーバー'}},
    'voices': {'vce-0001': {'state': 'proposed', 'provenance': 'owner', 'traits': [
        {'trait': 'わかりやすい', 'is': '専門的な中身も噛み砕いて言う', 'is_not': '子ども扱いはしない'},
        {'trait': '誠実', 'is': 'できることを具体的に言う', 'is_not': '大げさな約束はしない'},
        {'trait': '前向き', 'is': '解決の道筋を示す', 'is_not': '浮ついた煽りはしない'}]}},
}
CHANNEL = {'schema_version': 2, 'media': 'web', 'outputs': ['site/**/*.html'], 'ship': True, 'ship_command': 'mise run deploy'}
DRAFT = {'schema_version': 2, 'media': 'deck', 'outputs': ['decks/*.pptx'], 'ship': False}
OUTLINE = {
    'schema_version': 1, 'id': 'lp-municipal', 'channel': 'web',
    'brief': {'purpose': '自治体の情報部門からの資料請求を増やす', 'action': '資料請求', 'medium': 'LP', 'constraints': ['1 ページ']},
    'audience': 'aud-0001', 'outputs': ['site/index.html'],
    'pages': [{'id': 'top', 'label': 'トップ', 'sections': [
        {'role': 'problem', 'problem': 'prb-0001', 'copy': '既存システムの運用、重くありませんか'},
        {'role': 'benefit', 'message': 'msg-0001', 'copy': '運用の手間が半分になります'},
        {'role': 'proof', 'message': 'msg-0002', 'proofs': ['prf-0001'], 'copy': '導入は{{fct-0001}}'},
        {'role': 'cta', 'message': 'msg-0003', 'copy': '資料請求はこちら（{{fct-0002}}）'},
    ]}],
}
DECK = {
    'schema_version': 1, 'id': 'deck-smb', 'channel': 'deck',
    'brief': {'purpose': '中小企業向けの説明', 'action': '問い合わせ', 'medium': '資料', 'cut': ['msg-0001']},
    'audience': 'aud-0002', 'outputs': ['decks/smb.pptx'],
    'pages': [{'id': 's1', 'sections': [
        {'role': 'proof', 'message': 'msg-0002', 'proofs': ['prf-0001'], 'copy': '{{fct-0001}}が導入'},
        {'role': 'cta', 'message': 'msg-0003', 'copy': 'お問い合わせ {{fct-0002}}'}]}],
}


def dump(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding='utf-8')


def write_repo(root, ssot=SSOT, channels=(CHANNEL, DRAFT), outlines=(OUTLINE, DECK)):
    base = root / 'design/info_arch'
    dump(base / 'info_arch.yaml', ssot)
    for ch in channels:
        dump(base / 'channels' / f'{ch["media"]}.yaml', ch)
    for o in outlines:
        dump(base / 'outlines' / f'{o["id"]}.yaml', o)
    return base


def cli(*args):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), *args], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


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
    assert report['targets'] == {'channels': 2, 'outlines': 2, 'sections': 6}


def test_two_artifacts_share_content_but_differ_in_order():
    """Spec Requirements 12: same facts and messages, different priority and structure per brief."""
    def msgs(o):
        return [s['message'] for p in o['pages'] for s in p['sections'] if 'message' in s]
    def copy_text(o):
        return ' '.join(s['copy'] for p in o['pages'] for s in p['sections'])
    assert set(msgs(DECK)) <= set(msgs(OUTLINE))
    assert msgs(DECK)[0] != msgs(OUTLINE)[0]
    assert '{{fct-0001}}' in copy_text(DECK) and '{{fct-0001}}' in copy_text(OUTLINE)


def mutate(case, root):
    ssot, outline, deck = copy.deepcopy(SSOT), copy.deepcopy(OUTLINE), copy.deepcopy(DECK)
    channels = [copy.deepcopy(CHANNEL), copy.deepcopy(DRAFT)]
    sec = outline['pages'][0]['sections']
    if case == 'outline-ref-broken':
        sec[1]['message'] = 'msg-0099'
    elif case == 'outline-ref-tombstoned':
        sec[3]['message'] = 'msg-0004'
    elif case == 'outline-ref-empty':
        ssot['messages']['msg-0001'] = {'state': 'empty', 'provenance': 'proposal'}
    elif case == 'outline-placeholder-unknown':
        sec[2]['copy'] = '導入は{{fct-0099}}'
    elif case == 'outline-placeholder-malformed':
        sec[2]['copy'] = '導入は{{fct-1}}と{{ msg-0001 }}'
    elif case == 'outline-fact-literal':
        sec[2]['copy'] = '導入は100社以上'
    elif case == 'outline-cut-used':
        deck['brief']['cut'] = ['msg-0002']
    elif case == 'outline-channel-unknown':
        outline['channel'] = 'poster'
    elif case == 'outline-output-undeclared':
        outline['outputs'] = ['other/index.html']
    elif case == 'outline-id-mismatch':
        base = write_repo(root, ssot, channels, [outline, deck])
        (base / 'outlines/lp-municipal.yaml').rename(base / 'outlines/lp-renamed.yaml')
        return
    elif case == 'expired-proof-shipped':
        ssot['proofs']['prf-0001']['as_of'] = '2024-01-01'
    elif case == 'expired-fact-shipped':
        ssot['facts']['fct-0001']['expires'] = '2026-10-22'
    elif case == 'no-outlines':
        outline = deck = None
    elif case == 'no-channels':
        channels = []
    elif case == 'ssot-state-approved':
        ssot['messages']['msg-0001']['state'] = 'approved'
    elif case == 'ssot-ref-to-empty':
        ssot['messages']['msg-0002'] = {'state': 'empty', 'provenance': 'proposal'}
        deck['pages'][0]['sections'] = [deck['pages'][0]['sections'][1]]
        sec.pop(2)
    elif case == 'ssot-id-not-issued':
        ssot['ledger']['next']['fct'] = 2
    elif case == 'ssot-voice-multiple':
        ssot['ledger']['next']['vce'] = 3
        ssot['voices']['vce-0002'] = copy.deepcopy(ssot['voices']['vce-0001'])
    elif case == 'channel-ship-without-command':
        del channels[0]['ship_command']
    elif case == 'ssot-duplicate-key':
        write_repo(root, ssot, channels, [o for o in (outline, deck) if o])
        p = root / 'design/info_arch/info_arch.yaml'
        p.write_text(p.read_text(encoding='utf-8') + 'schema_version: 2\n', encoding='utf-8')
        return
    elif case == 'outline-cut-ref-broken':
        deck['brief']['cut'] = ['msg-0099']
    elif case == 'outline-cut-ref-empty':
        ssot['ledger']['next']['msg'] = 6
        ssot['messages']['msg-0005'] = {'state': 'empty', 'provenance': 'proposal'}
        deck['brief']['cut'] = ['msg-0005']
    elif case == 'outline-cut-tombstoned-used':
        deck['brief']['cut'] = ['msg-0004']
    elif case == 'outline-placeholder-multiline':
        sec[2]['copy'] = '導入は{{fct-001\n}}と{{ msg-0001\n}}'
    elif case == 'ssot-id-noncanonical':
        ssot['messages']['msg-00001'] = copy.deepcopy(ssot['messages']['msg-0001'])
    elif case == 'ssot-evergreen-with-date':
        ssot['facts']['fct-0002']['expires'] = '2020-01-01'
    else:
        raise KeyError(case)
    write_repo(root, ssot, channels, [o for o in (outline, deck) if o])


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
    ssot['ledger']['next']['fct'] = 4
    ssot['facts']['fct-0003'] = {'state': 'proposed', 'provenance': 'owner', 'label': '旧料金', 'value': '月額 1万円', 'expires': '2025-01-01'}
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 0 and 'IA_EXPIRED_UNSHIPPED' in codes(report, 'info'), report['findings']


def test_expiry_days_policy_overrides_default(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['policy'] = {'expiry_days': {'proof': 10}}
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_EXPIRED_SHIPPED' in codes(report), report['findings']


def test_message_outside_audience_priorities_is_a_warning(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['audiences']['aud-0001']['priorities'] = ['msg-0001', 'msg-0003']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 1 and 'IA_OUTLINE_OFF_PRIORITY' in codes(report, 'warn'), report['findings']


def test_impact_lists_outlines_and_outputs(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'fct-0001')
    assert rc == 0 and out['outlines'] == ['deck-smb', 'lp-municipal']
    assert out['outputs'] == ['decks/smb.pptx', 'site/index.html']
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'aud-0002')
    assert out['outlines'] == ['deck-smb']
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'prf-0001')
    assert out['outlines'] == ['deck-smb', 'lp-municipal']


def test_crash_is_exit_2(tmp_path):
    write_repo(tmp_path)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'check', '--root', str(tmp_path), '--date', '2026-13-45'],
                          capture_output=True, text=True, check=False)
    assert proc.returncode == 2 and json.loads(proc.stdout)['exit_code'] == 2


def test_template_is_schema_valid(tmp_path):
    import shutil
    shutil.copytree(SKILL / 'templates/design', tmp_path / 'design')
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


def test_proof_supporting_an_empty_message_fails_on_its_own(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['ledger']['next']['msg'] = 6
    ssot['messages']['msg-0005'] = {'state': 'empty', 'provenance': 'proposal'}
    ssot['proofs']['prf-0001']['supports'] = ['msg-0002', 'msg-0005']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_EMPTY_REFERENCED' in codes(report), report['findings']


def test_impact_on_unknown_id_is_exit_2(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'msg-9999')
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
