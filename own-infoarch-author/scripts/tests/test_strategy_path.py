"""Strategy layer (SPEC-information-architecture-standardization v5, requirements 13 and 15).

A message in an outline must trace, through the strategy atoms it serves, to the
segment of the outline's audience. The path is derived from references alone:
the deliverable cannot leave out a strength that supports a value on its path.
Shipped deliverables fail; drafts warn.
"""
import copy
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
from test_check import CHANNEL, DECK, DRAFT, OUTLINE, SSOT, check, codes, write_repo  # noqa: E402


def model(ssot=SSOT):
    return ia_check.Model(root=Path('.'), ssot=ssot)


def path_ids(ssot, mid, seg):
    return ia_check.strategy_path(model(ssot), mid, seg)[0]


def test_v2_ssot_is_rejected_without_compatibility(tmp_path):
    old = copy.deepcopy(SSOT)
    old['schema_version'] = 2
    write_repo(tmp_path, old)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_SCHEMA' in codes(report), report['findings']


def test_path_follows_only_the_audience_segment_and_pulls_in_strengths():
    ids = path_ids(SSOT, 'msg-0001', 'seg-0001')
    assert ids == ['alt-0001', 'gan-0001', 'job-0001', 'pan-0001', 'prm-0001', 'seg-0001', 'str-0001', 'val-0001']
    # msg-0002 serves values of both segments; each audience sees only its own branch
    assert path_ids(SSOT, 'msg-0002', 'seg-0002') == ['job-0002', 'pan-0002', 'seg-0002', 'val-0002']


def test_strength_for_an_off_path_value_is_not_on_the_path():
    ssot = copy.deepcopy(SSOT)
    ssot['strengths']['str-0001']['produces'] = ['val-0002']
    assert 'str-0001' not in path_ids(ssot, 'msg-0001', 'seg-0001')
    assert 'str-0001' in path_ids(ssot, 'msg-0002', 'seg-0002')


def test_path_hash_moves_with_any_path_atom_and_ignores_others():
    base = ia_check.strategy_path(model(), 'msg-0001', 'seg-0001')[1]
    for block, aid, fld in (('strengths', 'str-0001', 'text'), ('alternatives', 'alt-0001', 'text'),
                            ('jobs', 'job-0001', 'text'), ('segments', 'seg-0001', 'situation')):
        ssot = copy.deepcopy(SSOT)
        ssot[block][aid][fld] += '（改）'
        assert ia_check.strategy_path(model(ssot), 'msg-0001', 'seg-0001')[1] != base, aid
    ssot = copy.deepcopy(SSOT)
    ssot['values']['val-0002']['text'] += '（改）'
    assert ia_check.strategy_path(model(ssot), 'msg-0001', 'seg-0001')[1] == base


def test_untraced_message_fails_when_shipped_and_warns_in_a_draft(tmp_path):
    ssot = copy.deepcopy(SSOT)
    del ssot['messages']['msg-0003']['serves']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    shipped = [f for f in report['findings'] if f['code'] == 'IA_STRATEGY_UNTRACED' and f['level'] == 'fail']
    drafted = [f for f in report['findings'] if f['code'] == 'IA_STRATEGY_UNTRACED' and f['level'] == 'warn']
    assert rc == 2 and shipped and drafted, report['findings']
    assert all('lp-municipal' in f['message'] for f in shipped) and all('deck-smb' in f['message'] for f in drafted)


def test_message_of_another_segment_fails_when_shipped(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['audience'] = 'aud-0002'  # seg-0002, but msg-0001 serves only seg-0001
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_STRATEGY_SEGMENT_MISMATCH' in codes(report), report['findings']


def test_pain_section_of_another_segment_fails_when_shipped(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['pages'][0]['sections'][0]['pain'] = 'pan-0002'
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_STRATEGY_SEGMENT_MISMATCH' in codes(report), report['findings']


def test_audience_without_segment_is_a_schema_failure(tmp_path):
    ssot = copy.deepcopy(SSOT)
    del ssot['audiences']['aud-0001']['segment']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_SCHEMA' in codes(report), report['findings']


@pytest.mark.parametrize('block,aid,fld,target', [
    ('jobs', 'job-0001', 'segment', 'seg-0099'), ('pains', 'pan-0001', 'job', 'job-0099'),
    ('values', 'val-0001', 'relieves', ['pan-0099']), ('strengths', 'str-0001', 'versus', ['alt-0099']),
    ('promises', 'prm-0001', 'values', ['val-0099']), ('messages', 'msg-0001', 'serves', ['val-0099'])])
def test_broken_strategy_reference_fails(tmp_path, block, aid, fld, target):
    ssot = copy.deepcopy(SSOT)
    ssot[block][aid][fld] = target
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_REF_BROKEN' in codes(report), report['findings']
