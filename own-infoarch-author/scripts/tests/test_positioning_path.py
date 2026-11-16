"""The positioning an outline's messages stand on (v6 requirements 2 and 15; ISSUE-14 replaces the message side).

A message in an outline must trace, through the positioning atoms it serves, to the segment of the outline's
audience. The path is derived from arrows alone: the deliverable cannot leave out a strength that produces a value
on its path, or the arrows that tie them. Shipped deliverables fail; drafts warn.
"""
import copy
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
from test_check import DECK, OUTLINE, SSOT, check, codes, write_repo  # noqa: E402


def model(ssot=SSOT):
    return ia_check.Model(root=Path('.'), ssot=ssot)


def path_ids(ssot, mid, seg):
    return ia_check.positioning_path(model(ssot), mid, seg)[0]


@pytest.mark.parametrize('version', [2, 3])
def test_older_ssot_is_rejected_without_compatibility(tmp_path, version):
    old = copy.deepcopy(SSOT)
    old['schema_version'] = version
    write_repo(tmp_path, old, confirmed=False)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_SCHEMA' in codes(report), report['findings']


def test_path_follows_only_the_audience_segment_and_pulls_in_strengths_and_arrows():
    ids = path_ids(SSOT, 'msg-0001', 'seg-0001')
    assert ids == ['alt-0001', 'gan-0001', 'has_job:seg-0001:job-0001', 'job-0001',
                   'lacks:str-0001:alt-0001', 'need:job-0001:gan-0001', 'need:job-0001:pan-0001', 'pan-0001',
                   'produces:str-0001:val-0001', 'satisfies:val-0001:gan-0001', 'satisfies:val-0001:pan-0001',
                   'seg-0001', 'str-0001', 'val-0001']
    # msg-0002 serves values of both segments; each audience sees only its own branch
    assert path_ids(SSOT, 'msg-0002', 'seg-0002') == [
        'alt-0003', 'has_job:seg-0002:job-0002', 'job-0002', 'lacks:str-0002:alt-0003', 'need:job-0002:pan-0002',
        'pan-0002', 'produces:str-0002:val-0002', 'satisfies:val-0002:pan-0002', 'seg-0002', 'str-0002', 'val-0002']


def test_strength_for_an_off_path_value_is_not_on_the_path():
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'produces' and e['from'] == 'str-0001')]
    ssot['edges'].append({'kind': 'produces', 'from': 'str-0001', 'to': 'val-0003', 'state': 'proposed', 'provenance': 'owner'})
    assert 'str-0001' not in path_ids(ssot, 'msg-0001', 'seg-0001')


def test_path_hash_moves_with_any_path_element_and_ignores_others():
    base = ia_check.positioning_path(model(), 'msg-0001', 'seg-0001')[1]
    for block, aid, fld in (('strengths', 'str-0001', 'text'), ('alternatives', 'alt-0001', 'text'),
                            ('jobs', 'job-0001', 'text'), ('segments', 'seg-0001', 'situation')):
        ssot = copy.deepcopy(SSOT)
        ssot[block][aid][fld] += '（改）'
        assert ia_check.positioning_path(model(ssot), 'msg-0001', 'seg-0001')[1] != base, aid
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'lacks' and e['from'] == 'str-0001')]
    assert ia_check.positioning_path(model(ssot), 'msg-0001', 'seg-0001')[1] != base, 'arrow'
    ssot = copy.deepcopy(SSOT)
    ssot['strengths']['str-0001']['provenance'] = 'extracted:docs/plan.md'  # who wrote it is not what it says
    assert ia_check.positioning_path(model(ssot), 'msg-0001', 'seg-0001')[1] == base, 'provenance'
    for block, aid in (('values', 'val-0002'), ('values', 'val-0003'), ('alternatives', 'alt-0002')):
        ssot = copy.deepcopy(SSOT)
        ssot[block][aid]['text'] += '（改）'
        assert ia_check.positioning_path(model(ssot), 'msg-0001', 'seg-0001')[1] == base, aid


def test_untraced_message_fails_when_shipped_and_warns_in_a_draft(tmp_path):
    ssot = copy.deepcopy(SSOT)
    del ssot['messages']['msg-0003']['serves']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    shipped = [f for f in report['findings'] if f['code'] == 'IA_POSITIONING_UNTRACED' and f['level'] == 'fail']
    drafted = [f for f in report['findings'] if f['code'] == 'IA_POSITIONING_UNTRACED' and f['level'] == 'warn']
    assert rc == 2 and shipped and drafted, report['findings']
    assert all('lp-municipal' in f['message'] for f in shipped) and all('deck-smb' in f['message'] for f in drafted)


def test_message_of_another_segment_fails_when_shipped(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['audience'] = 'aud-0002'  # seg-0002, but msg-0001 serves only seg-0001
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_POSITIONING_SEGMENT_MISMATCH' in codes(report), report['findings']


def test_pain_section_of_another_segment_fails_when_shipped(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['pages'][0]['sections'][0]['pain'] = 'pan-0002'
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_POSITIONING_SEGMENT_MISMATCH' in codes(report), report['findings']


def test_audience_without_segment_is_a_schema_failure(tmp_path):
    ssot = copy.deepcopy(SSOT)
    del ssot['audiences']['aud-0001']['segment']
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_SCHEMA' in codes(report), report['findings']


@pytest.mark.parametrize('block,aid', [('jobs', 'job-0001'), ('values', 'val-0001'), ('categories', 'cat-0001')])
def test_positioning_atom_without_segment_is_a_schema_failure(tmp_path, block, aid):
    ssot = copy.deepcopy(SSOT)
    del ssot[block][aid]['segment']
    write_repo(tmp_path, ssot, confirmed=False)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_SCHEMA' in codes(report), report['findings']


@pytest.mark.parametrize('block,aid,fld,target', [
    ('jobs', 'job-0001', 'segment', 'seg-0099'), ('values', 'val-0001', 'segment', 'seg-0099'),
    ('proofs', 'prf-0003', 'segment', 'seg-0099'), ('messages', 'msg-0001', 'serves', ['val-0099'])])
def test_broken_positioning_reference_fails(tmp_path, block, aid, fld, target):
    ssot = copy.deepcopy(SSOT)
    ssot[block][aid][fld] = target
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_REF_BROKEN' in codes(report), report['findings']
