"""The positioning an outline's messages stand on (v6 requirements 2, 6 and 15).

A message in an outline points at one value of the positioning its target's segment stands on. The path is derived from arrows alone: the deliverable cannot leave out a strength that produces a value
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
    ids = path_ids(SSOT, ['val-0001'], 'seg-0001')
    assert ids == ['alt-0001', 'alt-0002', 'cat-0001',
                   'fulfills:alt-0001:gan-0001', 'fulfills:alt-0001:pan-0001',  # the comparison set and its degrees
                   'fulfills:alt-0002:gan-0001', 'fulfills:alt-0002:pan-0001',
                   'gan-0001', 'has_job:seg-0001:job-0001', 'job-0001',
                   'lacks:str-0001:alt-0001', 'member_of:seg-0001:cat-0001',
                   'need:job-0001:gan-0001', 'need:job-0001:pan-0001', 'pan-0001',
                   'prf-0001', 'prf-0003',  # evidence a message inherits, with its supports arrows
                   'produces:str-0001:val-0001', 'recalls:job-0001:cat-0001',
                   'satisfies:val-0001:gan-0001', 'satisfies:val-0001:pan-0001',
                   'seg-0001', 'str-0001', 'supports:prf-0001:val-0001',
                   'supports:prf-0003:fulfills:alt-0001:pan-0001', 'supports:prf-0003:str-0001', 'val-0001']
    # each segment's value sees only its own branch
    assert path_ids(SSOT, ['val-0002'], 'seg-0002') == [
        'alt-0003', 'cat-0002', 'fulfills:alt-0003:pan-0002', 'has_job:seg-0002:job-0002', 'job-0002',
        'lacks:str-0002:alt-0003', 'member_of:seg-0002:cat-0002', 'need:job-0002:pan-0002', 'pan-0002', 'prf-0002',
        'produces:str-0002:val-0002', 'recalls:job-0002:cat-0002', 'satisfies:val-0002:pan-0002', 'seg-0002',
        'str-0002', 'supports:prf-0002:val-0002', 'val-0002']


def test_strength_for_an_off_path_value_is_not_on_the_path():
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'produces' and e['from'] == 'str-0001')]
    ssot['edges'].append({'kind': 'produces', 'from': 'str-0001', 'to': 'val-0003', 'state': 'proposed', 'provenance': 'owner'})
    assert 'str-0001' not in path_ids(ssot, ['val-0001'], 'seg-0001')


def test_path_hash_moves_with_any_path_element_and_ignores_others():
    base = ia_check.positioning_path(model(), ['val-0001'], 'seg-0001')[1]
    for block, aid, fld in (('strengths', 'str-0001', 'text'), ('alternatives', 'alt-0001', 'text'),
                            ('jobs', 'job-0001', 'text'), ('segments', 'seg-0001', 'situation')):
        ssot = copy.deepcopy(SSOT)
        ssot[block][aid][fld] += '（改）'
        assert ia_check.positioning_path(model(ssot), ['val-0001'], 'seg-0001')[1] != base, aid
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'lacks' and e['from'] == 'str-0001')]
    assert ia_check.positioning_path(model(ssot), ['val-0001'], 'seg-0001')[1] != base, 'arrow'
    ssot = copy.deepcopy(SSOT)
    ssot['strengths']['str-0001']['provenance'] = 'extracted:docs/plan.md'  # who wrote it is not what it says
    assert ia_check.positioning_path(model(ssot), ['val-0001'], 'seg-0001')[1] == base, 'provenance'
    # alt-0002 is in val-0001's comparison set (a gap compares every alternative), so it is not in this list
    for block, aid in (('values', 'val-0002'), ('values', 'val-0003'), ('gains', 'gan-0002')):
        ssot = copy.deepcopy(SSOT)
        ssot[block][aid]['text'] += '（改）'
        assert ia_check.positioning_path(model(ssot), ['val-0001'], 'seg-0001')[1] == base, aid


def test_message_of_another_segment_fails(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['target']['segment'] = 'seg-0002'  # its messages point at seg-0001's values
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_OUTLINE_VALUE_OTHER_POSITIONING' in codes(report), report['findings']


def test_target_without_segment_is_a_schema_failure(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    del outline['target']['segment']
    write_repo(tmp_path, outlines=(outline, DECK))
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
    ('proofs', 'prf-0003', 'segment', 'seg-0099')])
def test_broken_positioning_reference_fails(tmp_path, block, aid, fld, target):
    ssot = copy.deepcopy(SSOT)
    ssot[block][aid][fld] = target
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_REF_BROKEN' in codes(report), report['findings']
