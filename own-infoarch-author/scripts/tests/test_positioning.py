"""Positioning (SPEC-information-architecture-standardization v6, requirements 2, 4, 5 and 8; Quality 1 and 3).

A positioning is every approved atom and arrow of one segment. Gaps and points of parity are derived from the
approved `fulfills` arrows, never recorded. Requirement 5 disqualifies an approved positioning; drafts are not
checked. The positioning id is a hash of the approved content only: the order things are written in the file and
where a copy came from never move it.
"""
import copy
import json
import random
import sys
from pathlib import Path

import pytest
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
import positioning  # noqa: E402
from test_check import SKILL, SSOT, check, cli, codes, dump, write_repo  # noqa: E402

TABLE = json.loads((SKILL / 'fixtures/positioning-red-cases.json').read_text(encoding='utf-8'))
CONDITIONS = {c['id']: c for c in TABLE['conditions']}
CASES = TABLE['cases']


def _key(edge):
    return f'{edge["kind"]}:{edge["from"]}:{edge["to"]}'


def apply(ops, ssot=SSOT):
    """Apply a red case's declarative operations to a copy of the base fixture. -> (ssot, keys left unapproved)."""
    ssot = copy.deepcopy(ssot)
    blocks = {p: b for b, p in positioning.ATOM_BLOCKS.items()} | {'prf': 'proofs'}
    unapproved = []
    for op in ops:
        kind = op['op']
        if kind == 'set_atom':
            ssot[blocks[op['id'].split('-')[0]]][op['id']] = op['atom']
        elif kind == 'set_field':
            ssot[blocks[op['id'].split('-')[0]]][op['id']][op['field']] = op['value']
        elif kind == 'set_next':
            ssot['ledger']['next'][op['prefix']] = op['value']
        elif kind == 'add_edge':
            ssot['edges'].append({'state': 'proposed', 'provenance': 'owner', **op['edge']})
        elif kind == 'remove_edge':
            before = len(ssot['edges'])
            ssot['edges'] = [e for e in ssot['edges'] if _key(e) != op['key']]
            assert len(ssot['edges']) == before - 1, op
        elif kind in ('set_field_edge', 'remove_field_edge'):
            [edge] = [e for e in ssot['edges'] if _key(e) == op['key']]
            if kind == 'set_field_edge':
                edge[op['field']] = op['value']
            else:
                del edge[op['field']]
        elif kind == 'set_block':
            ssot[op['block']] = op['value']
        elif kind == 'unapprove':
            unapproved.append(op['key'])
        else:
            raise KeyError(kind)
    return ssot, tuple(unapproved)


def write_case(root, case):
    ssot, unapproved = apply(case['ops'])
    write_repo(root, ssot, confirmed=False)
    from test_check import confirm
    confirm(root, ssot, skip=unapproved)


def model(ssot=SSOT, approve_all=True, skip=()):
    m = ia_check.Model(root=Path('.'), ssot=ssot)
    rows = {k: v for k, v in positioning.approval_targets(m).items() if k[len(positioning.APPROVAL_PREFIX):] not in skip}
    return m, (positioning.approved_keys(m, rows) if approve_all else set())


# ---- the table (Quality 1) ----------------------------------------------------------------------------------------

def test_every_condition_has_a_red_case_and_every_case_names_a_condition():
    covered = {c['condition'] for c in CASES}
    assert sorted(set(CONDITIONS) - covered) == []  # 対応の無い失格条件 0 件
    assert sorted(covered - set(CONDITIONS)) == []
    assert len({c['id'] for c in CASES}) == len(CASES)
    req5 = [c for c in CONDITIONS.values() if c['requirement'] == '5']
    assert len(req5) == 8  # spec v6 requirement 5 lists eight conditions; add a row when the spec adds one


def test_every_requirement_5_code_is_raised_by_the_checker():
    """A condition whose code the checker never raises would sit in the table forever green."""
    source = ''.join((SKILL / f'scripts/{n}').read_text(encoding='utf-8') for n in ('positioning.py', 'check.py'))
    for c in CONDITIONS.values():
        if c['code'].startswith('IA_POSITIONING_'):
            assert f"'{c['code']}'" in source, c


@pytest.mark.parametrize('case', CASES, ids=[c['id'] for c in CASES])
def test_red_case_fails_with_its_condition_code(case, tmp_path):
    write_case(tmp_path, case)
    rc, report = check(tmp_path)
    assert rc == 2, report['findings']
    assert CONDITIONS[case['condition']]['code'] in codes(report), codes(report)


def test_base_fixture_has_no_positioning_finding(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
    assert not [f for f in report['findings'] if f['code'].startswith('IA_POSITIONING_')], report['findings']


def test_unknown_operation_is_rejected():
    with pytest.raises(KeyError):
        apply([{'op': 'not-an-op'}])


# ---- only approved content is checked (requirement 5) ---------------------------------------------------------------

def test_drafts_are_not_disqualified(tmp_path):
    """A broken positioning nobody approved is a draft: requirement 5 does not look at it."""
    ssot, _ = apply([{'op': 'remove_edge', 'key': 'member_of:seg-0001:cat-0001'},
                     {'op': 'remove_edge', 'key': 'fulfills:alt-0002:pan-0001'}])
    write_repo(tmp_path, ssot, confirmed=False)
    rc, report = check(tmp_path)
    assert not [f for f in report['findings'] if f['code'].startswith('IA_POSITIONING_') and f['level'] == 'fail'
                and f['code'] != 'IA_POSITIONING_UNAPPROVED'], report['findings']


def test_unapproved_isolated_atom_is_ignored_and_leaves_the_id(tmp_path):
    case = next(c for c in CASES if c['id'] == 'isolated-alternative')
    ssot, _ = apply(case['ops'])
    m, keys = model(ssot, skip=('alt-0009',))
    base_m, base_keys = model()
    assert positioning.derive(m, keys)['seg-0001'].id == positioning.derive(base_m, base_keys)['seg-0001'].id
    report = ia_check.Report('check', '2026-11-16')
    positioning.check_positioning(m, report, keys)
    assert 'IA_POSITIONING_ISOLATED_ATOM' not in report.codes('fail')


def test_approval_follows_the_content_of_each_element():
    m, _ = model()
    targets = positioning.approval_targets(m)
    changed = copy.deepcopy(SSOT)
    changed['values']['val-0001']['text'] = '運用の手間がなくなる'
    after = positioning.approval_targets(ia_check.Model(root=Path('.'), ssot=changed))
    moved = sorted(k for k in targets if targets[k] != after.get(k))
    assert moved == ['ia:positioning:val-0001']  # one atom changed; its arrows and neighbours keep their approval
    assert positioning.approved_keys(ia_check.Model(root=Path('.'), ssot=changed), targets) == set(
        k[len(positioning.APPROVAL_PREFIX):] for k in targets) - {'val-0001'}


# ---- derivation (requirement 4) -----------------------------------------------------------------------------------

def test_gaps_and_points_of_parity_are_derived():
    m, keys = model()
    pos = positioning.derive(m, keys)
    assert sorted(pos) == ['seg-0001', 'seg-0002']
    one = pos['seg-0001']
    assert one.category == 'cat-0001'
    assert one.alternatives == ['alt-0001', 'alt-0002'] and one.category_alternatives == ['alt-0001']
    assert one.gaps == ['gan-0001', 'pan-0001'] and one.parity == ['gan-0002']
    two = pos['seg-0002']
    assert two.gaps == ['pan-0002'] and two.parity == []


def test_gap_compares_every_alternative_of_the_segment_not_only_the_category():
    """Q33: outsourcing (alt-0002) is not in the category, but it still decides whether pan-0001 is a gap."""
    ssot, _ = apply([{'op': 'set_field_edge', 'key': 'fulfills:alt-0002:pan-0001', 'field': 'degree', 'value': 'full'}])
    m, keys = model(ssot)
    one = positioning.derive(m, keys)['seg-0001']
    assert 'pan-0001' not in one.gaps and 'pan-0001' not in one.parity  # met by only some alternatives: neither


def test_missing_fulfillment_is_not_derived_as_none():
    ssot, _ = apply([{'op': 'remove_edge', 'key': 'fulfills:alt-0002:pan-0001'}])
    m, keys = model(ssot)
    one = positioning.derive(m, keys)['seg-0001']
    assert one.gaps is None and one.parity is None  # not derived; the checker disqualifies instead
    assert one.missing_fulfillment == [['alt-0002', 'pan-0001']]


def test_positioning_contains_only_its_segment_and_only_approved_elements():
    m, keys = model(skip=('val-0003', 'satisfies:val-0003:gan-0002'))
    one = positioning.derive(m, keys)['seg-0001']
    assert 'val-0003' not in one.atoms and 'satisfies:val-0003:gan-0002' not in one.edges
    assert not any(a.startswith(('job-0002', 'alt-0003', 'seg-0002')) for a in one.atoms)


# ---- ids (requirement 8, Quality 3) -------------------------------------------------------------------------------

def ids(ssot=SSOT, **kw):
    m, keys = model(ssot, **kw)
    return {seg: p.id for seg, p in positioning.derive(m, keys).items()}


def test_id_is_the_same_for_the_same_input():
    first = ids()
    assert first == ids() and all(len(v) == len('pos-') + 64 and v.startswith('pos-') for v in first.values())
    assert first['seg-0001'] != first['seg-0002']


def test_text_change_moves_only_that_segment():
    changed = copy.deepcopy(SSOT)
    changed['alternatives']['alt-0001']['text'] = '今のシステムを延命する'
    before, after = ids(), ids(changed)  # re-approved, as the owner would after reading the change
    assert after['seg-0001'] != before['seg-0001']
    assert after['seg-0002'] == before['seg-0002']  # alt-0003 is a copy, not a shared atom


def test_degree_change_moves_the_id_and_the_gaps():
    """Scenario 5: the owner corrects one fulfillment arrow."""
    ssot, _ = apply([{'op': 'set_field_edge', 'key': 'fulfills:alt-0001:gan-0002', 'field': 'degree', 'value': 'partial'}])
    m, keys = model(ssot)
    one = positioning.derive(m, keys)['seg-0001']
    assert one.parity == [] and ids(ssot)['seg-0001'] != ids()['seg-0001']


def test_file_order_does_not_move_the_id():
    shuffled = copy.deepcopy(SSOT)
    random.Random(7).shuffle(shuffled['edges'])
    for block in positioning.ATOM_BLOCKS:
        shuffled[block] = dict(reversed(list(shuffled[block].items())))
    shuffled['values']['val-0001'] = dict(reversed(list(shuffled['values']['val-0001'].items())))
    assert ids(shuffled) == ids()


def test_copy_provenance_does_not_move_the_id_or_the_approval():
    changed = copy.deepcopy(SSOT)
    del changed['alternatives']['alt-0003']['copied_from']
    m, _ = model()
    assert ids(changed) == ids()
    after = positioning.approval_targets(ia_check.Model(root=Path('.'), ssot=changed))
    assert after == positioning.approval_targets(m)


def test_provenance_does_not_move_the_id_or_the_approval():
    """Who wrote an element (provenance) is not what it says: re-labelling it must not move anything."""
    changed = copy.deepcopy(SSOT)
    changed['alternatives']['alt-0001']['provenance'] = 'extracted:docs/plan.md'
    changed['edges'][0]['provenance'] = 'proposal'
    m, _ = model()
    assert ids(changed) == ids()
    assert positioning.approval_targets(ia_check.Model(root=Path('.'), ssot=changed)) == positioning.approval_targets(m)


def test_supports_counts_only_when_its_target_arrow_is_in_the_positioning():
    case = next(c for c in CASES if c['id'] == 'evidence-supports-unpositioned-arrow')
    ssot, skip = apply(case['ops'])
    m, keys = model(ssot, skip=skip)
    one = positioning.derive(m, keys)['seg-0001']
    assert 'supports:prf-0003:fulfills:alt-0001:pan-0009' not in one.edges


def test_unapproved_draft_does_not_move_the_id():
    ssot, _ = apply([{'op': 'set_atom', 'id': 'gan-0009', 'atom': {'state': 'proposed', 'provenance': 'proposal',
                                                                   'segment': 'seg-0001', 'text': '下書きの得たいもの'}},
                     {'op': 'set_next', 'prefix': 'gan', 'value': 10}])
    assert ids(ssot, skip=('gan-0009',)) == ids()


def test_cli_ids_through_files_ignore_order_and_copy_provenance(tmp_path):
    """The real path: YAML files on disk, the owner's approvals file, the CLI."""
    write_repo(tmp_path / 'a')
    rc, a = cli('positioning', '--root', str(tmp_path / 'a'))
    assert rc == 0 and sorted(a['positionings']) == ['seg-0001', 'seg-0002'], a
    one = a['positionings']['seg-0001']
    assert one['gaps'] == ['gan-0001', 'pan-0001'] and one['parity'] == ['gan-0002'] and one['category'] == 'cat-0001'

    write_repo(tmp_path / 'b')
    path = tmp_path / 'b/design/info_arch/info_arch.yaml'
    data = yaml.safe_load(path.read_text(encoding='utf-8'))
    data['edges'] = list(reversed(data['edges']))
    data['alternatives']['alt-0003'].pop('copied_from')
    dump(path, data)  # approvals were recorded before the edit and still hold
    rc, b = cli('positioning', '--root', str(tmp_path / 'b'))
    assert rc == 0 and {k: v['id'] for k, v in b['positionings'].items()} == {k: v['id'] for k, v in a['positionings'].items()}

    data['values']['val-0002']['text'] = '少人数でも止まらずに回る'
    dump(path, data)
    rc, c = cli('positioning', '--root', str(tmp_path / 'b'))
    assert c['positionings']['seg-0001']['id'] == a['positionings']['seg-0001']['id']
    assert 'val-0002' in c['positionings']['seg-0002']['unapproved']  # the edit withdrew its approval
