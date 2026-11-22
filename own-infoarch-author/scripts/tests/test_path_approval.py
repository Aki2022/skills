"""Owner approval per value path (SPEC-information-architecture-standardization requirements 3, 4, 8, 9, 15, 17
and 18; ADR-20261118-ia-approval-per-value-path, biz_ops).

The owner approves one value together with the path that supports it: the claim (the value, the strengths that
produce it and the alternatives they lack), the basis of the difference (what the value satisfies, every
alternative of the comparison set with its degree arrows) and the premises (segment, jobs, member category and the
categories the jobs recall), plus the evidence on any of those. The row key is `ia:path:<value id>`; its hash binds
the content of every atom and arrow on the path, extracted from all current data (drafts included).
"""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
import outline as ol  # noqa: E402
import positioning as pos  # noqa: E402
from test_check import OUTLINE, SSOT, check, cli, codes, confirm, dump, write_repo  # noqa: E402

VALUE_PATHS = ('val-0001', 'val-0002', 'val-0003')


def model(ssot=SSOT):
    return ia_check.Model(root=Path('.'), ssot=ssot)


def paths(ssot=SSOT):
    return pos.paths(model(ssot))


def hashes(ssot=SSOT):
    return {v: p.hash for v, p in paths(ssot).items()}


def edit(ssot=SSOT, **kw):
    out = copy.deepcopy(ssot)
    for fn in kw.values():
        fn(out)
    return out


def _edge(ssot, key):
    [e] = [e for e in ssot['edges'] if f'{e["kind"]}:{e["from"]}:{e["to"]}' == key]
    return e


def _add(kind, frm, to, **kw):
    return lambda s: s['edges'].append({'kind': kind, 'from': frm, 'to': to, 'state': 'proposed',
                                        'provenance': 'owner', **kw})


def _remove(key):
    def go(s):
        before = len(s['edges'])
        s['edges'] = [e for e in s['edges'] if f'{e["kind"]}:{e["from"]}:{e["to"]}' != key]
        assert len(s['edges']) == before - 1, key
    return go


def unapproved(report, level='fail'):
    return sorted({v for f in report['findings'] if f['code'] == 'IA_POSITIONING_UNAPPROVED' and f['level'] == level
                   for v in f['detail']['values']})


# ---- the path (requirement 3) ---------------------------------------------------------------------------------

def test_gap_path_holds_claim_basis_premises_and_evidence():
    p = paths()['val-0001']
    assert p.segment == 'seg-0001' and p.comparison == ['alt-0001', 'alt-0002'] and p.targets == ['gan-0001', 'pan-0001']
    roles = p.roles
    assert {k for k, r in roles.items() if r == 'claim'} == {
        'val-0001', 'str-0001', 'produces:str-0001:val-0001', 'lacks:str-0001:alt-0001'}
    assert {k for k, r in roles.items() if r == 'basis'} == {
        'satisfies:val-0001:pan-0001', 'satisfies:val-0001:gan-0001', 'pan-0001', 'gan-0001', 'alt-0001', 'alt-0002',
        'fulfills:alt-0001:pan-0001', 'fulfills:alt-0002:pan-0001', 'fulfills:alt-0001:gan-0001',
        'fulfills:alt-0002:gan-0001'}
    assert {k for k, r in roles.items() if r == 'premise'} == {
        'seg-0001', 'has_job:seg-0001:job-0001', 'job-0001', 'need:job-0001:pan-0001', 'need:job-0001:gan-0001',
        'member_of:seg-0001:cat-0001', 'cat-0001', 'recalls:job-0001:cat-0001'}
    assert {k for k, r in roles.items() if r == 'evidence'} == {
        'prf-0001', 'supports:prf-0001:val-0001', 'prf-0003', 'supports:prf-0003:str-0001',
        'supports:prf-0003:fulfills:alt-0001:pan-0001'}
    # A need the value does not satisfy, and its degrees, stay off the path.
    assert not {'gan-0002', 'need:job-0001:gan-0002', 'fulfills:alt-0001:gan-0002'} & set(roles)
    assert p.problems == [] and p.id == 'pth-' + p.hash


def test_parity_path_compares_only_the_member_category():
    p = paths()['val-0003']
    assert p.comparison == ['alt-0001'] and p.targets == ['gan-0002']
    assert 'includes:cat-0001:alt-0001' in p.roles and 'alt-0002' not in p.roles
    assert 'fulfills:alt-0002:gan-0002' not in p.roles
    assert p.problems == []


def test_hash_ignores_file_order_and_metadata():
    shuffled = edit(order=lambda s: s['edges'].reverse(),
                    meta=lambda s: s['jobs']['job-0001'].update(provenance='ai'))
    assert hashes(shuffled) == hashes()


@pytest.mark.parametrize('key', sorted(set(paths()['val-0001'].roles)))
def test_changing_any_member_drops_the_path_hash(key):
    """Acceptance 1: every atom and arrow on the path is bound (the red is seen once per member)."""
    def change(s):
        block = next((b for b, p in pos.ATOM_BLOCKS.items() if key.startswith(p + '-')), 'proofs')
        atom = s.get(block, {}).get(key) if ':' not in key else None
        if atom is not None:
            field = 'text' if 'text' in atom else 'qualifier'
            atom[field] += '（改）'
            return
        e = _edge(s, key)
        e['note'] = 'changed'
    assert hashes(edit(c=change))['val-0001'] != hashes()['val-0001']


# ---- shared and non-shared changes (Quality 3, requirement 15) ------------------------------------------------

def _set_and_ship(tmp_path, ssot):
    dump(tmp_path / 'design/info_arch/info_arch.yaml', ssot)
    rc, report = check(tmp_path)
    m = ia_check.load(tmp_path, ia_check.Report('t', '2026-11-18'))
    approved = pos.approved_keys(m, ia_check.read_approvals(tmp_path, ia_check.Report('t', '2026-11-18')))
    return ol.set_ids(m, 'lp-municipal', approved)['set'], unapproved(report)


def test_non_shared_change_moves_neither_the_set_id_nor_shipping(tmp_path):
    write_repo(tmp_path)
    base_set, base_block = _set_and_ship(tmp_path, SSOT)
    assert base_block == []
    other_segment = edit(t=lambda s: s['values']['val-0002'].update(text='少人数でも確実に回る'))
    draft_value = edit(a=lambda s: s['values'].update({'val-0009': {
        'state': 'proposed', 'provenance': 'proposal', 'segment': 'seg-0001', 'kind': 'difference', 'text': '下書き'}}),
                       n=lambda s: s['ledger']['next'].update(val=10))
    off_path_degree = edit(d=lambda s: _edge(s, 'fulfills:alt-0002:gan-0002').update(degree='partial'))
    for changed in (other_segment, draft_value, off_path_degree):
        assert _set_and_ship(tmp_path, changed) == (base_set, [])


def test_shared_change_moves_the_set_id_and_stops_shipping(tmp_path):
    write_repo(tmp_path)
    base_set, _ = _set_and_ship(tmp_path, SSOT)
    shared_job = edit(t=lambda s: s['jobs']['job-0001'].update(text='既存システムを安く運用する'))
    new_set, blocked = _set_and_ship(tmp_path, shared_job)
    assert new_set != base_set and blocked == ['val-0001', 'val-0003']  # the job is a premise of both paths
    one_path = edit(t=lambda s: s['strengths']['str-0001'].update(text='更新が自動'))
    new_set, blocked = _set_and_ship(tmp_path, one_path)
    assert new_set != base_set and blocked == ['val-0001']


# ---- additions and removals reach the hash (acceptance 3) -----------------------------------------------------

ALT9 = {'state': 'proposed', 'provenance': 'proposal', 'segment': 'seg-0001', 'text': '紙の台帳で管理する'}


def _with_alt9(s):
    s['alternatives']['alt-0009'] = dict(ALT9)
    s['ledger']['next']['alt'] = 10


@pytest.mark.parametrize('name, change, moved, kept', [
    ('add alternative', _with_alt9, {'val-0001'}, {'val-0002', 'val-0003'}),
    ('remove alternative', lambda s: [s['alternatives'].pop('alt-0002'),
                                      s.update(edges=[e for e in s['edges'] if 'alt-0002' not in (e['from'], e['to'])])],
     {'val-0001'}, {'val-0002', 'val-0003'}),
    ('add degree', lambda s: [_with_alt9(s), _add('fulfills', 'alt-0009', 'pan-0001', degree='none')(s)],
     {'val-0001'}, {'val-0003'}),
    ('remove degree', _remove('fulfills:alt-0002:pan-0001'), {'val-0001'}, {'val-0003'}),
    ('change degree', lambda s: _edge(s, 'fulfills:alt-0001:gan-0002').update(degree='partial'), {'val-0003'},
     {'val-0001'}),
    ('change member category', lambda s: [s['categories'].update({'cat-0009': {
        'state': 'proposed', 'provenance': 'proposal', 'segment': 'seg-0001', 'text': '自治体向けクラウド'}}),
        s['ledger']['next'].update(cat=10), _remove('member_of:seg-0001:cat-0001')(s),
        _add('member_of', 'seg-0001', 'cat-0009')(s)], {'val-0001', 'val-0003'}, {'val-0002'}),
    ('add recalled category', lambda s: [s['categories'].update({'cat-0009': {
        'state': 'proposed', 'provenance': 'proposal', 'segment': 'seg-0001', 'text': '保守の委託先'}}),
        s['ledger']['next'].update(cat=10), _add('recalls', 'job-0001', 'cat-0009')(s)],
     {'val-0001', 'val-0003'}, {'val-0002'}),
    ('remove recalled category', _remove('recalls:job-0001:cat-0001'), {'val-0001', 'val-0003'}, {'val-0002'}),
    ('include alternative in category', _add('includes', 'cat-0001', 'alt-0002'), {'val-0003'}, {'val-0001'}),
])
def test_additions_and_removals_move_the_paths_they_belong_to(name, change, moved, kept):
    before, after = hashes(), hashes(edit(c=change))
    assert {v for v in VALUE_PATHS if before[v] != after[v]} == moved, name
    assert {v for v in kept if before[v] == after[v]} == kept, name


def test_added_alternative_without_degree_disqualifies_the_gap_path_only():
    ps = paths(edit(a=_with_alt9))
    # alt-0009 has no arrow yet, so it is also an isolated atom of the path
    assert [c for c, _ in ps['val-0001'].problems] == ['IA_POSITIONING_ISOLATED_ATOM', 'IA_POSITIONING_FULFILLMENT_MISSING']
    assert ps['val-0001'].missing_fulfillment == [['alt-0009', 'gan-0001'], ['alt-0009', 'pan-0001']]
    assert ps['val-0003'].problems == []  # alt-0009 is outside the member category


def test_disqualified_path_stops_only_the_outlines_that_point_at_it(tmp_path):
    ssot = edit(a=_with_alt9)
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    fails = [f for f in report['findings'] if f['code'] == 'IA_POSITIONING_FULFILLMENT_MISSING']
    assert rc == 2 and [f['level'] for f in fails] == ['fail'], report['findings']  # lp-municipal ships val-0001
    outline = copy.deepcopy(OUTLINE)
    for msg in outline['messages'].values():
        msg['value'] = 'val-0003'
        msg.pop('compares', None)
    write_repo(tmp_path, ssot, outlines=(outline,))
    rc, report = check(tmp_path)
    fails = [f for f in report['findings'] if f['code'] == 'IA_POSITIONING_FULFILLMENT_MISSING']
    assert rc == 1 and [f['level'] for f in fails] == ['warn'], report['findings']


# ---- approval rows (requirements 3 and 18) --------------------------------------------------------------------

def test_targets_are_one_row_per_value(tmp_path):
    write_repo(tmp_path, confirmed=False)
    rc, out = cli('targets', '--root', str(tmp_path))
    assert rc == 0
    assert sorted(out['approval_hashes']) == [f'ia:path:{v}' for v in VALUE_PATHS]
    assert out['must_approval_keys'] == ['ia:path:val-0001', 'ia:path:val-0003']  # deck-smb does not ship
    assert out['approval_hashes']['ia:path:val-0001'] == hashes()['val-0001']


def test_unapproved_path_fails_shipped_and_warns_draft(tmp_path):
    write_repo(tmp_path, confirmed=False)
    rc, report = check(tmp_path)
    assert rc == 2 and unapproved(report) == ['val-0001', 'val-0003'] and unapproved(report, 'warn') == ['val-0002']


def test_approved_paths_pass(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0 and not unapproved(report) and not unapproved(report, 'warn'), report['findings']


def _ledger(tmp_path):
    return yaml.safe_load((tmp_path / 'design/approvals.yaml').read_text(encoding='utf-8'))


def test_legacy_per_element_rows_are_not_counted(tmp_path):
    write_repo(tmp_path, confirmed=False)
    m = model()
    legacy = [{'key': pos.LEGACY_PREFIX + k, 'value_hash': pos.element_hash(k, x), 'approver': 'owner',
               'date': '2026-11-18'} for k, x in sorted({**pos.atoms(m), **pos.edges(m)}.items())]
    dump(tmp_path / 'design/approvals.yaml', legacy)
    rc, report = check(tmp_path)
    assert rc == 2 and unapproved(report) == ['val-0001', 'val-0003'], report['findings']
    assert 'IA_APPROVALS_LEGACY' in codes(report, 'warn')


def test_an_approval_is_not_reused_for_another_path(tmp_path):
    write_repo(tmp_path, confirmed=False)
    h = hashes()
    rows = [{'key': 'ia:path:val-0001', 'value_hash': h['val-0003'], 'approver': 'owner', 'date': '2026-11-18'},
            {'key': 'ia:path:val-0003', 'value_hash': h['val-0001'], 'approver': 'owner', 'date': '2026-11-18'}]
    dump(tmp_path / 'design/approvals.yaml', rows)
    rc, report = check(tmp_path)
    assert rc == 2 and unapproved(report) == ['val-0001', 'val-0003'], report['findings']


def test_compared_alternative_must_be_in_the_comparison_set(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['messages']['msg-0002']['compares'] = ['alt-0002']  # val-0003 is parity: only cat-0001's alt-0001
    write_repo(tmp_path, outlines=(outline,))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_OUTLINE_COMPARE_OUTSIDE_SET' in codes(report), report['findings']
    outline['messages']['msg-0001']['compares'] = ['alt-0002']  # val-0001 is a gap: every alternative compares
    outline['messages']['msg-0002'].pop('compares')
    write_repo(tmp_path, outlines=(outline,))
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']


def test_unapproved_unreferenced_path_only_warns(tmp_path):
    ssot = edit(a=lambda s: s['values'].update({'val-0009': {
        'state': 'proposed', 'provenance': 'proposal', 'segment': 'seg-0001', 'kind': 'difference', 'text': '下書き'}}),
                n=lambda s: s['ledger']['next'].update(val=10), e=_add('satisfies', 'val-0009', 'pan-0001'))
    write_repo(tmp_path, ssot, confirmed=False)
    confirm(tmp_path, ssot, skip=('val-0009',))
    rc, report = check(tmp_path)
    assert rc == 1 and 'IA_POSITIONING_PATH_PENDING' in codes(report, 'warn'), report['findings']


# ---- ids and differences (requirements 8 and 9) ---------------------------------------------------------------

def test_set_cli_names_the_paths(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('set', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert rc == 0
    h = hashes()
    assert out['paths'] == {'val-0001': 'pth-' + h['val-0001'], 'val-0003': 'pth-' + h['val-0003']}
    assert out['set'] == out['outline'] + '+pth-' + h['val-0001'] + '+pth-' + h['val-0003']
    assert out['positioning'].startswith('pos-')


def test_diff_separates_positioning_edits_from_outline_selection(tmp_path):
    a, b = tmp_path / 'a', tmp_path / 'b'
    write_repo(a)
    changed = edit(t=lambda s: s['strengths']['str-0001'].update(text='更新が自動'))
    variant = copy.deepcopy(OUTLINE)
    variant['messages']['msg-0002']['value'] = 'val-0001'  # val-0003 is no longer pointed at
    write_repo(b, changed, outlines=(variant,))
    rc, out = cli('diff', '--root', str(a), '--root-b', str(b), '--a', 'lp-municipal', '--b', 'lp-municipal')
    edits = [c for c in out['changes'] if c['layer'] == 'positioning']
    assert {(c['path'], c['element'], c['change']) for c in edits} == {('val-0001', 'str-0001', 'changed')}
    selection = [c for c in out['changes'] if c.get('kind') == 'selection']
    assert selection == [{'layer': 'outline', 'element': 'paths/val-0003', 'change': 'removed', 'kind': 'selection'}]


# ---- one-time migration of the per-element rows (requirement 17) ----------------------------------------------

def _git(root, *args):
    subprocess.run(['git', '-C', str(root), *args], check=True, capture_output=True, text=True)


def _legacy_repo(tmp_path, skip=(), trailer=True):
    write_repo(tmp_path, confirmed=False)
    m = model()
    rows = [{'key': pos.LEGACY_PREFIX + k, 'value_hash': pos.element_hash(k, x), 'approver': 'owner',
             'date': '2026-11-18'} for k, x in sorted({**pos.atoms(m), **pos.edges(m)}.items()) if k not in skip]
    rows.append({'key': 'ds:web:token:color.primary', 'value_hash': 'a' * 64, 'approver': 'owner', 'date': '2026-11-10'})
    dump(tmp_path / 'design/approvals.yaml', rows)
    _git(tmp_path, 'init', '-q')
    _git(tmp_path, 'add', '-A')
    message = 'record approvals' + ('\n\nOwner-Approval: 「positioningは承認する」 (2026-11-18)\n'
                                    'Owner-Approval: 「代替の案は承認する」 (2026-11-18)' if trailer else '')
    _git(tmp_path, 'commit', '-q', '-m', message)  # the ambient identity, as test_ship does


def test_migration_replaces_per_element_rows_with_path_rows(tmp_path):
    _legacy_repo(tmp_path)
    rc, out = cli('migrate-approvals', '--root', str(tmp_path))
    assert rc == 0 and out['written'] is False
    assert [r['key'] for r in out['migrated']] == [f'ia:path:{v}' for v in VALUE_PATHS] and out['unmigrated'] == []
    assert out['legacy_rows'] == len({**pos.atoms(model()), **pos.edges(model())})
    assert out['commit_trailers'] == ['Owner-Approval-Migrated: 「positioningは承認する」 (2026-11-18)',
                                      'Owner-Approval-Migrated: 「代替の案は承認する」 (2026-11-18)']
    assert (tmp_path / 'design/approvals.yaml').read_text(encoding='utf-8').count(pos.LEGACY_PREFIX) > 0  # dry run
    rc, out = cli('migrate-approvals', '--root', str(tmp_path), '--write')
    assert rc == 0 and out['written'] is True
    keys = [r['key'] for r in _ledger(tmp_path)]
    assert sorted(keys) == ['ds:web:token:color.primary'] + [f'ia:path:{v}' for v in VALUE_PATHS]
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']


def test_migration_leaves_a_path_with_a_gap_unapproved(tmp_path):
    _legacy_repo(tmp_path, skip=('fulfills:alt-0002:gan-0001',))  # on val-0001's path only
    rc, out = cli('migrate-approvals', '--root', str(tmp_path), '--write')
    assert rc == 0
    assert [r['key'] for r in out['migrated']] == ['ia:path:val-0002', 'ia:path:val-0003']
    assert out['unmigrated'] == [{'value': 'val-0001', 'reason': 'not-every-member-approved',
                                  'missing': ['fulfills:alt-0002:gan-0001']}]
    rc, report = check(tmp_path)
    assert rc == 2 and unapproved(report) == ['val-0001']


def test_migration_needs_the_owners_words(tmp_path):
    _legacy_repo(tmp_path, trailer=False)
    rc, out = cli('migrate-approvals', '--root', str(tmp_path), '--write')
    assert rc == 2 and out['written'] is False and out['migrated'] == []
    assert {u['reason'] for u in out['unmigrated']} == {'no-owner-words'}
