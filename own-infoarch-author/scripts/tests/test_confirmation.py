"""Owner confirmation of the strategy layer (SPEC-information-architecture-standardization v5, requirements 14, 15, 19).

Confirmation is a hash-bound row in design/approvals.yaml (key ia:strategy:<id>), recorded by the owner.
The hash binds the atom's content and the content of everything it depends on, so a change to the atom or
to anything below it withdraws the confirmation. Only atoms on a shipped path must be confirmed.
"""
import copy
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
from test_check import SSOT, check, cli, codes, confirm, dump, write_repo  # noqa: E402


def unconfirmed(report, level='fail'):
    return sorted({a for f in report['findings'] if f['code'] == 'IA_STRATEGY_UNCONFIRMED' and f['level'] == level
                   for a in f['detail']['atoms']})


def test_unconfirmed_strategy_fails_shipped_and_warns_draft(tmp_path):  # acceptance (b)
    write_repo(tmp_path, confirmed=False)
    rc, report = check(tmp_path)
    assert rc == 2
    assert 'val-0001' in unconfirmed(report) and 'seg-0001' in unconfirmed(report)
    assert 'val-0002' in unconfirmed(report, 'warn')  # deck-smb is a draft


def test_fully_confirmed_paths_pass(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0 and not unconfirmed(report) and not unconfirmed(report, 'warn'), report['findings']


def test_mixed_confirmation_names_only_the_missing_atom(tmp_path):  # acceptance (c) and (i)
    write_repo(tmp_path, confirmed=False)
    confirm(tmp_path, SSOT, skip=('str-0001',))
    rc, report = check(tmp_path)
    assert rc == 2 and unconfirmed(report) == ['str-0001'], report['findings']


def test_content_change_withdraws_confirmation(tmp_path):  # acceptance (e)
    write_repo(tmp_path)
    changed = copy.deepcopy(SSOT)
    changed['values']['val-0001']['text'] = '運用の手間がなくなる'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, report = check(tmp_path)
    assert rc == 2 and 'val-0001' in unconfirmed(report), report['findings']


def test_dependency_change_withdraws_confirmation_of_the_dependent(tmp_path):  # acceptance (f)
    write_repo(tmp_path)
    changed = copy.deepcopy(SSOT)
    changed['jobs']['job-0001']['text'] = '既存システムを安く運用する'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, report = check(tmp_path)
    # val-0001's own content did not change; its confirmation still falls because what it stands on changed
    assert rc == 2 and {'job-0001', 'val-0001', 'prm-0001'} <= set(unconfirmed(report)), report['findings']


def test_strength_change_withdraws_value_confirmation(tmp_path):  # acceptance (i)
    write_repo(tmp_path)
    changed = copy.deepcopy(SSOT)
    changed['strengths']['str-0001']['text'] = '別の強み'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, report = check(tmp_path)
    assert rc == 2 and {'str-0001', 'val-0001'} <= set(unconfirmed(report)), report['findings']


def test_unconfirmed_strength_of_an_off_path_value_does_not_stop(tmp_path):  # acceptance (i), the negative
    ssot = copy.deepcopy(SSOT)
    ssot['strengths']['str-0002'] = {'state': 'proposed', 'provenance': 'owner', 'text': '少人数向けの強み',
                                     'versus': ['alt-0001'], 'produces': ['val-0002']}
    ssot['ledger']['next']['str'] = 3
    write_repo(tmp_path, ssot, confirmed=False)
    confirm(tmp_path, ssot, skip=('str-0002', 'val-0002', 'prm-0002', 'job-0002', 'pan-0002', 'seg-0002'))
    rc, report = check(tmp_path)
    assert 'str-0002' not in unconfirmed(report)  # only the draft deck stands on val-0002
    assert rc == 1 and 'str-0002' in unconfirmed(report, 'warn'), report['findings']


def test_confirmation_hash_ignores_unrelated_atoms():
    m = ia_check.Model(root=Path('.'), ssot=SSOT)
    base = ia_check.strategy_targets(m)[0]['ia:strategy:val-0001']
    changed = copy.deepcopy(SSOT)
    changed['values']['val-0002']['text'] += '（改）'
    assert ia_check.strategy_targets(ia_check.Model(root=Path('.'), ssot=changed))[0]['ia:strategy:val-0001'] == base


def test_targets_cli_lists_every_strategy_atom_and_requires_the_shipped_paths(tmp_path):
    write_repo(tmp_path, confirmed=False)
    rc, out = cli('targets', '--root', str(tmp_path))
    assert rc == 0
    keys = set(out['approval_hashes'])
    assert {'ia:strategy:cpt-0001', 'ia:strategy:val-0002', 'ia:strategy:str-0001'} <= keys
    assert all(k.startswith('ia:strategy:') and len(v) == 64 for k, v in out['approval_hashes'].items())
    required = set(out['must_approval_keys'])
    assert 'ia:strategy:val-0001' in required and 'ia:strategy:str-0001' in required
    assert 'ia:strategy:val-0002' not in required and 'ia:strategy:cpt-0001' not in required  # draft-only / no message
