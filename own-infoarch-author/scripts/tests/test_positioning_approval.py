"""Owner approval of the positioning, one atom and one arrow at a time (v6 requirements 3 and 15).

Approval is a hash-bound row in design/approvals.yaml (key ia:positioning:<atom id> or
ia:positioning:<kind>:<from>:<to>), recorded by the owner. The hash binds that element's own content only, so a
change withdraws exactly the approval of what changed. Only elements on a shipped path must be approved to ship;
the positioning (requirement 4) is built from approved elements alone. ISSUE-13 adds per-element prepare, the
contract version and the ship report on top of this.
"""
import copy
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
from test_check import SSOT, check, cli, confirm, dump, write_repo  # noqa: E402


def unapproved(report, level='fail'):
    return sorted({a for f in report['findings'] if f['code'] == 'IA_POSITIONING_UNAPPROVED' and f['level'] == level
                   for a in f['detail']['atoms']})


def test_unapproved_positioning_fails_shipped_and_warns_draft(tmp_path):
    write_repo(tmp_path, confirmed=False)
    rc, report = check(tmp_path)
    assert rc == 2
    assert {'val-0001', 'seg-0001', 'produces:str-0001:val-0001'} <= set(unapproved(report))
    assert 'val-0002' in unapproved(report, 'warn')  # deck-smb is a draft


def test_fully_approved_paths_pass(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0 and not unapproved(report) and not unapproved(report, 'warn'), report['findings']


def test_one_missing_arrow_is_named_alone(tmp_path):
    write_repo(tmp_path, confirmed=False)
    confirm(tmp_path, SSOT, skip=('lacks:str-0001:alt-0001',))
    rc, report = check(tmp_path)
    assert rc == 2 and unapproved(report) == ['lacks:str-0001:alt-0001'], report['findings']


def test_content_change_withdraws_only_that_approval(tmp_path):
    write_repo(tmp_path)
    changed = copy.deepcopy(SSOT)
    changed['jobs']['job-0001']['text'] = '既存システムを安く運用する'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, report = check(tmp_path)
    # v5 withdrew everything standing on the job; v6 withdraws the job alone
    assert rc == 2 and unapproved(report) == ['job-0001'], report['findings']


def test_arrow_attribute_change_withdraws_that_arrow(tmp_path):
    write_repo(tmp_path)
    changed = copy.deepcopy(SSOT)
    [edge] = [e for e in changed['edges'] if e['kind'] == 'fulfills' and e['from'] == 'alt-0002' and e['to'] == 'pan-0001']
    edge['degree'] = 'none'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_POSITIONING_FULFILLMENT_MISSING' in [f['code'] for f in report['findings']], report['findings']


def test_unapproved_element_off_the_shipped_path_does_not_stop(tmp_path):
    write_repo(tmp_path, confirmed=False)
    confirm(tmp_path, SSOT, skip=('str-0002',))
    rc, report = check(tmp_path)
    assert 'str-0002' not in unapproved(report)  # only the draft deck stands on str-0002
    assert rc == 1 and 'str-0002' in unapproved(report, 'warn'), report['findings']


def test_targets_cli_lists_every_element_and_requires_the_shipped_paths(tmp_path):
    write_repo(tmp_path, confirmed=False)
    rc, out = cli('targets', '--root', str(tmp_path))
    assert rc == 0
    keys = set(out['approval_hashes'])
    assert {'ia:positioning:cat-0001', 'ia:positioning:val-0002', 'ia:positioning:member_of:seg-0001:cat-0001',
            'ia:positioning:supports:prf-0003:fulfills:alt-0001:pan-0001'} <= keys
    assert all(k.startswith('ia:positioning:') and len(v) == 64 for k, v in out['approval_hashes'].items())
    required = set(out['must_approval_keys'])
    assert {'ia:positioning:val-0001', 'ia:positioning:str-0001', 'ia:positioning:lacks:str-0001:alt-0001'} <= required
    assert 'ia:positioning:val-0002' not in required and 'ia:positioning:cat-0001' not in required  # draft-only / no message


def test_targets_keys_fit_the_design_approval_key_shape(tmp_path):
    """own-design-author's resolve_ia accepts ia:<ns>:<[A-Za-z0-9_.:/-]+>; ia_module's own pattern must accept them too."""
    import re
    import ia_module
    write_repo(tmp_path, confirmed=False)
    _, out = cli('targets', '--root', str(tmp_path))
    for key in out['approval_hashes']:
        assert re.fullmatch(r'ia:[a-z][a-z0-9-]*:[A-Za-z0-9_.:/-]+', key) and ia_module.KEY.fullmatch(key), key
