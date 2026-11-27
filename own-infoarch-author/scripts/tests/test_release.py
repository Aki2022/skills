"""Release approval (SPEC-deliverable-production-layers requirements 17 and 19): the owner's last gate binds the
script, the layout plan and the shipped files by hash. Changing one character of any of them after the approval
takes the approval off; changing a comp image does not (comps are not in git and never part of the release).
Every case runs the real CLI against a real git fixture (test_ship's `repo`: a shipped deck and a draft web page)."""
from pathlib import Path
import sys

import pytest
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_ship import TODAY, cli, codes, commit, repo, run_git, ship  # noqa: E402,F401  (repo is a fixture)
import layer_paths  # noqa: E402

from release_fixtures import KEY, OID, approve_release, expected_hash, plan_of, record, unapprove, write_plan  # noqa: E402


def approved_repo(root):
    """The fixture (already approved by test_ship's `repo`), changed back to the approved state and committed."""
    return root


def bare(root):
    unapprove(root)
    commit(root, 'no plan, no release approval')
    return root


def edit(path, old, new):
    text = Path(path).read_text(encoding='utf-8')
    assert old in text, (path, old)
    Path(path).write_text(text.replace(old, new, 1), encoding='utf-8')


# ---- the green shape ---------------------------------------------------------------------------------------------

def test_approved_release_passes_the_gate(repo):
    approved_repo(repo)
    code, out, _ = ship(repo)
    assert code == 0, out
    assert 'IA_RELEASE_UNAPPROVED' not in codes(out)


def test_release_command_reports_the_hash_that_the_approval_must_carry(repo):
    approved_repo(repo)
    code, out, _ = cli(repo, 'release', '--root', str(repo), '--date', TODAY)
    row = next(r for r in out['releases'] if r['script'] == OID)
    assert row['key'] == KEY and row['hash'] == expected_hash(repo) and row['state'] == 'approved'
    assert row['shipped'] is True and code == 0


# ---- the red cases: one character off the approved content takes the approval off --------------------------------

def test_no_release_approval_stops_a_shipped_script(repo):
    bare(repo)
    write_plan(repo)
    commit(repo, 'plan only')
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_RELEASE_UNAPPROVED' in codes(out), out


@pytest.mark.parametrize('what', ['script', 'plan', 'html'])
def test_one_character_changed_after_approval_takes_the_approval_off(repo, what):
    approved_repo(repo)
    if what == 'script':
        edit(layer_paths.path(repo, 'infoarch', 'scripts', f'{OID}.yaml'), '端的で親しみやすい調子', '端的で親しみやすい調子。')
    elif what == 'plan':
        edit(plan_of(repo), 'reviewed: true', 'reviewed: false')
    else:
        (repo / 'decks/smb.pptx').write_bytes(b'fixture deck.')
    commit(repo, f'change {what}')
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_RELEASE_UNAPPROVED' in codes(out), (what, out)
    state = cli(repo, 'release', '--root', str(repo), '--date', TODAY)[1]['releases']
    assert next(r for r in state if r['script'] == OID)['state'] == 'changed'


def test_changing_a_comp_image_does_not_take_the_approval_off(repo):
    approved_repo(repo)
    comp = layer_paths.path(repo, 'build', OID, 'comp')
    comp.mkdir(parents=True, exist_ok=True)
    (comp / 'top.png').write_bytes(b'first comp')
    (repo / '.gitignore').write_text('.shipment.json\n' + layer_paths.rel('build', OID, 'comp') + '/\n')
    commit(repo, 'comp is ignored')
    (comp / 'top.png').write_bytes(b'a different comp')  # never in git, so never in the release
    code, out, _ = ship(repo)
    assert code == 0, out
    # the gate reads the commit, where the comp is absent anyway; `release` reads the working tree, where it is not
    working = cli(repo, 'release', '--root', str(repo), '--date', TODAY)[1]['releases']
    assert next(r for r in working if r['script'] == OID)['state'] == 'approved'


def test_a_shipped_script_without_a_plan_cannot_be_approved(repo):
    bare(repo)
    record(repo, '0' * 64)  # an approval of something: there is no plan to bind
    commit(repo, 'approval without a plan')
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_RELEASE_UNAPPROVED' in codes(out), out
    state = cli(repo, 'release', '--root', str(repo), '--date', TODAY)[1]['releases']
    row = next(r for r in state if r['script'] == OID)
    assert row['state'] == 'blocked' and any('計画' in p for p in row['problems'])


def test_a_remaining_proposal_makes_no_approval_request(repo):
    """Requirement 17: while any proposal stays, the check does not build an approval request: `targets` carries no
    release row to approve, and the shipped script still stops at the gate. (No approval recorded yet.)"""
    bare(repo)
    write_plan(repo)
    path = layer_paths.path(repo, 'infoarch', 'scripts', f'{OID}.yaml')
    script = yaml.safe_load(path.read_text(encoding='utf-8'))
    script['copies']['cpy-0004'].pop('refs')  # the evidence copy loses its source: it stays only as a proposal
    script['copies']['cpy-0004']['proposal'] = True
    path.write_text(yaml.safe_dump(script, allow_unicode=True, sort_keys=False), encoding='utf-8')
    commit(repo, 'a proposal remains')
    code, targets, _ = cli(repo, 'targets', '--root', str(repo))
    assert code == 0 and KEY not in targets['approval_hashes']
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_SCRIPT_PROPOSAL_AT_SHIP' in codes(out)
    state = cli(repo, 'release', '--root', str(repo), '--date', TODAY)[1]['releases']
    assert next(r for r in state if r['script'] == OID)['state'] == 'blocked'


def test_a_proposal_added_after_the_approval_leaves_a_hash_no_file_has(repo):
    """The design layer refuses a ledger row that names no target, so the key stays; its hash is one no content has."""
    path = layer_paths.path(repo, 'infoarch', 'scripts', f'{OID}.yaml')
    script = yaml.safe_load(path.read_text(encoding='utf-8'))
    script['copies']['cpy-0004'].pop('refs')
    script['copies']['cpy-0004']['proposal'] = True
    path.write_text(yaml.safe_dump(script, allow_unicode=True, sort_keys=False), encoding='utf-8')
    commit(repo, 'a proposal after the approval')
    code, targets, _ = cli(repo, 'targets', '--root', str(repo))
    assert code == 0 and targets['approval_hashes'][KEY] != expected_hash(repo)
    assert next(r for r in cli(repo, 'release', '--root', str(repo), '--date', TODAY)[1]['releases'] if r['script'] == OID)['state'] == 'blocked'


def test_targets_carry_the_release_row_when_nothing_blocks_it(repo):
    bare(repo)
    write_plan(repo)
    commit(repo, 'plan')
    code, targets, _ = cli(repo, 'targets', '--root', str(repo))
    assert code == 0 and targets['approval_hashes'][KEY] == expected_hash(repo)


def test_a_draft_script_needs_no_release_approval(repo):
    code, out, _ = cli(repo, 'release', '--root', str(repo), '--date', TODAY)
    draft = next(r for r in out['releases'] if r['script'] == 'lp-municipal')
    assert draft['shipped'] is False and draft['state'] == 'not-shipped'
