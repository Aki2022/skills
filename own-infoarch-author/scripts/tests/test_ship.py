"""ISSUE-05: the IA ship gate runs on the shipped commit, the predicted merge tree and the ship_command entry.

No mock is ship evidence: every case runs the real CLI against a real git fixture. The real-push case needs
INFOARCH_TEST_VIBE_ROOT (an installed vibe-guard) and the rendered case INFOARCH_TEST_PLAYWRIGHT_DIR; when they are
not set the cases are reported as skipped, never as passed.
"""
import copy
import hashlib
import json
import os
from pathlib import Path
import shlex
import shutil
import subprocess
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_check import DECK, OUTLINE, SSOT, dump, write_repo  # noqa: E402

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
DESIGN_SCRIPTS = SKILL.parent / 'own-design-author/scripts'
sys.path.insert(0, str(DESIGN_SCRIPTS))
import judgment_ledger as jl  # noqa: E402

TODAY = '2026-11-11'
VIBE_ROOT = os.environ.get('INFOARCH_TEST_VIBE_ROOT')
PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')

# A shipped deck keeps the default cases browser-free: DOM checks only HTML, so these cases exercise check + judge.
SHIP_DECK = {'schema_version': 2, 'media': 'deck', 'outputs': ['decks/*.pptx'], 'ship': True,
             'ship_command': 'python3 deploy.py', 'merge_base': 'origin/main'}
DRAFT_WEB = {'schema_version': 2, 'media': 'web', 'outputs': ['site/**/*.html'], 'ship': False}
LP_DRAFT = {**copy.deepcopy(OUTLINE), 'channel': 'web'}
DEPLOY = ('import json,os\nfrom pathlib import Path\n'
          'Path(".shipment.json").write_text(json.dumps({k: os.environ[k] for k in ("INFOARCH_SHIP_COMMIT", "INFOARCH_SHIP_TREE")}))\n')


def run_git(root, *args, env=None, allowed=(0,)):
    proc = subprocess.run(['git', *args], cwd=root, env=env, capture_output=True, text=True)
    assert proc.returncode in allowed, proc.stdout + proc.stderr
    return proc


def cli(root, *args, env=None):
    proc = subprocess.run([str(SCRIPTS / 'infoarch-author'), *args], cwd=root, capture_output=True, text=True,
                          env=env or dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable))
    lines = proc.stdout.strip().splitlines()
    assert lines, proc.stderr
    return proc.returncode, json.loads(proc.stdout), proc


def ship(root, *args, env=None):
    return cli(root, 'ship', '--root', str(root), '--date', TODAY, *args, env=env)


def commit(root, message):
    run_git(root, 'add', '-A')
    run_git(root, 'commit', '-qm', message)
    return run_git(root, 'rev-parse', 'HEAD').stdout.strip()


def cover_ledger(root, score=0.0):
    """Record every current IA candidate in the shared ledger, as a Jev run would."""
    code, out, _ = cli(root, 'judge', '--root', str(root), '--date', TODAY, '--include-candidates')
    assert code in (0, 1), out
    ledger = {'schema_version': 3, 'policy': jl.load_policy(),
              'rows': [{'candidate': r, 'score': score} for r in out['judgment_candidates']],
              'holdout': {'source': 'unmeasured'}}
    (root / 'design/judgments.json').write_text(json.dumps(ledger, ensure_ascii=False), encoding='utf-8')


def fresh_ssot():
    ssot = copy.deepcopy(SSOT)
    ssot['proofs']['prf-0001']['as_of'] = '2026-10-01'
    ssot['facts']['fct-0001']['as_of'] = '2026-10-01'
    return ssot


@pytest.fixture
def repo(tmp_path):
    root = tmp_path / 'repo'
    root.mkdir()
    run_git(root, 'init', '-q', '-b', 'main')
    run_git(root, 'config', 'user.name', 'Fixture Owner')  # the email comes from the runner's git identity
    deck = {**copy.deepcopy(DECK), 'channel': 'deck'}
    write_repo(root, ssot=fresh_ssot(), channels=(SHIP_DECK, DRAFT_WEB), outlines=(LP_DRAFT, deck))
    (root / 'decks').mkdir()
    (root / 'decks/smb.pptx').write_bytes(b'fixture deck')
    (root / 'deploy.py').write_text(DEPLOY)
    (root / '.gitignore').write_text('.shipment.json\n')
    cover_ledger(root)
    commit(root, 'fixture base')
    run_git(root, 'update-ref', 'refs/remotes/origin/main', 'HEAD')
    return root


def expire_shipped_proof(root):
    ssot = fresh_ssot()
    ssot['proofs']['prf-0001']['as_of'] = '2024-01-01'  # DECK (shipped) references prf-0001
    dump(root / 'design/info_arch/info_arch.yaml', ssot)


def codes(out):
    return [f['code'] for stage in out.get('stages', []) for f in stage.get('findings', [])]


# ---- gate on the committed source ----------------------------------------------------------

def test_green_gate_binds_the_committed_sha_and_never_deploys(repo):
    code, out, _ = ship(repo)
    assert code == 0, out
    assert out['commit'] == run_git(repo, 'rev-parse', 'HEAD').stdout.strip()
    assert [s['stage'] for s in out['stages']] == ['tree', 'check', 'dom', 'judge']
    assert out['shipped_outlines'] == ['deck-smb'] and out['channels'] == {'deck': ['deck-smb']}
    assert out['stages'][0]['targets']['outputs'] == 1  # the deck (not HTML) is still checked to be in the commit
    assert out['deployment'] == 'not-requested' and not (repo / '.shipment.json').exists()
    judge = out['stages'][3]
    assert judge['uncovered'] == [] and judge['screened'] > 0


def test_expired_proof_referenced_by_a_shipped_outline_stops_the_gate(repo):
    expire_shipped_proof(repo)
    commit(repo, 'expired proof')
    code, out, _ = ship(repo)
    assert code == 2, out
    assert 'IA_EXPIRED_SHIPPED' in codes(out)
    assert [s['stage'] for s in out['stages']] == ['tree', 'check']  # stops before dom, judge and deploy


def test_dirty_tree_is_not_shipped(repo):
    (repo / 'decks/smb.pptx').write_bytes(b'changed but not committed')
    code, out, _ = ship(repo)
    assert code == 2 and 'clean committed source' in ' '.join(out['errors'])


def test_zero_ship_channels_is_not_a_pass(repo):
    dump(repo / 'design/info_arch/channels/deck.yaml', {**SHIP_DECK, 'ship': False})
    commit(repo, 'nothing ships')
    code, out, _ = ship(repo)
    assert code == 2 and 'no ship:true IA channel' in ' '.join(out['errors'])


def test_a_ship_channel_with_zero_outlines_is_not_a_pass(repo):
    """Another channel shipping something does not make an empty one green."""
    dump(repo / 'design/info_arch/channels/flyer.yaml', {**SHIP_DECK, 'media': 'flyer', 'outputs': ['flyers/*.pdf']})
    commit(repo, 'a shipping channel nobody uses')
    code, out, _ = ship(repo)
    assert code == 2 and any('flyer' in e and 'zero' in e for e in out['errors']), out


def test_a_shipped_output_must_be_a_committed_file(repo):
    """A non-HTML output is checked too, against the commit: an ignored copy on disk does not count."""
    run_git(repo, 'rm', '-q', '--cached', 'decks/smb.pptx')
    (repo / '.gitignore').write_text('.shipment.json\ndecks/\n')
    commit(repo, 'deck ignored, still on disk')
    assert (repo / 'decks/smb.pptx').is_file()
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_SHIP_OUTPUT_MISSING' in codes(out), out


def test_gate_reads_inputs_from_the_commit_not_from_ignored_files(repo):
    """An ignored ledger on disk would cover every judgment; the commit has none, so the shipped outline is uncovered."""
    run_git(repo, 'rm', '-q', '--cached', 'design/judgments.json')
    (repo / '.gitignore').write_text('.shipment.json\ndesign/judgments.json\n')
    commit(repo, 'ledger ignored, still on disk')
    assert (repo / 'design/judgments.json').is_file()
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_SHIP_UNCOVERED' in codes(out), out


@pytest.mark.parametrize('path', ['decks/smb.pptx', 'design/info_arch/info_arch.yaml'])
def test_symlinks_in_shipped_paths_are_refused(repo, tmp_path, path):
    target = repo / path
    outside = tmp_path / ('outside-' + target.name)
    outside.write_bytes(target.read_bytes())
    target.unlink()
    target.symlink_to(outside)
    commit(repo, 'shipped path is a symlink')
    code, out, _ = ship(repo)
    assert code == 2 and 'IA_SHIP_SYMLINK' in codes(out), out


def test_uncovered_judgment_of_a_shipped_outline_stops_but_a_draft_does_not(repo):
    changed = {**copy.deepcopy(DECK), 'channel': 'deck'}
    changed['pages'][0]['sections'][1]['copy'] = 'いますぐお問い合わせ {{fct-0002}}'
    dump(repo / 'design/info_arch/outlines/deck-smb.yaml', changed)
    commit(repo, 'shipped copy changed without a judgment')
    code, out, _ = ship(repo)
    assert code == 2, out
    assert 'IA_SHIP_UNCOVERED' in codes(out)
    cover_ledger(repo)
    draft = copy.deepcopy(LP_DRAFT)
    draft['pages'][0]['sections'][1]['copy'] = '運用の手間がかなり減ります'
    dump(repo / 'design/info_arch/outlines/lp-municipal.yaml', draft)
    commit(repo, 'draft copy changed without a judgment')
    code, out, _ = ship(repo)
    assert code == 1, out  # reported as uncovered, but a draft that does not ship does not stop the shipment
    assert 'IA_JUDGE_UNCOVERED' in codes(out) and 'IA_SHIP_UNCOVERED' not in codes(out)


def test_review_scores_go_to_a_human_without_stopping(repo, tmp_path):
    cover_ledger(repo, score=jl.load_policy()['theta'])
    sha = commit(repo, 'every judgment at theta')
    code, out, _ = ship(repo)
    assert code == 1 and 'IA_JUDGE_REVIEW' in codes(out), out
    # The dispatcher stops a push on any nonzero exit, so pre-push reports warnings but exits 0 (only 2 stops).
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(refs_file(tmp_path, sha)))
    assert code == 0 and out['exit_code'] == 0, out
    assert any('IA_JUDGE_REVIEW' in w for w in out['warnings']), out


# ---- ship_command is the formal entry ---------------------------------------------------------

def test_ship_command_runs_only_after_a_green_gate_on_the_same_commit(repo):
    code, out, _ = ship(repo, '--channel', 'deck')
    assert code == 0 and out['deployment'] == 'command-completed', out
    shipped = json.loads((repo / '.shipment.json').read_text())
    assert shipped['INFOARCH_SHIP_COMMIT'] == out['commit']
    assert shipped['INFOARCH_SHIP_TREE'] == run_git(repo, 'rev-parse', 'HEAD^{tree}').stdout.strip()
    expire_shipped_proof(repo)
    commit(repo, 'expired proof')
    code, out, _ = ship(repo, '--channel', 'deck')
    assert code == 2 and 'IA_EXPIRED_SHIPPED' in codes(out)
    assert json.loads((repo / '.shipment.json').read_text()) == shipped  # the red commit never reached deploy


def test_unknown_or_draft_channel_is_refused(repo):
    for name in ('nowhere', 'web'):
        code, out, _ = ship(repo, '--channel', name)
        assert code == 2 and 'unknown or nonshipping IA channel' in ' '.join(out['errors'])


def test_failed_ship_command_is_a_failure(repo):
    dump(repo / 'design/info_arch/channels/deck.yaml', {**SHIP_DECK, 'ship_command': 'python3 -c "raise SystemExit(3)"'})
    commit(repo, 'failing deploy')
    code, out, _ = ship(repo, '--channel', 'deck')
    assert code == 2 and 'ship_command failed' in ' '.join(out['errors'])


# ---- predicted merge (pre-push) ----------------------------------------------------------------

def refs_file(tmp_path, *shas):
    path = tmp_path / 'push-refs'
    path.write_text(''.join(f'refs/heads/topic {sha} refs/heads/topic {"0" * 40}\n' for sha in shas))
    return path


def test_pushing_an_expired_reference_stops_on_the_pushed_commit(repo, tmp_path):
    run_git(repo, 'checkout', '-qb', 'topic')
    expire_shipped_proof(repo)
    sha = commit(repo, 'expired proof on a topic')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(refs_file(tmp_path, sha)))
    assert code == 2, out
    assert 'pushed commit IA ship gate failed' in out['errors']
    check = out['checks'][0]
    assert check['predicted_merge'] is False and check['commit'] == check['source_commit'] == sha
    assert 'IA_EXPIRED_SHIPPED' in codes(check)
    assert len(run_git(repo, 'worktree', 'list', '--porcelain').stdout.split('worktree ')) == 2  # snapshot removed


def test_two_green_branches_fail_on_their_predicted_merge(repo, tmp_path):
    """The topic is green alone; main has since expired the proof the topic starts to ship."""
    base = run_git(repo, 'rev-parse', 'HEAD').stdout.strip()
    ssot = fresh_ssot()
    ssot['proofs']['prf-0003'] = {'state': 'proposed', 'provenance': 'owner', 'text': '満足度調査',
                                  'source': '調査票', 'supports': ['msg-0002'], 'as_of': '2026-10-01'}
    ssot['ledger']['next']['prf'] = 4
    dump(repo / 'design/info_arch/info_arch.yaml', ssot)
    cover_ledger(repo)
    base = commit(repo, 'second proof, used by nobody')
    run_git(repo, 'checkout', '-qb', 'topic')
    deck = {**copy.deepcopy(DECK), 'channel': 'deck'}
    deck['pages'][0]['sections'][0]['proofs'] = ['prf-0001', 'prf-0003']
    dump(repo / 'design/info_arch/outlines/deck-smb.yaml', deck)
    cover_ledger(repo)
    topic = commit(repo, 'ship the second proof')
    code, out, _ = ship(repo)
    assert code == 0, out  # green on its own commit
    run_git(repo, 'checkout', '-q', 'main')
    ssot['proofs']['prf-0003']['as_of'] = '2024-01-01'
    dump(repo / 'design/info_arch/info_arch.yaml', ssot)
    commit(repo, 'second proof is old (nobody ships it yet)')
    code, out, _ = ship(repo)
    assert code == 0, out  # green on main too: the expired proof is not shipped there
    run_git(repo, 'update-ref', 'refs/remotes/origin/main', 'HEAD')
    run_git(repo, 'checkout', '-q', 'topic')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(refs_file(tmp_path, topic)))
    assert code == 2 and 'predicted merge IA ship gate failed' in out['errors'], out
    own, merged = out['checks']
    assert not own['predicted_merge'] and own['exit_code'] == 0  # the topic commit alone is green
    assert merged['predicted_merge'] and 'IA_EXPIRED_SHIPPED' in codes(merged)
    assert run_git(repo, 'rev-parse', 'HEAD').stdout.strip() == topic and base


def test_pre_push_checks_the_pushed_commit_itself(repo, tmp_path):
    """An old red commit pushed to a new branch: its merge into the fixed main is green, the commit is not."""
    expire_shipped_proof(repo)
    bad = commit(repo, 'expired proof')
    dump(repo / 'design/info_arch/info_arch.yaml', fresh_ssot())
    commit(repo, 'refresh the proof date')
    run_git(repo, 'update-ref', 'refs/remotes/origin/main', 'HEAD')
    refs = tmp_path / 'old-bad'
    refs.write_text(f'refs/heads/old-bad {bad} refs/heads/old-bad {"0" * 40}\n')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(refs))
    assert code == 2, out
    own = [c for c in out['checks'] if not c['predicted_merge']]
    assert own and own[0]['commit'] == bad and 'IA_EXPIRED_SHIPPED' in codes(own[0]), out


def test_refs_pushed_together_are_merged_with_each_other(repo, tmp_path):
    """main and topic are pushed at once: topic is green against origin/main but red against the pushed main."""
    ssot = fresh_ssot()
    ssot['proofs']['prf-0003'] = {'state': 'proposed', 'provenance': 'owner', 'text': '満足度調査',
                                  'source': '調査票', 'supports': ['msg-0002'], 'as_of': '2026-10-01'}
    ssot['ledger']['next']['prf'] = 4
    dump(repo / 'design/info_arch/info_arch.yaml', ssot)
    cover_ledger(repo)
    old = commit(repo, 'second proof, used by nobody')
    run_git(repo, 'update-ref', 'refs/remotes/origin/main', 'HEAD')
    run_git(repo, 'checkout', '-qb', 'topic')
    deck = {**copy.deepcopy(DECK), 'channel': 'deck'}
    deck['pages'][0]['sections'][0]['proofs'] = ['prf-0001', 'prf-0003']
    dump(repo / 'design/info_arch/outlines/deck-smb.yaml', deck)
    cover_ledger(repo)
    topic = commit(repo, 'ship the second proof')
    run_git(repo, 'checkout', '-q', 'main')
    ssot['proofs']['prf-0003']['as_of'] = '2024-01-01'
    dump(repo / 'design/info_arch/info_arch.yaml', ssot)
    main = commit(repo, 'second proof is old (nobody ships it on main)')
    alone = tmp_path / 'alone'
    alone.write_text(f'refs/heads/topic {topic} refs/heads/topic {"0" * 40}\n')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(alone))
    assert code == 0, out  # origin/main has not moved yet
    together = tmp_path / 'together'
    together.write_text(f'refs/heads/main {main} refs/heads/main {old}\n'
                        f'refs/heads/topic {topic} refs/heads/topic {"0" * 40}\n')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(together))
    assert code == 2, out
    red = [c for c in out['checks'] if 'IA_EXPIRED_SHIPPED' in codes(c)]
    assert red and red[0]['predicted_merge'] and red[0]['source_commit'] == topic and red[0]['merge_base_commit'] == main, out


def test_green_push_passes_and_deleted_refs_only_skip(repo, tmp_path):
    sha = run_git(repo, 'rev-parse', 'HEAD').stdout.strip()
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(refs_file(tmp_path, sha)))
    assert code == 0 and len(out['checks']) == 1, out
    deleted = tmp_path / 'deleted'
    deleted.write_text(f'(delete) {"0" * 40} refs/heads/old {sha}\n')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(deleted))
    assert code == 0 and out['checks'] == [] and out['skipped'], out


def test_pre_push_without_a_merge_base_fails_closed(repo, tmp_path):
    row = dict(SHIP_DECK)
    row.pop('merge_base')
    dump(repo / 'design/info_arch/channels/deck.yaml', row)
    sha = commit(repo, 'no merge base')
    code, out, _ = ship(repo, '--pre-push', '--refs-file', str(refs_file(tmp_path, sha)))
    assert code == 2 and 'IA channel without merge_base' in ' '.join(out['errors']), out


def test_pre_push_without_refs_is_refused(repo, tmp_path):
    empty = tmp_path / 'empty'
    empty.write_text('')
    for extra in ((), ('--refs-file', str(empty))):
        code, out, _ = ship(repo, '--pre-push', *extra)
        assert code == 2, out


def test_merge_base_must_be_a_plain_ref(repo):
    dump(repo / 'design/info_arch/channels/deck.yaml', {**SHIP_DECK, 'merge_base': '--output=x'})
    code, out, _ = cli(repo, 'check', '--root', str(repo), '--date', TODAY)
    assert code == 2 and 'IA_SCHEMA' in ' '.join(f['code'] for f in out['findings']), out


# ---- pre-push extension and doctor -------------------------------------------------------------

EXTENSION = SKILL / 'hooks/pre-push.d/61-infoarch-ship'


def run_extension(root, *args, env=None):
    return subprocess.run([str(EXTENSION), *args], cwd=root, capture_output=True, text=True, env=env)


def test_extension_is_opt_in_and_fails_closed_once_enabled(repo, tmp_path):
    run_git(repo, 'checkout', '-qb', 'topic')  # vibe-guard blocks commits to main once .vibe-guard/ exists
    refs = refs_file(tmp_path, run_git(repo, 'rev-parse', 'HEAD').stdout.strip())
    assert run_extension(repo, str(refs)).returncode == 0  # no marker: this repository has not enabled the gate
    (repo / '.vibe-guard').mkdir()
    (repo / '.vibe-guard/infoarch-ship.enabled').write_text('yes\n')
    assert run_extension(repo, str(refs)).returncode == 2  # an invalid marker never reads as disabled
    shutil.copy2(SKILL / 'templates/infoarch-ship.enabled', repo / '.vibe-guard/infoarch-ship.enabled')
    assert run_extension(repo).returncode == 2  # refs file missing
    env = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable, INFOARCH_SHIP_COMMAND=str(SCRIPTS / 'infoarch-author'))
    assert run_extension(repo, str(refs), env=env).returncode == 0
    expire_shipped_proof(repo)
    red = refs_file(tmp_path, commit(repo, 'expired proof'))
    blocked = run_extension(repo, str(red), env=env)
    assert blocked.returncode == 2 and 'IA_EXPIRED_SHIPPED' in blocked.stdout


def enable_marker(root):
    (root / '.vibe-guard').mkdir(exist_ok=True)
    shutil.copy2(SKILL / 'templates/infoarch-ship.enabled', root / '.vibe-guard/infoarch-ship.enabled')


def gate_env():
    return dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable, INFOARCH_SHIP_COMMAND=str(SCRIPTS / 'infoarch-author'))


def test_deleting_the_worktree_marker_does_not_disable_a_committed_gate(repo, tmp_path):
    run_git(repo, 'checkout', '-qb', 'topic')
    enable_marker(repo)
    commit(repo, 'enable the IA ship gate')
    expire_shipped_proof(repo)
    red = commit(repo, 'expired proof')
    (repo / '.vibe-guard/infoarch-ship.enabled').unlink()  # local deletion only: the pushed commit still enables it
    blocked = run_extension(repo, str(refs_file(tmp_path, red)), env=gate_env())
    assert blocked.returncode == 2 and 'IA_EXPIRED_SHIPPED' in blocked.stdout, blocked.stdout + blocked.stderr
    # From a checkout whose HEAD and remote-tracking branches never had the marker: the pushed commit alone decides.
    run_git(repo, 'checkout', '-q', 'main')
    blocked = run_extension(repo, str(refs_file(tmp_path, red)), env=gate_env())
    assert blocked.returncode == 2 and 'IA_EXPIRED_SHIPPED' in blocked.stdout, blocked.stdout + blocked.stderr


def test_removing_the_marker_in_a_commit_is_reported_and_still_gated(repo, tmp_path):
    run_git(repo, 'checkout', '-qb', 'topic')
    enable_marker(repo)
    enabled = commit(repo, 'enable the IA ship gate')
    run_git(repo, 'rm', '-q', '.vibe-guard/infoarch-ship.enabled')
    expire_shipped_proof(repo)
    red = commit(repo, 'disable the gate and expire a shipped proof')
    refs = tmp_path / 'refs'
    # Only the commit this push replaces on the remote has the marker (origin/main never had it).
    refs.write_text(f'refs/heads/topic {red} refs/heads/topic {enabled}\n')
    blocked = run_extension(repo, str(refs), env=gate_env())
    assert blocked.returncode == 2 and 'IA_EXPIRED_SHIPPED' in blocked.stdout, blocked.stdout + blocked.stderr
    assert 'infoarch-ship.enabled' in blocked.stderr  # the removal is visible, not silent
    dump(repo / 'design/info_arch/info_arch.yaml', fresh_ssot())
    green = commit(repo, 'refresh the proof date')
    refs.write_text(f'refs/heads/topic {green} refs/heads/topic {enabled}\n')
    passed = run_extension(repo, str(refs), env=gate_env())
    assert passed.returncode == 0 and 'infoarch-ship.enabled' in passed.stderr, passed.stdout + passed.stderr


def test_a_new_branch_without_the_marker_is_gated_while_the_merge_base_has_it(repo, tmp_path):
    run_git(repo, 'checkout', '-qb', 'topic')
    enable_marker(repo)
    commit(repo, 'enable the IA ship gate')
    run_git(repo, 'update-ref', 'refs/remotes/origin/main', 'HEAD')  # the merge base now commits the marker
    run_git(repo, 'checkout', '-qb', 'escape')
    run_git(repo, 'rm', '-q', '.vibe-guard/infoarch-ship.enabled')
    expire_shipped_proof(repo)
    red = commit(repo, 'disable the gate on a new branch')
    blocked = run_extension(repo, str(refs_file(tmp_path, red)), env=gate_env())  # remote sha is all zeros
    assert blocked.returncode == 2 and 'IA_EXPIRED_SHIPPED' in blocked.stdout, blocked.stdout + blocked.stderr
    assert 'infoarch-ship.enabled' in blocked.stderr


def hooks_harness(tmp_path, dispatcher):
    """A vibe-guard-shaped root: ROOT/bin/vibe-guard (the scanner) and ROOT/git-hooks/{pre-push,pre-push.d/}."""
    guard = tmp_path / 'guard'
    hooks = guard / 'git-hooks'
    (hooks / 'pre-push.d').mkdir(parents=True)
    (guard / 'bin').mkdir()
    (guard / 'bin/vibe-guard').write_text('#!/bin/sh\ncat >/dev/null\n')
    (guard / 'bin/vibe-guard').chmod(0o755)
    (hooks / 'pre-push').write_text(dispatcher.replace('@ROOT@', shlex.quote(str(guard))))
    (hooks / 'pre-push').chmod(0o755)
    config = tmp_path / 'gitconfig'
    config.write_text('[core]\n hooksPath = ' + str(hooks) + '\n')
    return hooks, dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable, GIT_CONFIG_GLOBAL=str(config))


# The shape of vibe-guard's dispatcher: errexit makes any extension's nonzero exit stop the push.
DISPATCHER = ('#!/usr/bin/env bash\nset -euo pipefail\nROOT=@ROOT@\nrefs_file="$(mktemp)"\n'
              'trap \'rm -f "$refs_file"\' EXIT\ntee "$refs_file" | "$ROOT/bin/vibe-guard" scan-push "$@"\n'
              'd="$ROOT/git-hooks/pre-push.d"\nfor ext in "$d"/*; do\n  [[ -x "$ext" ]] || continue\n'
              '  "$ext" "$refs_file"\ndone\n')


def test_doctor_disqualifies_missing_hook_marker_and_ship_command(repo, tmp_path):
    hooks, env = hooks_harness(tmp_path, DISPATCHER)
    code, out, _ = ship(repo, '--doctor', env=env)
    assert code == 2 and any('marker' in e for e in out['errors']) and any('extension' in e for e in out['errors']), out
    shutil.copy2(EXTENSION, hooks / 'pre-push.d/61-infoarch-ship')
    (repo / '.vibe-guard').mkdir()
    shutil.copy2(SKILL / 'templates/infoarch-ship.enabled', repo / '.vibe-guard/infoarch-ship.enabled')
    code, out, _ = ship(repo, '--doctor', env=env)
    assert code == 0, out
    assert out['channels'] == ['deck'] and out['scope'] == 'configuration-diagnostic'
    # A ship:true channel with no ship_command (schema-invalid, read raw by the doctor) and no merge_base.
    dump(repo / 'design/info_arch/channels/deck.yaml', {k: v for k, v in SHIP_DECK.items() if k not in ('ship_command', 'merge_base')})
    code, out, _ = ship(repo, '--doctor', env=env)
    assert code == 2
    assert any('deck' in e and 'ship_command' in e for e in out['errors'])
    assert any('deck' in e and 'merge_base' in e for e in out['errors'])
    (hooks / 'pre-push.d/61-infoarch-ship').write_text('#!/bin/sh\nexit 0\n')
    dump(repo / 'design/info_arch/channels/deck.yaml', SHIP_DECK)
    code, out, _ = ship(repo, '--doctor', env=env)
    assert code == 2 and any('extension' in e for e in out['errors'])  # a stale or edited copy is not the gate


def test_doctor_rejects_a_dispatcher_that_drops_the_scanner(repo, tmp_path):
    hooks, env = hooks_harness(tmp_path, '#!/bin/sh\nfor ext in pre-push.d/*; do "$ext"; done\n')
    shutil.copy2(EXTENSION, hooks / 'pre-push.d/61-infoarch-ship')
    (repo / '.vibe-guard').mkdir()
    shutil.copy2(SKILL / 'templates/infoarch-ship.enabled', repo / '.vibe-guard/infoarch-ship.enabled')
    code, out, _ = ship(repo, '--doctor', env=env)
    assert code == 2 and any('dispatcher' in e for e in out['errors']), out


@pytest.mark.parametrize('swallow', [
    DISPATCHER.replace('"$ext" "$refs_file"\n', '"$ext" "$refs_file" || true\n'),
    DISPATCHER.replace('set -euo pipefail\n', ''),  # without errexit only the last extension's exit survives
])
def test_doctor_rejects_a_dispatcher_that_does_not_propagate_the_gate_exit(repo, tmp_path, swallow):
    assert swallow != DISPATCHER
    hooks, env = hooks_harness(tmp_path, swallow)
    shutil.copy2(EXTENSION, hooks / 'pre-push.d/61-infoarch-ship')
    enable_marker(repo)
    code, out, _ = ship(repo, '--doctor', env=env)
    assert code == 2 and any('exit' in e for e in out['errors']), out


# ---- real push through the installed vibe-guard -------------------------------------------------

def copy_guard_hooks(guard, harness):
    target = harness / 'git-hooks'
    shutil.copytree(guard / 'git-hooks', target, symlinks=False)
    for path in (target, *target.rglob('*')):
        path.chmod(path.stat().st_mode | 0o200)
    return target


@pytest.mark.skipif(not VIBE_ROOT, reason='INFOARCH_TEST_VIBE_ROOT is not set: real git push through vibe-guard not run (reported as skipped, not passed)')
def test_real_push_stops_on_the_predicted_merge_and_scanner_still_runs(repo, tmp_path):
    guard = Path(VIBE_ROOT).resolve()
    global_config = Path(os.environ.get('GIT_CONFIG_GLOBAL', str(Path.home() / '.gitconfig')))
    before = hashlib.sha256(global_config.read_bytes()).hexdigest() if global_config.exists() else None
    harness = tmp_path / 'guard'
    (harness / 'bin').mkdir(parents=True)
    hooks = copy_guard_hooks(guard, harness)
    hook = hooks / 'pre-push'
    lines = hook.read_text().splitlines()
    assert any(line.startswith('ROOT=') for line in lines)
    hook.write_text('\n'.join('ROOT=' + shlex.quote(str(harness)) if line.startswith('ROOT=') else line for line in lines) + '\n')
    scanner = tmp_path / 'scanner-called'
    wrapper = harness / 'bin/vibe-guard'
    wrapper.write_text('#!/bin/sh\nset -eu\nprintf called >> ' + shlex.quote(str(scanner)) +
                       '\nexec ' + shlex.quote(str(guard / 'bin/vibe-guard')) + ' "$@"\n')
    wrapper.chmod(0o755)
    for other in (hooks / 'pre-push.d').iterdir():
        other.unlink()  # only the scanner and this gate: other repositories' extensions are not under test
    shutil.copy2(EXTENSION, hooks / 'pre-push.d/61-infoarch-ship')
    config = tmp_path / 'gitconfig'
    include = '[include]\n path = ' + str(global_config) + '\n' if global_config.exists() else ''
    config.write_text(include + '[core]\n hooksPath = ' + str(hooks) + '\n')
    env = dict(os.environ, GIT_CONFIG_GLOBAL=str(config), VIBE_GUARD_ROOT=str(guard), DESIGN_AUTHOR_PYTHON=sys.executable,
               INFOARCH_SHIP_COMMAND=str(SCRIPTS / 'infoarch-author'))
    run_git(repo, 'checkout', '-qb', 'topic')  # vibe-guard blocks commits to main once .vibe-guard/ exists
    (repo / '.vibe-guard').mkdir()
    shutil.copy2(SKILL / 'templates/infoarch-ship.enabled', repo / '.vibe-guard/infoarch-ship.enabled')
    commit(repo, 'enable the IA ship gate')
    run_git(repo, 'update-ref', 'refs/remotes/origin/main', 'HEAD')
    remote = tmp_path / 'remote.git'
    subprocess.run(['git', 'init', '--bare', '-q', str(remote)], check=True)
    run_git(repo, 'remote', 'add', 'fixture', str(remote))
    expire_shipped_proof(repo)
    commit(repo, 'expired proof on a topic')
    blocked = run_git(repo, 'push', 'fixture', 'HEAD:refs/heads/topic', env=env, allowed=(1,))
    assert 'IA_EXPIRED_SHIPPED' in blocked.stdout + blocked.stderr
    assert scanner.exists() and not (remote / 'refs/heads/topic').exists()
    code, out, _ = ship(repo, '--channel', 'deck', env=env)
    assert code == 2 and not (repo / '.shipment.json').exists()  # the ship_command entry stops on the same commit
    dump(repo / 'design/info_arch/info_arch.yaml', fresh_ssot())
    commit(repo, 'refresh the proof date')
    run_git(repo, 'push', 'fixture', 'HEAD:refs/heads/topic', env=env)
    assert (remote / 'refs/heads/topic').read_text().strip() == run_git(repo, 'rev-parse', 'HEAD').stdout.strip()
    assert scanner.read_text().count('called') >= 2
    if before is not None:
        assert hashlib.sha256(global_config.read_bytes()).hexdigest() == before
    assert not run_git(repo, 'config', '--local', '--get', 'core.hooksPath', allowed=(1,)).stdout


# ---- rendered outputs (dom) -----------------------------------------------------------------------

@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_shipped_html_is_rendered_and_its_first_screens_are_judged(repo, tmp_path):
    env = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    dump(repo / 'design/info_arch/channels/web.yaml', {**SHIP_DECK, 'media': 'web', 'outputs': ['site/**/*.html'],
                                                        'viewports': [{'name': 'mobile', 'width': 375, 'height': 812}]})
    code, out, proc = cli(repo, 'skeleton', '--root', str(repo), '--outline', 'lp-municipal', '--out', str(repo), env=env)
    assert code == 0, proc.stdout
    written = sorted((repo / 'site').rglob('*.html'))
    assert written, proc.stdout
    commit(repo, 'ship the LP')
    code, out, _ = ship(repo, env=env)
    assert code == 2 and 'IA_SHIP_UNCOVERED' in codes(out), out  # the LP's first screen has no judgment yet
    dom_stage = [s for s in out['stages'] if s['stage'] == 'dom'][0]
    assert dom_stage['exit_code'] == 0 and dom_stage['first_screens'] >= 1
    dom_out = tmp_path / 'dom.json'
    code, data, proc = cli(repo, 'dom', '--root', str(repo), '--outline', 'lp-municipal', '--date', TODAY, env=env)
    dom_out.write_text(proc.stdout)
    code, data, _ = cli(repo, 'judge', '--root', str(repo), '--date', TODAY, '--dom-report', str(dom_out), '--include-candidates', env=env)
    ledger = {'schema_version': 3, 'policy': jl.load_policy(),
              'rows': [{'candidate': r, 'score': 0.0} for r in data['judgment_candidates']], 'holdout': {'source': 'unmeasured'}}
    (repo / 'design/judgments.json').write_text(json.dumps(ledger, ensure_ascii=False), encoding='utf-8')
    commit(repo, 'record the LP judgments')
    code, out, _ = ship(repo, env=env)
    assert code == 0, out
    html = written[0]
    original = html.read_text()
    # Below the first screen, so only the text outside every section changes (not the emphasis input).
    html.write_text(original.replace('</body>', '<div style="height:4000px"></div><footer><p>業界No.1 の実績</p></footer></body>'))
    commit(repo, 'a claim outside every section')
    code, out, _ = ship(repo, env=env)
    assert code == 2 and 'IA_SHIP_UNCOVERED' in codes(out), out  # the added text is judged before it ships
    html.write_text(original.replace('data-ia-msg', 'data-x'))
    commit(repo, 'anchors lost')
    code, out, _ = ship(repo, env=env)
    assert code == 2 and [s['stage'] for s in out['stages']][-1] == 'dom', out


# ---- Step 2.5 (maintenance) -----------------------------------------------------------------------

def maintain(root, *extra):
    return cli(root, 'maintain', '--root', str(root), '--date', TODAY, *extra)


def test_maintain_skips_in_one_line_without_info_arch(tmp_path):
    run_git(tmp_path, 'init', '-q')
    code, out, proc = maintain(tmp_path)
    assert code == 0 and proc.stdout.count('\n') == 1, proc.stdout
    assert out['status'] == 'skipped' and 'design/info_arch' in out['reason']


def test_maintain_lists_due_evidence_and_changed_ssot(repo):
    ssot = fresh_ssot()
    ssot['proofs']['prf-0001']['expires'] = '2026-11-20'  # within the default window
    dump(repo / 'design/info_arch/info_arch.yaml', ssot)
    code, out, proc = maintain(repo)
    assert proc.stdout.count('\n') == 1, proc.stdout
    assert code == 1, out
    assert out['due'] == [{'id': 'prf-0001', 'expires': '2026-11-20', 'outlines': ['deck-smb', 'lp-municipal'], 'shipped': True}]
    assert out['changed_paths'] == ['design/info_arch/info_arch.yaml']
    assert out['impacted_outlines'] == ['deck-smb', 'lp-municipal']
    code, out, _ = maintain(repo, '--within-days', '3')
    assert out['due'] == [] and code == 0, out  # a change alone is listed, not a warning
    assert out['changed_paths'] == ['design/info_arch/info_arch.yaml']


def test_maintain_is_clean_when_nothing_changed_or_falls_due(repo):
    code, out, _ = maintain(repo)
    assert code == 0 and out['status'] == 'checked' and out['due'] == [] and out['changed_paths'] == [], out


def test_maintain_fails_on_expired_shipped_evidence(repo):
    expire_shipped_proof(repo)
    commit(repo, 'expired proof')
    code, out, _ = maintain(repo)
    assert code == 2 and any('IA_EXPIRED_SHIPPED' in e for e in out['errors']), out


def test_session_close_runs_step_2_5_without_docs_governance():
    """Step 2.5 is keyed on design/ and design/info_arch/, not on docs/: no docs path may jump past it."""
    text = (SKILL.parent / 'own-session-close/SKILL.md').read_text()
    assert 'go straight to Step 3' not in text
    no_docs = [line for line in text.splitlines() if line.startswith('- **No `docs/` governance**')]
    assert len(no_docs) == 1 and 'Step 2.5' in no_docs[0], no_docs


# ---- the PATH entry is a symlink to the wrapper ----------------------------------------------------

def test_wrapper_runs_through_absolute_and_relative_symlinks(tmp_path):
    bindir = tmp_path / 'bin'
    bindir.mkdir()
    absolute = bindir / 'infoarch-author'
    absolute.symlink_to(SCRIPTS / 'infoarch-author')
    chained = tmp_path / 'other' / 'ia'
    chained.parent.mkdir()
    chained.symlink_to(os.path.relpath(absolute, chained.parent))  # relative link to the absolute link
    for link in (absolute, chained):
        proc = subprocess.run([str(link), 'maintain', '--root', str(tmp_path), '--date', TODAY], cwd=tmp_path,
                              capture_output=True, text=True, env=dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable))
        assert proc.returncode == 0, proc.stdout + proc.stderr
        assert json.loads(proc.stdout)['status'] == 'skipped'
