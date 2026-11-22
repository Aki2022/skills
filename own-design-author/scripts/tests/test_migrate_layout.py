"""`design-author migrate-layout`: the one-time move of a repository's layers from the old `design/` tree to
`ux/positioning`, `ux/infoarch`, `ux/design`, the ledger and the judgment record directly under `ux/`.

The move changes places only: every file keeps its bytes, a layer the repository does not have is not created."""
import hashlib
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import layer_paths as lp  # noqa: E402
import migrate_layout as ml  # noqa: E402
from layer_testkit import SKILLS, git_init, run  # noqa: E402

LEGACY = 'design'  # the tree being migrated away from; spelled here because this test is about it


def write(repo, rel, text='x\n'):
    path = repo / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)
    return path


def full_repo(tmp_path, name='repo', with_ia=True, with_positioning=True):
    repo = git_init(tmp_path / name)
    for key in ('user.email', 'user.name'):
        subprocess.run(['git', '-C', str(repo), 'config', key, 'fixture'], check=True)  # any string: git only needs one
    write(repo, f'{LEGACY}/design.yaml', 'a: 1\n')
    write(repo, f'{LEGACY}/DESIGN.md', 'prose mentions design/assets once\n')
    write(repo, f'{LEGACY}/NOTICE', 'n\n')
    write(repo, f'{LEGACY}/tokens/brand.tokens.json', '{}\n')
    write(repo, f'{LEGACY}/profiles/web.json', '{}\n')
    write(repo, f'{LEGACY}/checks/a.yaml', 'targets: [x]\n')
    write(repo, f'{LEGACY}/checks/assets.yaml', '{"rules": [{"id": "logo", "kind": "source-hash", "targets": ["design/assets/logo.png"]}]}\n')
    write(repo, f'{LEGACY}/assets/logo.png', 'png\n')
    write(repo, f'{LEGACY}/render-environment.json', '{}\n')
    write(repo, f'{LEGACY}/approvals.yaml', '[]\n')
    write(repo, f'{LEGACY}/judgments.json', '{}\n')
    write(repo, f'{LEGACY}/approval-request.json', '{}\n')
    if with_ia:
        write(repo, f'{LEGACY}/info_arch/INFO_ARCH.md', 'ia\n')
        write(repo, f'{LEGACY}/info_arch/channels/web.yaml', 'c: 1\n')
        write(repo, f'{LEGACY}/info_arch/outlines/lp.yaml', 'o: 1\n')
        write(repo, f'{LEGACY}/info_arch/accepted_reviews.yaml', '[]\n')
    if with_positioning:
        write(repo, f'{LEGACY}/info_arch/info_arch.yaml', 'schema_version: 5\n')
    write(repo, '.gitignore', f'node_modules/\n{LEGACY}/.resolved/\n{LEGACY}/assets.json\n# {LEGACY}/ kept in comments\n')
    subprocess.run(['git', '-C', str(repo), 'add', '-A'], check=True)
    subprocess.run(['git', '-C', str(repo), 'commit', '-q', '-m', 'fixture'], check=True)
    # ignored generated files that are not tracked
    write(repo, f'{LEGACY}/.resolved/web.json', '{}\n')
    write(repo, f'{LEGACY}/assets.json', '{}\n')
    return repo


def tree(repo, root):
    base = repo / root
    if not base.exists():
        return {}
    return {p.relative_to(repo).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted(base.rglob('*')) if p.is_file()}


def test_every_old_path_has_one_new_place_named_by_the_layout_rule():
    assert ml.map_path(f'{LEGACY}/info_arch/info_arch.yaml') == lp.rel('positioning_file')
    assert ml.map_path(f'{LEGACY}/info_arch/outlines/lp.yaml') == lp.rel('infoarch', 'outlines', 'lp.yaml')
    assert ml.map_path(f'{LEGACY}/info_arch/INFO_ARCH.md') == lp.rel('infoarch', 'INFO_ARCH.md')
    assert ml.map_path(f'{LEGACY}/approvals.yaml') == lp.rel('ledger')
    assert ml.map_path(f'{LEGACY}/judgments.json') == lp.rel('judgments')
    assert ml.map_path(f'{LEGACY}/design.yaml') == lp.rel('design_config')
    assert ml.map_path(f'{LEGACY}/.resolved/web.json') == lp.rel('resolved', 'web.json')
    assert ml.map_path(f'{LEGACY}/tokens/brand.tokens.json') == lp.rel('design_dir', 'tokens', 'brand.tokens.json')
    with pytest.raises(ValueError):
        ml.map_path('docs/x.md')


def test_dry_run_changes_nothing_and_says_what_would_move(tmp_path):
    repo = full_repo(tmp_path)
    before = tree(repo, LEGACY)
    status = subprocess.run(['git', '-C', str(repo), 'status', '--porcelain'], capture_output=True, text=True).stdout
    plan = ml.plan(repo)
    assert plan['applied'] is False
    assert len(plan['moves']) == len(before) and len(before) >= 15
    assert tree(repo, LEGACY) == before and not (repo / lp.LAYERS_ROOT).exists()
    assert subprocess.run(['git', '-C', str(repo), 'status', '--porcelain'], capture_output=True, text=True).stdout == status


def test_apply_moves_every_file_with_its_bytes_and_nothing_stays(tmp_path):
    repo = full_repo(tmp_path)
    before = tree(repo, LEGACY)
    result = ml.apply(repo)
    assert result['applied'] is True
    assert not (repo / LEGACY).exists(), 'the old tree must be gone, empty directories included'
    after = tree(repo, lp.LAYERS_ROOT)
    assert sorted(v for k, v in after.items() if not k.endswith('checks/assets.yaml')) == sorted(
        v for k, v in before.items() if not k.endswith('checks/assets.yaml')), 'the same bytes, one for one'
    assert len(after) == len(before)
    for old, digest in before.items():
        if old.endswith('checks/assets.yaml'):
            continue  # its asset targets are rewritten on purpose (see the next test)
        assert after[ml.map_path(old)] == digest
    assert (repo / lp.rel('positioning_file')).read_text() == 'schema_version: 5\n'
    assert (repo / lp.rel('infoarch', 'channels', 'web.yaml')).is_file()
    assert (repo / lp.rel('ledger')).is_file() and (repo / lp.rel('judgments')).is_file()


def test_tracked_files_are_renamed_in_git_and_ignored_ones_are_carried_along(tmp_path):
    repo = full_repo(tmp_path)
    ml.apply(repo)
    status = subprocess.run(['git', '-C', str(repo), 'status', '--porcelain'], capture_output=True, text=True).stdout
    assert 'R  ' in status and ' D ' not in status.replace('RD', '')
    ignored = subprocess.run(['git', '-C', str(repo), 'status', '--porcelain', '--ignored'], capture_output=True, text=True).stdout
    assert (repo / lp.rel('resolved', 'web.json')).is_file()
    assert (repo / lp.rel('assets_index')).is_file()
    assert lp.rel('resolved') in ignored or lp.rel('assets_index') in ignored


def test_asset_rule_targets_follow_the_assets_and_nothing_else_in_the_file_changes(tmp_path):
    repo = full_repo(tmp_path)
    result = ml.apply(repo)
    new = (repo / lp.rel('checks', 'assets.yaml')).read_text()
    assert f'"{lp.rel("assets")}/logo.png"' in new and f'"{LEGACY}/assets/' not in new
    assert json.loads(new)['rules'][0]['id'] == 'logo'
    assert result['rule_target_rewrites'] == [{'file': lp.rel('checks', 'assets.yaml'), 'count': 1}]
    other = (repo / lp.rel('checks', 'a.yaml')).read_text()
    assert other == 'targets: [x]\n'


def test_the_rule_rewrite_is_restored_when_the_move_fails_late(tmp_path, monkeypatch):
    repo = full_repo(tmp_path)
    original = (repo / LEGACY / 'checks' / 'assets.yaml').read_bytes()
    real = ml._remove_empty_dirs
    calls = {'n': 0}

    def late(base):
        calls['n'] += 1
        if calls['n'] == 1:
            raise OSError('late failure')
        return real(base)
    monkeypatch.setattr(ml, '_remove_empty_dirs', late)
    with pytest.raises(ml.MigrationError):
        ml.apply(repo)
    assert (repo / LEGACY / 'checks' / 'assets.yaml').read_bytes() == original


def test_gitignore_lines_follow_the_files_and_comments_stay(tmp_path):
    repo = full_repo(tmp_path)
    ml.apply(repo)
    text = (repo / '.gitignore').read_text().splitlines()
    assert lp.rel('resolved') + '/' in text and lp.rel('assets_index') in text
    assert 'node_modules/' in text and f'# {LEGACY}/ kept in comments' in text
    assert f'{LEGACY}/.resolved/' not in text and f'{LEGACY}/assets.json' not in text


def test_a_layer_the_repository_does_not_have_is_not_created(tmp_path):
    repo = full_repo(tmp_path, 'design-only', with_ia=False, with_positioning=False)
    ml.apply(repo)
    assert not (repo / lp.rel('infoarch')).exists()
    assert not (repo / lp.rel('positioning_dir')).exists()
    assert (repo / lp.rel('design_config')).is_file()


def test_positioning_without_the_rest_of_the_infoarch(tmp_path):
    repo = full_repo(tmp_path, 'pos-only', with_ia=False, with_positioning=True)
    ml.apply(repo)
    assert (repo / lp.rel('positioning_file')).is_file()
    assert not (repo / lp.rel('infoarch')).exists()


def test_prose_that_names_the_old_place_is_reported_and_left_alone(tmp_path):
    repo = full_repo(tmp_path)
    result = ml.apply(repo)
    notes = [m for m in result['prose_mentions'] if m['file'].endswith('DESIGN.md')]
    assert notes and notes[0]['lines'] == 1
    assert (repo / lp.rel('design_dir', 'DESIGN.md')).read_text() == 'prose mentions design/assets once\n'


@pytest.mark.parametrize('case', ['already_moved', 'nothing_to_move', 'both'])
def test_refuses_where_there_is_nothing_to_migrate_or_it_would_overwrite(tmp_path, case):
    repo = git_init(tmp_path / case)
    if case == 'already_moved':
        write(repo, lp.rel('design_config'))
    elif case == 'both':
        write(repo, f'{LEGACY}/design.yaml')
        write(repo, lp.rel('design_config'))
    with pytest.raises(ml.MigrationError):
        ml.plan(repo)
    with pytest.raises(ml.MigrationError):
        ml.apply(repo)


def test_a_failure_in_the_middle_puts_everything_back(tmp_path, monkeypatch):
    repo = full_repo(tmp_path)
    before = tree(repo, LEGACY)
    real = ml._move
    calls = {'n': 0}

    def flaky(repo_, src, dst, tracked):
        calls['n'] += 1
        if calls['n'] == 6:
            raise OSError('disk full')
        return real(repo_, src, dst, tracked)
    monkeypatch.setattr(ml, '_move', flaky)
    with pytest.raises(ml.MigrationError):
        ml.apply(repo)
    assert tree(repo, LEGACY) == before
    assert not (repo / lp.LAYERS_ROOT).exists()
    assert (repo / '.gitignore').read_text().splitlines()[1] == f'{LEGACY}/.resolved/'


def test_a_failure_after_the_gitignore_was_rewritten_restores_it_too(tmp_path, monkeypatch):
    repo = full_repo(tmp_path)
    before = tree(repo, LEGACY)
    ignore = (repo / '.gitignore').read_text()
    real = ml._remove_empty_dirs
    calls = {'n': 0}

    def late(base):
        calls['n'] += 1
        if calls['n'] == 1:
            raise OSError('late failure')
        return real(base)
    monkeypatch.setattr(ml, '_remove_empty_dirs', late)
    with pytest.raises(ml.MigrationError):
        ml.apply(repo)
    assert (repo / '.gitignore').read_text() == ignore
    assert tree(repo, LEGACY) == before


def test_cli_dry_run_then_apply_and_a_second_apply_fails(tmp_path):
    repo = full_repo(tmp_path)
    script = SKILLS / 'own-design-author/scripts/design-author'
    dry = run([script, 'migrate-layout', '--repo', repo])
    assert dry.returncode == 0, dry.stdout + dry.stderr
    out = json.loads(dry.stdout)
    assert out['applied'] is False and out['moves']
    assert (repo / LEGACY / 'design.yaml').exists()
    done = run([script, 'migrate-layout', '--repo', repo, '--apply'])
    assert done.returncode == 0, done.stdout + done.stderr
    assert json.loads(done.stdout)['applied'] is True
    again = run([script, 'migrate-layout', '--repo', repo, '--apply'])
    assert again.returncode != 0
