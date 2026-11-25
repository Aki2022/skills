"""own-web-build stops on a repository whose layers moved to a layout this version does not read (ISSUE-03 of the ux move).

Every command line entry of the skill is run on a moved repository ('foreign_only': only the foreign root; 'both': the
foreign root next to the current one). Each must fail (non-zero) with LAYER_LAYOUT_FOREIGN_ROOT, never skip, never
report "0 found", and never write.

Entries that never look at the layers are out of the guard: none (the factory scripts all name a repository)."""
import json
import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-design-author/scripts/tests'))
from layer_testkit import (ALL_SKILLS, FOREIGN_CODE, LAYOUT, ROOT, SKILLS, assert_foreign_failure, copy_skills,  # noqa: E402
                           git_init, moved_repo, run)

SCRIPTS = SKILLS / 'own-web-build/scripts'
KINDS = ['foreign_only', 'both']
RESOLVED = LAYOUT['entries']['resolved']
OVERRIDES = LAYOUT['entries']['layout_overrides']


def _html(tmp_path):
    page = tmp_path / 'page.html'
    page.write_text('<!doctype html><html lang="ja"><title>x</title><body>x</body></html>')
    return page


@pytest.mark.parametrize('kind', KINDS)
def test_audit_html_stops_on_a_moved_repository_even_when_the_manifest_is_missing(tmp_path, kind):
    """The manifest of a moved repository does not exist at the path this version reads: still the layout message,
    not a bare ENOENT or 'invalid or unavailable input'."""
    repo = moved_repo(tmp_path, kind)
    manifest = repo / RESOLVED / 'web.json'
    assert not manifest.exists()
    result = run([sys.executable, SCRIPTS / 'audit_html.py', _html(tmp_path), '--manifest', manifest])
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_audit_html_stops_on_a_manifest_that_exists_in_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    manifest = repo / RESOLVED / 'web.json'
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text('{}')
    result = run([sys.executable, SCRIPTS / 'audit_html.py', _html(tmp_path), '--manifest', manifest])
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_audit_screen_stops_on_a_moved_repository_given_repo(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run([sys.executable, SCRIPTS / 'audit_screen.py', '--screen', tmp_path / 'screen.yaml',
                  '--html', _html(tmp_path), '--repo', repo], cwd=tmp_path)
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_audit_screen_stops_on_a_moved_repository_given_overrides_only(tmp_path, kind):
    """No --repo and the working directory is elsewhere: the overrides file names the repository."""
    repo = moved_repo(tmp_path, kind)
    elsewhere = tmp_path / 'elsewhere'
    elsewhere.mkdir()
    result = run([sys.executable, SCRIPTS / 'audit_screen.py', '--screen', tmp_path / 'screen.yaml',
                  '--html', _html(tmp_path), '--overrides', repo / OVERRIDES], cwd=elsewhere)
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_audit_screen_stops_when_run_from_inside_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run([sys.executable, SCRIPTS / 'audit_screen.py', '--screen', tmp_path / 'screen.yaml',
                  '--html', _html(tmp_path)], cwd=repo)
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_build_screen_stops_on_a_moved_repository_and_writes_nothing(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    out = tmp_path / 'out' / 'page.html'
    result = run([sys.executable, SCRIPTS / 'build_screen.py', '--repo', repo, '--screen', 'scr-x', '--out', out])
    assert_foreign_failure(result)
    assert not out.exists()


@pytest.mark.parametrize('kind', KINDS)
def test_measure_screen_stops_on_a_moved_repository_before_looking_for_a_browser(tmp_path, kind):
    """The guard fires before Playwright is looked for: a machine with no runtime at all still reports the layout."""
    repo = moved_repo(tmp_path, kind)
    env = {'OWN_PLAYWRIGHT_HOME': str(tmp_path / 'no-runtime'), 'SCREEN_PLAYWRIGHT_DIR': str(repo)}
    result = run(['node', SCRIPTS / 'measure_screen.mjs', _html(tmp_path), '[]'], cwd=tmp_path, env=env)
    assert_foreign_failure(result)
    result = run(['node', SCRIPTS / 'measure_screen.mjs', _html(tmp_path), '[]'], cwd=repo,
                 env={'OWN_PLAYWRIGHT_HOME': str(tmp_path / 'no-runtime')})
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_preflight_stops_on_a_moved_repository_and_says_why(tmp_path, kind):
    """preflight swallows an L0 answer it cannot parse; the layout message must reach the caller regardless."""
    repo = moved_repo(tmp_path, kind)
    result = run([sys.executable, SCRIPTS / 'preflight.py', '--repo', repo])
    assert_foreign_failure(result)
    answer = json.loads(result.stdout)
    assert answer['exit_code'] == 2 and any(FOREIGN_CODE in e for e in answer['errors'])


def test_preflight_guards_even_when_the_l0_command_is_missing(tmp_path):
    """The guard is preflight's own, not borrowed from L0: with no L0 command the message is still the layout one."""
    repo = moved_repo(tmp_path, 'foreign_only')
    result = run([sys.executable, SCRIPTS / 'preflight.py', '--repo', repo], env={'DESIGN_AUTHOR_COMMAND': str(tmp_path / 'missing')})
    assert_foreign_failure(result)


def test_a_repository_with_neither_root_is_not_a_layout_failure(tmp_path):
    """The guard is not a blanket refusal: a plain repository is stopped for its own reasons, not for the layout."""
    repo = git_init(tmp_path / 'plain')
    result = run([sys.executable, SCRIPTS / 'build_screen.py', '--repo', repo, '--screen', 'scr-x', '--out', tmp_path / 'p.html'])
    assert FOREIGN_CODE not in result.stdout + result.stderr


def test_the_allocation_rules_name_the_overrides_file_by_layer_entry_not_by_path():
    rules = json.loads((SKILLS / 'own-design-author/references/allocation-rules.json').read_text(encoding='utf-8'))
    override = rules['override']
    assert 'file' not in override, 'a literal path is back in the rules'
    assert override['file_entry'] in LAYOUT['entries']
    assert LAYOUT['entries'][override['file_entry']] == OVERRIDES


def test_a_renamed_layer_root_is_followed_by_the_guard_and_the_help(tmp_path):
    """Rename the layers in a copy of the skills: the copy's audit_screen/audit_html answer with the new names."""
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
        data['foreign_roots'] = ['zzforeign']
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    repo = git_init(tmp_path / 'repo')
    (repo / 'zzforeign').mkdir()
    for args in (['audit_screen.py', '--screen', 'x', '--html', 'x', '--repo', repo],
                 ['audit_html.py', _html(tmp_path), '--manifest', repo / RESOLVED.replace(ROOT, renamed, 1) / 'web.json']):
        result = run([sys.executable, skills / 'own-web-build/scripts' / args[0], *args[1:]], cwd=tmp_path)
        assert_foreign_failure(result)
    result = run([sys.executable, skills / 'own-web-build/scripts/audit_screen.py', '--help'])
    assert renamed + '/' in result.stdout and (ROOT + '/') not in result.stdout, result.stdout


@pytest.mark.parametrize('script', ['plan_check.py', 'comp_request.py'])
@pytest.mark.parametrize('kind', KINDS)
def test_the_plan_and_comp_entries_stop_on_a_moved_repository_and_write_nothing(tmp_path, kind, script):
    repo = moved_repo(tmp_path, kind)
    before = sorted(p.relative_to(repo).as_posix() for p in repo.rglob('*') if '.git' not in p.parts)
    result = run([sys.executable, SCRIPTS / script, '--repo', repo, '--deliverable', 'lp-x'])
    assert_foreign_failure(result)
    assert sorted(p.relative_to(repo).as_posix() for p in repo.rglob('*') if '.git' not in p.parts) == before
