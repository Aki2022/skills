"""own-design-author follows the layer-paths rule and refuses a repository it cannot read (ISSUE-01 of the ux move)."""
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from layer_testkit import (ALL_SKILLS, FOREIGN_CODE, LAYOUT, ROOT, assert_foreign_failure, copy_skills, git_init,
                           moved_repo, run, SKILLS)

CATALOG_ITEM = {'id': 'dads-000001', 'match_key': 'dads-000001', 'document': 'fixture', 'title': 'x',
                'headings': ['x'], 'text': 'A synthetic fixture policy.', 'source_url': 'https://example.org/',
                'evidence_types': ['token']}
COMMANDS = ('check', 'maintain', 'sync')


def _catalogs(tmp_path):
    catalogs = tmp_path / 'catalogs'
    catalogs.mkdir()
    (catalogs / 'test-v1.json').write_text(json.dumps({'schema_version': 1, 'version': 'test-v1', 'items': [CATALOG_ITEM]}))
    (catalogs / 'id-ledger.json').write_text(json.dumps({'schema_version': 1, 'next_id': 2, 'entries': [
        {'id': 'dads-000001', 'match_key': 'dads-000001', 'retired': False}]}))
    return catalogs


@pytest.mark.parametrize('kind', ['foreign_only', 'both'])
@pytest.mark.parametrize('command', COMMANDS)
def test_design_author_stops_on_a_moved_repository(tmp_path, kind, command):
    repo = moved_repo(tmp_path, kind)
    result = run([SKILLS / 'own-design-author/scripts/design-author', command, '--repo', repo])
    assert_foreign_failure(result)


def test_design_author_init_refuses_a_moved_repository_and_creates_nothing(tmp_path):
    repo = moved_repo(tmp_path, 'foreign_only')
    result = run([SKILLS / 'own-design-author/scripts/design-author', 'init', '--repo', repo, '--preset', 'lp',
                  '--catalog-version', 'test-v1', '--catalog-dir', _catalogs(tmp_path)])
    assert_foreign_failure(result)
    assert not (repo / ROOT).exists()


def _initialised(tmp_path, skills, name):
    repo = git_init(tmp_path / name)
    subprocess.run(['git', '-C', str(repo), 'remote', 'add', 'origin', 'https://example.org/fixture.git'], check=True)
    init = run([skills / 'own-design-author/scripts/design-author', 'init', '--repo', repo, '--preset', 'lp',
                '--catalog-version', 'test-v1', '--catalog-dir', _catalogs(tmp_path)])
    return repo, init


def test_a_renamed_layer_root_is_followed_end_to_end(tmp_path):
    """Rename every entry's root in a copy of the skills: init writes the new place, check reads it, nothing is left at
    the old one. A command that still spells the old root would create or read it."""
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    repo, init = _initialised(tmp_path, skills, 'repo-renamed')
    assert init.returncode == 0, init.stdout + init.stderr
    assert (repo / renamed / 'design.yaml').is_file()
    assert not (repo / ROOT).exists(), 'a command still wrote the old root'
    check = run([skills / 'own-design-author/scripts/design-author', 'check', '--repo', repo])
    baseline_skills = copy_skills(tmp_path / 'skills-baseline', ALL_SKILLS)
    base_repo, base_init = _initialised(tmp_path, baseline_skills, 'repo-baseline')
    assert base_init.returncode == 0, base_init.stdout + base_init.stderr
    base = run([baseline_skills / 'own-design-author/scripts/design-author', 'check', '--repo', base_repo])
    assert not (repo / ROOT).exists(), 'check recreated the old root'
    assert check.returncode == base.returncode, (check.stdout, base.stdout)
    assert base.returncode in (0, 1), 'baseline must be a real verdict, not an error: ' + base.stdout


def test_a_broken_rule_stops_the_command(tmp_path):
    """The rule is read, not duplicated: remove an entry from a copy and the command fails instead of guessing."""
    def mutate(data):
        del data['entries']['ledger']
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS)
    repo, init = _initialised(tmp_path, skills, 'repo-broken')
    assert init.returncode == 0, init.stdout + init.stderr
    broken = copy_skills(tmp_path / 'skills-broken', ALL_SKILLS, mutate)
    result = run([broken / 'own-design-author/scripts/design-author', 'check', '--repo', repo])
    assert result.returncode != 0, result.stdout
    ok = run([skills / 'own-design-author/scripts/design-author', 'check', '--repo', repo])
    assert ok.returncode in (0, 1), ok.stdout
