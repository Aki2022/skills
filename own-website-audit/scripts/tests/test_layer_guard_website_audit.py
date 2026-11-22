"""own-website-audit/scripts/measure_lp.mjs stops on a repository whose layers moved (ISSUE-03 of the ux move).

The guard fires before any browser starts and without Playwright: a machine with no runtime reports the layout, and a
manifest that does not exist at the path this version reads still gets the layout message, not a bare ENOENT.

Entries that never look at the layers are out of the guard: none (measure_lp.mjs is the skill's only script)."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-design-author/scripts/tests'))
from layer_testkit import (ALL_SKILLS, FOREIGN_CODE, LAYOUT, ROOT, SKILLS, assert_foreign_failure, copy_skills,  # noqa: E402
                           git_init, moved_repo, run)

SCRIPT = SKILLS / 'own-website-audit/scripts/measure_lp.mjs'
KINDS = ['foreign_only', 'both']
RESOLVED = LAYOUT['entries']['resolved']
NO_RUNTIME = {'OWN_PLAYWRIGHT_DIR': '/nonexistent-playwright', 'OWN_PLAYWRIGHT_HOME': '/nonexistent-playwright-home'}


@pytest.mark.parametrize('kind', KINDS)
def test_measure_lp_stops_on_a_moved_repository_with_a_missing_manifest(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    manifest = repo / RESOLVED / 'web.json'
    assert not manifest.exists()
    result = run(['node', SCRIPT, 'https://example.invalid/', '--manifest', manifest], cwd=tmp_path, env=NO_RUNTIME)
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_measure_lp_stops_on_a_moved_repository_with_an_existing_manifest(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    manifest = repo / RESOLVED / 'web.json'
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text('{}')
    result = run(['node', SCRIPT, 'https://example.invalid/', '--manifest', manifest], cwd=tmp_path, env=NO_RUNTIME)
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_measure_lp_stops_with_the_fixed_environment_flag_too(tmp_path, kind):
    """--environment makes the script look up the repository root a second time; both roads must stop the same way."""
    repo = moved_repo(tmp_path, kind)
    result = run(['node', SCRIPT, 'https://example.invalid/', '--manifest', repo / RESOLVED / 'web.json',
                  '--environment', repo / 'render-environment.json'], cwd=tmp_path, env=NO_RUNTIME)
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_measure_lp_stops_when_run_from_inside_a_moved_repository_with_the_default_manifest(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run(['node', SCRIPT, 'https://example.invalid/'], cwd=repo, env=NO_RUNTIME)
    assert_foreign_failure(result)


def test_a_manifest_that_is_not_in_the_resolved_place_keeps_its_own_message(tmp_path):
    """The guard is not a blanket refusal: a plain repository and a wrong path still get the old, specific message."""
    repo = git_init(tmp_path / 'plain')
    result = run(['node', SCRIPT, 'https://example.invalid/', '--manifest', repo / 'web.json'], cwd=tmp_path, env=NO_RUNTIME)
    assert result.returncode == 2
    assert FOREIGN_CODE not in result.stdout + result.stderr
    assert 'resolved manifest unavailable' in result.stderr


def test_the_usage_line_and_the_default_manifest_follow_a_renamed_layer_root(tmp_path):
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    result = run(['node', skills / 'own-website-audit/scripts/measure_lp.mjs'], cwd=tmp_path, env=NO_RUNTIME)
    assert result.returncode == 2
    assert RESOLVED.replace(ROOT, renamed, 1) + '/web.json' in result.stderr and (ROOT + '/') not in result.stderr, result.stderr
