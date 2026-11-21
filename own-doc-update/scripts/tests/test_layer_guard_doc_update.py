"""validate_repo_docs.py follows the layer-paths rule and refuses a repository it cannot read (ISSUE-03 of the ux move).

The validator reads the design master (design config) for the specs back-reference check. A repository that already
moved to a layout this version does not read must make it report LAYER_LAYOUT_FOREIGN_ROOT and exit non-zero (not read
around it and report "no design master"). A repository with no layers and no foreign root behaves as before (the
existing test_design_links.py pins that).

Entries that never look at the layers are out of the guard: every other script of the skill (create_*, archive_*,
docs_hygiene, index_*, dictionary, ownership, init_repo_docs) never reads the layer directory."""
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-design-author/scripts/tests'))
from layer_testkit import (ALL_SKILLS, FOREIGN, FOREIGN_CODE, LAYOUT, ROOT, SKILLS, assert_foreign_failure, copy_skills,  # noqa: E402
                           git_init, moved_repo, run)

VALIDATOR = SKILLS / 'own-doc-update/scripts/validate_repo_docs.py'
CONFIG = LAYOUT['entries']['design_config']
MASTER = LAYOUT['entries']['design_dir']


@pytest.mark.parametrize('kind', ['foreign_only', 'both'])
def test_validator_stops_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run([sys.executable, VALIDATOR, repo])
    assert_foreign_failure(result)


def test_validator_stops_on_a_moved_repository_that_is_otherwise_valid_and_says_the_layout_error_once(tmp_path):
    """A repository whose docs are complete must still fail for the layout, not pass for lack of anything to check."""
    repo = git_init(tmp_path / 'full')
    init = subprocess.run([sys.executable, SKILLS / 'own-doc-update/scripts/init_repo_docs.py', str(repo)],
                          capture_output=True, text=True)
    clean = run([sys.executable, VALIDATOR, repo])
    assert FOREIGN_CODE not in clean.stdout + clean.stderr, (init.stdout + init.stderr, clean.stdout)
    (repo / FOREIGN[0]).mkdir()
    (repo / FOREIGN[0] / '.keep').write_text('')
    result = run([sys.executable, VALIDATOR, repo])
    assert_foreign_failure(result)
    assert (result.stdout + result.stderr).count(FOREIGN_CODE) == 1, result.stdout


def _spec_repo(tmp_path, root_name):
    repo = tmp_path / 'repo'
    (repo / root_name).mkdir(parents=True)
    (repo / 'docs/specs').mkdir(parents=True)
    return repo


def _findings(skills_dir, repo):
    code = ("import importlib.util,sys;"
            "spec=importlib.util.spec_from_file_location('v',sys.argv[1]);m=importlib.util.module_from_spec(spec);"
            "spec.loader.exec_module(m);e=[];m.validate_design_references(__import__('pathlib').Path(sys.argv[2]),e);"
            "print('\\n'.join(e))")
    return run([sys.executable, '-c', code, skills_dir / 'own-doc-update/scripts/validate_repo_docs.py', repo])


def test_the_design_master_is_read_from_the_place_the_rule_names(tmp_path):
    """Rename the layer root in a copy of the skills: the copy reads the new place and prints the new name."""
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    moved_config = CONFIG.replace(ROOT, renamed, 1)
    repo = _spec_repo(tmp_path, renamed)
    (repo / moved_config).parent.mkdir(parents=True, exist_ok=True)
    (repo / 'docs/specs/visual.md').write_text(f'# Policy\n\n[Master](../../{moved_config})\n')
    (repo / moved_config).write_text(json.dumps({'spec': 'docs/specs/visual.md'}))
    ok = _findings(skills, repo)
    assert ok.returncode == 0 and ok.stdout.strip() == '', ok.stdout + ok.stderr
    (repo / moved_config).unlink()
    red = _findings(skills, repo)
    assert f'{moved_config}: missing master referenced by specs' in red.stdout, red.stdout + red.stderr
    assert f'{ROOT}/' not in red.stdout


def test_the_validator_without_the_resolver_fails_loudly_and_never_falls_back(tmp_path):
    """own-doc-update copied alone (no own-design-author next to it): a clear message, a non-zero exit, no literal path."""
    alone = copy_skills(tmp_path / 'alone', ('own-doc-update',))
    repo = _spec_repo(tmp_path, ROOT)
    (repo / CONFIG).parent.mkdir(parents=True, exist_ok=True)
    (repo / CONFIG).write_text(json.dumps({'spec': 'docs/specs/visual.md'}))
    result = run([sys.executable, alone / 'own-doc-update/scripts/validate_repo_docs.py', repo])
    text = result.stdout + result.stderr
    assert result.returncode != 0, text
    assert 'layer_paths' in text and 'own-design-author' in text, text
