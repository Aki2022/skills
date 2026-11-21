"""own-positioning-author follows the layer-paths rule and refuses a repository it cannot read (ISSUE-03 of the ux move).

Every subcommand of positioning-author has a row below: a repository that already moved to a layout this version does
not read (a foreign root; with or without the current root beside it) stops the command with LAYER_LAYOUT_FOREIGN_ROOT,
instead of reporting "the SSOT is missing" or reading around it."""
import json
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-design-author/scripts/tests'))
from layer_testkit import (ALL_SKILLS, ROOT, SKILLS, assert_foreign_failure, copy_skills, moved_repo,  # noqa: E402
                           run)
from ia_fixtures import SSOT, confirm, dump  # noqa: E402

CLI = SKILLS / 'own-positioning-author/scripts/positioning-author'
# None is explicitly out of the guard: every subcommand takes --root and reads a layer.
OUT_OF_THE_GUARD = {}


def _file(tmp_path, name, text):
    path = tmp_path / name
    path.write_text(text, encoding='utf-8')
    return path


COMMANDS = {
    'check': lambda tmp: ['check'],
    'check-strict': lambda tmp: ['check', '--strict'],
    'positioning': lambda tmp: ['positioning'],
    'targets': lambda tmp: ['targets'],
    'migrate-approvals': lambda tmp: ['migrate-approvals'],
    'migrate-approvals-write': lambda tmp: ['migrate-approvals', '--write'],
    'review': lambda tmp: ['review', '--request', _file(tmp, 'r.json', '{"rows": []}')],
    'request': lambda tmp: ['request', '--out', tmp / 'out'],
    'request-read-decisions': lambda tmp: ['request', '--out', tmp / 'out', '--read-decisions', _file(tmp, 'd.json', '[]')],
}


def _subcommands():
    """The subcommands the CLI declares, read from its own --help so a new one cannot be forgotten."""
    text = run([CLI, '--help']).stdout
    choices = re.search(r'\{([a-z0-9,-]+)\}', text)
    return set(choices.group(1).split(',')) if choices else set()


def test_every_subcommand_has_a_guard_row(tmp_path):
    covered = {argv(tmp_path)[0] for argv in COMMANDS.values()} | set(OUT_OF_THE_GUARD)
    declared = _subcommands()
    assert declared, 'the parser was not read (0 subcommands is not a pass)'
    assert declared <= covered, f'subcommands without a guard test: {sorted(declared - covered)}'


@pytest.mark.parametrize('kind', ['foreign_only', 'both'])
@pytest.mark.parametrize('name', sorted(COMMANDS))
def test_positioning_author_stops_on_a_moved_repository(tmp_path, kind, name):
    repo = moved_repo(tmp_path, kind)
    result = run([CLI, *COMMANDS[name](tmp_path), '--root', repo])
    assert_foreign_failure(result)


def test_the_library_roads_are_guarded_too(tmp_path):
    """positioning_model.load_ssot and positioning.read_approvals are entered by own-infoarch-author as libraries."""
    repo = moved_repo(tmp_path, 'foreign_only')
    code = (
        'import sys\n'
        f'sys.path[:0] = [{str(SKILLS / "own-positioning-author/scripts")!r}]\n'
        'import ia_common, positioning, positioning_model as pm\n'
        'from pathlib import Path\n'
        'from layer_paths import LayerLayoutError\n'
        f'repo = Path({str(repo)!r})\n'
        'report = ia_common.Report("t", "2026-10-23")\n'
        'for name, call in (("load_ssot", lambda: pm.load_ssot(repo, report)),\n'
        '                   ("read_approvals", lambda: positioning.read_approvals(repo, report, ia_common.load_yaml, ia_common.YamlError))):\n'
        '    try:\n'
        '        call()\n'
        '    except LayerLayoutError:\n'
        '        continue\n'
        '    print(name, "did not raise"); sys.exit(4)\n'
    )
    import subprocess
    result = subprocess.run([sys.executable, '-c', code], capture_output=True, text=True)
    assert result.returncode == 0, result.stdout + result.stderr


def test_a_renamed_layer_root_is_followed_end_to_end(tmp_path):
    """Rename every entry's root in a copy of the skills, relocate a green fixture there: `check` reads the new place and
    reaches the same verdict as the unrenamed copy on the same fixture; the owner's approvals are read from the new
    ledger place (approved count equals the baseline's), and nothing is left at the old root."""
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    baseline = copy_skills(tmp_path / 'skills-baseline', ALL_SKILLS)
    base_repo, moved = tmp_path / 'repo-baseline', tmp_path / 'repo-renamed'
    for repo in (base_repo, moved):
        dump(repo / ROOT / 'info_arch/info_arch.yaml', SSOT)
        confirm(repo, SSOT)
    (moved / ROOT).rename(moved / renamed)
    assert not (moved / ROOT).exists()
    verdicts = {}
    for label, skill_dir, repo in (('baseline', baseline, base_repo), ('renamed', skills, moved)):
        result = run([skill_dir / 'own-positioning-author/scripts/positioning-author', 'check', '--strict', '--root', repo])
        report = json.loads(result.stdout)
        verdicts[label] = (result.returncode, report['counts'], report['targets'])
    assert verdicts['baseline'][0] in (0, 1) and verdicts['baseline'][2] == {'values': 3, 'approved': 3}, verdicts
    assert verdicts['renamed'] == verdicts['baseline'], verdicts
    assert not (moved / ROOT).exists(), 'a command recreated the old root'
