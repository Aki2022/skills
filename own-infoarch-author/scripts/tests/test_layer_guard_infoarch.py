"""own-infoarch-author follows the layer-paths rule and refuses a repository it cannot read (ISSUE-03 of the ux move).

Every subcommand of infoarch-author has a row below: a repository that already moved to a layout this version does not
read (a foreign root; with or without the current root beside it) stops the command with LAYER_LAYOUT_FOREIGN_ROOT.
`maintain` is the road that used to skip silently when the tree was absent: on a moved repository it must fail."""
import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-design-author/scripts/tests'))
from layer_testkit import (ALL_SKILLS, ROOT, SKILLS, assert_foreign_failure, copy_skills, moved_repo,  # noqa: E402
                           run)

sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_check as tc  # noqa: E402  (the green fixture repository writer)

CLI = SKILLS / 'own-infoarch-author/scripts/infoarch-author'
DATE = '2026-10-23'

# Every subcommand of infoarch-author. A row is (name, argv builder); the builder gets (repo, tmp_path).
# None is explicitly out of the guard: every subcommand takes --root and reads a layer, so nothing is listed here.
OUT_OF_THE_GUARD = {}


def _file(tmp_path, name, text):
    path = tmp_path / name
    path.write_text(text, encoding='utf-8')
    return path


COMMANDS = {
    'check': lambda repo, tmp: ['check'],
    'blueprint': lambda repo, tmp: ['blueprint'],
    'request': lambda repo, tmp: ['request', '--out', tmp / 'out'],
    'request-read-decisions': lambda repo, tmp: ['request', '--out', tmp / 'out',
                                                 '--read-decisions', _file(tmp, 'd.json', '[]')],
    'dom': lambda repo, tmp: ['dom', '--outline', 'x'],
    'judge': lambda repo, tmp: ['judge'],
    'accept-reviews': lambda repo, tmp: ['accept-reviews'],
    'evaluate': lambda repo, tmp: ['evaluate'],
    'ship': lambda repo, tmp: ['ship'],
    'ship-doctor': lambda repo, tmp: ['ship', '--doctor'],
    'ship-pre-push': lambda repo, tmp: ['ship', '--pre-push', '--refs-file', _file(tmp, 'refs', '')],
    'ship-channel': lambda repo, tmp: ['ship', '--channel', 'web'],
    'maintain': lambda repo, tmp: ['maintain'],
    'skeleton-outline': lambda repo, tmp: ['skeleton', '--outline', 'x'],
    'skeleton-screen': lambda repo, tmp: ['skeleton', '--screen', 's'],
    'impact': lambda repo, tmp: ['impact', '--id', 'seg-0001'],
    'set': lambda repo, tmp: ['set', '--outline', 'x'],
    'diff': lambda repo, tmp: ['diff', '--a', 'x', '--b', 'y'],
    'targets': lambda repo, tmp: ['targets'],
    'release': lambda repo, tmp: ['release'],
    'migrate-v6': lambda repo, tmp: ['migrate-v6'],
    'migrate-v6-write': lambda repo, tmp: ['migrate-v6', '--write'],
    'migrate-dictionary': lambda repo, tmp: ['migrate-dictionary', '--map', _file(tmp, 'm.json', '{}')],
    'migrate-dictionary-write': lambda repo, tmp: ['migrate-dictionary', '--map', _file(tmp, 'm.json', '{}'), '--write'],
    'answer-key': lambda repo, tmp: ['answer-key', '--out', tmp / 'ak'],
    'answer-key-measure': lambda repo, tmp: ['answer-key', '--measure', _file(tmp, 'l.json', '{}'),
                                             '--key', _file(tmp, 'k.json', '{}')],
}


def _subcommands():
    """The subcommands the CLI declares, read from its own --help so a new one cannot be forgotten."""
    text = run([CLI, '--help']).stdout
    choices = re.search(r'\{([a-z0-9,-]+)\}', text)
    return set(choices.group(1).split(',')) if choices else set()


def test_every_subcommand_has_a_guard_row(tmp_path):
    covered = {argv(None, tmp_path)[0] for argv in COMMANDS.values()} | set(OUT_OF_THE_GUARD)
    declared = _subcommands()
    assert declared, 'the parser was not read (0 subcommands is not a pass)'
    assert declared <= covered, f'subcommands without a guard test: {sorted(declared - covered)}'


@pytest.mark.parametrize('kind', ['foreign_only', 'both'])
@pytest.mark.parametrize('name', sorted(COMMANDS))
def test_infoarch_author_stops_on_a_moved_repository(tmp_path, kind, name):
    repo = moved_repo(tmp_path, kind)
    argv = COMMANDS[name](repo, tmp_path)
    result = run([CLI, *argv, '--root', repo])
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', ['foreign_only', 'both'])
def test_the_second_checkout_of_diff_is_guarded_too(tmp_path, kind):
    moved = moved_repo(tmp_path, kind)
    plain = tmp_path / 'plain'
    plain.mkdir()
    result = run([CLI, 'diff', '--a', 'x', '--b', 'y', '--root', plain, '--root-b', moved])
    assert_foreign_failure(result)


def test_maintain_on_a_moved_repository_is_not_a_skip(tmp_path):
    """The silent skip ("design/info_arch/ absent" -> exit 0) is the "0 found" failure this move closes."""
    repo = moved_repo(tmp_path, 'foreign_only')
    result = run([CLI, 'maintain', '--root', repo])
    assert result.returncode != 0, result.stdout
    assert '"status": "skipped"' not in result.stdout, result.stdout


def test_maintain_still_skips_a_repository_that_has_no_tree_at_all(tmp_path):
    """Behaviour that must not change: a repository that never had the layer is a skip, not a failure."""
    from layer_testkit import git_init
    repo = git_init(tmp_path / 'plain')
    result = run([CLI, 'maintain', '--root', repo])
    assert result.returncode == 0, result.stdout + result.stderr
    assert json.loads(result.stdout)['status'] == 'skipped'


def test_the_library_roads_are_guarded_too(tmp_path):
    """ia_module.resolve (called by own-design-author's L0) and ia_ship.* are entered as libraries, not through the CLI."""
    repo = moved_repo(tmp_path, 'foreign_only')
    code = (
        'import sys\n'
        f'sys.path[:0] = [{str(SKILLS / "own-infoarch-author/scripts")!r}]\n'
        'import ia_ship, ia_module\n'
        f'repo = {str(repo)!r}\n'
        'from layer_paths import LayerLayoutError\n'
        'for name, call in (("resolve", lambda: ia_module.resolve(root=repo, approvals=[], profiles=[])),\n'
        '                   ("doctor", lambda: ia_ship.doctor(repo)),\n'
        '                   ("maintain", lambda: ia_ship.maintain(repo, "2026-10-23")),\n'
        '                   ("gate", lambda: ia_ship.gate(repo, "2026-10-23")),\n'
        '                   ("ship_channel", lambda: ia_ship.ship_channel(repo, "web", "2026-10-23")),\n'
        '                   ("pre_push", lambda: ia_ship.pre_push(repo, "refs", "2026-10-23"))):\n'
        '    try:\n'
        '        call()\n'
        '    except LayerLayoutError:\n'
        '        continue\n'
        '    except Exception as exc:\n'
        '        print(name, "raised", type(exc).__name__, exc); sys.exit(3)\n'
        '    print(name, "did not raise"); sys.exit(4)\n'
    )
    result = subprocess.run([sys.executable, '-c', code], capture_output=True, text=True)
    assert result.returncode == 0, result.stdout + result.stderr


def test_a_renamed_layer_root_is_followed_end_to_end(tmp_path):
    """Rename every entry's root in a copy of the skills, relocate a green fixture there: `check` reads the new place
    and reaches the same verdict as the unrenamed copy on the same fixture; nothing is left or created at the old root."""
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    baseline = copy_skills(tmp_path / 'skills-baseline', ALL_SKILLS)
    base_repo, moved = tmp_path / 'repo-baseline', tmp_path / 'repo-renamed'
    tc.write_repo(base_repo)
    tc.write_repo(moved)
    (moved / ROOT).rename(moved / renamed)
    assert not (moved / ROOT).exists()
    args = ['check', '--date', DATE]
    verdicts = {}
    for label, skill_dir, repo in (('baseline', baseline, base_repo), ('renamed', skills, moved)):
        result = run([skill_dir / 'own-infoarch-author/scripts/infoarch-author', *args, '--root', repo])
        report = json.loads(result.stdout)
        verdicts[label] = (result.returncode, report['exit_code'], report['counts'], report['targets'])
    assert verdicts['baseline'][0] in (0, 1), 'the baseline must be a real verdict: ' + repr(verdicts['baseline'])
    assert verdicts['baseline'][3]['outlines'] > 0, 'the baseline checked nothing: ' + repr(verdicts['baseline'])
    assert verdicts['renamed'] == verdicts['baseline'], verdicts
    assert not (moved / ROOT).exists(), 'a command recreated the old root'
