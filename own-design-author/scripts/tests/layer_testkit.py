"""Shared helpers for the layer-layout tests of every skill (import with sys.path pointing at this directory).

- `moved_repo(tmp_path, kind)`: a git repository that already moved to a layout this version does not read
  ('foreign_only' = only the foreign root, 'both' = foreign root and the current root).
- `copy_skills(dest, names, mutate)`: a copy of the skills (code and references, no tests) whose layer-paths.json can
  be rewritten, to prove a command follows the rule instead of spelling a path.
- `run(cmd, cwd, env)`: run a command, never raising.
"""
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
LAYOUT_FILE = SKILLS / 'own-design-author/references/layer-paths.json'
LAYOUT = json.loads(LAYOUT_FILE.read_text())
ROOT = LAYOUT['entries']['layers']
FOREIGN = LAYOUT['foreign_roots']
FOREIGN_CODE = 'LAYER_LAYOUT_FOREIGN_ROOT'
ALL_SKILLS = ('own-design-author', 'own-infoarch-author', 'own-positioning-author', 'own-doc-update',
              'own-web-build', 'own-website-audit')


def git_init(path):
    path.mkdir(parents=True, exist_ok=True)
    subprocess.run(['git', 'init', '-q', str(path)], check=True)
    return path


def moved_repo(tmp_path, kind):
    repo = git_init(tmp_path / ('repo-' + kind))
    for name in FOREIGN:
        (repo / name).mkdir()
        (repo / name / '.keep').write_text('')
    if kind == 'both':
        (repo / ROOT).mkdir()
        (repo / ROOT / '.keep').write_text('')
    elif kind != 'foreign_only':
        raise ValueError(kind)
    return repo


def copy_skills(dest, names=ALL_SKILLS, mutate=None):
    dest = Path(dest)
    for name in names:
        shutil.copytree(SKILLS / name, dest / name, symlinks=True,
                        ignore=shutil.ignore_patterns('tests', '__pycache__', 'node_modules', '.git', '.pytest_cache'))
    if mutate:
        file = dest / 'own-design-author/references/layer-paths.json'
        data = json.loads(file.read_text())
        mutate(data)
        file.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    return dest


def run(cmd, cwd=None, env=None, timeout=120):
    merged = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable)
    merged.update(env or {})
    return subprocess.run([str(c) for c in cmd], cwd=cwd, env=merged, capture_output=True, text=True, timeout=timeout)


def assert_foreign_failure(result):
    """A command run on a moved repository fails (non-zero) and says why, on stdout or stderr."""
    text = result.stdout + result.stderr
    assert result.returncode != 0, 'the command passed on a repository this version does not read:\n' + text
    assert FOREIGN_CODE in text, 'it failed, but not with the layout message:\n' + text
