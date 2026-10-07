"""project_update の関連付け系テストが共有する fixture vault。

公開 repo に置くので、client 名・実 project key は書かない（proj_a / proj_b の合成名だけ）。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1]
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

PROJECT_NOTE = """---
title: project_{key}
project: {key}
client: Client {key}
status: active
---

# {key}

## overview

人が書いた物語部分。render は触らない。

## next actions

- 手書きの項目
"""

LIST_PROJECT = """# projects

| name | client |
| --- | --- |
| proj_a | Client proj_a |
| proj_b | Client proj_b |
"""


@pytest.fixture
def vault(tmp_path: Path) -> Path:
    root = tmp_path / "vault"
    for directory in ("project", "record", "presentation", "setting/list", "setting/template"):
        (root / directory).mkdir(parents=True)
    for key in ("proj_a", "proj_b"):
        (root / "project" / f"project_{key}.md").write_text(
            PROJECT_NOTE.format(key=key), encoding="utf-8"
        )
    (root / "setting/list/list_project.md").write_text(LIST_PROJECT, encoding="utf-8")
    return root


def write_note(path: Path, frontmatter: str, body: str = "本文。\n") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"---\n{frontmatter}---\n{body}", encoding="utf-8")
    return path


def snapshot(root: Path) -> dict[str, bytes]:
    return {
        str(p.relative_to(root)): p.read_bytes()
        for p in sorted(root.rglob("*"))
        if p.is_file() and not p.is_symlink()
    }
