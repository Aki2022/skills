"""書き込み経路の絶対パス検査が、読む経路と同じ強さであることを固定する。

2026-09-29 の実測: `_check_local_paths` は `validate`（_validate_record 経由）と
`bootstrap`（_preflight 内のループ）では record に対して呼ばれるが、`backfill` が通る
`prepare_record` では呼ばれなかった。結果として **validate が拒否する入力を backfill が
受け入れて書き込む**状態だった。読むモードより書くモードが緩いのは、
「成功に見えて誤る」側の失敗なので fail-closed に揃える。

SPEC-project-update の「全経路に共通」Requirement 1 に対応する。
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "project_update.py"

PROJECT_NOTE = """---
project: acme
client: Acme Inc
---

# Acme

## overview

## people

## stakeholders

## next actions

## open issues / risks

## minutes

## related notes

## documents
"""

LIST_PROJECT = """# projects

| name | client |
| --- | --- |
| acme | Acme Inc |
"""

RECORD_CLEAN = """---
title: kickoff
---

普通の本文。相対リンクだけを持つ。

[資料](../attachments/deck.pdf)
"""

# 実測した 4 件と同じ形: recorder が frontmatter 直後の本文に書いた一時作業領域の絶対パス。
# `/private/tmp/...` は vibe-guard の local-info パターンには当たらないので、
# テストを書くために allowlist 例外を足す必要が無い（実物の形でそのまま固定できる）。
RECORD_WITH_ABSOLUTE_PATH = """---
title: kickoff
---

普通の本文。

source_transcript: /private/tmp/recorder/work/20260601_abc123/transcript/transcript.normalized.json
"""

RECORD_WITH_FILE_URL = """---
title: kickoff
---

普通の本文。

資料は file:///tmp/deck.pdf に置いた。
"""

# validate は「既に project に紐付いた」record を検査するので、YAML と backlink を持つ形にする。
_LINKED = """---
title: kickoff
project: acme
---

> project: [project_acme](../project/project_acme.md)

普通の本文。

{leak}
"""
LINKED_WITH_ABSOLUTE_PATH = _LINKED.format(leak="source_transcript: /private/tmp/recorder/work/20260601_abc123/transcript/transcript.normalized.json")
LINKED_WITH_FILE_URL = _LINKED.format(leak="資料は file:///tmp/deck.pdf に置いた。")


def build_vault(tmp_path: Path, record_body: str) -> tuple[Path, Path]:
    vault = tmp_path / "vault"
    (vault / "project").mkdir(parents=True)
    (vault / "record").mkdir(parents=True)
    (vault / "setting" / "list").mkdir(parents=True)
    (vault / "project" / "project_acme.md").write_text(PROJECT_NOTE, encoding="utf-8")
    (vault / "setting" / "list" / "list_project.md").write_text(LIST_PROJECT, encoding="utf-8")
    record = vault / "record" / "kickoff.md"
    record.write_text(record_body, encoding="utf-8")
    return vault, record


def run(vault: Path, record: Path, mode: str, *extra: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT), "--vault", str(vault), "--project-key", "acme",
         "--record", str(record), "--mode", mode, *extra],
        capture_output=True, text=True,
    )


@pytest.mark.parametrize("body", [LINKED_WITH_ABSOLUTE_PATH, LINKED_WITH_FILE_URL])
def test_validate_rejects_local_path(tmp_path: Path, body: str) -> None:
    """読む経路は拒否する（既存の挙動。対照として固定する）。"""
    vault, record = build_vault(tmp_path, body)
    result = run(vault, record, "validate")
    assert result.returncode == 3, result.stdout + result.stderr
    assert "local absolute path or file:// URL detected" in result.stderr


@pytest.mark.parametrize("body", [RECORD_WITH_ABSOLUTE_PATH, RECORD_WITH_FILE_URL])
def test_backfill_rejects_local_path_without_writing(tmp_path: Path, body: str) -> None:
    """書く経路も同じ強さで拒否し、ファイルを書き換えない。"""
    vault, record = build_vault(tmp_path, body)
    before = record.read_bytes()
    result = run(vault, record, "backfill", "--apply")
    assert result.returncode == 3, result.stdout + result.stderr
    assert "local absolute path or file:// URL detected" in result.stderr
    assert record.read_bytes() == before, "拒否したのに record が書き換わっている"


def test_backfill_still_writes_a_clean_record(tmp_path: Path) -> None:
    """陰性対照: 絶対パスが無い record は従来どおり書き込める（過剰拒否していない）。"""
    vault, record = build_vault(tmp_path, RECORD_CLEAN)
    before = record.read_text(encoding="utf-8")
    result = run(vault, record, "backfill", "--apply")
    assert result.returncode == 0, result.stdout + result.stderr
    after = record.read_text(encoding="utf-8")
    assert after != before, "clean な record が書き換わっていない"
    assert "project:\n  - acme\n" in after  # 書き手は常にリスト（SPEC-project-association）
