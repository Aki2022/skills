"""`--mode backfill --source model`: 判定器の自動適用が `project_source: model` で書く（SPEC-project-association 要件 13）。

model は手動・legacy・出所不明の所属を上書きしない（拒否して何も書かない）。自分（model）の所属には追加できる。
人が明示した backfill（既定の manual）は、これまでどおり model に勝つ。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from conftest import snapshot
from test_apply_list_schema import backlink, body_of, fm, legacy_vault, record, run  # noqa: F401  (fixture と helper を共有する)

MODEL = ["--source", "model"]


def apply_model(vault: Path, key: str, *records: str) -> int:
    return run(vault, "backfill", key, *records, extra=[*MODEL, "--apply"])


def test_model_writes_a_list_with_source_model_and_a_backlink(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: kickoff\n")
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 0
    assert fm(path)["project"] == ["proj_a"] and fm(path)["project_source"] == "model"
    assert body_of(path).startswith(backlink("proj_a") + "\n")


def test_model_appends_to_its_own_association_and_keeps_the_source(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "project:\n  - proj_a\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    assert apply_model(legacy_vault, "proj_b", "record/r1.md") == 0
    assert fm(path)["project"] == ["proj_a", "proj_b"] and fm(path)["project_source"] == "model"
    assert body_of(path).startswith(backlink("proj_a") + "\n" + backlink("proj_b") + "\n")


def test_model_on_a_key_it_already_wrote_is_a_noop(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "project:\n  - proj_a\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 0
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    "frontmatter,body",
    [
        ("project:\n  - proj_a\nproject_source: manual\n", backlink("proj_a") + "\n\n本文。\n"),
        ("project:\n  - proj_a\nproject_source: legacy\n", backlink("proj_a") + "\n\n本文。\n"),
        ("project:\n  - proj_a\n", backlink("proj_a") + "\n\n本文。\n"),  # 出所が無い既存の所属も、人が付けたかもしれないので守る
        ("project: proj_a\n", backlink("proj_a") + "\n\n本文。\n"),
        ("project: []\nproject_source: manual\n", "本文。\n"),  # 人が「所属なし」と決めた
    ],
    ids=["manual", "legacy", "no-source", "scalar-no-source", "manual-empty"],
)
def test_model_never_overwrites_a_human_or_legacy_or_unknown_association(legacy_vault: Path, frontmatter: str, body: str, capsys) -> None:
    path = record(legacy_vault, "r1", frontmatter, body)
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_b", "record/r1.md") == 2  # CONFLICT
    assert path.read_bytes() == before
    assert "never overwrites" in capsys.readouterr().err


def test_a_conflict_in_one_record_writes_nothing_for_the_others(legacy_vault: Path) -> None:
    fresh = record(legacy_vault, "fresh", "title: x\n")
    manual = record(legacy_vault, "manual", "project:\n  - proj_a\nproject_source: manual\n", backlink("proj_a") + "\n\n本文。\n")
    before = snapshot(legacy_vault)
    assert apply_model(legacy_vault, "proj_b", "record/fresh.md", "record/manual.md") == 2
    assert snapshot(legacy_vault) == before and "project" not in fm(fresh)
    assert fm(manual)["project_source"] == "manual"


def test_model_dry_run_is_the_default_and_writes_nothing(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: kickoff\n")
    before = path.read_bytes()
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=MODEL) == 0
    assert path.read_bytes() == before


def test_a_human_backfill_still_beats_a_model_association(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: kickoff\n")
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 0
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    assert fm(path)["project_source"] == "manual"


def test_model_keeps_crlf_line_endings(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: kickoff\n", newline="\r\n")
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 0
    raw = path.read_bytes()
    assert b"\r\n" in raw and b"\n" not in raw.replace(b"\r\n", b"")


def test_an_unknown_source_value_is_refused(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "title: kickoff\n")
    with pytest.raises(SystemExit):
        run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--source", "guess"])


def test_model_on_a_key_it_already_wrote_does_not_touch_a_commented_frontmatter(legacy_vault: Path) -> None:
    """書き直さない（書き直すと YAML コメントを落とすので CONFLICT になる）。既に載っている key の再適用は何もしない。"""
    path = record(legacy_vault, "r1", "project:\n  - proj_a  # 人のメモ\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 0
    assert path.read_bytes() == before
