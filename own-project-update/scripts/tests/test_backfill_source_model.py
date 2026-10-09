"""`--mode backfill --source model`: 判定器の自動適用が `project_source: model` で書く（SPEC-project-association 要件 13）。

model は手動・legacy・出所不明の所属を上書きしない（拒否して何も書かない）。自分（model）の所属には追加できる。
人が明示した backfill（既定の manual）は、これまでどおり model に勝つ。
"""

from __future__ import annotations

from pathlib import Path

import pytest

import project_update as pu
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


# ---- PR #36 の独立レビューの指摘 ---------------------------------------------------------------------------


def test_duplicate_project_source_keys_stop_a_model_write(legacy_vault: Path) -> None:
    """PyYAML は後勝ちで読むが、書き直しは同名キーを 1 つにまとめる。人の manual 行を黙って消さない。"""
    path = record(legacy_vault, "r1", "project:\n  - proj_a\nproject_source: manual\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_b", "record/r1.md") == 2 and path.read_bytes() == before


def test_duplicate_project_keys_stop_a_model_write(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "project:\n  - proj_a\nproject:\n  - proj_a\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_b", "record/r1.md") == 2 and path.read_bytes() == before


@pytest.mark.parametrize("mode", ["validate", "bootstrap"])
def test_source_model_is_refused_outside_backfill(legacy_vault: Path, mode: str) -> None:
    record(legacy_vault, "r1", "title: x\n")
    assert run(legacy_vault, mode, "proj_a", "record/r1.md", extra=MODEL) == 3  # 何も書かれないのに成功に見えない


@pytest.mark.parametrize("value", ["Manual", "MODEL", "Model", "guess", "0", "no", "false"])
def test_only_the_exact_word_model_counts_as_the_models_own_source(legacy_vault: Path, value: str) -> None:
    path = record(legacy_vault, "r1", f"project:\n  - proj_a\nproject_source: {value}\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_b", "record/r1.md") == 2 and path.read_bytes() == before


def test_a_falsy_looking_source_with_no_project_is_still_protected(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "project: []\nproject_source: no\n", "本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 2 and path.read_bytes() == before


def test_a_manual_record_that_already_lists_the_key_is_still_a_conflict_for_the_model(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "project:\n  - proj_a\nproject_source: manual\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert apply_model(legacy_vault, "proj_a", "record/r1.md") == 2 and path.read_bytes() == before


def test_prepare_record_defaults_to_manual_and_rejects_an_unknown_source(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: x\n")
    change = pu.prepare_record(path, "proj_a")
    assert "project_source: manual" in change.after
    with pytest.raises(pu.ValidationError):
        pu.prepare_record(path, "proj_a", None, "Model")
