"""apply（backfill）がリスト形式の project・project_source・summary・project ごとの backlink 行を書く。

SPEC-project-association Requirement 2・4・14 と WS-20261003-project-association の ISSUE-01 の受入:
読み手は旧来の文字列 `project:` を 1 要素のリストとして読み、既存の所属は保持して追加する
（`_assert_no_other_project` の単一前提を捨てる）。書き手は常にリストで書き、手動の指定は model に勝つ。
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest
import yaml

import project_update as pu
from conftest import snapshot

FULL_PROJECT_NOTE = """---
project: {key}
client: Client {key}
---

# {key}

## overview

## people

## stakeholders

## next actions

## open issues / risks

## minutes

## related notes

## documents
"""

LIST_TABLE = """# projects

| name | client |
| --- | --- |
| proj_a | Client proj_a |
| proj_b | Client proj_b |
"""


@pytest.fixture
def legacy_vault(tmp_path: Path) -> Path:
    root = tmp_path / "vault"
    for directory in ("project", "record", "setting/list"):
        (root / directory).mkdir(parents=True)
    for key in ("proj_a", "proj_b"):
        (root / "project" / f"project_{key}.md").write_text(FULL_PROJECT_NOTE.format(key=key), encoding="utf-8")
    (root / "setting/list/list_project.md").write_text(LIST_TABLE, encoding="utf-8")
    return root


def backlink(key: str) -> str:
    return f"> project: [project_{key}](../project/project_{key}.md)"


def record(vault: Path, name: str, frontmatter: str, body: str = "本文。\n", newline: str = "\n") -> Path:
    path = vault / "record" / f"{name}.md"
    text = f"---\n{frontmatter}---\n{body}"
    path.write_bytes(text.replace("\n", newline).encode("utf-8"))
    return path


def run(vault: Path, mode: str, key: str, *records: str, extra: list[str] | None = None) -> int:
    args = ["--vault", str(vault), "--mode", mode, "--project-key", key]
    for item in records:
        args += ["--record", item]
    return pu.main(args + (extra or []))


def fm(path: Path) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8").split("---\n", 2)[1])


def body_of(path: Path) -> str:
    return path.read_text(encoding="utf-8").split("---\n", 2)[2]


# ---- 書く: 常にリスト・project_source・backlink ---------------------------------------


def test_backfill_writes_list_project_source_and_backlink(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: kickoff\n")
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    assert fm(path)["project"] == ["proj_a"]  # 1 件でもリスト
    assert fm(path)["project_source"] == "manual"
    assert body_of(path).startswith(backlink("proj_a") + "\n")


def test_backfill_dry_run_is_default_and_writes_nothing(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "title: kickoff\n")
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md") == 0
    assert snapshot(legacy_vault) == before


def test_summary_option_writes_summary_and_conflicting_summary_stops(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: kickoff\n")
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply", "--summary", "立ち上げ会議"]) == 0
    assert fm(path)["summary"] == "立ち上げ会議"
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply", "--summary", "別の要約"]) == 2
    assert snapshot(legacy_vault) == before  # 既存の summary を黙って書き換えない
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply", "--summary", "立ち上げ会議"]) == 0


def test_summary_must_be_one_line_without_local_path(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "title: kickoff\n")
    before = snapshot(legacy_vault)
    for bad in ("一行目\n二行目", "see /private/tmp/someone/x", "   "):
        assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply", "--summary", bad]) == 3
    assert snapshot(legacy_vault) == before


def test_rerun_is_a_noop(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "title: kickoff\n")
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    first = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    assert snapshot(legacy_vault) == first


# ---- 読む: 旧来の文字列 project を 1 要素のリストとして ------------------------------------


def test_legacy_scalar_project_with_its_backlink_is_a_noop_for_the_same_key(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: old\nproject: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    before = path.read_bytes()
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    assert path.read_bytes() == before  # 旧形式のまま触らない（移行は migrate の仕事）


def test_validate_accepts_scalar_and_list_forms(legacy_vault: Path) -> None:
    record(legacy_vault, "old", "project: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    record(
        legacy_vault, "new", "project:\n  - proj_a\n  - proj_b\nproject_source: model\n",
        backlink("proj_a") + "\n" + backlink("proj_b") + "\n\n本文。\n",
    )
    assert run(legacy_vault, "validate", "proj_a", "record/old.md", "record/new.md") == 0
    assert run(legacy_vault, "validate", "proj_b", "record/new.md") == 0


def test_validate_rejects_a_missing_backlink_for_a_listed_project(legacy_vault: Path) -> None:
    record(legacy_vault, "r", "project:\n  - proj_a\n  - proj_b\n", backlink("proj_a") + "\n\n本文。\n")
    assert run(legacy_vault, "validate", "proj_a", "record/r.md") == 3


def test_validate_rejects_a_record_that_does_not_list_the_key(legacy_vault: Path) -> None:
    record(legacy_vault, "r", "project:\n  - proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    assert run(legacy_vault, "validate", "proj_b", "record/r.md") == 3


# ---- 既存の所属を保持して追加（旧: 別 project があると停止） ------------------------------


def test_adding_a_second_project_to_a_scalar_record_keeps_the_first(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: old\nproject: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply"]) == 0
    assert fm(path)["project"] == ["proj_a", "proj_b"]
    assert body_of(path).startswith(backlink("proj_a") + "\n" + backlink("proj_b") + "\n")
    assert fm(path)["title"] == "old"  # 他のキーは保たれる


def test_manual_beats_model_when_a_project_is_added(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "project:\n  - proj_a\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply"]) == 0
    assert fm(path)["project_source"] == "manual"


def test_explicit_backfill_of_an_existing_key_promotes_model_to_manual_but_leaves_legacy_alone(legacy_vault: Path) -> None:
    """人が明示的に採用したら model に勝つ（SPEC-project-association Req 13）。legacy・未記載は触らない。"""
    model = record(legacy_vault, "m", "project:\n  - proj_a\nproject_source: model\n", backlink("proj_a") + "\n\n本文。\n")
    legacy = record(legacy_vault, "l", "project:\n  - proj_a\nproject_source: legacy\n", backlink("proj_a") + "\n\n本文。\n")
    bare = record(legacy_vault, "b", "project: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    legacy_before, bare_before = legacy.read_bytes(), bare.read_bytes()
    assert run(legacy_vault, "backfill", "proj_a", "record/m.md", "record/l.md", "record/b.md", extra=["--apply"]) == 0
    assert fm(model)["project_source"] == "manual"
    assert legacy.read_bytes() == legacy_before and bare.read_bytes() == bare_before


@pytest.mark.parametrize(
    "frontmatter,body",
    [
        ("project: proj_a\n", "> project: [project_proj_a](../elsewhere/x.md)\n\n本文。\n"),  # 崩れた backlink
        ("project: proj_a\n", backlink("proj_x") + "\n\n本文。\n"),  # list に無い project の backlink
        ("project: proj_a\n", "\n本文。\n"),  # backlink が無い（読み手の矛盾）
    ],
    ids=["malformed-backlink", "backlink-for-unlisted-project", "missing-backlink"],
)
def test_inconsistent_record_stops_instead_of_being_guessed(legacy_vault: Path, frontmatter: str, body: str) -> None:
    record(legacy_vault, "r1", frontmatter, body)
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply"]) == 2
    assert snapshot(legacy_vault) == before


def test_bootstrap_no_longer_refuses_a_record_that_belongs_to_another_project(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "project: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    (legacy_vault / "project/project_proj_b.md").unlink()
    (legacy_vault / "setting/list/list_project.md").write_text("# p\n\n| name | client |\n| --- | --- |\n", encoding="utf-8")
    assert run(legacy_vault, "bootstrap", "proj_b", "record/r1.md") == 0


def test_bootstrap_still_refuses_when_the_key_is_already_present_without_a_note(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "project: proj_b\n", backlink("proj_b") + "\n\n本文。\n")
    (legacy_vault / "project/project_proj_b.md").unlink()
    (legacy_vault / "setting/list/list_project.md").write_text("# p\n\n| name | client |\n| --- | --- |\n", encoding="utf-8")
    assert run(legacy_vault, "bootstrap", "proj_b", "record/r1.md") == 2


# ---- 書き方の保証 -------------------------------------------------------------------


def test_crlf_record_keeps_crlf_when_a_project_is_added(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: x\n", "本文。\n", newline="\r\n")
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    raw = path.read_bytes()
    assert raw.count(b"\r\n") == raw.count(b"\n") > 0


def test_local_path_in_a_record_still_stops_the_write(legacy_vault: Path) -> None:
    record(legacy_vault, "r1", "title: x\n", "see /private/tmp/someone/x\n")
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 3
    assert snapshot(legacy_vault) == before


def test_two_records_are_written_all_or_nothing(legacy_vault: Path) -> None:
    record(legacy_vault, "ok", "title: x\n")
    record(legacy_vault, "bad", "project: proj_a\n", "> project: [project_proj_a](../x/y.md)\n\n本文。\n")
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_b", "record/ok.md", "record/bad.md", extra=["--apply"]) == 2
    assert snapshot(legacy_vault) == before


def test_backlinks_are_written_in_list_order_with_one_line_per_project(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: x\n")
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply"]) == 0
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    assert fm(path)["project"] == ["proj_b", "proj_a"]
    assert body_of(path).startswith(backlink("proj_b") + "\n" + backlink("proj_a") + "\n")
    assert body_of(path).count("> project:") == 2


# ---- レビュー指摘（PR #29）の回帰 -----------------------------------------------------------


@pytest.mark.parametrize(
    "frontmatter",
    [
        "project:\n  - proj_a\n  # 手書きの注意書き\n  - proj_c\nproject_source: manual\n",
        "project: proj_a  # 誰の案件\n",
    ],
    ids=["comment-inside-the-list", "trailing-comment"],
)
def test_yaml_comments_on_the_project_key_stop_the_write_instead_of_being_dropped(
    legacy_vault: Path, frontmatter: str
) -> None:
    record(legacy_vault, "r1", frontmatter, backlink("proj_a") + "\n" + backlink("proj_c") + "\n\n本文。\n" if "proj_c" in frontmatter else backlink("proj_a") + "\n\n本文。\n")
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply"]) == 2
    assert snapshot(legacy_vault) == before


def test_comments_on_other_keys_are_kept(legacy_vault: Path) -> None:
    path = record(legacy_vault, "r1", "title: x  # 題\n# 全体のメモ\n")
    assert run(legacy_vault, "backfill", "proj_a", "record/r1.md", extra=["--apply"]) == 0
    text = path.read_text(encoding="utf-8")
    assert "title: x  # 題" in text and "# 全体のメモ" in text


@pytest.mark.parametrize("key_form", ['"project"', "'project'"])
def test_quoted_project_key_is_replaced_not_duplicated(legacy_vault: Path, key_form: str) -> None:
    """PyYAML は重複キーを後勝ちで黙って読む。引用符つきのキーを見落として 2 つ目を書くと、validate が通るのに誤る。"""
    path = record(legacy_vault, "r1", f"{key_form}: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    assert run(legacy_vault, "backfill", "proj_b", "record/r1.md", extra=["--apply"]) == 0
    frontmatter_text = path.read_text(encoding="utf-8").split("---\n", 2)[1]
    keys = [l for l in frontmatter_text.splitlines() if re.match(r"^[\"']?project[\"']?\s*:", l)]
    assert len(keys) == 1, keys
    assert fm(path)["project"] == ["proj_a", "proj_b"]


def test_summary_with_several_records_is_refused(legacy_vault: Path) -> None:
    record(legacy_vault, "a", "title: a\n")
    record(legacy_vault, "b", "title: b\n")
    before = snapshot(legacy_vault)
    assert run(legacy_vault, "backfill", "proj_a", "record/a.md", "record/b.md", extra=["--apply", "--summary", "一件の要約"]) == 3
    assert snapshot(legacy_vault) == before  # 1 会議の要約を全 record に黙って書かない


@pytest.mark.parametrize("mode", ["validate", "bootstrap"])
def test_summary_is_backfill_only(legacy_vault: Path, mode: str) -> None:
    record(legacy_vault, "a", "project: proj_a\n", backlink("proj_a") + "\n\n本文。\n")
    assert run(legacy_vault, mode, "proj_a", "record/a.md", extra=["--summary", "x"]) == 3
