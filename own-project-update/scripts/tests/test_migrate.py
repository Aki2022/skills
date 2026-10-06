"""migrate: 既存 vault を新契約へ 1 回で移す（dry-run → 差分 → 人間 → 適用）。

SPEC-project-association の「配置と移行」と WS-20261003-project-association の ISSUE-03 の受入:
dry-run が対象件数を出し、適用後に render が no-op になり（= 3 出力が揃っている）、migrate の再実行が no-op。
情報を黙って落とさない: 表現できない内容は `## migrated notes` に残し、既存の summary は上書きしない。
"""

from __future__ import annotations

import os
import re
from pathlib import Path

import pytest
import yaml

import project_update as pu
from conftest import snapshot, write_note

PROJECT_A = """---
title: project_proj_a
project: proj_a
client: Client A
status: active
---

# proj_a

## overview

人が書いた本文。

## minutes

<!-- 新しい順。topics は 1 行 -->

| date | minutes | topics |
| ---- | ------- | ------ |
| 2026-08-20 | [r2](../record/r2.md) | 二回目の議題 |
| 2026-07-01 | [r1](../record/r1.md) | 初回の議題 |
| 2026-06-01 | [gone](../record/gone.md) | 消えた record の議題 |

> 2026-08-22 に r3 を r2 へ集約した。

## related notes

- [r4](../record/r4.md) — 他案件の会議でこの案件に触れた

## documents

- [local]()
- [deck](../presentation/deck.md)
- [site](https://example.com/x)
"""

PROJECT_B = """---
title: project_proj_b
project: proj_b
client: Client B
status: active
---

# proj_b

## overview

## minutes

| date | minutes | topics |
| ---- | ------- | ------ |
| 2026-05-01 | [r4](../record/r4.md) | B の議題 |
"""

LIST = """# list of projects

## list

| name | client | partner | status | last meeting | path |
| --- | --- | --- | --- | --- | --- |
| proj_a | Client A | 協力会社 | active | 2026-08-20 | [project\\_proj\\_a](../../project/project_proj_a.md) |
| proj_b | Client B |  | active | 2026-05-01 | [project_proj_b](../../project/project_proj_b.md) |
"""

TEMPLATE = """---
title:
project:
client:
client_aliases: []
status: active
started:
last_updated:
tags:
  - project
---

# {{project}}

## overview

- 案件の一言説明:

## minutes

<!-- 新しい順。 -->

| date | minutes | topics |
| ---- | ------- | ------ |
|      |         |        |

## related notes

<!-- 他クライアントの議事録 -->

## documents

- [local]()
"""


def backlink(key: str) -> str:
    return f"> project: [project_{key}](../project/project_{key}.md)"


@pytest.fixture
def old_vault(tmp_path: Path) -> Path:
    root = tmp_path / "vault"
    for directory in ("project", "record", "presentation", "setting/list", "setting/template"):
        (root / directory).mkdir(parents=True)
    (root / "project/project_proj_a.md").write_text(PROJECT_A, encoding="utf-8")
    (root / "project/project_proj_b.md").write_text(PROJECT_B, encoding="utf-8")
    (root / "setting/list/list_project.md").write_text(LIST, encoding="utf-8")
    (root / "setting/template/template_project.md").write_text(TEMPLATE, encoding="utf-8")
    write_note(root / "record/r1.md", "title: r1\ndate: 2026-07-01\nproject: proj_a\n", backlink("proj_a") + "\n\n本文1。\n")
    write_note(root / "record/r2.md", "title: r2\ndate: 2026-08-20\n", "本文2。\n")
    write_note(root / "record/r4.md", "title: r4\ndate: 2026-05-01\nproject: proj_b\n", backlink("proj_b") + "\n\n本文4。\n")
    write_note(root / "presentation/deck.md", "title: deck\nkind: presentation\ndate: 2026-09-01\n", "資料。\n")
    return root


def migrate(vault: Path, *extra: str) -> int:
    return pu.main(["migrate", "--vault", str(vault), *extra])


def plan_id(capsys: pytest.CaptureFixture[str]) -> str:
    match = re.search(r"^PLAN_ID (\w+)$", capsys.readouterr().out, re.M)
    assert match, "dry-run must print PLAN_ID"
    return match.group(1)


def apply_it(vault: Path, capsys: pytest.CaptureFixture[str], *extra: str) -> int:
    assert migrate(vault, *extra) == 0
    pid = plan_id(capsys)
    return migrate(vault, "--apply", "--plan-id", pid, *extra)


def fm(path: Path) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8").split("---\n", 2)[1])


def body(path: Path) -> str:
    return path.read_text(encoding="utf-8").split("---\n", 2)[2]


# ---- dry-run とプラン ID ----------------------------------------------------------------


def test_dry_run_writes_nothing_and_reports_counts_and_a_plan_id(old_vault: Path, capsys) -> None:
    before = snapshot(old_vault)
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert snapshot(old_vault) == before
    assert re.search(r"^PLAN_ID \w+$", out, re.M) and "DRY_RUN" in out
    for label in ("project_notes=2", "records_to_update=", "list_rows=2"):
        assert label in out


def test_apply_requires_the_plan_id_of_the_reviewed_dry_run(old_vault: Path, capsys) -> None:
    assert migrate(old_vault) == 0
    pid = plan_id(capsys)
    before = snapshot(old_vault)
    assert migrate(old_vault, "--apply") == 3  # plan id 無し
    assert migrate(old_vault, "--apply", "--plan-id", "deadbeef") == 2  # 見た差分と違う
    assert snapshot(old_vault) == before
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 0


def test_plan_id_changes_when_the_vault_changes_after_review(old_vault: Path, capsys) -> None:
    assert migrate(old_vault) == 0
    pid = plan_id(capsys)
    write_note(old_vault / "record/late.md", "title: late\n", "後から増えた。\n")
    (old_vault / "record/r1.md").write_text(
        (old_vault / "record/r1.md").read_text(encoding="utf-8").replace("本文1。", "本文1 改。"), encoding="utf-8"
    )
    # 差分の元になる record は変えていないが、project ノートを変えると計画が変わる
    note = old_vault / "project/project_proj_a.md"
    note.write_text(note.read_text(encoding="utf-8").replace("初回の議題", "初回の議題 改"), encoding="utf-8")
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 2


# ---- record: リスト化・legacy・summary -------------------------------------------------------


def test_records_get_list_project_legacy_source_summary_and_backlinks(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    r1, r2, r4 = (old_vault / f"record/{n}.md" for n in ("r1", "r2", "r4"))
    assert fm(r1)["project"] == ["proj_a"] and fm(r1)["project_source"] == "legacy" and fm(r1)["summary"] == "初回の議題"
    # minutes に載っている record は、その project に属する証拠として legacy で紐付く
    assert fm(r2)["project"] == ["proj_a"] and fm(r2)["project_source"] == "legacy" and fm(r2)["summary"] == "二回目の議題"
    assert body(r2).startswith(backlink("proj_a") + "\n")
    # related notes は「言及」であって所属ではない: 既定では紐付けない（既存の所属 proj_b のまま）
    assert fm(r4)["project"] == ["proj_b"] and fm(r4)["project_source"] == "legacy"
    assert fm(r4)["summary"] == "B の議題"  # 自分の project の minutes の topics が summary


def test_related_notes_are_associated_only_when_the_human_asks_for_it(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys, "--related-notes", "associate") == 0
    r4 = old_vault / "record/r4.md"
    assert fm(r4)["project"] == ["proj_b", "proj_a"]  # 既存の所属を保持して追加
    assert body(r4).startswith(backlink("proj_b") + "\n" + backlink("proj_a") + "\n")


def test_default_related_notes_do_not_move_last_meeting(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r5.md", "title: r5\ndate: 2026-09-30\nproject: proj_b\n", backlink("proj_b") + "\n\n本文5。\n")
    note = old_vault / "project/project_proj_a.md"
    note.write_text(note.read_text(encoding="utf-8").replace("- [r4](../record/r4.md)", "- [r5](../record/r5.md)"), encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    row = next(l for l in (old_vault / "setting/list/list_project.md").read_text(encoding="utf-8").splitlines() if l.startswith("| proj_a"))
    assert "2026-09-30" not in row  # 言及しただけの record の日付は proj_a の最終会議日にならない


def test_a_minutes_row_for_a_record_without_frontmatter_is_kept_not_dropped(old_vault: Path, capsys) -> None:
    (old_vault / "record/nofm.md").write_text("frontmatter の無いメモ。\n", encoding="utf-8")
    note = old_vault / "project/project_proj_a.md"
    note.write_text(note.read_text(encoding="utf-8").replace("| 2026-06-01 | [gone](../record/gone.md)", "| 2026-06-02 | [nofm](../record/nofm.md) | メモの議題 |\n| 2026-06-01 | [gone](../record/gone.md)"), encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    section = text.split("## migrated notes", 1)[1].split("<!-- generated:associated-notes begin -->", 1)[0]
    assert "メモの議題" in section  # 紐付けも summary も書けない record の行は、情報を落とさず残す
    assert (old_vault / "record/nofm.md").read_text(encoding="utf-8") == "frontmatter の無いメモ。\n"


def test_bare_bullets_and_empty_links_are_dropped_as_template_placeholders(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    note.write_text(note.read_text(encoding="utf-8") + "\n## related notes\n\n-\n\n## documents\n\n- [その他]()\n", encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    assert "migrated notes" not in note.read_text(encoding="utf-8")


def test_no_double_blank_line_before_the_generated_region(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    text = (old_vault / "project/project_proj_a.md").read_text(encoding="utf-8")
    assert "\n\n\n<!-- generated" not in text


def test_existing_summary_is_never_overwritten_and_is_reported(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r1.md", "title: r1\nproject: proj_a\nsummary: 手で書いた要約\n", backlink("proj_a") + "\n\n本文1。\n")
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert "SUMMARY_KEPT r1.md" in out
    pid = re.search(r"^PLAN_ID (\w+)$", out, re.M).group(1)
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 0
    assert fm(old_vault / "record/r1.md")["summary"] == "手で書いた要約"


def test_document_note_listed_under_documents_is_associated_as_legacy(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    deck = old_vault / "presentation/deck.md"
    assert fm(deck)["project"] == ["proj_a"] and fm(deck)["project_source"] == "legacy"
    assert body(deck).startswith(backlink("proj_a") + "\n")


# ---- project ノート: 旧 3 節 → generated 領域、残すもの -----------------------------------------


def test_old_sections_become_one_generated_region_and_other_text_is_kept(old_vault: Path, capsys) -> None:
    before = (old_vault / "project/project_proj_a.md").read_text(encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    text = (old_vault / "project/project_proj_a.md").read_text(encoding="utf-8")
    for heading in ("## minutes", "## related notes", "## documents"):
        assert heading not in text
    assert text.count("<!-- generated:associated-notes begin -->") == 1
    assert "人が書いた本文。" in text and "## overview" in text
    region = text.split("<!-- generated:associated-notes begin -->", 1)[1]
    for name in ("r1.md", "r2.md", "deck.md"):  # r4 は related notes（言及）なので既定では領域に載せない
        assert f"../record/{name}" in region or f"../presentation/{name}" in region
    assert "../record/r4.md" not in region and "../record/r4.md" in text  # 言及は原文のまま migrated notes に残る
    assert before.split("---\n", 2)[2].split("## minutes")[0] in text  # frontmatter 以外で minutes より前は byte-identical


def test_content_that_cannot_be_represented_is_preserved_in_migrated_notes(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    text = (old_vault / "project/project_proj_a.md").read_text(encoding="utf-8")
    assert "## migrated notes" in text
    section = text.split("## migrated notes", 1)[1].split("<!-- generated:associated-notes begin -->", 1)[0]
    assert "2026-08-22 に r3 を r2 へ集約した。" in section  # 引用メモ
    assert "他案件の会議でこの案件に触れた" in section  # related notes の説明文
    assert "https://example.com/x" in section  # 外部リンク
    assert "消えた record の議題" in section  # 実体の無い record の行
    assert "[local]()" not in text  # 空リンクのテンプレ placeholder だけは捨てる（情報が無い）


def test_partner_moves_from_the_list_into_frontmatter(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "project/project_proj_a.md")["partner"] == "協力会社"
    assert "partner" not in fm(old_vault / "project/project_proj_b.md")


def test_template_loses_the_old_sections_and_gains_the_new_keys(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    text = (old_vault / "setting/template/template_project.md").read_text(encoding="utf-8")
    for heading in ("## minutes", "## related notes", "## documents"):
        assert heading not in text
    assert "partner:" in text and "scope:" in text
    assert "<!-- generated:associated-notes begin -->" in text and "## overview" in text


# ---- 適用後の状態: render が no-op・再実行が no-op ------------------------------------------------


def test_after_apply_render_is_a_noop_and_migrate_is_a_noop(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    capsys.readouterr()
    first = snapshot(old_vault)
    assert pu.main(["render", "--vault", str(old_vault), "--apply"]) == 0
    assert snapshot(old_vault) == first  # 3 出力が揃っている
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out
    assert migrate(old_vault, "--apply", "--plan-id", "anything") == 0  # 変更が無ければ plan id は要らない
    assert snapshot(old_vault) == first


def test_list_project_is_regenerated_with_partner_and_last_meeting(old_vault: Path, capsys) -> None:
    assert apply_it(old_vault, capsys) == 0
    text = (old_vault / "setting/list/list_project.md").read_text(encoding="utf-8")
    assert "協力会社" in text and "2026-08-20" in text


# ---- 安全 -------------------------------------------------------------------------------------


def test_failure_midway_restores_everything(old_vault: Path, capsys, monkeypatch: pytest.MonkeyPatch) -> None:
    assert migrate(old_vault) == 0
    pid = plan_id(capsys)
    before = snapshot(old_vault)
    real_replace = os.replace
    calls = {"n": 0}

    def flaky(src, dst, *a, **k):
        calls["n"] += 1
        if calls["n"] == 5:
            raise OSError("simulated Drive I/O failure")
        return real_replace(src, dst, *a, **k)

    monkeypatch.setattr(os, "replace", flaky)
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 3
    monkeypatch.undo()
    assert snapshot(old_vault) == before


def test_a_record_with_a_backlink_but_no_project_is_repaired_and_reported(old_vault: Path, capsys) -> None:
    """実 vault に、backlink はあるのに frontmatter の project が無い record がある。補修して dry-run に必ず出す。"""
    write_note(old_vault / "record/r1.md", "title: r1\ndate: 2026-07-01\n", backlink("proj_a") + "\n\n本文1。\n")
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert "REPAIRED_RECORD r1.md" in out
    pid = re.search(r"^PLAN_ID (\w+)$", out, re.M).group(1)
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 0
    assert fm(old_vault / "record/r1.md")["project"] == ["proj_a"] and fm(old_vault / "record/r1.md")["project_source"] == "legacy"


def test_a_record_whose_backlink_names_another_project_gets_both_as_legacy_and_is_reported(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r1.md", "title: r1\nproject: proj_a\n", backlink("proj_b") + "\n\n本文。\n")
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert "REPAIRED_RECORD r1.md" in out
    pid = re.search(r"^PLAN_ID (\w+)$", out, re.M).group(1)
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 0
    assert fm(old_vault / "record/r1.md")["project"] == ["proj_a", "proj_b"]


def test_a_malformed_backlink_still_stops_the_whole_migration(old_vault: Path) -> None:
    write_note(old_vault / "record/r1.md", "title: r1\nproject: proj_a\n", "> project: [project_proj_a](../elsewhere/x.md)\n\n本文。\n")
    before = snapshot(old_vault)
    assert migrate(old_vault) == 2
    assert snapshot(old_vault) == before


def test_diff_dir_receives_one_diff_per_changed_file_outside_the_vault(old_vault: Path, tmp_path: Path) -> None:
    out = tmp_path / "diffs"
    assert migrate(old_vault, "--diff-dir", str(out)) == 0
    names = sorted(p.name for p in out.rglob("*.diff"))
    assert "project_proj_a.md.diff" in names and "r2.md.diff" in names and "template_project.md.diff" in names
    assert (out / "SUMMARY.md").is_file()
    assert not list(old_vault.rglob("*.diff"))


def test_a_list_that_render_cannot_read_back_stops_the_migration_instead_of_overwriting_it(tmp_path: Path) -> None:
    root = tmp_path / "vault"
    for directory in ("project", "record", "setting/list"):
        (root / directory).mkdir(parents=True)
    (root / "project/project_p.md").write_text(
        "---\ntitle: project_p\nproject: p\nclient: C\nstatus: active\n---\n\n# p\n\n## overview\n", encoding="utf-8"
    )
    (root / "setting/list/list_project.md").write_text("手書きのメモだけ\n", encoding="utf-8")
    before = snapshot(root)
    assert migrate(root) == 2  # 情報を黙って落とさない（render と同じ fail-closed）
    assert snapshot(root) == before


def test_a_backlink_with_an_escaped_underscore_label_is_read_and_rewritten_in_canonical_form(old_vault: Path, capsys) -> None:
    """実 vault の 4 件は `[project\\_東芝](…)` のように Markdown のエスケープが付いている。意味は同じなので壊れた扱いにしない。"""
    escaped = "> project: [project\\_proj_a](../project/project_proj_a.md)"
    write_note(old_vault / "record/r1.md", "title: r1\nproject: proj_a\n", escaped + "\n\n本文1。\n")
    assert apply_it(old_vault, capsys) == 0
    assert body(old_vault / "record/r1.md").startswith(backlink("proj_a") + "\n")  # 正規形に揃う
    assert pu.main(["--vault", str(old_vault), "--mode", "validate", "--project-key", "proj_a", "--record", "record/r1.md"]) in (0, 3)


def test_a_backlink_in_the_middle_of_the_body_is_moved_to_the_top_not_duplicated(old_vault: Path, capsys) -> None:
    """実 vault の record に、backlink が本文の途中にあるものがある。先頭に足して元を残すと二重になる。"""
    write_note(
        old_vault / "record/r1.md",
        "title: r1\ndate: 2026-07-01\n",
        "# 冒頭の見出し\n\n前の文。\n\n" + backlink("proj_a") + "\n\n後の文。\n",
    )
    assert apply_it(old_vault, capsys) == 0
    text = body(old_vault / "record/r1.md")
    assert text.startswith(backlink("proj_a") + "\n")
    assert text.count("> project:") == 1
    assert "前の文。" in text and "後の文。" in text and "# 冒頭の見出し" in text
    capsys.readouterr()
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out  # 再実行は no-op


def test_validate_accepts_a_project_note_after_migration(old_vault: Path, capsys) -> None:
    """移行後の project ノートには旧 3 節が無い。旧来の validate が必須見出しとして拒否してはいけない。"""
    assert apply_it(old_vault, capsys) == 0
    legacy_headings = """## people

## stakeholders

## next actions

## open issues / risks
"""
    note = old_vault / "project/project_proj_a.md"
    note.write_text(note.read_text(encoding="utf-8").replace("## overview", "## overview\n\n" + legacy_headings, 1), encoding="utf-8")
    assert pu.main(["--vault", str(old_vault), "--mode", "validate", "--project-key", "proj_a", "--record", "record/r1.md"]) == 0
