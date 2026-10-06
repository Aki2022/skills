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
    note.write_text(note.read_text(encoding="utf-8") + "\n## related notes\n\n-\n\n## documents\n\n- [local]()\n", encoding="utf-8")
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


# ---- 独立レビュー（PR #31）の指摘: 黙って落とさない -------------------------------------------------


def migrated_notes(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    assert "## migrated notes" in text, text
    return text.split("## migrated notes", 1)[1].split("<!-- generated:associated-notes begin -->", 1)[0]


def edit(path: Path, old: str, new: str) -> None:
    text = path.read_text(encoding="utf-8")
    assert old in text
    path.write_text(text.replace(old, new, 1), encoding="utf-8")


def test_a_human_html_comment_in_an_old_section_is_kept_but_the_templates_own_comment_is_dropped(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## minutes\n\n<!-- 新しい順。 -->\n<!-- 重要: 田中さんの確認待ち。売上 1,000 万円 -->\n<!-- 複数行\n売上の注記 -->後ろの文\n")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "田中さんの確認待ち" in kept and "売上の注記" in kept and "後ろの文" in kept
    assert "新しい順" not in kept  # テンプレ自身の注釈コメントは placeholder


def test_a_fenced_block_in_an_old_section_is_kept_verbatim_and_not_parsed_as_a_table(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(
        note,
        "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n",
        "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n\n```\n| 2026-07-01 | [r1](../record/r1.md) | in fence |\n\n## documents\ncode here\n```\n",
    )
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "```\n| 2026-07-01 | [r1](../record/r1.md) | in fence |\n\n## documents\ncode here\n```" in kept
    assert "summary" not in fm(old_vault / "record/r1.md") or fm(old_vault / "record/r1.md")["summary"] != "in fence"
    assert "proj_b" not in fm(old_vault / "record/r1.md")["project"]


def test_every_minutes_datum_survives_even_when_it_cannot_become_the_summary(old_vault: Path, capsys) -> None:
    """2 つ目の topic・余分な列・record の日付と違う表の日付・手書きの summary と食い違う topic は、行ごと残す。"""
    pa = old_vault / "project/project_proj_a.md"
    edit(pa, "| 2026-07-01 | [r1](../record/r1.md) | 初回の議題 |", "| 2026-07-02 | [r1](../record/r1.md) | 初回の議題 | 担当: alice |\n| 2026-07-01 | [r1](../record/r1.md) | 初回の別の議題 |")
    pb = old_vault / "project/project_proj_b.md"
    edit(pb, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n| 2026-05-01 | [r1](../record/r1.md) | B 側から見た議題 |")
    write_note(old_vault / "record/r2.md", "title: r2\ndate: 2026-08-20\nsummary: 手で書いた要約\n", "本文2。\n")
    assert apply_it(old_vault, capsys) == 0
    kept_a = migrated_notes(pa)
    assert "担当: alice" in kept_a and "初回の別の議題" in kept_a
    assert "二回目の議題" in kept_a  # r2 の手書き summary と違うので topic を行ごと残す
    assert "B 側から見た議題" in migrated_notes(pb)
    assert fm(old_vault / "record/r2.md")["summary"] == "手で書いた要約"


def test_blank_lines_in_preserved_content_are_kept_so_paragraphs_do_not_merge(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## minutes\n\n> a\n\n> b\n\n段落その 1。\n\n段落その 2。\n")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "> a\n\n> b\n\n段落その 1。\n\n段落その 2。" in kept
    assert "\n\n\n" not in kept


def test_an_existing_generated_region_inside_an_old_section_is_regenerated_not_copied_as_notes(old_vault: Path, capsys) -> None:
    assert pu.main(["render", "--vault", str(old_vault), "--accept-list-changes", "--apply"]) == 0  # 旧 vault に render を先にかけた状態
    capsys.readouterr()
    note = old_vault / "project/project_proj_a.md"
    assert "generated:associated-notes begin" in note.read_text(encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    assert text.count("generated:associated-notes begin") == 1
    assert "| kind | date | note | summary |" not in migrated_notes(note)


def test_a_lone_region_marker_in_an_old_section_stops_the_migration(old_vault: Path) -> None:
    note = old_vault / "project/project_proj_a.md"
    edit(note, "## documents\n", "## documents\n\n<!-- generated:associated-notes begin -->\n")
    before = snapshot(old_vault)
    assert migrate(old_vault) == 2
    assert snapshot(old_vault) == before


@pytest.mark.parametrize("heading", ["## minutes (2026)", "### minutes", "## Minutes"])
def test_heading_variants_of_the_old_sections_are_migrated_too(old_vault: Path, capsys, heading: str) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", f"{heading}\n")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    assert "| date | minutes | topics |" not in text
    assert "proj_b" in fm(old_vault / "record/r4.md")["project"]


def test_a_scalar_project_is_converted_to_a_list_even_when_project_source_exists(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r4.md", "title: r4\ndate: 2026-05-01\nproject: proj_b\nproject_source: manual\n", backlink("proj_b") + "\n\n本文4。\n")
    assert apply_it(old_vault, capsys) == 0
    data = fm(old_vault / "record/r4.md")
    assert data["project"] == ["proj_b"] and data["project_source"] == "manual"


def test_notes_outside_the_top_level_record_directory_are_converted_too(old_vault: Path, capsys) -> None:
    write_note(old_vault / "presentation/old.md", "title: old\nkind: presentation\nproject: proj_a\n", backlink("proj_a") + "\n\n資料。\n")
    write_note(old_vault / "record/sub/deep.md", "title: deep\nproject: proj_b\n", backlink("proj_b") + "\n\n深い。\n")
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "presentation/old.md")["project"] == ["proj_a"]
    assert fm(old_vault / "record/sub/deep.md")["project"] == ["proj_b"]
    capsys.readouterr()
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out


def test_a_record_with_a_bom_can_be_applied_and_keeps_its_bom(old_vault: Path, capsys) -> None:
    path = old_vault / "record/r2.md"
    path.write_bytes(b"\xef\xbb\xbf" + path.read_bytes())
    assert apply_it(old_vault, capsys) == 0
    assert path.read_bytes().startswith(b"\xef\xbb\xbf---")
    assert fm_bom(path)["project"] == ["proj_a"]


def fm_bom(path: Path) -> dict:
    return yaml.safe_load(path.read_text(encoding="utf-8").lstrip("﻿").split("---\n", 2)[1])


def test_a_bom_project_note_is_refused_with_a_message_that_names_the_bom(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    note.write_bytes(b"\xef\xbb\xbf" + note.read_bytes())
    before = snapshot(old_vault)
    assert migrate(old_vault) in (2, 3)
    assert "BOM" in capsys.readouterr().err
    assert snapshot(old_vault) == before


@pytest.mark.parametrize("empty", ["partner:", "partner: ''", "partner: null"])
def test_partner_is_moved_from_the_list_when_the_frontmatter_key_exists_but_is_empty(old_vault: Path, capsys, empty: str) -> None:
    edit(old_vault / "project/project_proj_a.md", "status: active\n", f"status: active\n{empty}\n")
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "project/project_proj_a.md")["partner"] == "協力会社"


def test_a_diff_dir_inside_the_vault_is_refused_and_nothing_is_written(old_vault: Path) -> None:
    before = snapshot(old_vault)
    assert migrate(old_vault, "--diff-dir", str(old_vault / "record" / "diffs")) in (2, 3)
    assert snapshot(old_vault) == before


def test_stale_diffs_are_removed_when_the_plan_changes_and_a_foreign_directory_is_refused(old_vault: Path, tmp_path: Path, capsys) -> None:
    out = tmp_path / "diffs"
    assert apply_it(old_vault, capsys, "--diff-dir", str(out)) == 0
    assert list(out.rglob("*.diff"))
    capsys.readouterr()
    assert migrate(old_vault, "--diff-dir", str(out)) == 0  # NOOP
    assert not list(out.rglob("*.diff"))  # 古い diff を残さない
    foreign = tmp_path / "foreign"
    foreign.mkdir()
    (foreign / "keep.txt").write_text("人の持ち物\n", encoding="utf-8")
    assert migrate(old_vault, "--diff-dir", str(foreign)) in (2, 3)
    assert (foreign / "keep.txt").exists()


def test_a_newline_only_change_shows_up_in_the_diff(old_vault: Path, tmp_path: Path) -> None:
    path = old_vault / "record/r2.md"
    path.write_bytes(path.read_bytes().replace(b"\n", b"\r\n"))
    out = tmp_path / "diffs"
    assert migrate(old_vault, "--diff-dir", str(out)) == 0
    assert (out / "record/r2.md.diff").read_text(encoding="utf-8").strip()


def test_a_template_with_a_note_style_comment_is_rewritten_and_the_dropped_comments_are_reported(old_vault: Path, capsys) -> None:
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert "TEMPLATE_COMMENTS_DROPPED" in out and "新しい順" in out


def test_an_existing_migrated_notes_section_is_extended_not_duplicated(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## migrated notes\n\n### from minutes\n\n> 前の移行で残したメモ\n\n## minutes\n\n> 新しく見つけたメモ\n\n")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    assert text.count("## migrated notes") == 1
    assert "前の移行で残したメモ" in text and "新しく見つけたメモ" in text
    capsys.readouterr()
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out


@pytest.mark.parametrize(
    "row",
    ["| 2026-05-01 | [r4](../record/r4.md) |", "| 2026-05-01 | [r4](<../record/r4.md>) | 角括弧形 |", "| 2026-05-01 | [r4](../record/r4.md#anchor) | アンカー |", "| 2026-05-01 | [r4](../record/r4.md?x=1) | クエリ |"],
)
def test_minutes_links_with_an_anchor_a_query_or_angle_brackets_still_associate(old_vault: Path, capsys, row: str) -> None:
    write_note(old_vault / "record/r4.md", "title: r4\ndate: 2026-05-01\n", "本文4。\n")
    edit(old_vault / "project/project_proj_b.md", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", row)
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "record/r4.md")["project"] == ["proj_b"]


def test_a_filename_with_ascii_parentheses_associates(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r5 (1).md", "title: r5\ndate: 2026-05-02\n", "本文5。\n")
    edit(old_vault / "project/project_proj_b.md", "| 2026-05-01 |", "| 2026-05-02 | [r5](<../record/r5 (1).md>) | 括弧つき |\n| 2026-05-01 |")
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "record/r5 (1).md")["project"] == ["proj_b"]


def test_a_backlink_with_several_escaped_underscores_is_read(tmp_path: Path, capsys, old_vault: Path) -> None:
    for name in ("proj_x_y", ):
        (old_vault / f"project/project_{name}.md").write_text(PROJECT_B.replace("proj_b", name), encoding="utf-8")
    escaped = "> project: [project\\_proj\\_x\\_y](../project/project_proj_x_y.md)"
    write_note(old_vault / "record/r6.md", "title: r6\nproject: proj_x_y\n", escaped + "\n\n本文6。\n")
    assert apply_it(old_vault, capsys) == 0
    assert body(old_vault / "record/r6.md").startswith("> project: [project_proj_x_y](../project/project_proj_x_y.md)\n")


def test_a_symlinked_record_directory_does_not_crash(old_vault: Path, tmp_path: Path) -> None:
    outside = tmp_path / "outside_records"
    outside.mkdir()
    write_note(outside / "x.md", "title: x\nproject: proj_a\n", backlink("proj_a") + "\n\n外。\n")
    (old_vault / "record/linked").symlink_to(outside, target_is_directory=True)
    code = migrate(old_vault)
    assert code in (0, 2, 3)  # traceback（exit 1）にしない


# ---- 2 回目の独立レビューの指摘: 行の一部だけが表現できるときは行ごと残す ----------------------------------


def test_a_documents_line_with_an_empty_link_and_other_content_is_kept(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- 旧: [local]() / 新: [deck](https://example.com/d) 重要\n- [local]() 補足: 田中さんへ共有済み\n- [local]()\n\n## minutes\n")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "https://example.com/d" in kept and "田中さんへ共有済み" in kept
    assert kept.count("[local]()") == 2  # 純粋な空リンクだけが placeholder


def test_a_minutes_cell_with_two_links_or_surrounding_text_keeps_the_row_and_associates_every_record(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r3.md", "title: r3\ndate: 2026-05-02\n", "本文3。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) / [r3](../record/r3.md) (draft, 田中さん欠席) | B の議題 |")
    assert apply_it(old_vault, capsys) == 0
    assert "(draft, 田中さん欠席)" in migrated_notes(note)
    assert fm(old_vault / "record/r3.md")["project"] == ["proj_b"] and fm(old_vault / "record/r4.md")["project"] == ["proj_b"]


def test_text_before_or_after_the_link_of_a_documents_item_is_kept(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", '## documents\n\n- 最新版: [deck](../presentation/deck.md "表題")\n- 1. [deck](../presentation/deck.md)\n\n## minutes\n')
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "最新版:" in kept and '"表題"' in kept and "1. [deck]" in kept
    assert "proj_b" in fm(old_vault / "presentation/deck.md")["project"]


def test_a_link_label_that_is_not_the_filename_or_title_is_kept(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "[r4](../record/r4.md)", "[Kickoff w/ Tanaka](../record/r4.md)")
    assert apply_it(old_vault, capsys) == 0
    assert "Kickoff w/ Tanaka" in migrated_notes(note)


def test_a_preserved_table_keeps_its_header_and_separator(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n| name | link |\n| --- | --- |\n| 見積 | [x](https://example.com/x) |\n\n## minutes\n")
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 | 担当: bob |")
    edit(note, "| date | minutes | topics |\n| ---- | ------- | ------ |", "| date | minutes | topics | owner |\n| ---- | ------- | ------ | ----- |")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "| name | link |\n| --- | --- |\n| 見積 |" in kept
    assert "| date | minutes | topics | owner |\n| ---- | ------- | ------ | ----- |\n| 2026-05-01 |" in kept


def test_a_link_into_a_hidden_directory_is_not_associated_and_the_row_is_kept(old_vault: Path, capsys) -> None:
    write_note(old_vault / ".trash/r9.md", "title: r9\n", "ごみ箱。\n")
    before = (old_vault / ".trash/r9.md").read_text(encoding="utf-8")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n| 2026-04-01 | [r9](../.trash/r9.md) | ごみ箱の議題 |")
    assert apply_it(old_vault, capsys) == 0
    assert "ごみ箱の議題" in migrated_notes(note)
    assert (old_vault / ".trash/r9.md").read_text(encoding="utf-8") == before


def test_accept_list_changes_lists_what_it_drops_in_the_dry_run(old_vault: Path, capsys) -> None:
    edit(old_vault / "project/project_proj_a.md", "status: active\n", "status: active\npartner: 別会社\n")
    assert migrate(old_vault) == 2  # 既定では止まる
    capsys.readouterr()
    assert migrate(old_vault, "--accept-list-changes") == 0
    out = capsys.readouterr().out
    assert "LIST_VALUE_DROPPED" in out and "協力会社" in out


def test_a_diff_dir_with_a_foreign_summary_file_is_refused_and_foreign_diffs_survive(old_vault: Path, tmp_path: Path) -> None:
    foreign = tmp_path / "book"
    foreign.mkdir()
    (foreign / "SUMMARY.md").write_text("# Summary\n\n- [章](ch1.md)\n", encoding="utf-8")
    (foreign / "patch.diff").write_text("人の patch\n", encoding="utf-8")
    assert migrate(old_vault, "--diff-dir", str(foreign)) in (2, 3)
    assert (foreign / "patch.diff").exists()


def test_only_the_diffs_the_previous_dry_run_wrote_are_removed(old_vault: Path, tmp_path: Path) -> None:
    out = tmp_path / "diffs"
    assert migrate(old_vault, "--diff-dir", str(out)) == 0
    (out / "mine.diff").write_text("人の diff\n", encoding="utf-8")
    assert migrate(old_vault, "--diff-dir", str(out)) == 0
    assert (out / "mine.diff").exists()


def test_a_backlink_example_inside_a_code_fence_is_not_a_membership(old_vault: Path, capsys) -> None:
    fence = "```\n" + backlink("proj_b") + "\n```\n"
    write_note(old_vault / "record/r2.md", "title: r2\ndate: 2026-08-20\n", "本文2。\n\n" + fence)
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "record/r2.md")["project"] == ["proj_a"]
    assert fence in body(old_vault / "record/r2.md")


def test_rows_in_an_indented_code_fence_are_not_parsed(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    sample = "- 例:\n\n    ```\n    | 2026-07-07 | [r7](../record/r7.md) | sample |\n    ```\n"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n\n" + sample)
    assert apply_it(old_vault, capsys) == 0
    assert "    | 2026-07-07 | [r7](../record/r7.md) | sample |" in migrated_notes(note)
    assert "project" not in fm(old_vault / "record/r7.md")


def test_a_mid_body_backlink_of_an_already_converted_record_is_moved_to_the_top(old_vault: Path, capsys) -> None:
    write_note(
        old_vault / "record/r4.md",
        "title: r4\ndate: 2026-05-01\nproject:\n  - proj_b\nproject_source: manual\nsummary: 手書き\n",
        "# 冒頭\n\n" + backlink("proj_b") + "\n\n本文4。\n",
    )
    assert apply_it(old_vault, capsys) == 0
    assert body(old_vault / "record/r4.md").startswith(backlink("proj_b") + "\n")


def test_an_old_section_nested_under_another_old_section_is_migrated_and_rerun_is_a_noop(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [deck](../presentation/deck.md)\n\n### minutes\n")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    assert "### minutes" not in text and "| date | minutes | topics |" not in text
    assert fm(old_vault / "record/r4.md")["project"] == ["proj_b"]
    assert "proj_b" in fm(old_vault / "presentation/deck.md")["project"]
    capsys.readouterr()
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out


@pytest.mark.parametrize("heading", ["## minutes ##", "   ## minutes"])
def test_closing_hashes_and_small_indentation_of_an_old_heading_are_recognised(old_vault: Path, capsys, heading: str) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", f"{heading}\n")
    assert apply_it(old_vault, capsys) == 0
    assert "| date | minutes | topics |" not in note.read_text(encoding="utf-8")


def test_a_level_one_heading_ends_an_old_section(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    note.write_text(note.read_text(encoding="utf-8") + "\n# 付録\n\n付録の本文。\n", encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    assert "# 付録\n\n付録の本文。" in text
    assert "migrated notes" not in text  # 付録は旧節の一部ではない


def test_a_migrated_notes_heading_nested_in_an_old_section_is_not_taken_as_the_target(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| date | minutes | topics |", "### migrated notes\n\n> 入れ子のメモ\n\n| date | minutes | topics |")
    assert apply_it(old_vault, capsys) == 0
    text = note.read_text(encoding="utf-8")
    assert "入れ子のメモ" in text
    capsys.readouterr()
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out


@pytest.mark.parametrize("cell", ["[r4](../record/r4.md) (draft, 田中さん欠席)", "(draft) [r4](../record/r4.md)", '[r4](../record/r4.md "表題")'])
def test_a_single_link_with_surrounding_text_or_a_title_in_the_minutes_cell_keeps_the_row(old_vault: Path, capsys, cell: str) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "[r4](../record/r4.md)", cell)
    assert apply_it(old_vault, capsys) == 0
    assert cell in migrated_notes(note)
    assert fm(old_vault / "record/r4.md")["project"] == ["proj_b"]


# ---- 3 回目の独立レビューの指摘 ----------------------------------------------------------------------


def test_a_data_row_followed_by_a_dash_row_is_not_swallowed_as_a_header(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-07-07 | [r7](../record/r7.md) | UNIQUE-7 |\n| --- | --- | --- |\n| 2026-05-01 | [r4](../record/r4.md) | B の議題 |")
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "record/r7.md")["project"] == ["proj_b"]
    assert fm(old_vault / "record/r7.md")["summary"] == "UNIQUE-7"


def test_a_custom_header_of_a_minutes_table_is_kept_even_when_every_row_is_pure(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| date | minutes | topics |\n| ---- | ------- | ------ |", "| 担当 | 期限 |\n| --- | --- |\n| 田中 | 来週 |\n\n| date | minutes | topics |\n| ---- | ------- | ------ |")
    assert apply_it(old_vault, capsys) == 0
    assert "| 担当 | 期限 |\n| --- | --- |\n| 田中 | 来週 |" in migrated_notes(note)


def test_a_minutes_table_with_another_column_order_is_read_by_header_name(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| date | minutes | topics |\n| ---- | ------- | ------ |\n| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| date | topics | minutes |\n| ---- | ------ | ------- |\n| 2026-07-07 | 順序違いの議題 | [r7](../record/r7.md) |")
    assert apply_it(old_vault, capsys) == 0
    assert fm(old_vault / "record/r7.md")["summary"] == "順序違いの議題"


def test_a_minutes_table_with_an_unknown_header_is_kept_verbatim_and_nothing_is_associated(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| date | minutes | topics |\n| ---- | ------- | ------ |\n| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 日付 | 資料 | 備考 |\n| --- | --- | --- |\n| 2026-07-07 | [r7](../record/r7.md) | 備考 |")
    assert apply_it(old_vault, capsys) == 0
    assert "| 2026-07-07 | [r7](../record/r7.md) | 備考 |" in migrated_notes(note)


@pytest.mark.parametrize("key", ["foo (bar)", "a]b", "x | y", "a/b"])
def test_a_project_key_that_cannot_be_written_as_a_backlink_is_never_written(old_vault: Path, capsys, key: str) -> None:
    write_note(old_vault / "record/r8.md", f"title: r8\nproject: '{key}'\n", "本文8。\n")
    before = (old_vault / "record/r8.md").read_text(encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0  # 対応する project ノートが無いので書き換えず報告する
    assert (old_vault / "record/r8.md").read_text(encoding="utf-8") == before
    capsys.readouterr()
    assert migrate(old_vault) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out


def test_a_project_note_whose_key_cannot_be_a_backlink_stops_the_dry_run(old_vault: Path) -> None:
    """key が既存の project ノートのものなら、書くことになるので止める（render も同じ key を拒否する）。"""
    (old_vault / "project/project_a (b).md").write_text(PROJECT_B.replace("proj_b", "a (b)"), encoding="utf-8")
    before = snapshot(old_vault)
    assert migrate(old_vault) in (2, 3)
    assert snapshot(old_vault) == before


def test_the_rows_of_a_record_with_an_unknown_project_are_kept_not_consumed(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r9.md", "title: r9\ndate: 2026-07-09\nproject: ghost\n", "本文9。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n| 2026-07-09 | [r9](../record/r9.md) | 幽霊の議題 |")
    edit(note, "## minutes\n", "## documents\n\n- [r9](../record/r9.md)\n\n## minutes\n")
    before = (old_vault / "record/r9.md").read_text(encoding="utf-8")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "| 2026-07-09 | [r9](../record/r9.md) | 幽霊の議題 |" in kept and "- [r9](../record/r9.md)" in kept
    assert (old_vault / "record/r9.md").read_text(encoding="utf-8") == before


def test_a_minutes_table_with_an_unrecognised_header_is_reported(old_vault: Path, capsys) -> None:
    edit(old_vault / "project/project_proj_b.md", "| date | minutes | topics |\n| ---- | ------- | ------ |", "| 日付 | 議事録 | 議題 |\n| --- | --- | --- |")
    assert migrate(old_vault) == 0
    assert "UNRECOGNIZED_TABLE proj_b" in capsys.readouterr().out


def test_an_empty_link_with_a_label_the_template_does_not_use_is_kept(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [local]()\n- [契約書の保管場所]()\n\n## minutes\n")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "- [契約書の保管場所]()" in kept and "[local]()" not in kept


def test_a_backlink_report_is_not_emitted_for_a_record_that_is_left_alone(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r9.md", "title: r9\nproject: ghost\n", "本文9。\n")
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert "UNKNOWN_PROJECT" in out and "ADDED_BACKLINK r9.md" not in out


def test_a_note_whose_project_has_no_project_note_is_left_alone_and_reported(old_vault: Path, capsys) -> None:
    write_note(old_vault / "other/memo.md", "title: memo\nproject: プロジェクト管理の進め方について\n", "メモ。\n")
    before = (old_vault / "other/memo.md").read_text(encoding="utf-8")
    assert migrate(old_vault) == 0
    out = capsys.readouterr().out
    assert "UNKNOWN_PROJECT other/memo.md" in out and "CHANGE other/memo.md" not in out
    assert apply_it(old_vault, capsys) == 0
    assert (old_vault / "other/memo.md").read_text(encoding="utf-8") == before


@pytest.mark.parametrize("target", ["../presentation/deck.md#決定事項", "../presentation/deck.md?view=x"])
def test_an_anchor_or_query_in_a_pure_item_keeps_the_line(old_vault: Path, capsys, target: str) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", f"## documents\n\n- [deck]({target})\n\n## minutes\n")
    assert apply_it(old_vault, capsys) == 0
    assert f"[deck]({target})" in migrated_notes(note)


def test_children_and_continuation_lines_keep_their_parent_item(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [deck](../presentation/deck.md)\n  - 版: v3、提出済\n  続きの説明文\n\n## minutes\n")
    assert apply_it(old_vault, capsys) == 0
    assert "- [deck](../presentation/deck.md)\n  - 版: v3、提出済\n  続きの説明文" in migrated_notes(note)


def test_a_documents_table_row_and_a_numbered_item_with_a_record_link_associate_and_are_kept(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    write_note(old_vault / "record/r8.md", "title: r8\ndate: 2026-07-08\n", "本文8。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n| name | link |\n| --- | --- |\n| 議事 | [r7](../record/r7.md) |\n\n1. [r8](../record/r8.md)\n\n## minutes\n")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "| 議事 | [r7](../record/r7.md) |" in kept
    assert "r8" not in kept  # `1. [r8](…)` だけの項目は純粋なので消える（紐付けは残る）
    assert fm(old_vault / "record/r7.md")["project"] == ["proj_b"] and fm(old_vault / "record/r8.md")["project"] == ["proj_b"]


def test_previous_diff_removal_never_leaves_the_diff_directory(old_vault: Path, tmp_path: Path) -> None:
    out = tmp_path / "diffs"
    assert migrate(old_vault, "--diff-dir", str(out)) == 0
    victim_dir = tmp_path / "outside"
    victim_dir.mkdir()
    (victim_dir / "victim.md.diff").write_text("外の diff\n", encoding="utf-8")
    (out / "evil").symlink_to(victim_dir, target_is_directory=True)
    summary = out / "SUMMARY.md"
    summary.write_text(summary.read_text(encoding="utf-8").replace("## files\n", "## files\n\n- evil/victim.md\n", 1), encoding="utf-8")
    assert migrate(old_vault, "--diff-dir", str(out)) == 0
    assert (victim_dir / "victim.md.diff").exists()


# レビューで見つかったテストの穴（変異しても緑だったもの）


def test_a_row_whose_table_date_differs_from_the_record_date_is_kept(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n| 2026-07-09 | [r7](../record/r7.md) | 日付違い |")
    assert apply_it(old_vault, capsys) == 0
    assert "| 2026-07-09 | [r7](../record/r7.md) | 日付違い |" in migrated_notes(note)


def test_text_after_a_known_comment_and_an_unclosed_comment_are_kept(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## minutes\n\n<!--  新しい順。  --> 人が書いたメモ\n<!-- 閉じていない手書きの注意\n")
    assert apply_it(old_vault, capsys) == 0
    kept = migrated_notes(note)
    assert "人が書いたメモ" in kept and "閉じていない手書きの注意" in kept
    assert "新しい順" not in kept  # 空白の違いは同じ文面


def test_a_longer_fence_is_not_closed_by_a_shorter_one(old_vault: Path, capsys) -> None:
    write_note(old_vault / "record/r7.md", "title: r7\ndate: 2026-07-07\n", "本文7。\n")
    note = old_vault / "project/project_proj_b.md"
    sample = "````\n```\n| 2026-07-07 | [r7](../record/r7.md) | t |\n````\n"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n\n" + sample)
    assert apply_it(old_vault, capsys) == 0
    assert sample.rstrip("\n") in migrated_notes(note)
    assert "project" not in fm(old_vault / "record/r7.md")


def test_an_empty_link_label_written_inside_a_template_comment_counts_as_a_placeholder(old_vault: Path, capsys) -> None:
    template = old_vault / "setting/template/template_project.md"
    edit(template, "## documents\n\n- [local]()\n", "## documents\n\n<!-- 例: [ローカル資料]() の形で書く。\n絶対パスは書かない -->\n")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [ローカル資料]()\n\n## minutes\n")
    assert apply_it(old_vault, capsys) == 0
    assert "ローカル資料" not in note.read_text(encoding="utf-8")


# ---- 人間が決める入力: --placeholder-label / --template-from / --add-frontmatter -------------------------


def test_a_placeholder_label_named_by_the_human_is_dropped(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [旧メモ]()\n- [契約書の保管場所]()\n\n## minutes\n")
    assert migrate(old_vault) == 0  # 決めていないうちは、template に無い label の空リンクは残る
    capsys.readouterr()
    assert apply_it(old_vault, capsys, "--placeholder-label", "旧メモ") == 0
    kept = migrated_notes(note)
    assert "[旧メモ]()" not in kept and "[契約書の保管場所]()" in kept  # 決めていない label は残る


def test_the_placeholder_label_is_part_of_the_plan_so_the_apply_must_repeat_it(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [旧メモ]()\n\n## minutes\n")
    assert migrate(old_vault, "--placeholder-label", "旧メモ") == 0
    pid = plan_id(capsys)
    assert migrate(old_vault, "--apply", "--plan-id", pid) == 2  # 付け忘れたら別のプラン（CONFLICT）


NEW_TEMPLATE = """---
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

<!-- ローカルの絶対パスを書かないこと（vibe-guard が vault 全体のバックアップを止める）。資料へのリンクは vault からの相対パス・URL -->

## overview

- 案件の一言説明:
"""


def write_template_source(tmp_path: Path, text: str = NEW_TEMPLATE) -> Path:
    path = tmp_path / "new_template.md"
    path.write_text(text, encoding="utf-8")
    return path


def test_template_from_replaces_the_template_and_still_adds_partner_scope_and_the_region(old_vault: Path, tmp_path: Path, capsys) -> None:
    tmpl = old_vault / "setting/template/template_project.md"
    edit(tmpl, "<!-- 他クライアントの議事録 -->", "<!-- 他クライアントの議事録 -->\n\n手で書いた注意書き（旧節の中の本文）")  # 旧節に本文があり TEMPLATE_MANUAL になる状態
    assert migrate(old_vault) == 0
    assert "TEMPLATE_MANUAL" in capsys.readouterr().out
    source = write_template_source(tmp_path)
    assert apply_it(old_vault, capsys, "--template-from", str(source)) == 0
    text = tmpl.read_text(encoding="utf-8")
    assert "ローカルの絶対パスを書かないこと" in text and "## minutes" not in text and "## documents" not in text
    assert "partner:" in text and "scope:" in text and "generated:associated-notes begin" in text
    capsys.readouterr()
    assert migrate(old_vault, "--template-from", str(source)) == 0
    assert "NOOP no migration changes required" in capsys.readouterr().out


def test_template_from_refuses_a_file_that_still_has_the_old_sections_or_a_local_path(old_vault: Path, tmp_path: Path) -> None:
    before = snapshot(old_vault)
    assert migrate(old_vault, "--template-from", str(write_template_source(tmp_path, NEW_TEMPLATE + "\n## minutes\n"))) in (2, 3)
    local = "/" + "Users" + "/someone/doc.md"
    assert migrate(old_vault, "--template-from", str(write_template_source(tmp_path, NEW_TEMPLATE + f"\n{local}\n"))) in (2, 3)
    assert snapshot(old_vault) == before


def test_template_from_refuses_a_vault_without_a_template(old_vault: Path, tmp_path: Path) -> None:
    (old_vault / "setting/template/template_project.md").unlink()
    assert migrate(old_vault, "--template-from", str(write_template_source(tmp_path))) in (2, 3)


def test_add_frontmatter_gives_a_frontmatter_less_note_a_title_and_a_date_and_lets_its_row_associate(old_vault: Path, capsys) -> None:
    (old_vault / "record/20260728_メモ.md").write_text("## 経緯\n\n- 本文。\n", encoding="utf-8")
    (old_vault / "record/メモ日付なし.md").write_text("本文だけ。\n", encoding="utf-8")
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n| 2026-07-28 | [20260728_メモ](../record/20260728_メモ.md) | メモの議題 |")
    assert migrate(old_vault) == 0
    assert "UNLINKED_ROW" in capsys.readouterr().out  # 付けなければ今までどおり残るだけ
    assert apply_it(old_vault, capsys, "--add-frontmatter", "record/20260728_メモ.md", "--add-frontmatter", "record/メモ日付なし.md") == 0
    first = old_vault / "record/20260728_メモ.md"
    data = fm(first)
    assert data["title"] == "20260728_メモ" and str(data["date"]) == "2026-07-28"
    assert data["project"] == ["proj_b"] and data["summary"] == "メモの議題"
    assert body(first).endswith("## 経緯\n\n- 本文。\n")  # 本文はそのまま
    second = old_vault / "record/メモ日付なし.md"
    assert fm(second) == {"title": "メモ日付なし"}  # 日付はファイル名に無いので作らない
    assert second.read_text(encoding="utf-8").endswith("本文だけ。\n")
    capsys.readouterr()
    assert migrate(old_vault, "--add-frontmatter", "record/20260728_メモ.md", "--add-frontmatter", "record/メモ日付なし.md") == 0  # 再実行は no-op
    assert "NOOP no migration changes required" in capsys.readouterr().out


@pytest.mark.parametrize("bad", ["record/missing.md", "../outside.md", "project/project_proj_a.md", "record"])
def test_add_frontmatter_refuses_a_note_that_is_missing_or_is_outside(old_vault: Path, bad: str) -> None:
    before = snapshot(old_vault)
    assert migrate(old_vault, "--add-frontmatter", bad) in (2, 3)
    assert snapshot(old_vault) == before


def test_add_frontmatter_skips_a_note_that_already_has_one_and_says_so(old_vault: Path, capsys) -> None:
    assert migrate(old_vault, "--add-frontmatter", "record/r1.md") == 0
    out = capsys.readouterr().out
    assert "FRONTMATTER_PRESENT record/r1.md" in out and "FRONTMATTER_ADDED" not in out


def test_add_frontmatter_never_touches_an_existing_file_outside_the_vault(old_vault: Path, tmp_path: Path) -> None:
    outside = tmp_path / "outside.md"
    outside.write_text("外の note。\n", encoding="utf-8")
    for arg in ("../outside.md", str(outside)):
        assert migrate(old_vault, "--add-frontmatter", arg) in (2, 3)
    assert outside.read_text(encoding="utf-8") == "外の note。\n"


# ---- 独立レビュー（PR #34）の指摘 -------------------------------------------------------------------


def named(vault: Path, rel: str, content: bytes) -> Path:
    path = vault / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return path


def test_add_frontmatter_refuses_a_path_that_goes_through_a_symlinked_directory(old_vault: Path, tmp_path: Path) -> None:
    (old_vault / "record/pl").symlink_to(old_vault / "project", target_is_directory=True)
    named(old_vault, "project/README.md", b"project dir doc\n")
    outside = tmp_path / "ext"
    outside.mkdir()
    (outside / "x.md").write_text("外。\n", encoding="utf-8")
    (old_vault / "record/ext").symlink_to(outside, target_is_directory=True)
    for rel in ("record/pl/README.md", "record/ext/x.md"):
        before = snapshot(old_vault)
        assert migrate(old_vault, "--add-frontmatter", rel) == 3  # 例外（traceback）にせず、検査で止める
        assert snapshot(old_vault) == before
    assert (outside / "x.md").read_text(encoding="utf-8") == "外。\n"


@pytest.mark.parametrize("rel", ["Project/README.md", "SETTING/x.md", "record/.hidden/x.md"])
def test_add_frontmatter_guards_use_the_names_on_disk_not_the_spelling_typed(old_vault: Path, rel: str) -> None:
    named(old_vault, "project/README.md", b"doc\n")
    named(old_vault, "setting/x.md", b"doc\n")
    named(old_vault, "record/.hidden/x.md", b"doc\n")
    before = snapshot(old_vault)
    assert migrate(old_vault, "--add-frontmatter", rel) == 3
    assert snapshot(old_vault) == before


def test_add_frontmatter_takes_the_title_and_the_association_from_the_name_on_disk(old_vault: Path, capsys) -> None:
    named(old_vault, "record/20260728_Memo.md", "本文。\n".encode())
    note = old_vault / "project/project_proj_b.md"
    edit(note, "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |", "| 2026-05-01 | [r4](../record/r4.md) | B の議題 |\n| 2026-07-28 | [20260728_Memo](../record/20260728_Memo.md) | メモ |")
    assert apply_it(old_vault, capsys, "--add-frontmatter", "record/20260728_MEMO.md") == 0  # 綴り（大文字小文字）が違っても同じファイル
    data = fm(old_vault / "record/20260728_Memo.md")
    assert data["title"] == "20260728_Memo" and data["project"] == ["proj_b"]


def test_a_title_that_cannot_round_trip_through_yaml_is_refused_not_a_crash(old_vault: Path) -> None:
    named(old_vault, "record/a b.md", "本文。\n".encode())
    before = snapshot(old_vault)
    assert migrate(old_vault, "--add-frontmatter", "record/a b.md") == 3
    assert snapshot(old_vault) == before


@pytest.mark.parametrize(
    "name, content, expect_date",
    [
        ("20261340_x.md", "本文\n", None),  # 存在しない日付は付けない
        ("20260728_crlf.md", "行 1\r\n行 2\r\n", "2026-07-28"),
        ("20260728_empty.md", "", "2026-07-28"),
    ],
)
def test_add_frontmatter_edge_cases_keep_the_body_byte_for_byte(old_vault: Path, capsys, name: str, content: str, expect_date: str | None) -> None:
    path = named(old_vault, f"record/{name}", content.encode())
    assert apply_it(old_vault, capsys, "--add-frontmatter", f"record/{name}") == 0
    raw = path.read_bytes().decode("utf-8")
    assert raw.endswith(content)
    data = yaml.safe_load(raw.split("---", 2)[1])
    assert (str(data.get("date")) if "date" in data else None) == expect_date
    if "\r\n" in content:
        assert "\r\n" in raw.split("---", 2)[1]  # frontmatter も CRLF


def test_a_note_that_starts_with_a_horizontal_rule_is_left_alone_with_an_honest_message(old_vault: Path, capsys) -> None:
    path = named(old_vault, "record/rule.md", "---\n区切り線で始まる本文\n".encode())
    assert migrate(old_vault, "--add-frontmatter", "record/rule.md") == 0
    out = capsys.readouterr().out
    assert "FRONTMATTER_PRESENT record/rule.md" in out and "starts like a frontmatter" in out
    assert path.read_bytes() == "---\n区切り線で始まる本文\n".encode()


@pytest.mark.parametrize("kind", ["bom", "symlink", "non_utf8"])
def test_add_frontmatter_refuses_a_bom_a_symlinked_note_and_a_non_utf8_note(old_vault: Path, tmp_path: Path, kind: str) -> None:
    if kind == "bom":
        named(old_vault, "record/x.md", b"\xef\xbb\xbfbody\n")
    elif kind == "non_utf8":
        named(old_vault, "record/x.md", b"\xff\xfe body\n")
    else:
        target = tmp_path / "t.md"
        target.write_text("t\n", encoding="utf-8")
        (old_vault / "record/x.md").symlink_to(target)
    before = snapshot(old_vault)
    assert migrate(old_vault, "--add-frontmatter", "record/x.md") == 3
    assert snapshot(old_vault) == before


def test_template_from_refuses_a_bom_and_unbalanced_region_markers(old_vault: Path, tmp_path: Path) -> None:
    before = snapshot(old_vault)
    bom = tmp_path / "bom.md"
    bom.write_bytes(b"\xef\xbb\xbf" + NEW_TEMPLATE.encode())
    assert migrate(old_vault, "--template-from", str(bom)) == 3
    for body in ("<!-- generated:associated-notes begin -->\n", "<!-- generated:associated-notes end -->\n<!-- generated:associated-notes begin -->\n", "<!-- generated:associated-notes begin -->\n<!-- generated:associated-notes end -->\n" * 2):
        assert migrate(old_vault, "--template-from", str(write_template_source(tmp_path, NEW_TEMPLATE + "\n" + body))) == 3
    assert snapshot(old_vault) == before


def test_template_from_keeps_dropping_the_project_comments_equal_to_the_old_templates_and_reports(old_vault: Path, tmp_path: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## minutes\n\n<!-- 新しい順。 -->\n")  # 旧 template の旧節にあるコメントと同文
    assert apply_it(old_vault, capsys, "--template-from", str(write_template_source(tmp_path))) == 0
    assert "新しい順" not in note.read_text(encoding="utf-8")
    assert migrate(old_vault, "--template-from", str(write_template_source(tmp_path))) == 0


def test_template_replaced_is_reported_and_an_unused_placeholder_label_is_reported(old_vault: Path, tmp_path: Path, capsys) -> None:
    assert migrate(old_vault, "--template-from", str(write_template_source(tmp_path)), "--placeholder-label", "誤字のラベル") == 0
    out = capsys.readouterr().out
    assert "TEMPLATE_REPLACED" in out and "PLACEHOLDER_LABEL_UNUSED 誤字のラベル" in out


def test_a_blank_placeholder_label_is_refused(old_vault: Path) -> None:
    assert migrate(old_vault, "--placeholder-label", "  ") in (2, 3)


def test_the_label_is_matched_after_stripping_whitespace(old_vault: Path, capsys) -> None:
    note = old_vault / "project/project_proj_b.md"
    edit(note, "## minutes\n", "## documents\n\n- [旧メモ]()\n\n## minutes\n")
    assert apply_it(old_vault, capsys, "--placeholder-label", " 旧メモ ") == 0
    assert "旧メモ" not in note.read_text(encoding="utf-8")
