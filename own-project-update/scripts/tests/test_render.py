"""render v0: project ノートの associated-notes 領域だけを書き直す。

SPEC-project-association の Requirement 5・6・7 と、WS-20261002-document-publish の
ISSUE-02 の受入（2 回目が byte-identical・領域外の本文が不変・旧形式の文字列 project も集計・
部分出力を残さない）を固定する。
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

import project_update as pu
from conftest import snapshot, write_note

BEGIN = "<!-- generated:associated-notes begin -->"
END = "<!-- generated:associated-notes end -->"


def run(vault: Path, *extra: str) -> int:
    return pu.main(["render", "--vault", str(vault), *extra])


def region_of(path: Path) -> str:
    text = path.read_text(encoding="utf-8")
    return text.split(BEGIN, 1)[1].split(END, 1)[0]


def seed_notes(vault: Path) -> None:
    write_note(
        vault / "record/20260701_kickoff.md",
        "title: kickoff\ndate: 2026-07-01\nproject: proj_a\nsummary: 立ち上げ会議\n",
    )
    write_note(
        vault / "record/20260715_review.md",
        "title: review\ndate: 2026-07-15\nproject:\n  - proj_a\n  - proj_b\nsummary: 中間レビュー\n",
    )
    write_note(
        vault / "presentation/20260726_proposal.md",
        "title: proposal\nkind: presentation\ndate: 2026-07-26\nproject:\n  - proj_a\nsummary: 提案資料\n",
    )


def test_dry_run_is_default_and_writes_nothing(vault: Path) -> None:
    seed_notes(vault)
    before = snapshot(vault)
    assert run(vault) == 0
    assert snapshot(vault) == before


def test_apply_appends_region_and_keeps_narrative_unchanged(vault: Path) -> None:
    seed_notes(vault)
    note = vault / "project/project_proj_a.md"
    original = note.read_text(encoding="utf-8")

    assert run(vault, "--apply") == 0

    text = note.read_text(encoding="utf-8")
    assert text.startswith(original)  # 既存の見出し・本文は 1 byte も変えない
    assert BEGIN in text and END in text
    region = region_of(note)
    for link in (
        "../record/20260701_kickoff.md",
        "../record/20260715_review.md",
        "../presentation/20260726_proposal.md",
    ):
        assert region.count(f"]({link})") == 1
    assert "立ち上げ会議" in region and "提案資料" in region


def test_second_run_is_byte_identical(vault: Path) -> None:
    seed_notes(vault)
    assert run(vault, "--apply") == 0
    first = snapshot(vault)
    assert run(vault, "--apply") == 0
    assert snapshot(vault) == first


def test_legacy_scalar_project_and_list_project_are_both_collected(vault: Path) -> None:
    seed_notes(vault)
    assert run(vault, "--apply") == 0
    # review は proj_a / proj_b の両方、kickoff（旧形式の文字列）は proj_a だけ
    assert "20260715_review.md" in region_of(vault / "project/project_proj_b.md")
    assert "20260701_kickoff.md" not in region_of(vault / "project/project_proj_b.md")
    assert "20260701_kickoff.md" in region_of(vault / "project/project_proj_a.md")


def test_rows_sorted_by_kind_then_date_descending(vault: Path) -> None:
    seed_notes(vault)
    write_note(
        vault / "record/20260801_later.md", "title: later\ndate: 2026-08-01\nproject: proj_a\n"
    )
    assert run(vault, "--apply") == 0
    region = region_of(vault / "project/project_proj_a.md")
    order = [
        region.index("20260726_proposal.md"),  # kind=presentation が record より先（昇順）
        region.index("20260801_later.md"),
        region.index("20260715_review.md"),
        region.index("20260701_kickoff.md"),
    ]
    assert order == sorted(order)


def test_project_notes_and_setting_templates_are_not_collected(vault: Path) -> None:
    seed_notes(vault)
    write_note(vault / "setting/template/template_note.md", "project: proj_a\n")
    assert run(vault, "--apply") == 0
    region = region_of(vault / "project/project_proj_a.md")
    assert "project_proj_a.md" not in region
    assert "template_note.md" not in region


def test_summary_pipe_and_newline_do_not_break_the_table(vault: Path) -> None:
    write_note(
        vault / "record/20260701_x.md",
        'title: x\ndate: 2026-07-01\nproject: proj_a\nsummary: "a | b\\nc"\n',
    )
    assert run(vault, "--apply") == 0
    row = [l for l in region_of(vault / "project/project_proj_a.md").splitlines() if "20260701_x" in l]
    assert len(row) == 1
    assert "a \\| b c" in row[0]


def test_existing_region_is_rewritten_and_text_around_it_is_kept(vault: Path) -> None:
    seed_notes(vault)
    note = vault / "project/project_proj_a.md"
    note.write_text(
        note.read_text(encoding="utf-8")
        + f"\n{BEGIN}\n古い内容\n{END}\n\n## 手書きの後続\n\n残す。\n",
        encoding="utf-8",
    )
    assert run(vault, "--apply") == 0
    text = note.read_text(encoding="utf-8")
    assert "古い内容" not in text
    assert text.count(BEGIN) == 1
    assert text.endswith(f"{END}\n\n## 手書きの後続\n\n残す。\n")


def test_no_notes_and_no_region_leaves_the_project_notes_alone_and_the_second_run_is_a_noop(vault: Path) -> None:
    notes_before = {p.name: p.read_bytes() for p in (vault / "project").glob("*.md")}
    assert run(vault, "--apply") == 0
    assert {p.name: p.read_bytes() for p in (vault / "project").glob("*.md")} == notes_before  # 領域は新設しない
    first = snapshot(vault)  # list_project.md は生成ビューとして書き直される（render 完全版）
    assert run(vault, "--apply") == 0
    assert snapshot(vault) == first


@pytest.mark.parametrize(
    "tail",
    [f"{BEGIN}\nx\n", f"{END}\n{BEGIN}\n", f"{BEGIN}\n{END}\n{BEGIN}\n{END}\n"],
    ids=["begin-only", "end-before-begin", "duplicated"],
)
def test_broken_markers_stop_without_writing(vault: Path, tail: str) -> None:
    seed_notes(vault)
    note = vault / "project/project_proj_a.md"
    note.write_text(note.read_text(encoding="utf-8") + "\n" + tail, encoding="utf-8")
    before = snapshot(vault)
    assert run(vault, "--apply") == 3
    assert snapshot(vault) == before


def test_unparsable_frontmatter_with_project_line_stops(vault: Path) -> None:
    seed_notes(vault)
    (vault / "record/broken.md").write_text(
        "---\nproject: proj_a\ntitle: [unclosed\n---\n本文\n", encoding="utf-8"
    )
    before = snapshot(vault)
    assert run(vault, "--apply") == 3
    assert snapshot(vault) == before


def test_unparsable_frontmatter_without_project_line_is_skipped(vault: Path) -> None:
    seed_notes(vault)
    (vault / "record/other.md").write_text("---\ntitle: [unclosed\n---\n本文\n", encoding="utf-8")
    assert run(vault, "--apply") == 0


def test_local_path_in_summary_is_refused(vault: Path) -> None:
    write_note(
        vault / "record/20260701_x.md",
        "title: x\ndate: 2026-07-01\nproject: proj_a\nsummary: see /private/tmp/someone/file\n",
    )
    before = snapshot(vault)
    assert run(vault, "--apply") == 3
    assert snapshot(vault) == before


def test_io_failure_midway_leaves_no_partial_output(
    vault: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    seed_notes(vault)
    before = snapshot(vault)
    real_replace = os.replace
    calls = {"n": 0}

    def flaky(src, dst, *a, **k):
        calls["n"] += 1
        if calls["n"] == 2:
            raise OSError("simulated Drive I/O failure")
        return real_replace(src, dst, *a, **k)

    monkeypatch.setattr(os, "replace", flaky)
    assert run(vault, "--apply") == 3
    monkeypatch.undo()
    assert snapshot(vault) == before
    # 一時ファイルも残さない
    assert not [p for p in vault.rglob(".*project-update-*")]


def test_unknown_project_key_is_reported_not_silently_dropped(
    vault: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    write_note(vault / "record/20260701_x.md", "title: x\ndate: 2026-07-01\nproject: typo_key\n")
    assert run(vault, "--apply") == 0
    assert "typo_key" in capsys.readouterr().err


def test_project_key_filter_limits_writes(vault: Path) -> None:
    seed_notes(vault)
    before_b = (vault / "project/project_proj_b.md").read_bytes()
    assert run(vault, "--project-key", "proj_a", "--apply") == 0
    assert (vault / "project/project_proj_b.md").read_bytes() == before_b
    assert BEGIN in (vault / "project/project_proj_a.md").read_text(encoding="utf-8")


def test_unknown_project_key_option_stops(vault: Path) -> None:
    assert run(vault, "--project-key", "nope", "--apply") == 3


# ---- レビュー指摘（PR #26）の回帰 --------------------------------------------------


def test_bom_prefixed_note_is_collected(vault: Path) -> None:
    (vault / "record").mkdir(exist_ok=True)
    (vault / "record/20260701_bom.md").write_text(
        "\ufeff---\ntitle: bom\ndate: 2026-07-01\nproject: proj_a\n---\n本文\n", encoding="utf-8"
    )
    assert run(vault, "--apply") == 0
    assert "20260701_bom.md" in region_of(vault / "project/project_proj_a.md")


def test_unclosed_frontmatter_that_declares_project_stops(vault: Path) -> None:
    (vault / "record/20260701_open.md").write_text("---\ntitle: x\nproject: proj_a\n本文\n", encoding="utf-8")
    before = snapshot(vault)
    assert run(vault, "--apply") == 3
    assert snapshot(vault) == before


def test_loose_and_integer_dates_are_normalised_for_ordering(vault: Path) -> None:
    write_note(vault / "record/a.md", "title: a\ndate: 2026-7-5\nproject: proj_a\n")
    write_note(vault / "record/b.md", "title: b\ndate: '2026/08/01'\nproject: proj_a\n")
    write_note(vault / "record/c.md", "title: c\ndate: 20260710\nproject: proj_a\n")
    assert run(vault, "--apply") == 0
    region = region_of(vault / "project/project_proj_a.md")
    assert "2026-07-05" in region and "2026-08-01" in region and "2026-07-10" in region
    assert region.index("b.md") < region.index("c.md") < region.index("a.md")


def test_crlf_project_note_keeps_crlf(vault: Path) -> None:
    seed_notes(vault)
    note = vault / "project/project_proj_a.md"
    note.write_bytes(note.read_bytes().replace(b"\n", b"\r\n"))
    assert run(vault, "--apply") == 0
    raw = note.read_bytes()
    assert BEGIN.encode() in raw
    assert raw.count(b"\r\n") == raw.count(b"\n")  # LF だけの行が混ざらない
    first = snapshot(vault)
    assert run(vault, "--apply") == 0
    assert snapshot(vault) == first


def test_render_refuses_to_write_over_a_note_that_changed_meanwhile(
    vault: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    seed_notes(vault)
    note = vault / "project/project_proj_a.md"
    real = pu._write_batch

    def racing(v, changes):  # 計画後・書き込み前に別の書き手が同じ project ノートを更新した状況
        note.write_text(note.read_text(encoding="utf-8") + "\n別セッション\n", encoding="utf-8")
        return real(v, changes)

    monkeypatch.setattr(pu, "_write_batch", racing)
    assert run(vault, "--apply") == 2
    assert "別セッション" in note.read_text(encoding="utf-8")
    assert BEGIN not in note.read_text(encoding="utf-8")


# ---- 実 vault で見つかった: UTF-8 として読めない note ------------------------------------


def test_a_note_that_is_not_utf8_and_declares_no_project_is_skipped_with_a_warning(
    vault: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    """実 vault の clip に、途中で切れたマルチバイトを含む note が 1 本あるだけで render 全体が止まっていた。"""
    seed_notes(vault)
    (vault / "record/broken_bytes.md").write_bytes(
        "---\ntitle: broken\n---\n本文".encode("utf-8") + b"\xe3\x81" + b"\n"  # 途中で切れた「あ」
    )
    assert run(vault, "--apply") == 0
    assert "broken_bytes.md" in capsys.readouterr().err  # 黙って飛ばさない
    assert "20260701_kickoff.md" in region_of(vault / "project/project_proj_a.md")


def test_a_non_utf8_note_that_declares_a_project_stops_the_render(vault: Path) -> None:
    seed_notes(vault)
    (vault / "record/broken_project.md").write_bytes(
        "---\ntitle: x\nproject: proj_a\n---\n本文".encode("utf-8") + b"\xe3\x81\n"
    )
    before = snapshot(vault)
    assert run(vault, "--apply") == 3  # project を名乗るのに読めないノートは、黙って表から消さない
    assert snapshot(vault) == before


def test_a_non_utf8_note_whose_frontmatter_is_unreadable_stops_the_render(vault: Path) -> None:
    seed_notes(vault)
    (vault / "record/broken_frontmatter.md").write_bytes(b"---\ntitle: \xe3\x81\nproject: proj_a\n---\n\xe3\n")
    assert run(vault, "--apply") == 3
