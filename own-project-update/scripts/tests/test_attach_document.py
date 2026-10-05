"""attach-document: 生成資料のテキストを vault に document note として公開する。

SPEC-document-publish の Requirement 1〜12 のうち、own-project-update が執行する部分
（document note 契約・リスト schema・冪等・人間編集で停止・部分適用なし・artifact link）を固定する。
assess は未実装なので「提案なし」で進むことも固定する。
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest
import yaml

import project_update as pu
from conftest import snapshot, write_note

NAME = "20260726_proposal"
BEGIN = "<!-- generated:associated-notes begin -->"
END = "<!-- generated:associated-notes end -->"


@pytest.fixture
def material(tmp_path: Path) -> dict[str, Path]:
    src = tmp_path / "src"
    src.mkdir()
    (src / "digest.md").write_text("- 提案の骨子\n- 二つ目のメッセージ\n", encoding="utf-8")
    (src / "outline.md").write_text("# アウトライン\n\n## 1. 現状\n\n本文。\n", encoding="utf-8")
    (src / "notes.md").write_text("話す順番のメモ。\n", encoding="utf-8")
    return {"digest": src / "digest.md", "outline": src / "outline.md", "notes": src / "notes.md"}


def attach(vault: Path, material: dict[str, Path], *extra: str, name: str = NAME) -> int:
    return pu.main(
        [
            "attach-document",
            "--vault", str(vault),
            "--kind", "presentation",
            "--name", name,
            "--source-repo", "deck_repo",
            "--source-path", f"decks/{name}",
            "--digest-file", str(material["digest"]),
            "--outline-file", str(material["outline"]),
            "--notes-file", str(material["notes"]),
            *extra,
        ]
    )


def frontmatter(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    return yaml.safe_load(text.split("---\n", 2)[1])


def doc(vault: Path, name: str = NAME) -> Path:
    return vault / "presentation" / f"{name}.md"


def region_of(path: Path) -> str:
    return path.read_text(encoding="utf-8").split(BEGIN, 1)[1].split(END, 1)[0]


# ---- 基本の書き込み -------------------------------------------------------------


def test_dry_run_is_default_and_writes_nothing(vault: Path, material) -> None:
    before = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a") == 0
    assert snapshot(vault) == before


def test_apply_writes_document_note_with_list_schema(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    data = frontmatter(doc(vault))
    assert data["kind"] == "presentation"
    assert data["project"] == ["proj_a"]  # 1 件でもリスト
    assert data["project_source"] == "manual"
    assert data["summary"] == "提案の骨子"  # ダイジェスト 1 行目（箇条書き記号は除く）
    assert data["date"].isoformat() == "2026-07-26"  # name の yyyymmdd から
    assert data["source_repo"] == "deck_repo"
    assert data["source_path"] == f"decks/{NAME}"
    assert len(data["source_hash"]) == 64 and len(data["content_hash"]) == 64
    assert data["published_at"]
    text = doc(vault).read_text(encoding="utf-8")
    body = text.split("---\n", 2)[2]
    assert body.startswith("> project: [project_proj_a](../project/project_proj_a.md)\n")
    assert text.index("提案の骨子") < text.index("# アウトライン") < text.index("話す順番のメモ")


def test_apply_renders_project_note_in_the_same_run(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    region = region_of(vault / "project/project_proj_a.md")
    assert region.count(f"](../presentation/{NAME}.md)") == 1
    assert "提案の骨子" in region
    # 紐付けていない project には触らない
    assert BEGIN not in (vault / "project/project_proj_b.md").read_text(encoding="utf-8")


def test_rerun_is_a_noop_including_published_at(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    first = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    assert snapshot(vault) == first


# ---- 人間編集・衝突・入力検査 ------------------------------------------------------


def test_human_edit_of_body_stops_and_keeps_file(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    note = doc(vault)
    note.write_text(note.read_text(encoding="utf-8") + "\n人間が足したメモ\n", encoding="utf-8")
    material["digest"].write_text("- 変わった骨子\n", encoding="utf-8")  # 再生成を要する変更
    before = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 2
    assert snapshot(vault) == before


def test_adding_a_second_project_is_not_seen_as_human_edit(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    assert attach(vault, material, "--project", "proj_b", "--apply") == 0
    assert frontmatter(doc(vault))["project"] == ["proj_a", "proj_b"]
    body = doc(vault).read_text(encoding="utf-8").split("---\n", 2)[2]
    assert body.count("> project:") == 2
    assert NAME in region_of(vault / "project/project_proj_a.md")
    assert NAME in region_of(vault / "project/project_proj_b.md")
    # 3 回目（同じ指定）は no-op
    first = snapshot(vault)
    assert attach(vault, material, "--project", "proj_b", "--apply") == 0
    assert snapshot(vault) == first


def test_source_change_regenerates_and_keeps_foreign_frontmatter(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    note = doc(vault)
    text = note.read_text(encoding="utf-8")
    note.write_text(text.replace("kind: presentation\n", "kind: presentation\ntags:\n  - keep\n", 1), encoding="utf-8")
    old_hash = frontmatter(note)["source_hash"]
    material["digest"].write_text("- 改訂後の骨子\n", encoding="utf-8")
    assert attach(vault, material, "--apply") == 0  # project は既存を引き継ぐ（指定なしでも決定済み）
    data = frontmatter(note)
    assert data["source_hash"] != old_hash
    assert data["summary"] == "改訂後の骨子"
    assert data["tags"] == ["keep"]
    assert data["project"] == ["proj_a"]
    assert "改訂後の骨子" in region_of(vault / "project/project_proj_a.md")


def test_legacy_scalar_project_is_read_and_preserved(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    note = doc(vault)
    note.write_text(
        note.read_text(encoding="utf-8").replace("project:\n- proj_a\n", "project: proj_a\n").replace(
            "project:\n  - proj_a\n", "project: proj_a\n"
        ),
        encoding="utf-8",
    )
    assert frontmatter(note)["project"] == "proj_a"  # 前提: 旧形式になっている
    assert attach(vault, material, "--project", "proj_b", "--apply") == 0
    assert frontmatter(note)["project"] == ["proj_a", "proj_b"]


def test_undecided_association_stops_before_writing(vault: Path, material) -> None:
    before = snapshot(vault)
    assert attach(vault, material, "--apply") == 2
    assert snapshot(vault) == before


def test_no_project_creates_note_without_project_and_skips_render(vault: Path, material) -> None:
    project_before = (vault / "project/project_proj_a.md").read_bytes()
    assert attach(vault, material, "--no-project", "--apply") == 0
    data = frontmatter(doc(vault))
    assert "project" not in data and "project_source" not in data
    assert (vault / "project/project_proj_a.md").read_bytes() == project_before


def test_dry_run_reports_no_proposals_when_assess_is_unavailable(
    vault: Path, material, capsys: pytest.CaptureFixture[str]
) -> None:
    assert attach(vault, material, "--hint", "client=Client proj_a") == 0
    out = capsys.readouterr().out
    assert "ASSESS unavailable" in out
    assert "HINT client=Client proj_a" in out


def test_nonexistent_project_is_refused(vault: Path, material) -> None:
    before = snapshot(vault)
    assert attach(vault, material, "--project", "nope", "--apply") == 3
    assert snapshot(vault) == before


def test_project_and_no_project_are_exclusive(vault: Path, material) -> None:
    with pytest.raises(SystemExit):
        attach(vault, material, "--project", "proj_a", "--no-project")


@pytest.mark.parametrize("bad", ["/private/tmp/someone/deck", "file:///x"])
def test_local_path_in_material_is_refused(vault: Path, material, bad: str) -> None:
    material["outline"].write_text(f"# t\n\n参照 {bad}\n", encoding="utf-8")
    before = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 3
    assert snapshot(vault) == before


@pytest.mark.parametrize("source_path", ["/abs/path", "../up", "~/x", "a/../../b"])
def test_source_path_must_be_repo_relative(vault: Path, material, source_path: str) -> None:
    rc = pu.main(
        [
            "attach-document", "--vault", str(vault), "--kind", "presentation", "--name", NAME,
            "--source-repo", "deck_repo", "--source-path", source_path,
            "--digest-file", str(material["digest"]), "--project", "proj_a", "--apply",
        ]
    )
    assert rc == 3


def test_missing_kind_directory_and_reserved_kind_are_refused(vault: Path, material) -> None:
    for kind in ("report", "project", "setting"):
        rc = pu.main(
            [
                "attach-document", "--vault", str(vault), "--kind", kind, "--name", NAME,
                "--source-repo", "deck_repo", "--source-path", "x",
                "--digest-file", str(material["digest"]), "--project", "proj_a", "--apply",
            ]
        )
        assert rc == 3, kind
    assert not (vault / "report").exists()


def test_name_without_date_prefix_needs_explicit_date(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply", name="undated") == 3
    assert attach(vault, material, "--project", "proj_a", "--date", "2026-08-02", "--apply", name="undated") == 0
    assert frontmatter(doc(vault, "undated"))["date"].isoformat() == "2026-08-02"


def test_same_name_from_another_source_is_a_conflict(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    rc = pu.main(
        [
            "attach-document", "--vault", str(vault), "--kind", "presentation", "--name", NAME,
            "--source-repo", "other_repo", "--source-path", "x",
            "--digest-file", str(material["digest"]), "--project", "proj_a", "--apply",
        ]
    )
    assert rc == 2


def test_note_with_tampered_content_hash_field_missing_stops(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    note = doc(vault)
    note.write_text(
        "".join(l for l in note.read_text(encoding="utf-8").splitlines(True) if not l.startswith("content_hash:")),
        encoding="utf-8",
    )
    material["digest"].write_text("- 変更\n", encoding="utf-8")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 2


def test_failure_during_write_leaves_vault_unchanged(
    vault: Path, material, monkeypatch: pytest.MonkeyPatch
) -> None:
    before = snapshot(vault)
    real_replace = os.replace
    calls = {"n": 0}

    def flaky(src, dst, *a, **k):
        calls["n"] += 1
        if calls["n"] == 2:  # document note は書けたが project ノートで失敗
            raise OSError("simulated I/O failure")
        return real_replace(src, dst, *a, **k)

    monkeypatch.setattr(os, "replace", flaky)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 3
    monkeypatch.undo()
    assert snapshot(vault) == before  # 新規作成した document note も巻き戻る


# ---- artifact link ----------------------------------------------------------------


def test_artifact_link_is_relative_symlink_and_linked_from_body(
    vault: Path, material, tmp_path: Path
) -> None:
    deck = tmp_path / "deckdir" / f"{NAME}.pptx"
    deck.parent.mkdir()
    deck.write_bytes(b"pptx")
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    link = vault / "presentation" / f"{NAME}.pptx"
    assert link.is_symlink()
    target = os.readlink(link)
    assert not os.path.isabs(target)
    assert (link.parent / target).resolve() == deck.resolve()
    assert f"]({NAME}.pptx)" in doc(vault).read_text(encoding="utf-8")
    # 再実行は no-op（symlink も同じ）
    first = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    assert snapshot(vault) == first and os.readlink(link) == target


def test_unresolvable_artifact_skips_link_but_publishes(
    vault: Path, material, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(tmp_path / "missing.pptx"), "--apply") == 0
    assert doc(vault).is_file()
    assert not (vault / "presentation" / f"{NAME}.pptx").exists()
    assert f"]({NAME}.pptx)" not in doc(vault).read_text(encoding="utf-8")
    assert "ARTIFACT skipped" in capsys.readouterr().out


def test_regular_file_at_artifact_path_is_not_clobbered(vault: Path, material, tmp_path: Path) -> None:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    occupied = vault / "presentation" / f"{NAME}.pptx"
    occupied.write_bytes(b"human file")
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    assert occupied.read_bytes() == b"human file" and not occupied.is_symlink()


def test_artifact_in_linked_worktree_resolves_to_main_checkout(
    vault: Path, material, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    main = tmp_path / "main_checkout"
    wt = tmp_path / "wt_checkout"
    (main / "decks").mkdir(parents=True)
    (wt / "decks").mkdir(parents=True)
    (main / "decks" / "d.pptx").write_bytes(b"main")
    (wt / "decks" / "d.pptx").write_bytes(b"main")

    def fake_git(args: list[str], cwd: Path) -> str:
        if args[:2] == ["rev-parse", "--show-toplevel"]:
            return str(wt)
        if args[:2] == ["worktree", "list"]:
            return f"worktree {main}\nHEAD abc\nbranch refs/heads/main\n\nworktree {wt}\nHEAD def\n"
        raise AssertionError(args)

    monkeypatch.setattr(pu, "_git_output", fake_git)
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(wt / "decks" / "d.pptx"), "--apply") == 0
    link = vault / "presentation" / f"{NAME}.pptx"
    assert (link.parent / os.readlink(link)).resolve() == (main / "decks" / "d.pptx").resolve()


def test_worktree_artifact_that_differs_from_main_is_not_linked(
    vault: Path, material, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    main, wt = tmp_path / "main_checkout", tmp_path / "wt_checkout"
    for root, payload in ((main, b"old"), (wt, b"new, unmerged")):
        (root / "decks").mkdir(parents=True)
        (root / "decks" / "d.pptx").write_bytes(payload)

    def fake_git(args: list[str], cwd: Path) -> str:
        if args[:2] == ["rev-parse", "--show-toplevel"]:
            return str(wt)
        return f"worktree {main}\n\nworktree {wt}\n"

    monkeypatch.setattr(pu, "_git_output", fake_git)
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(wt / "decks" / "d.pptx"), "--apply") == 0
    assert not (vault / "presentation" / f"{NAME}.pptx").exists()
    assert "not merged yet" in capsys.readouterr().out


# ---- レビュー指摘（PR #26）の回帰 --------------------------------------------------


def test_rerun_with_artifact_adds_the_link_line_to_an_unchanged_source(
    vault: Path, material, tmp_path: Path
) -> None:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    assert f"]({NAME}.pptx)" not in doc(vault).read_text(encoding="utf-8")
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    assert f"]({NAME}.pptx)" in doc(vault).read_text(encoding="utf-8")
    assert (vault / "presentation" / f"{NAME}.pptx").is_symlink()
    first = snapshot(vault)  # content_hash が更新されているので、次の再実行は人間編集と誤認しない
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    assert snapshot(vault) == first


def test_regeneration_without_artifact_option_keeps_a_live_link_line(
    vault: Path, material, tmp_path: Path
) -> None:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    material["digest"].write_text("- 改訂\n", encoding="utf-8")
    assert attach(vault, material, "--apply") == 0
    assert f"]({NAME}.pptx)" in doc(vault).read_text(encoding="utf-8")


def test_hand_removed_association_clears_the_stale_row(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    note = doc(vault)
    text = note.read_text(encoding="utf-8")
    note.write_text(text.replace("project:\n  - proj_a\n", "project: []\n"), encoding="utf-8")
    assert attach(vault, material, "--no-project", "--apply") == 0
    assert NAME not in region_of(vault / "project/project_proj_a.md")


def test_long_summary_round_trips_without_folding(vault: Path, material) -> None:
    long = "長い要約 " * 60
    material["digest"].write_text(f"- {long}\n", encoding="utf-8")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    assert frontmatter(doc(vault))["summary"] == long.strip()
    first = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    assert snapshot(vault) == first


@pytest.mark.parametrize(
    "leak",
    [
        "`/home/someone/x`",
        '"/private/tmp/x"',
        "'/Volumes/Drive/x'",
        "(~/Library/CloudStorage/GoogleDrive-a/b)",
        "C:\\Users\\x\\deck",
        "file:///x",
    ],
)
def test_ingest_rejects_quoted_and_drive_paths(vault: Path, material, leak: str) -> None:
    material["outline"].write_text(f"# t\n\n参照 {leak}\n", encoding="utf-8")
    before = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 3
    assert snapshot(vault) == before


def test_ordinary_urls_and_paths_are_not_mistaken_for_local_paths(vault: Path, material) -> None:
    material["outline"].write_text("# t\n\nhttps://example.com/home/page と docs/private/x\n", encoding="utf-8")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0


@pytest.mark.parametrize("bad", ["20260726_a(b)", "20260726_a[b]", "20260727_x#y", "20260727_x%y"])
def test_names_that_break_links_are_refused_with_a_name_message(
    vault: Path, material, bad: str, capsys: pytest.CaptureFixture[str]
) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply", name=bad) == 3
    assert "project key" not in capsys.readouterr().err


def test_crlf_document_note_keeps_crlf_when_a_project_is_added(vault: Path, material) -> None:
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    note = doc(vault)
    note.write_bytes(note.read_bytes().replace(b"\n", b"\r\n"))
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0  # 同内容は NOOP
    assert attach(vault, material, "--project", "proj_b", "--apply") == 0
    raw = note.read_bytes()
    assert raw.count(b"\r\n") == raw.count(b"\n") > 0


def test_symlink_failure_publishes_without_a_dangling_link_line(
    vault: Path, material, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys
) -> None:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")

    def refuse(link: Path, target: str) -> None:
        raise OSError("symlinks are refused here")

    monkeypatch.setattr(pu, "_install_link", refuse)
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 0
    assert f"]({NAME}.pptx)" not in doc(vault).read_text(encoding="utf-8")
    assert "ARTIFACT skipped" in capsys.readouterr().out


def test_batch_failure_removes_the_link_it_just_created(
    vault: Path, material, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    deck = tmp_path / "deck.pptx"
    deck.write_bytes(b"pptx")
    before = snapshot(vault)
    real_replace = os.replace
    calls = {"n": 0}

    def flaky(src, dst, *a, **k):
        calls["n"] += 1
        if calls["n"] == 3:  # 1 = symlink、2 = document note、3 = project ノートで失敗
            raise OSError("simulated I/O failure")
        return real_replace(src, dst, *a, **k)

    monkeypatch.setattr(os, "replace", flaky)
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(deck), "--apply") == 3
    monkeypatch.undo()
    assert snapshot(vault) == before
    assert not (vault / "presentation" / f"{NAME}.pptx").is_symlink()


def test_write_refuses_when_the_file_changed_after_planning(vault: Path) -> None:
    path = vault / "project/project_proj_a.md"
    planned = path.read_text(encoding="utf-8")
    change = pu.Change(path, planned, planned + "\nplanned", verify_before=True)
    path.write_text(planned + "\n別セッションが書いた\n", encoding="utf-8")
    with pytest.raises(pu.ConflictError):
        pu._write_batch(vault, [change])
    assert "別セッションが書いた" in path.read_text(encoding="utf-8")


@pytest.mark.parametrize(
    "prose",
    [
        "SCIM endpoint GET /Users/{id}",
        "the /home/ route and (/home/) tab",
        "regex a:\\d+ と C:\\ drive root",
        "CloudStorageClient and Store in CloudStorage buckets",
        "GET /users/123 と /users/[id]/edit と ^/users/(\\d+)$",
        "profile:/settings と user_profile:/me と GoogleDrive-style な UI",
        "~/.config/app/config.toml に置く",
        "a:/api/users と /Users/Shared と /api/v1/users/123",
        "https://example.com/Users-guide と 16:9 と TCP/IP",
        "My Drive/Shared is a product name",
    ],
)
def test_ingest_does_not_reject_ordinary_technical_prose(vault: Path, material, prose: str) -> None:
    material["outline"].write_text(f"# t\n\n{prose}\n", encoding="utf-8")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0


def test_changing_the_artifact_extension_replaces_the_link_line(vault: Path, material, tmp_path: Path) -> None:
    pptx, key = tmp_path / "deck.pptx", tmp_path / "deck.key"
    pptx.write_bytes(b"a")
    key.write_bytes(b"b")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    for artifact in (pptx, key, pptx):
        assert attach(vault, material, "--project", "proj_a", "--artifact", str(artifact), "--apply") == 0
    text = doc(vault).read_text(encoding="utf-8")
    assert text.count(f"]({NAME}.pptx)") == 1 and f"]({NAME}.key)" not in text
    first = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--artifact", str(pptx), "--apply") == 0
    assert snapshot(vault) == first  # 冪等（行が増え続けない）


def test_an_existing_note_with_an_api_path_summary_does_not_block_render(vault: Path) -> None:
    write_note(
        vault / "record/20260701_scim.md",
        "title: scim\ndate: 2026-07-01\nproject: proj_a\nsummary: GET /Users/{id} を設計\n",
    )
    assert pu.main(["render", "--vault", str(vault), "--apply"]) == 0


def test_impossible_dates_are_left_as_text(vault: Path) -> None:
    write_note(vault / "record/a.md", "title: a\ndate: 20261399\nproject: proj_a\n")
    assert pu.main(["render", "--vault", str(vault), "--apply"]) == 0
    assert "2026-13-99" not in (vault / "project/project_proj_a.md").read_text(encoding="utf-8")


@pytest.mark.parametrize(
    "leak",
    [
        "|<U>/alice/x|",
        "|a|<U>/alice/x|b|",
        "path:<U>/alice/Documents/a.pdf",
        "パス：<U>/alice/Documents/a.pdf",
        "ファイルは<U>/alice/x.md にある",
        "資料を<U>/alice/Documents/a.pdf",
        "HOME=<U>/alice",
        "export OUT=<U>/alice/out",
        "a.md,<U>/alice/b.md",
        "x;<U>/alice/x",
        "「<U>/alice/x」",
        "（<U>/alice/x）",
        "**<U>/alice/x**",
        "(see <U>/alice)",
        "[<U>/alice]",
        "{<U>/alice}",
        "/Volumes/Macintosh HD<U>/alice/x",
        "/Volumes/Backup",
        "/mnt/c<U>/alice/x",
        "c:\\users\\alice\\x",
        "D:\\work\\client\\x",
        "~/Google Drive/clientX/a.pdf",
        "~/OneDrive - Corp/a.pdf",
        "file:<U>/alice/x",
        "Library/Mobile Documents/com~apple~CloudDocs/x",
        "/var/folders/ab/cd/T/x",
        "[docs](/home/guide)",
        "%2F<U>%2Falice%2Fx",
        "<U>%2Falice",
        "file%3A%2F%2F%2F<U>%2Falice",
        "%252F<U>%252Falice",
        "\\<U>\\/alice",
        "&#47;<U>&#47;alice",
        "$HOME/code/x と ${HOME}/code/x と %USERPROFILE%\\Documents",
        "~alice/code/x",
        "~/obsidian_code と ~/Pictures と ~/Developer",
        "<U>\n/alice",
        "<U>/\nalice",
        "//<U>//alice",
        "/System/Volumes/Data<U>/alice/x",
        "/cygdrive/c<U>/alice",
        "Macintosh HD<U>/alice",
        "/export/home/alice",
        "sftp://alice@host/home/alice/x",
        "\\\\wsl$\\Ubuntu\\home\\alice",
        "\\\\server\\share\\alice\\x",
        "smb://server/share",
        "C:\\My Documents\\alice\\x",
        "C：\\Users\\alice",
        "C:\\\\Users\\\\alice\\\\Desktop",
        "c:/users/alice/x",
        "OneDrive - Contoso/Documents/x.pptx",
        "共有ドライブ/営業/x と Google ドライブ/x と iCloud Drive/x と Dropbox (Personal)/x",
        "CloudStorage/Dropbox/x",
        "GoogleDrive-alice@example.com/My Drive/x",
    ],
)
def test_ingest_rejects_realistic_leaks(vault: Path, material, leak: str) -> None:
    """偽陽性側（技術文を通す）だけでなく、偽陰性側（漏れを通さない）も固定する。"""
    leak = leak.replace("<U>", "/" + "Users")  # 検査に引っかからないよう、ソース上では連結して持つ
    material["outline"].write_text(f"# t\n\n{leak}\n", encoding="utf-8")
    before = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 3
    assert snapshot(vault) == before


def test_hint_with_a_path_after_equals_is_refused(vault: Path, material) -> None:
    assert attach(vault, material, "--hint", "path=/" + "Users/alice/x") == 3


def test_whitespace_only_lines_do_not_make_the_check_quadratic() -> None:
    import time

    started = time.monotonic()
    assert not pu._has_local_path(" \n" * 3000 + "ordinary text")
    assert time.monotonic() - started < 5


@pytest.mark.parametrize("separator", ["\x0b", "\x0c", "\x1c", "\x85", " ", " "])
def test_unicode_line_separators_in_material_do_not_break_idempotence(
    vault: Path, material, separator: str
) -> None:
    """PowerPoint の Shift+Enter は python-pptx で \\v になる。splitlines が分割して hash が合わなくなる。"""
    material["notes"].write_text(f"一行目{separator}二行目\n", encoding="utf-8")
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0
    text = doc(vault).read_text(encoding="utf-8")
    assert separator not in text and "一行目\n二行目" in text
    first = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--apply") == 0  # 偽の CONFLICT にならない
    assert snapshot(vault) == first


@pytest.mark.parametrize("separator", ["\x0b", "\x85", " "])
def test_title_with_a_line_separator_is_refused_before_it_reaches_the_note(
    vault: Path, material, separator: str
) -> None:
    """title は frontmatter の 1 行値。行区切りが入ると公開は通るのに次回の読み込みが YAML エラーになる。"""
    before = snapshot(vault)
    assert attach(vault, material, "--project", "proj_a", "--title", f"前{separator}後", "--apply") == 3
    assert snapshot(vault) == before


# ---- ダイジェストを任意にする（--summary） -----------------------------------------------


def attach_without_digest(vault: Path, material, *extra: str, name: str = NAME) -> int:
    return pu.main(
        [
            "attach-document", "--vault", str(vault), "--kind", "presentation", "--name", name,
            "--source-repo", "deck_repo", "--source-path", f"decks/{name}",
            "--outline-file", str(material["outline"]),
            *extra,
        ]
    )


def test_summary_option_replaces_the_digest_and_no_digest_section_is_written(vault: Path, material) -> None:
    assert attach_without_digest(vault, material, "--summary", "提案スライド（全18枚）", "--no-project", "--apply") == 0
    data = frontmatter(doc(vault))
    assert data["summary"] == "提案スライド（全18枚）"
    text = doc(vault).read_text(encoding="utf-8")
    assert "## digest" not in text and "## outline" in text
    first = snapshot(vault)
    assert attach_without_digest(vault, material, "--summary", "提案スライド（全18枚）", "--no-project", "--apply") == 0
    assert snapshot(vault) == first  # 冪等


def test_summary_is_required_when_there_is_no_digest(vault: Path, material) -> None:
    before = snapshot(vault)
    assert attach_without_digest(vault, material, "--no-project", "--apply") == 3
    assert snapshot(vault) == before


def test_something_to_publish_is_required(vault: Path, material) -> None:
    rc = pu.main(
        [
            "attach-document", "--vault", str(vault), "--kind", "presentation", "--name", NAME,
            "--source-repo", "deck_repo", "--source-path", "x", "--summary", "題", "--no-project", "--apply",
        ]
    )
    assert rc == 3  # digest も outline も無い


@pytest.mark.parametrize(
    "bad",
    ["一行目\n二行目", "see /private/tmp/someone/x", "   ", "\x0b"],
    ids=["multi-line", "local-path", "blank", "separator-only"],
)
def test_bad_summary_is_refused(vault: Path, material, bad: str) -> None:
    before = snapshot(vault)
    assert attach_without_digest(vault, material, "--summary", bad, "--no-project", "--apply") == 3
    assert snapshot(vault) == before


def test_digest_still_wins_the_summary_when_both_are_absent_of_summary_option(vault: Path, material) -> None:
    assert attach(vault, material, "--no-project", "--apply") == 0
    assert frontmatter(doc(vault))["summary"] == "提案の骨子"


def test_summary_option_beats_the_digest_first_line(vault: Path, material) -> None:
    assert attach(vault, material, "--summary", "明示した題", "--no-project", "--apply") == 0
    assert frontmatter(doc(vault))["summary"] == "明示した題"
    assert "## digest" in doc(vault).read_text(encoding="utf-8")  # digest を渡せば節は残る
