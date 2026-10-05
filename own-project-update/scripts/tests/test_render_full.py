"""render 完全版: generated 領域に加えて list_project.md と Raycast 一覧を、同じ入力から全体生成する。

SPEC-project-association Requirement 5・6 と WS-20261003-project-association の ISSUE-02 の受入:
3 出力は 2 回実行で byte-identical、領域外の本文は不変、全出力を揃えられない時は何も書かない（部分出力なし）。
list_project.md の最終会議日は record だけから計算する（資料の公開では更新しない）。
"""

from __future__ import annotations

import os
import re
import stat
from pathlib import Path

import pytest

import project_update as pu
from conftest import PROJECT_NOTE, snapshot, write_note

RAYCAST = """#!/bin/bash
# @raycast.schemaVersion 1
# @raycast.title Start
# @raycast.argument1 { "type": "text", "placeholder": "title" }
# @raycast.argument2 {"type": "dropdown", "placeholder": "project", "data": [{"title": "old", "value": "old"}]}
echo "$1 $2"
"""


def run(vault: Path, *extra: str) -> int:
    return pu.main(["render", "--vault", str(vault), *extra])


def list_file(vault: Path) -> Path:
    return vault / "setting/list/list_project.md"


def note_with(vault: Path, key: str, frontmatter: str) -> None:
    (vault / "project" / f"project_{key}.md").write_text(
        f"---\ntitle: project_{key}\nproject: {key}\n{frontmatter}---\n\n# {key}\n\n## overview\n\n人が書いた本文。\n",
        encoding="utf-8",
    )


@pytest.fixture
def raycast(tmp_path: Path) -> Path:
    path = tmp_path / "raycast_start.sh"
    path.write_text(RAYCAST, encoding="utf-8")
    path.chmod(0o755)
    return path


def seed(vault: Path) -> None:
    note_with(vault, "proj_a", "client: Client A\npartner: 協力会社\nstatus: active\n")
    note_with(vault, "proj_b", "client: proj_b\nstatus: active\n")
    write_note(vault / "record/20260701_a.md", "title: a\ndate: 2026-07-01\nproject: proj_a\n")
    write_note(vault / "record/20260820_a2.md", "title: a2\ndate: 2026-08-20\nproject:\n  - proj_a\n  - proj_b\n")
    write_note(vault / "presentation/20260901_deck.md", "title: deck\nkind: presentation\ndate: 2026-09-01\nproject:\n  - proj_a\n")


def rows(vault: Path) -> list[list[str]]:
    out = []
    for line in list_file(vault).read_text(encoding="utf-8").splitlines():
        cells = [c.strip() for c in re.split(r"(?<!\\)\|", line.strip().strip("|"))] if line.startswith("|") else []
        if cells and cells[0] not in ("name",) and not set(cells[0]) <= {"-", " "}:
            out.append(cells)
    return out


# ---- list_project.md ---------------------------------------------------------------


def test_list_project_is_generated_from_frontmatter_and_record_dates(vault: Path) -> None:
    seed(vault)
    assert run(vault, "--apply", "--accept-list-changes") == 0
    text = list_file(vault).read_text(encoding="utf-8")
    assert "| name | client | partner | status | last meeting | path |" in text
    by_name = {r[0]: r for r in rows(vault)}
    # 最終会議日は record だけ。資料（2026-09-01）では更新しない
    assert by_name["proj_a"][1:5] == ["Client A", "協力会社", "active", "2026-08-20"]
    assert by_name["proj_b"][1:5] == ["proj_b", "", "active", "2026-08-20"]
    assert by_name["proj_a"][5] == "[project_proj_a](../../project/project_proj_a.md)"


def test_rows_are_ordered_active_first_then_last_meeting_desc_then_name(vault: Path) -> None:
    seed(vault)
    note_with(vault, "proj_c", "client: C\nstatus: closed\n")
    note_with(vault, "proj_d", "client: D\nstatus: active\n")
    write_note(vault / "record/20260930_d.md", "title: d\ndate: 2026-09-30\nproject: proj_d\n")
    write_note(vault / "record/20261001_c.md", "title: c\ndate: 2026-10-01\nproject: proj_c\n")  # closed でも最終会議日は最新
    assert run(vault, "--apply", "--accept-list-changes") == 0
    assert [r[0] for r in rows(vault)] == ["proj_d", "proj_a", "proj_b", "proj_c"]  # active が先、closed は日付が新しくても後ろ


def test_partner_list_is_joined_and_cells_are_escaped(vault: Path) -> None:
    note_with(vault, "proj_a", "client: 'A | B'\npartner:\n  - 社一\n  - 社二\nstatus: active\n")
    assert run(vault, "--apply", "--accept-list-changes") == 0
    row = next(r for r in rows(vault) if r[0] == "proj_a")
    assert row[1] == "A \\| B" and row[2] == "社一, 社二"


def test_second_run_is_byte_identical_for_every_output(vault: Path, raycast: Path) -> None:
    seed(vault)
    assert run(vault, "--apply", "--accept-list-changes", "--raycast-script", str(raycast)) == 0
    first = snapshot(vault)
    raycast_first = raycast.read_bytes()
    assert run(vault, "--apply", "--raycast-script", str(raycast)) == 0  # 2 回目は --accept 不要（差が無い）
    assert snapshot(vault) == first and raycast.read_bytes() == raycast_first


def test_existing_list_rows_that_would_lose_information_stop_the_render(vault: Path) -> None:
    seed(vault)
    list_file(vault).write_text(
        "# projects\n\n| name | client | partner | status | last meeting | path |\n| --- | --- | --- | --- | --- | --- |\n"
        "| proj_a | Client A | 別の協力会社 | active | 2026-09-01 | [x](y) |\n",
        encoding="utf-8",
    )
    before = snapshot(vault)
    assert run(vault, "--apply") == 2  # partner の差・last meeting の後退は黙って落とさない
    assert snapshot(vault) == before
    assert run(vault, "--apply", "--accept-list-changes") == 0


def test_list_without_setting_directory_is_an_error(vault: Path) -> None:
    seed(vault)
    list_file(vault).unlink()
    list_file(vault).parent.rmdir()
    before = snapshot(vault)
    assert run(vault, "--apply") == 3
    assert snapshot(vault) == before


# ---- Raycast 一覧 ------------------------------------------------------------------


def test_raycast_line_lists_active_projects_in_list_order(vault: Path, raycast: Path) -> None:
    seed(vault)
    note_with(vault, "proj_c", "client: C\nstatus: closed\n")
    assert run(vault, "--apply", "--accept-list-changes", "--raycast-script", str(raycast)) == 0
    text = raycast.read_text(encoding="utf-8")
    line = next(l for l in text.splitlines() if l.startswith("# @raycast.argument2"))
    assert line == (
        '# @raycast.argument2 {"type": "dropdown", "placeholder": "project", "data": '
        '[{"title": "proj_a (Client A)", "value": "proj_a"}, {"title": "proj_b", "value": "proj_b"}]}'
    )
    assert "proj_c" not in line  # closed は出さない
    assert text.count("# @raycast.argument2") == 1 and 'echo "$1 $2"' in text  # 他の行は不変


def test_raycast_script_keeps_its_executable_bit(vault: Path, raycast: Path) -> None:
    seed(vault)
    assert run(vault, "--apply", "--accept-list-changes", "--raycast-script", str(raycast)) == 0
    assert stat.S_IMODE(raycast.stat().st_mode) == 0o755


@pytest.mark.parametrize("managed_lines", [0, 2], ids=["none", "two"])
def test_raycast_script_must_carry_exactly_one_managed_line_and_nothing_is_written(
    vault: Path, tmp_path: Path, managed_lines: int
) -> None:
    seed(vault)
    script = tmp_path / "bad.sh"
    line = '# @raycast.argument2 {"type": "dropdown", "placeholder": "project", "data": []}\n'
    script.write_text("#!/bin/bash\n" + line * managed_lines + "echo hi\n", encoding="utf-8")
    before_vault, before_script = snapshot(vault), script.read_bytes()
    assert run(vault, "--apply", "--accept-list-changes", "--raycast-script", str(script)) == 3
    assert snapshot(vault) == before_vault and script.read_bytes() == before_script  # 全出力を揃えられなければ何も書かない


def test_no_active_project_is_an_error_for_the_raycast_output(vault: Path, raycast: Path) -> None:
    note_with(vault, "proj_a", "client: A\nstatus: closed\n")
    for key in ("proj_b",):
        (vault / "project" / f"project_{key}.md").unlink()
    before = snapshot(vault)
    assert run(vault, "--apply", "--accept-list-changes", "--raycast-script", str(raycast)) == 3
    assert snapshot(vault) == before


def test_without_a_raycast_script_the_output_is_reported_as_skipped(vault: Path, capsys) -> None:
    seed(vault)
    assert run(vault, "--apply", "--accept-list-changes") == 0
    assert "RAYCAST skipped" in capsys.readouterr().out


# ---- 部分出力を残さない -------------------------------------------------------------


def test_io_failure_midway_restores_every_output(vault: Path, raycast: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    seed(vault)
    before_vault, before_script = snapshot(vault), raycast.read_bytes()
    real_replace = os.replace
    calls = {"n": 0}

    def flaky(src, dst, *a, **k):
        calls["n"] += 1
        if calls["n"] == 4:  # project ノート 2 本 + list を書いたあと Raycast で失敗
            raise OSError("simulated Drive I/O failure")
        return real_replace(src, dst, *a, **k)

    monkeypatch.setattr(os, "replace", flaky)
    assert run(vault, "--apply", "--accept-list-changes", "--raycast-script", str(raycast)) == 3
    monkeypatch.undo()
    assert snapshot(vault) == before_vault and raycast.read_bytes() == before_script


def test_project_key_filter_limits_only_the_regions(vault: Path) -> None:
    seed(vault)
    before_b = (vault / "project/project_proj_b.md").read_bytes()
    assert run(vault, "--project-key", "proj_a", "--apply", "--accept-list-changes") == 0
    assert (vault / "project/project_proj_b.md").read_bytes() == before_b  # 領域は絞る
    assert {r[0] for r in rows(vault)} == {"proj_a", "proj_b"}  # 一覧は全体


def test_dry_run_writes_nothing_and_reports_all_three_outputs(vault: Path, raycast: Path, capsys) -> None:
    seed(vault)
    before, before_script = snapshot(vault), raycast.read_bytes()
    assert run(vault, "--accept-list-changes", "--raycast-script", str(raycast)) == 0
    out = capsys.readouterr().out
    assert "LIST " in out and "RAYCAST " in out and "RENDER project=proj_a" in out and "DRY_RUN" in out
    assert snapshot(vault) == before and raycast.read_bytes() == before_script
