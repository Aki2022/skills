"""prepare_publish_materials.py: ⑤の vault 公開素材（公開用 outline・題・出典リポジトリ名）を作る。

スライド見出しやキーメッセージは解析しない（デッキごとに書式が違い、解析すると黙って間違えるため）。
outline.md はそのまま vault に載せ、note の summary は outline の題 1 行にする。
opt-out の記録は outline.md 先頭の `<!-- vault_publish: publish|opted_out -->` 1 行だけ。
"""

import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "prepare_publish_materials.py"

OUTLINE = """<!-- vault_publish: publish -->
# outline.md — デモ提案スライド（全3枚）

説明文。

# スライド1: 表紙

# スライド2: 現状

## キーメッセージ（1行）

現状は三つの課題を抱える
"""


def run(outline: Path, out_dir: Path, *extra: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(outline), "--out-dir", str(out_dir), *extra],
        capture_output=True,
        text=True,
    )


def write(tmp_path: Path, text: str, name: str = "outline.md") -> Path:
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


def git(cwd: Path, *args: str) -> None:
    subprocess.run(["git", "-C", str(cwd), *args], check=True, capture_output=True)


@pytest.fixture
def repo(tmp_path: Path) -> Path:
    """deck repo の代わり。remote の URL だけで足りるので commit は作らない。"""
    root = tmp_path / "deck_checkout"
    root.mkdir()
    git(root, "init", "-q")
    git(root, "remote", "add", "origin", "https://example.com/org/deck_repo.git")
    return root


def test_outputs_title_choice_and_repo_as_json_and_writes_only_the_published_outline(
    repo: Path, tmp_path: Path
) -> None:
    out = tmp_path / "out"
    result = run(write(repo, OUTLINE), out)
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout) == {
        "vault_publish": "publish",
        "summary": "デモ提案スライド（全3枚）",
        "source_repo": "deck_repo",
    }
    assert sorted(p.name for p in out.iterdir()) == ["outline.md"]  # digest は作らない
    assert (out / "outline.md").read_text(encoding="utf-8") == OUTLINE.split("\n", 1)[1]


def test_slide_headings_of_any_format_never_change_the_result(repo: Path, tmp_path: Path) -> None:
    """レビューで見つかった「見出しの書式次第で黙って別物になる」問題を、解析をやめて塞ぐ。"""
    weird = OUTLINE + "\n# スライド7b: 付録\n\n## キーメッセージ\n\n付録の主張\n\n# S9:\n\n# Slide 1｜x\n"
    plain = run(write(repo, OUTLINE), tmp_path / "a")
    odd = run(write(repo, weird), tmp_path / "b")
    assert json.loads(plain.stdout) == json.loads(odd.stdout)
    assert "付録の主張" in (tmp_path / "b" / "outline.md").read_text(encoding="utf-8")  # 本文は欠けずに載る


def test_opted_out_writes_nothing_and_reports_the_choice(repo: Path, tmp_path: Path) -> None:
    out = tmp_path / "out"
    result = run(write(repo, OUTLINE.replace("publish -->", "opted_out -->", 1)), out)
    assert result.returncode == 0
    assert json.loads(result.stdout)["vault_publish"] == "opted_out"
    assert not out.exists()


def test_missing_marker_is_reported_as_null_so_step5_can_ask(repo: Path, tmp_path: Path) -> None:
    result = run(write(repo, OUTLINE.split("\n", 1)[1]), tmp_path / "out")
    assert json.loads(result.stdout)["vault_publish"] is None


def test_bom_is_ignored_for_the_marker_and_the_title(repo: Path, tmp_path: Path) -> None:
    path = repo / "outline.md"
    path.write_bytes(b"\xef\xbb\xbf" + OUTLINE.encode("utf-8"))
    result = run(path, tmp_path / "out")
    assert json.loads(result.stdout)["vault_publish"] == "publish"
    assert b"\xef\xbb\xbf" not in (tmp_path / "out" / "outline.md").read_bytes()


def test_marker_prepended_in_front_of_a_bom_title_still_finds_the_title(repo: Path, tmp_path: Path) -> None:
    """⑤が既存デッキの先頭へ行を足すと、BOM 付きの H1 が 2 行目になる。"""
    body = "# outline.md — BOM の題\n".encode("utf-8")
    path = repo / "outline.md"
    path.write_bytes(b"<!-- vault_publish: publish -->\n\xef\xbb\xbf" + body)
    assert json.loads(run(path, tmp_path / "out").stdout)["summary"] == "BOM の題"


def test_two_markers_are_an_error(repo: Path, tmp_path: Path) -> None:
    text = OUTLINE + "\n<!-- vault_publish: opted_out -->\n"
    result = run(write(repo, text), tmp_path / "out")
    assert result.returncode != 0 and "vault_publish" in result.stderr


@pytest.mark.parametrize("value", ["maybe", "Publish", "publish,"])
def test_invalid_marker_value_is_an_error(repo: Path, tmp_path: Path, value: str) -> None:
    result = run(write(repo, OUTLINE.replace("publish -->", f"{value} -->", 1)), tmp_path / "out")
    assert result.returncode != 0 and "vault_publish" in result.stderr


def test_missing_title_needs_a_fallback_instead_of_picking_a_random_heading(repo: Path, tmp_path: Path) -> None:
    text = "<!-- vault_publish: publish -->\n本文だけ。\n\n# スライド1: 表紙\n"
    outline = write(repo, text)
    assert run(outline, tmp_path / "a").returncode != 0
    result = run(outline, tmp_path / "b", "--fallback-title", "20260105_demo")
    assert json.loads(result.stdout)["summary"] == "20260105_demo"


@pytest.mark.parametrize(
    "url,expected",
    [
        ("https://example.com/org/deck_repo.git", "deck_repo"),
        ("https://example.com/org/deck_repo", "deck_repo"),
        ("https://example.com/org/deck_repo/", "deck_repo"),
        ("git@example.com:org/deck_repo.git", "deck_repo"),
        ("git@example.com:deck_repo.git", "deck_repo"),
    ],
)
def test_source_repo_comes_from_the_remote_url_in_every_common_form(
    repo: Path, tmp_path: Path, url: str, expected: str
) -> None:
    git(repo, "remote", "set-url", "origin", url)
    result = run(write(repo, OUTLINE), tmp_path / "out")
    assert json.loads(result.stdout)["source_repo"] == expected


def test_source_repo_without_a_remote_is_the_main_checkout_directory_name(tmp_path: Path) -> None:
    root = tmp_path / "plain_checkout"
    root.mkdir()
    git(root, "init", "-q")
    result = run(write(root, OUTLINE), tmp_path / "out")
    assert json.loads(result.stdout)["source_repo"] == "plain_checkout"


def test_not_a_git_repository_is_an_error(tmp_path: Path) -> None:
    result = run(write(tmp_path, OUTLINE), tmp_path / "out")
    assert result.returncode != 0 and "git" in result.stderr


def test_missing_file_is_an_error(repo: Path, tmp_path: Path) -> None:
    assert run(repo / "nope.md", tmp_path / "out").returncode != 0
