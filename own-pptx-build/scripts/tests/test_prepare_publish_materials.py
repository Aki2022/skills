"""prepare_publish_materials.py: ⑤の vault 公開素材を outline.md から機械的に作る。

digest.md や deck_meta.json を別ファイルとして残さず、公開のたびに outline.md から作る（LLM を使わない）。
opt-out の記録は outline.md 先頭の `<!-- vault_publish: publish|opted_out -->` 1 行だけ。
"""

import json
import subprocess
import sys
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "prepare_publish_materials.py"

OUTLINE = """<!-- vault_publish: publish -->
# outline.md — デモ提案スライド（全3枚）

説明文。

---

# スライド1: 表紙

型: 表紙

## タイトル（メイン）

デモ提案

# スライド2: 現状

## タイトル（tracker）

現状

## キーメッセージ（1行）

現状は三つの課題を抱える

## 本文

- 詳細。

# スライド3: 効果

## キーメッセージ

導入で三割改善する
"""


def run(outline: Path, out_dir: Path) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(outline), "--out-dir", str(out_dir)],
        capture_output=True,
        text=True,
    )


def write(tmp_path: Path, text: str) -> Path:
    path = tmp_path / "outline.md"
    path.write_text(text, encoding="utf-8")
    return path


def test_digest_has_title_line_then_key_messages_without_the_cover(tmp_path: Path) -> None:
    out = tmp_path / "out"
    result = run(write(tmp_path, OUTLINE), out)
    assert result.returncode == 0, result.stderr
    assert (out / "digest.md").read_text(encoding="utf-8") == (
        "デモ提案スライド（全3枚）\n- S2 現状: 現状は三つの課題を抱える\n- S3 効果: 導入で三割改善する\n"
    )


def test_reports_the_recorded_choice_as_json(tmp_path: Path) -> None:
    result = run(write(tmp_path, OUTLINE), tmp_path / "out")
    assert json.loads(result.stdout) == {"vault_publish": "publish", "slides": 3}


def test_opted_out_is_reported_and_nothing_is_written(tmp_path: Path) -> None:
    out = tmp_path / "out"
    result = run(write(tmp_path, OUTLINE.replace("publish -->", "opted_out -->", 1)), out)
    assert result.returncode == 0
    assert json.loads(result.stdout)["vault_publish"] == "opted_out"
    assert not out.exists()  # opt-out なら vault に出す素材を作らない


def test_missing_marker_is_reported_as_unset_so_step5_can_ask(tmp_path: Path) -> None:
    text = OUTLINE.split("\n", 1)[1]
    result = run(write(tmp_path, text), tmp_path / "out")
    assert json.loads(result.stdout)["vault_publish"] is None


def test_published_outline_drops_the_marker_line_only(tmp_path: Path) -> None:
    out = tmp_path / "out"
    run(write(tmp_path, OUTLINE), out)
    published = (out / "outline.md").read_text(encoding="utf-8")
    assert "vault_publish" not in published
    assert published == OUTLINE.split("\n", 1)[1]


def test_invalid_marker_value_is_an_error(tmp_path: Path) -> None:
    result = run(write(tmp_path, OUTLINE.replace("publish -->", "maybe -->", 1)), tmp_path / "out")
    assert result.returncode != 0
    assert "vault_publish" in result.stderr


def test_outline_without_any_key_message_is_an_error_not_an_empty_digest(tmp_path: Path) -> None:
    text = "<!-- vault_publish: publish -->\n# outline.md — 空\n\n# スライド1: 表紙\n\n本文。\n"
    result = run(write(tmp_path, text), tmp_path / "out")
    assert result.returncode != 0
    assert "key message" in result.stderr


def test_missing_file_is_an_error(tmp_path: Path) -> None:
    assert run(tmp_path / "nope.md", tmp_path / "out").returncode != 0
