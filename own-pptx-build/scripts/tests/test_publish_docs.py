"""ISSUE-04（WS-20261002-document-publish）の受入: pipeline.md の①と⑤に vault 公開の工程がある。

文書の一致検査なので、見出しの節ごとに切り出して「その節の中に」あることを見る
（ファイル全体で grep すると、別の節の記述で通ってしまう）。追加の記録ファイル（digest.md・deck_meta.json）を
作らない設計なので、それらが手順に戻ってこないことも固定する。
"""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PIPELINE = (ROOT / "references" / "pipeline.md").read_text(encoding="utf-8")
SKILL = (ROOT / "SKILL.md").read_text(encoding="utf-8")
EDIT_MODE = (ROOT / "references" / "edit-mode.md").read_text(encoding="utf-8")


def section(text: str, heading_prefix: str) -> str:
    pattern = re.compile(rf"^## {re.escape(heading_prefix)}.*?(?=^## |\Z)", re.S | re.M)
    match = pattern.search(text)
    assert match, f"section not found: {heading_prefix}"
    return match.group(0)


def test_step1_asks_opt_out_once_and_records_it_in_outline_md() -> None:
    step1 = section(PIPELINE, "① ")
    assert "opt-out" in step1
    assert "<!-- vault_publish: publish -->" in step1 and "opted_out" in step1
    assert "outline.md` の先頭 1 行" in step1
    assert "聞かない" in step1  # 既に行があれば聞かない


def test_step5_calls_attach_document_with_confirmation_before_cleanup() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "attach-document" in step5 and "--apply" in step5
    assert "提案" in step5 and "catalog" in step5
    assert step5.index("attach-document") < step5.index("cleanup_deck.py")
    assert "prepare_publish_materials.py" in step5
    assert "export_speaker_notes.py" in step5


def test_step5_proposal_dry_run_omits_project_and_branches_on_document_line() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "--project` も `--apply` も付けずに dry-run" in step5
    assert "DOCUMENT CREATE" in step5 and "DOCUMENT UPDATE" in step5
    assert "聞かない" in step5  # 再納品は聞かない
    assert "記録が無い" in step5  # 行の無い既存デッキは⑤で補って聞く


def test_step5_derives_source_repo_from_remote_not_the_worktree_directory() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "source_repo" in step5 and "git remote" in step5
    assert "worktree" in step5 and "ディレクトリ名" in step5


def test_step5_explains_success_outputs_and_missing_notes() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "NOOP no changes required" in step5
    assert "`--notes-file` を付けずに" in step5  # ノートの無い pptx
    assert ".env" in step5 and "人間に確認" in step5


def test_step5_passes_the_outline_and_a_summary_not_a_digest_file() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert '--summary "' in step5 and '--outline-file "$M/outline.md"' in step5
    assert "--digest-file" not in step5
    assert "--fallback-title" in step5
    assert "mktemp -d" in step5 and "デッキ repo の外" in step5


def test_no_extra_record_files_come_back() -> None:
    for name, text in (("pipeline", PIPELINE), ("skill", SKILL), ("edit-mode", EDIT_MODE)):
        assert "digest.md" not in text, name
        assert "deck_meta" not in text, name
        assert "process/digest.md" not in text, name
        assert "project_confirmed" not in text, name


def test_skill_table_mentions_both_steps() -> None:
    row1 = next(l for l in SKILL.splitlines() if l.startswith("| ① "))
    row5 = next(l for l in SKILL.splitlines() if l.startswith("| ⑤ "))
    assert "opt-out" in row1 and "vault_publish" in row1
    assert "attach-document" in row5
