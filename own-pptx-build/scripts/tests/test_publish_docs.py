"""ISSUE-04（WS-20261002-document-publish）の受入: pipeline.md の①と⑤に vault 公開の工程がある。

文書の一致検査なので、見出しの節ごとに切り出して「その節の中に」あることを見る
（ファイル全体で grep すると、別の節の記述で通ってしまう）。
"""

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PIPELINE = (ROOT / "references" / "pipeline.md").read_text(encoding="utf-8")
SKILL = (ROOT / "SKILL.md").read_text(encoding="utf-8")


def section(text: str, heading_prefix: str) -> str:
    pattern = re.compile(rf"^## {re.escape(heading_prefix)}.*?(?=^## |\Z)", re.S | re.M)
    match = pattern.search(text)
    assert match, f"section not found: {heading_prefix}"
    return match.group(0)


def test_step1_asks_opt_out_and_records_deck_meta() -> None:
    step1 = section(PIPELINE, "① ")
    assert "opt-out" in step1
    assert "process/deck_meta.json" in step1
    assert "vault_publish" in step1 and "opted_out" in step1
    assert "process/digest.md" in step1


def test_step5_calls_attach_document_with_confirmation_before_cleanup() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "attach-document" in step5
    assert "--apply" in step5
    # 提案（key と確率。assess 未実装なら catalog 一覧）を人間に見せて確定してから apply する
    assert "提案" in step5 and "catalog" in step5
    assert step5.index("attach-document") < step5.index("cleanup_deck.py")
    assert "export_speaker_notes.py" in step5


def test_step5_reuses_recorded_choice_on_redelivery_and_backfills_missing_record() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "再納品" in step5 and "聞かない" in step5
    assert "記録が無い" in step5


def test_deck_meta_never_records_local_paths() -> None:
    step1 = section(PIPELINE, "① ")
    assert "ローカルパス" in step1


def test_skill_table_mentions_both_steps() -> None:
    row1 = next(l for l in SKILL.splitlines() if l.startswith("| ① "))
    row5 = next(l for l in SKILL.splitlines() if l.startswith("| ⑤ "))
    assert "opt-out" in row1
    assert "attach-document" in row5


def test_step1_records_source_repo_and_project_confirmation_state() -> None:
    step1 = section(PIPELINE, "① ")
    assert "source_repo" in step1
    assert "project_confirmed" in step1 and "project_hint" in step1


def test_step5_derives_source_repo_from_remote_not_the_worktree_directory() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "remote get-url origin" in step5
    assert "worktree" in step5 and "ディレクトリ名" in step5


def test_step5_proposal_dry_run_omits_project_and_explains_success_outputs() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert "NOOP no changes required" in step5  # 再実行の成功は APPLIED が出ない
    assert "ASSESS skipped" in step5 and "ASSOCIATION undecided" in step5
    assert "`--project` を付けずに" in step5  # 提案を見せる dry-run には --project を付けない


def test_step5_tells_how_to_find_the_vault_without_reading_env() -> None:
    step5 = section(PIPELINE, "⑤ ")
    assert ".env" in step5 and "人間に確認" in step5
