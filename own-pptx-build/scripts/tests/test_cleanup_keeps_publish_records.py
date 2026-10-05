"""cleanup_deck.py は vault 公開の記録（deck_meta.json・digest.md・speaker_notes.md）を消さない。

これらは edit-mode の再納品で確定値を再利用するための再構築ソース。消すと、再納品のたびに
opt-out と project を聞き直すことになる（SPEC-document-publish）。
"""

import importlib.util
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "cleanup_deck.py"
SPEC = importlib.util.spec_from_file_location("cleanup_deck", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


def test_publish_records_are_kept_and_intermediates_are_still_removed(tmp_path: Path) -> None:
    deck = tmp_path / "20260101_demo"
    process = deck / "process"
    process.mkdir(parents=True)
    keep = ["deck_meta.json", "digest.md", "speaker_notes.md", "outline.md", "feedback.json"]
    for name in keep:
        (process / name).write_text("x", encoding="utf-8")
    (process / "mockup_01.png").write_bytes(b"x")
    (process / "preview-1.png").write_bytes(b"x")

    targeted = {Path(t).name for t in MODULE.targets(str(deck))}

    assert {"mockup_01.png", "preview-1.png"} <= targeted  # 検査が空振りしていないこと
    assert not targeted & set(keep)
