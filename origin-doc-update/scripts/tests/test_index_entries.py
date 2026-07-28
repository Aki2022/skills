import importlib.util
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "index_entries.py"
SPEC = importlib.util.spec_from_file_location("index_entries", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class RemoveIndexEntryTest(unittest.TestCase):
    """00_index.md の active 参照行を、実際に使われている記法すべてで消せること。

    00_index.template.md はベアパス記法 (`- docs/issues/X.md — ...`) を示すが、
    実運用の index は markdown link 記法に育っていることが多い。片方しか消せないと
    archive 後も index に完了済み issue が残り、毎回手で消す羽目になる。
    """

    def remove(self, content: str, entry_id: str = "ISSUE-20260728-foo"):
        return MODULE.remove_index_entry(content, "docs/issues", entry_id)

    def test_removes_bare_path_form(self):
        content = "## Active Issues\n\n- docs/issues/ISSUE-20260728-foo.md — 未完了\n"
        new, removed = self.remove(content)
        self.assertEqual(removed, 1)
        self.assertNotIn("ISSUE-20260728-foo", new)

    def test_removes_link_with_full_path_label(self):
        content = (
            "## Current Focus\n\n"
            "- [docs/issues/ISSUE-20260728-foo.md](issues/ISSUE-20260728-foo.md) — 説明\n"
        )
        new, removed = self.remove(content)
        self.assertEqual(removed, 1)
        self.assertNotIn("ISSUE-20260728-foo", new)

    def test_removes_link_with_id_label(self):
        content = (
            "## Active Issues\n\n"
            "- [ISSUE-20260728-foo](issues/ISSUE-20260728-foo.md) — 未着手\n"
        )
        new, removed = self.remove(content)
        self.assertEqual(removed, 1)
        self.assertNotIn("ISSUE-20260728-foo", new)

    def test_removes_every_section_referencing_the_entry(self):
        content = (
            "## Current Focus\n\n"
            "- [docs/issues/ISSUE-20260728-foo.md](issues/ISSUE-20260728-foo.md) — 説明\n\n"
            "## Active Issues\n\n"
            "- [ISSUE-20260728-foo](issues/ISSUE-20260728-foo.md) — 未着手\n"
        )
        new, removed = self.remove(content)
        self.assertEqual(removed, 2)
        self.assertNotIn("ISSUE-20260728-foo", new)

    def test_keeps_other_entries(self):
        content = (
            "- [ISSUE-20260728-foo](issues/ISSUE-20260728-foo.md) — 消える\n"
            "- [ISSUE-20260728-bar](issues/ISSUE-20260728-bar.md) — 残る\n"
        )
        new, removed = self.remove(content)
        self.assertEqual(removed, 1)
        self.assertIn("ISSUE-20260728-bar", new)

    def test_does_not_match_longer_id_sharing_a_prefix(self):
        content = "- [ISSUE-20260728-foo-extra](issues/ISSUE-20260728-foo-extra.md) — 残る\n"
        new, removed = self.remove(content)
        self.assertEqual(removed, 0)
        self.assertEqual(new, content)

    def test_keeps_prose_mention_that_is_not_a_link_target(self):
        content = (
            "- [ISSUE-20260728-bar](issues/ISSUE-20260728-bar.md) — "
            "後続。ISSUE-20260728-foo の判断待ち\n"
        )
        new, removed = self.remove(content)
        self.assertEqual(removed, 0)
        self.assertEqual(new, content)

    def test_reports_zero_when_absent(self):
        content = "## Active Issues\n\n- [ISSUE-other](issues/ISSUE-other.md) — 未完了\n"
        new, removed = self.remove(content)
        self.assertEqual(removed, 0)
        self.assertEqual(new, content)

    def test_handles_workstream_directory(self):
        content = "- [WS-20260728-x](workstreams/WS-20260728-x.md) — 進行中\n"
        new, removed = MODULE.remove_index_entry(
            content, "docs/workstreams", "WS-20260728-x"
        )
        self.assertEqual(removed, 1)
        self.assertEqual(new, "")


if __name__ == "__main__":
    unittest.main()
