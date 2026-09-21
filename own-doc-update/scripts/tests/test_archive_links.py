"""Archiving must not leave relative links behind, in either direction.

Measured seven times in one session before this existed: moving a document into
`archive/` puts it one directory deeper, so every relative link *inside* it is
one `../` short, and every link *to* it from elsewhere points at a path that no
longer exists. Neither was rewritten, and nothing failed at the time -- the
breakage surfaced only when someone followed a link, or when the validator's
link check was run by hand.
"""

import sys
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))

import archive_links as LINKS


class ReanchorTest(unittest.TestCase):
    """Links inside the document that moved."""

    def test_a_parent_link_gains_one_level(self):
        moved = LINKS.reanchor(
            "see [ws](../workstreams/WS-1.md)",
            old_dir="docs/issues",
            new_dir="docs/issues/archive",
        )
        self.assertEqual(moved, "see [ws](../../workstreams/WS-1.md)")

    def test_a_sibling_link_becomes_a_parent_link(self):
        moved = LINKS.reanchor(
            "see [other](ISSUE-2.md)",
            old_dir="docs/issues",
            new_dir="docs/issues/archive",
        )
        self.assertEqual(moved, "see [other](../ISSUE-2.md)")

    def test_an_absolute_or_external_destination_is_untouched(self):
        text = "[a](https://example.com/x.md) [b](/etc/x.md) [c](#anchor)"
        self.assertEqual(
            LINKS.reanchor(text, old_dir="docs/issues", new_dir="docs/issues/archive"),
            text,
        )

    def test_a_link_inside_an_inline_code_span_is_untouched(self):
        text = "write `[x](../placeholder.md)` to opt out"
        self.assertEqual(
            LINKS.reanchor(text, old_dir="docs/issues", new_dir="docs/issues/archive"),
            text,
        )

    def test_an_anchor_and_a_title_survive_the_rewrite(self):
        moved = LINKS.reanchor(
            '[a](../specs/prd.md#goals "Title")',
            old_dir="docs/issues",
            new_dir="docs/issues/archive",
        )
        self.assertEqual(moved, '[a](../../specs/prd.md#goals "Title")')


class ReferrerTest(unittest.TestCase):
    """Links from other documents to the one that moved."""

    def test_a_referrer_is_repointed_at_the_archive(self):
        rewritten = LINKS.repoint(
            "see [i](../issues/ISSUE-1.md)",
            referrer_dir="docs/workstreams",
            old_path="docs/issues/ISSUE-1.md",
            new_path="docs/issues/archive/ISSUE-1.md",
        )
        self.assertEqual(rewritten, "see [i](../issues/archive/ISSUE-1.md)")

    def test_a_referrer_at_another_depth_is_repointed_too(self):
        rewritten = LINKS.repoint(
            "see [i](../../issues/ISSUE-1.md)",
            referrer_dir="docs/workstreams/archive",
            old_path="docs/issues/ISSUE-1.md",
            new_path="docs/issues/archive/ISSUE-1.md",
        )
        self.assertEqual(rewritten, "see [i](../../issues/archive/ISSUE-1.md)")

    def test_a_link_to_a_different_document_is_untouched(self):
        text = "see [other](../issues/ISSUE-2.md)"
        self.assertEqual(
            LINKS.repoint(
                text,
                referrer_dir="docs/workstreams",
                old_path="docs/issues/ISSUE-1.md",
                new_path="docs/issues/archive/ISSUE-1.md",
            ),
            text,
        )

    def test_nothing_to_do_returns_the_text_unchanged_and_is_detectable(self):
        text = "no links here"
        self.assertIs(
            LINKS.repoint(
                text,
                referrer_dir="docs",
                old_path="docs/issues/ISSUE-1.md",
                new_path="docs/issues/archive/ISSUE-1.md",
            ),
            text,
        )


if __name__ == "__main__":
    unittest.main()


class PlanTest(unittest.TestCase):
    """The whole move, over a real directory tree."""

    def _repo(self, tmp):
        root = Path(tmp)
        (root / "docs" / "issues" / "archive").mkdir(parents=True)
        (root / "docs" / "workstreams").mkdir(parents=True)
        (root / "docs" / "specs").mkdir(parents=True)
        (root / "docs" / "specs" / "prd.md").write_text("# prd\n")
        (root / "docs" / "issues" / "ISSUE-1.md").write_text(
            "see [prd](../specs/prd.md) and [ws](../workstreams/WS-1.md)\n"
        )
        (root / "docs" / "workstreams" / "WS-1.md").write_text(
            "see [i](../issues/ISSUE-1.md) and [other](../issues/ISSUE-2.md)\n"
        )
        (root / "docs" / "issues" / "ISSUE-2.md").write_text("nothing\n")
        return root

    def test_both_directions_are_planned(self):
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            root = self._repo(tmp)
            content = (root / "docs" / "issues" / "ISSUE-1.md").read_text()
            reanchored, updates = LINKS.plan_link_updates(
                str(root),
                "docs/issues/ISSUE-1.md",
                "docs/issues/archive/ISSUE-1.md",
                content,
            )
            self.assertIn("../../specs/prd.md", reanchored)
            self.assertIn("../../workstreams/WS-1.md", reanchored)

            self.assertEqual(len(updates), 1, updates)
            path, original, rewritten = updates[0]
            self.assertTrue(path.endswith("WS-1.md"))
            self.assertIn("../issues/archive/ISSUE-1.md", rewritten)
            self.assertIn("../issues/ISSUE-2.md", rewritten)
            self.assertNotEqual(original, rewritten)

    def test_only_the_links_to_the_moved_document_change(self):
        """Moving ISSUE-2 must repoint the link to IT and leave its neighbour alone.

        This test first asserted no updates at all, and the planner was right and
        the test was wrong: `WS-1.md` links to both issues.
        """
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            root = self._repo(tmp)
            content = (root / "docs" / "issues" / "ISSUE-2.md").read_text()
            _reanchored, updates = LINKS.plan_link_updates(
                str(root),
                "docs/issues/ISSUE-2.md",
                "docs/issues/archive/ISSUE-2.md",
                content,
            )
            self.assertEqual(len(updates), 1, updates)
            _path, _original, rewritten = updates[0]
            self.assertIn("../issues/archive/ISSUE-2.md", rewritten)
            self.assertIn("../issues/ISSUE-1.md", rewritten)

    def test_a_document_nobody_links_to_plans_no_updates(self):
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            root = self._repo(tmp)
            (root / "docs" / "issues" / "ISSUE-3.md").write_text("alone\n")
            _reanchored, updates = LINKS.plan_link_updates(
                str(root),
                "docs/issues/ISSUE-3.md",
                "docs/issues/archive/ISSUE-3.md",
                "alone\n",
            )
            self.assertEqual(updates, [])
