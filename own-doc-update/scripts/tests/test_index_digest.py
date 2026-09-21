"""index_digest.py — the SessionStart hook must inject something that arrives.

Measured 2026-09-19 across every session on this machine: Claude Code hands a
hook's output to the agent inline only up to about 10,000 CHARACTERS (largest
delivered inline 8,957; smallest spilled 10,019, and that spilled one was this
very index injection). Past it the output is written to a file and the agent
gets a 2 KB preview, so "read this first" runs against nothing and nothing
reports a failure.

Every test here is about that one property: what the hook prints must fit, must
still route to every document, and must say what it left out.
"""
import importlib.util
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))
SPEC = importlib.util.spec_from_file_location("index_digest", SCRIPTS / "index_digest.py")
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)

HOOK = SCRIPTS / "hooks" / "session_start.sh"


def index(front_matter: str = "", body: str = "") -> str:
    return f"---\n{front_matter}---\n\n# 00 Index\n\n{body}"


class DigestFitsTest(unittest.TestCase):
    def test_a_real_sized_index_is_cut_to_the_budget(self):
        rows = "".join(
            f"- [ISSUE-2026{n:04d}-a-fairly-long-slug-like-the-real-ones]"
            f"(issues/ISSUE-2026{n:04d}-a-fairly-long-slug-like-the-real-ones.md) — "
            + "説明が延々と続く " * 20
            + "\n"
            for n in range(90)
        )
        text = index("updated_at: 2026-09-20\ncurrent_focus: " + "物語 " * 3000 + "\n",
                     "## Active Issues\n\n" + rows)
        self.assertGreater(len(text), 20000)

        digest = MODULE.build_digest(text)

        self.assertLessEqual(len(digest), MODULE.DIGEST_MAX_CHARS)
        self.assertLess(MODULE.DIGEST_MAX_CHARS, 8957, "budget must sit under the measured inline maximum")

    def test_the_budget_is_spent_on_routing_not_on_narrative(self):
        text = index("updated_at: 2026-09-20\n",
                     "## Active Issues\n\n"
                     "- [ISSUE-a](issues/ISSUE-a.md) — one line\n\n"
                     + "進捗の物語がここに延々と書かれている。\n" * 200)

        digest = MODULE.build_digest(text)

        self.assertIn("issues/ISSUE-a.md", digest)
        self.assertNotIn("進捗の物語", digest)


class EveryDocumentStaysReachableTest(unittest.TestCase):
    def test_every_link_target_survives_when_descriptions_can_be_cut(self):
        rows = "".join(
            f"- [ISSUE-{n:03d}](issues/ISSUE-{n:03d}.md) — " + "説明" * 120 + "\n"
            for n in range(40)
        )
        digest = MODULE.build_digest(index("updated_at: 2026-09-20\n", "## Active Issues\n\n" + rows))

        for n in range(40):
            self.assertIn(f"issues/ISSUE-{n:03d}.md", digest)

    def test_a_link_inside_prose_is_routed_too(self):
        # The real index lists some entries as prose paragraphs, not bullets.
        text = index(
            "updated_at: 2026-09-20\n",
            "## Active Issues\n\n"
            "consumer feedback（2026-08-09 の実測）— "
            "[docs/issues/ISSUE-p](issues/ISSUE-p.md): 受け渡しで情報が落ちる\n",
        )

        digest = MODULE.build_digest(text)

        self.assertIn("issues/ISSUE-p.md", digest)

    def test_several_links_on_one_line_all_survive(self):
        text = index(
            "updated_at: 2026-09-20\n",
            "## Active Issues\n\n"
            "- 上記 2 件は [WS-x](workstreams/WS-x.md) と [WS-y](workstreams/WS-y.md) が束ねた\n",
        )

        digest = MODULE.build_digest(text)

        self.assertIn("workstreams/WS-x.md", digest)
        self.assertIn("workstreams/WS-y.md", digest)


class WhatIsLeftOutIsStatedTest(unittest.TestCase):
    def test_archive_rows_go_first_and_the_count_is_printed(self):
        rows = "".join(
            f"- [ISSUE-{n:03d}](issues/archive/ISSUE-{n:03d}.md) — done\n" for n in range(120)
        ) + "- [ISSUE-live](issues/ISSUE-live.md) — still open\n"
        text = index("updated_at: 2026-09-20\n", "## Active Issues\n\n" + rows)

        digest = MODULE.build_digest(text, max_chars=1200)

        self.assertLessEqual(len(digest), 1200)
        self.assertIn("issues/ISSUE-live.md", digest, "current work must outlive archived work")
        self.assertRegex(digest, r"archive[^\n]*\d+")

    def test_current_work_keeps_its_description_before_history_keeps_its_path(self):
        """What the budget buys, in order: current work with meaning, then history.

        The first run against the real index spent all 8,000 characters listing
        94 paths with no descriptions at all, 42 of them archived. A session
        opens on current work; archived documents are named in the full file and
        are one directory listing away.
        """
        rows = "".join(
            f"- [ISSUE-{n:03d}](issues/archive/ISSUE-{n:03d}.md) — 完了・archive 済み。"
            + "結論の説明" * 12 + "\n"
            for n in range(60)
        ) + "".join(
            f"- [ISSUE-live-{n}](issues/ISSUE-live-{n}.md) — いま動いている作業の説明。"
            + "詳しい事情" * 12 + "\n"
            for n in range(12)
        )
        text = index("updated_at: 2026-09-20\n", "## Active Issues\n\n" + rows)

        digest = MODULE.build_digest(text, max_chars=3000)

        self.assertLessEqual(len(digest), 3000)
        for n in range(12):
            self.assertIn(f"issues/ISSUE-live-{n}.md", digest)
        self.assertIn("いま動いている作業の説明", digest)
        self.assertRegex(digest, r"archive[^\n]*\d+")

    def test_the_full_file_is_always_named(self):
        digest = MODULE.build_digest(index("updated_at: 2026-09-20\n", "## Specs\n\n- [s](specs/s.md) — x\n"))
        self.assertIn("docs/00_index.md", digest)

    def test_dropping_rows_reports_how_many(self):
        rows = "".join(f"- [ISSUE-{n:03d}](issues/ISSUE-{n:03d}.md) — open\n" for n in range(200))
        digest = MODULE.build_digest(index("updated_at: 2026-09-20\n", "## Active Issues\n\n" + rows),
                                     max_chars=900)

        self.assertLessEqual(len(digest), 900)
        self.assertRegex(digest, r"\d+ 件")


class FrontMatterAndPolicyTest(unittest.TestCase):
    def test_a_huge_current_focus_is_shortened_not_dropped(self):
        text = index("updated_at: 2026-09-20\ncurrent_focus: " + "物語 " * 3000 + "\n",
                     "## Specs\n\n- [s](specs/s.md) — x\n")

        digest = MODULE.build_digest(text)

        self.assertIn("updated_at: 2026-09-20", digest)
        self.assertIn("current_focus", digest)
        self.assertIn("物語", digest)
        self.assertLessEqual(len(digest), MODULE.DIGEST_MAX_CHARS)

    def test_the_hygiene_pointer_suffix_does_not_eat_the_budget(self):
        # docs_hygiene.py ends a shortened row with "…（全文: log/index-YYYYMM.md）".
        # The digest already names where the full text lives, so repeating it on
        # every row spends the budget on the same sentence 50 times.
        text = index(
            "updated_at: 2026-09-20\n",
            "## Active Issues\n\n"
            "- [ISSUE-a](issues/ISSUE-a.md) — 上流待ちで止まっている …（全文: log/index-202609.md）\n",
        )

        digest = MODULE.build_digest(text)

        self.assertIn("上流待ちで止まっている", digest)
        self.assertNotIn("全文: log/index-202609.md", digest)

    def test_a_section_with_no_links_is_policy_and_is_kept(self):
        # Read Policy tells the agent how to read; it carries no link and must
        # not be discarded with the progress narrative.
        text = index("updated_at: 2026-09-20\n",
                     "## Read Policy\n\nRead this file first. 完了した行は消さない。\n\n"
                     "## Specs\n\n- [s](specs/s.md) — x\n")

        digest = MODULE.build_digest(text)

        self.assertIn("Read this file first", digest)
        self.assertIn("完了した行は消さない", digest)


class OnlyCompressWhenItHasToTest(unittest.TestCase):
    """A small index arrives as itself; compressing it would lose fidelity for nothing."""

    def test_an_index_that_fits_is_passed_through_unchanged(self):
        text = index("updated_at: 2026-09-20\n",
                     "## Read Policy\n\nRead this file first.\n\n"
                     "## Specs\n\n- [s](specs/s.md) — **bold** description with `code`\n")

        block = MODULE.render_injection(text)

        self.assertIn(text, block)
        self.assertIn("**bold**", block)
        self.assertNotIn("routing digest", block)

    def test_an_index_that_does_not_fit_arrives_as_a_marked_digest(self):
        rows = "".join(
            f"- [ISSUE-{n:03d}](issues/ISSUE-{n:03d}.md) — " + "説明" * 120 + "\n"
            for n in range(60)
        )
        block = MODULE.render_injection(index("updated_at: 2026-09-20\n",
                                              "## Active Issues\n\n" + rows))

        self.assertIn("routing digest", block)
        self.assertLessEqual(len(block), MODULE.DIGEST_MAX_CHARS)
        self.assertIn("issues/ISSUE-059.md", block)


class NeverErrorsTest(unittest.TestCase):
    def test_an_empty_index_produces_something_short_and_valid(self):
        digest = MODULE.build_digest("")
        self.assertIn("docs/00_index.md", digest)
        self.assertLessEqual(len(digest), MODULE.DIGEST_MAX_CHARS)

    def test_an_index_without_front_matter_still_routes(self):
        digest = MODULE.build_digest("# 00 Index\n\n## Specs\n\n- [s](specs/s.md) — x\n")
        self.assertIn("specs/s.md", digest)

    def test_the_same_input_gives_the_same_output(self):
        text = index("updated_at: 2026-09-20\n", "## Specs\n\n- [s](specs/s.md) — x\n")
        self.assertEqual(MODULE.build_digest(text), MODULE.build_digest(text))


class HookTest(unittest.TestCase):
    """The hook's contract: fast, never errors, silent when there is nothing."""

    def run_hook(self, cwd: Path) -> subprocess.CompletedProcess:
        return subprocess.run(
            ["bash", str(HOOK)],
            input=f'{{"cwd": "{cwd}"}}',
            capture_output=True,
            text=True,
            timeout=20,
        )

    def test_a_repo_without_an_index_stays_silent(self):
        with tempfile.TemporaryDirectory() as directory:
            result = self.run_hook(Path(directory))
            self.assertEqual(result.returncode, 0)
            self.assertEqual(result.stdout.strip(), "")

    def test_a_large_index_is_injected_as_a_digest_that_fits(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "docs").mkdir()
            rows = "".join(
                f"- [ISSUE-{n:03d}](issues/ISSUE-{n:03d}.md) — " + "説明" * 120 + "\n"
                for n in range(60)
            )
            (root / "docs/00_index.md").write_text(
                index("updated_at: 2026-09-20\ncurrent_focus: " + "物語 " * 2000 + "\n",
                      "## Active Issues\n\n" + rows)
            )

            result = self.run_hook(root)

            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertLessEqual(len(result.stdout), 8957,
                                 "the whole hook output, not just the digest, has to arrive")
            self.assertIn("docs/00_index.md", result.stdout)
            self.assertIn("issues/ISSUE-059.md", result.stdout)


if __name__ == "__main__":
    unittest.main()
