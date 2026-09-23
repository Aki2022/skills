"""docs_hygiene.py — mechanical fixes are applied, judgment items are only reported.

Every check is exercised red first: the fixture is built broken, the tool must
see it, and the fix must leave the validator with no new errors.
"""
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))
SPEC = importlib.util.spec_from_file_location("docs_hygiene", SCRIPTS / "docs_hygiene.py")
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)
import validate_repo_docs as VALIDATOR  # noqa: E402

ISSUE = """---
schema_version: 2
id: {id}
status: {status}
created_at: 2026-07-01
updated_at: {updated}
branch: {branch}
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: docs only
---

# {id}

## Acceptance

- verify: machine — true

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Moved to docs/issues/archive/
"""


def git(root: Path, *args: str, when: str = "") -> str:
    """Run git; `when` pins BOTH author and committer dates (the tool reads %cs)."""
    env = dict(os.environ, GIT_AUTHOR_NAME="t", GIT_AUTHOR_EMAIL="t@example.com",
               GIT_COMMITTER_NAME="t", GIT_COMMITTER_EMAIL="t@example.com")
    if when:
        env["GIT_AUTHOR_DATE"] = env["GIT_COMMITTER_DATE"] = when
    return subprocess.run(
        ["git", "-C", str(root), *args], check=True, capture_output=True, text=True, env=env
    ).stdout


def plumbing_commit(root: Path, message: str, when: str = "") -> None:
    """フックを通さずに履歴を作る。`git commit` は使わない。

    これが要るのは、テスト対象そのものが「不正な docs が履歴に入っている状態」だから。
    front matter の無いファイルを `--fix` が git 日付から補えることを検査するには、
    front matter の無い状態で commit されていなければならない。一方 pre-commit の
    docs-validator は `docs/00_index.md` を持つ repo で発火し、使い捨ての fixture repo も
    その条件に当たるので、通常の commit では fixture を作れない（実測 2026-09-22）。

    write-tree → commit-tree → update-ref は `--no-verify` でもフック経路の書き換えでも
    ないので、規約の禁止列挙のどれにも当たらない。vibe-guard 自身のテスト
    (reinstall の tests/docs-validator.test.sh) が同じ目的で同じ手を使っている。
    commit-tree は GIT_*_DATE を尊重するので、`git()` と同じ日付固定が効く。
    """
    tree = git(root, "write-tree").strip()
    parent = git(root, "rev-parse", "HEAD").strip()
    commit = git(root, "commit-tree", tree, "-p", parent, "-m", message, when=when).strip()
    git(root, "update-ref", "HEAD", commit)


class HygieneFixture(unittest.TestCase):
    def make_repo(self) -> Path:
        root = Path(tempfile.mkdtemp())
        for path in (
            "docs/adrs", "docs/specs", "docs/issues/archive",
            "docs/workstreams/archive", "docs/guides",
        ):
            (root / path).mkdir(parents=True)
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-09-01\ncurrent_focus: []\n---\n\n# 00 Index\n\n"
            "## Read Policy\n\nRead this file first.\n\n"
            "## Active Issues\n\n"
        )
        git(root, "init", "-q", "-b", "main")
        git(root, "add", ".")
        git(root, "commit", "-q", "-m", "init")
        return root

    def add_issue(self, root: Path, id_: str, status: str, updated: str,
                  branch: str = "", index: bool = True) -> Path:
        path = root / "docs/issues" / f"{id_}.md"
        path.write_text(ISSUE.format(id=id_, status=status, updated=updated,
                                     branch=branch or id_))
        if index:
            with (root / "docs/00_index.md").open("a") as fh:
                fh.write(f"- [{id_}](issues/{id_}.md) — one line\n")
        return path


class ArchiveCompleteTest(HygieneFixture):
    def test_complete_issue_is_archived_and_index_row_removed(self):
        root = self.make_repo()
        self.add_issue(root, "ISSUE-20260801-done", "complete", "2026-08-01")
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertFalse((root / "docs/issues/ISSUE-20260801-done.md").exists())
        self.assertTrue((root / "docs/issues/archive/ISSUE-20260801-done.md").exists())
        self.assertNotIn("ISSUE-20260801-done", (root / "docs/00_index.md").read_text())
        self.assertEqual(report["fixes"]["A1_archived"]["count"], 1)

    def test_archive_blocked_is_reported_not_forced(self):
        root = self.make_repo()
        path = self.add_issue(root, "ISSUE-20260801-half", "complete", "2026-08-01")
        path.write_text(path.read_text().replace("- [x] Guides updated", "- [ ] Guides updated"))
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertTrue(path.exists())
        self.assertEqual(report["fixes"]["A1_archived"]["count"], 0)
        self.assertEqual(len(report["fixes"]["A1_archive_blocked"]["items"]), 1)

    def test_dry_run_changes_nothing(self):
        root = self.make_repo()
        path = self.add_issue(root, "ISSUE-20260801-done", "complete", "2026-08-01")
        before = path.read_text()
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        self.assertEqual(path.read_text(), before)
        self.assertEqual(report["fixes"]["A1_archived"]["count"], 1)  # would-fix count


class IndexNarrativeTest(HygieneFixture):
    def test_long_line_and_prose_block_move_to_log_and_link_survives(self):
        root = self.make_repo()
        self.add_issue(root, "ISSUE-20260901-a", "active", "2026-09-01", index=False)
        long = "- [ISSUE-20260901-a](issues/ISSUE-20260901-a.md) — " + "物語 " * 400
        prose = "\n".join(f"narrative line {i} without any link" for i in range(5))
        index = root / "docs/00_index.md"
        index.write_text(index.read_text() + long + "\n\n## Progress\n\n" + prose + "\n")
        self.assertGreater(len(long), MODULE.INDEX_MAX_LINE_CHARS)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        text = index.read_text()
        self.assertTrue(all(len(line) <= MODULE.INDEX_MAX_LINE_CHARS for line in text.splitlines()))
        self.assertIn("](issues/ISSUE-20260901-a.md)", text)
        self.assertNotIn("narrative line 3", text)
        log = root / "docs/log/index-202609.md"
        self.assertTrue(log.exists())
        self.assertIn("narrative line 3", log.read_text())
        self.assertIn("物語 物語", log.read_text())
        self.assertIn("log/index-202609.md", text)
        self.assertEqual(report["fixes"]["A2_index_lines_moved"]["count"], 6)
        # Idempotent: a second run moves nothing.
        again = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertEqual(again["fixes"]["A2_index_lines_moved"]["count"], 0)

    def test_table_rows_are_shortened_per_cell_not_broken(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        row = "| [ISSUE-x](issues/ISSUE-x.md) | " + "x" * 700 + " | active |"
        index.write_text(index.read_text() + "\n| id | 状態 | s |\n|---|---|---|\n" + row + "\n")
        MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        rows = [l for l in index.read_text().splitlines() if l.startswith("| [ISSUE-x]")]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].count("|"), row.count("|"))
        self.assertLessEqual(len(rows[0]), MODULE.INDEX_MAX_LINE_CHARS)

    def test_short_prose_such_as_read_policy_stays(self):
        root = self.make_repo()
        MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertIn("Read this file first.", (root / "docs/00_index.md").read_text())
        self.assertFalse((root / "docs/log").exists())


class StatusNormalizeTest(HygieneFixture):
    def test_aliases_are_rewritten_and_noted(self):
        root = self.make_repo()
        p1 = self.add_issue(root, "ISSUE-20260901-o", "open", "2026-09-01")
        p2 = self.add_issue(root, "ISSUE-20260901-r", "resolved", "2026-09-01")
        p3 = self.add_issue(root, "ISSUE-20260901-h", "in-progress", "2026-09-01")
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertIn("status: active\n", p1.read_text())
        self.assertIn("hygiene_note: status normalized from open on 2026-09-19", p1.read_text())
        self.assertIn("status: in_progress\n", p3.read_text())
        # resolved → complete → archived in the same run
        self.assertFalse(p2.exists())
        self.assertEqual(report["fixes"]["A3_status_normalized"]["count"], 3)
        errors, _ = VALIDATOR.validate_repo(root)
        self.assertEqual([e for e in errors if "status" in e], [])


class FrontMatterFillTest(HygieneFixture):
    def test_spec_gets_full_envelope_from_git_and_guide_gets_minimal(self):
        root = self.make_repo()
        spec = root / "docs/specs/legacy-thing.md"
        spec.write_text("# Legacy\n\nbody\n")
        guide = root / "docs/guides/old.md"
        guide.write_text("# Old\n")
        git(root, "add", ".")
        plumbing_commit(root, "add", when="2026-06-14T00:00:00")
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        fm = VALIDATOR.parse_front_matter(spec)
        self.assertEqual(fm["id"], "SPEC-legacy-thing")
        self.assertEqual(fm["status"], "active")
        self.assertEqual(fm["created_at"], "2026-06-14")
        self.assertEqual(fm["updated_at"], "2026-06-14")
        self.assertIn("hygiene_note", fm)
        gfm = VALIDATOR.parse_front_matter(guide)
        self.assertEqual(gfm["updated_at"], "2026-06-14")
        self.assertEqual(report["fixes"]["A4_front_matter_added"]["count"], 2)
        errors, _ = VALIDATOR.validate_repo(root)
        self.assertEqual([e for e in errors if "legacy-thing" in e or "old.md" in e], [])


class ReportOnlyTest(HygieneFixture):
    def test_stale_and_orphaned_issues_are_reported_with_counts(self):
        root = self.make_repo()
        old = (date(2026, 9, 19) - timedelta(days=45)).isoformat()
        very_old = (date(2026, 9, 19) - timedelta(days=90)).isoformat()
        self.add_issue(root, "ISSUE-20260701-gone", "in_progress", old, branch="feat/gone")
        self.add_issue(root, "ISSUE-20260601-fresh", "pending", "2026-09-18", branch="main")
        self.add_issue(root, "ISSUE-20260501-dead", "active", very_old, branch="main")
        (root / "docs/guides/g.md").write_text(
            "---\nupdated_at: 2026-09-01\n---\n# G\n\nRun `npm run report` then see "
            "`.github/workflows/daily-report.yml` and `scripts/present.py` and `src/missing/file.ts`.\n"
            # The link baseline below claims this link is broken. It has to actually
            # be broken: a baseline entry that is already resolved is itself an
            # error the validator reports, and the pre-commit docs-validator then
            # refuses to commit the fixture (measured 2026-09-22).
            "\nSee [nowhere](nowhere.md).\n"
            + "".join(f"\n### 2026-08-{d:02d} 決定\n\ntext\n" for d in range(1, 8))
        )
        (root / "scripts").mkdir()
        (root / "scripts/present.py").write_text("")
        (root / "package.json").write_text(json.dumps({"scripts": {"build": "x"}}))
        (root / "docs/setup").mkdir()
        (root / "docs/setup/README.md").write_text("---\nupdated_at: 2026-01-01\n---\n# s\n")
        (root / "docs/validator-link-baseline.txt").write_text("docs/guides/g.md\tnowhere.md\n")
        git(root, "add", ".")
        git(root, "commit", "-q", "-m", "more", when="2026-06-01T00:00:00")

        report = MODULE.run(root, fix=False, report=True, today=date(2026, 9, 19))
        r = report["report"]
        self.assertEqual([i["file"] for i in r["R1_stale_branch_gone"]["items"]],
                         ["docs/issues/ISSUE-20260701-gone.md"])
        self.assertIn("docs/issues/ISSUE-20260501-dead.md",
                      [i["file"] for i in r["R2_untouched_60d"]["items"]])
        dead = {i["ref"] for i in r["R3_dead_references"]["items"]}
        self.assertEqual(dead, {"npm run report", ".github/workflows/daily-report.yml",
                                "src/missing/file.ts"})
        self.assertEqual([i["file"] for i in r["R4_history_in_guide_or_spec"]["items"]],
                         ["docs/guides/g.md"])
        self.assertEqual([i["dir"] for i in r["R5_non_canonical_dirs"]["items"]], ["docs/setup"])
        self.assertEqual(r["R6_baseline_debt"]["items"][0]["count"], 1)
        log = root / "docs/log/hygiene-20260919.md"
        self.assertTrue(log.exists())
        self.assertIn("R3_dead_references", log.read_text())
        self.assertIn("| 0 |", log.read_text().replace("| 0 ", "| 0 "))  # zero rows are printed too


class CliTest(HygieneFixture):
    def test_cli_json_names_the_repo_and_exits_zero(self):
        root = self.make_repo()
        out = subprocess.run(
            [sys.executable, str(SCRIPTS / "docs_hygiene.py"), str(root), "--json"],
            capture_output=True, text=True, check=True,
        ).stdout
        payload = json.loads(out)
        self.assertEqual(payload["repo"], str(root.resolve()))
        self.assertIn("fixes", payload)

    def test_repo_without_docs_index_is_a_clear_noop(self):
        root = Path(tempfile.mkdtemp())
        proc = subprocess.run(
            [sys.executable, str(SCRIPTS / "docs_hygiene.py"), str(root)],
            capture_output=True, text=True,
        )
        self.assertEqual(proc.returncode, 2)
        self.assertIn("00_index.md", proc.stderr)


if __name__ == "__main__":
    unittest.main()


class BudgetAndReferenceTest(HygieneFixture):
    def test_index_over_ceiling_is_shortened_progressively_and_originals_logged(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        rows = "".join(
            f"- [ISSUE-{n:04d}](issues/ISSUE-{n:04d}.md) — " + "説明" * 150 + "\n" for n in range(120)
        )
        index.write_text(index.read_text() + rows)
        self.assertGreater(index.stat().st_size, MODULE.INDEX_MAX_BYTES)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        item = report["fixes"]["A2_index_lines_moved"]
        self.assertLessEqual(index.stat().st_size, MODULE.INDEX_MAX_BYTES)
        self.assertFalse(item["over_ceiling_after_fix"])
        self.assertIn(item["description_cap"], MODULE.DESC_CAPS)
        text = index.read_text()
        self.assertEqual(text.count("](issues/ISSUE-"), 120)
        self.assertIn("説明説明", (root / "docs/log/index-202609.md").read_text())

    def test_long_front_matter_scalar_is_moved_to_log_and_kept_valid(self):
        # 2026-09-19: a real index carried a 12.9 KB `current_focus:` line. The
        # body pass skipped front matter, so --fix left both errors standing.
        root = self.make_repo()
        index = root / "docs/00_index.md"
        story = "物語 " * 1500 + "[ISSUE-x](issues/ISSUE-x.md) と [t](../tests/t.cjs)"
        index.write_text(index.read_text().replace("current_focus: []", f"current_focus: {story}"))
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        text = index.read_text()
        fm = VALIDATOR.parse_front_matter(index)
        self.assertIsNotNone(fm)
        self.assertLessEqual(len(str(fm.get("current_focus"))), MODULE.INDEX_MAX_LINE_CHARS)
        self.assertIn("log/index-202609.md", str(fm.get("current_focus")))
        self.assertTrue(all(len(line) <= MODULE.INDEX_MAX_LINE_CHARS for line in text.splitlines()))
        log = (root / "docs/log/index-202609.md").read_text()
        self.assertIn("current_focus", log)
        self.assertIn("物語 物語", log)
        # links moved one directory deeper still resolve
        self.assertIn("](../issues/ISSUE-x.md)", log)
        self.assertIn("](../../tests/t.cjs)", log)
        self.assertGreaterEqual(report["fixes"]["A2_index_lines_moved"]["count"], 1)
        again = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertEqual(again["fixes"]["A2_index_lines_moved"]["count"], 0)

    def test_budget_pass_recuts_lines_the_split_already_shortened(self):
        # 2026-09-19: rows over 500 chars were cut to 200 by the split and then
        # skipped by the budget pass because they carried the log suffix, so a
        # 39 KB index was reported "cannot be met" at cap 40 while every row
        # still had a 200-char tail.
        root = self.make_repo()
        index = root / "docs/00_index.md"
        rows = "".join(
            f"- [ISSUE-{n:04d}](issues/ISSUE-{n:04d}.md) — " + "説明" * 300 + "\n" for n in range(60)
        )
        index.write_text(index.read_text() + rows)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        item = report["fixes"]["A2_index_lines_moved"]
        self.assertLessEqual(index.stat().st_size, MODULE.INDEX_MAX_BYTES)
        self.assertFalse(item["over_ceiling_after_fix"])
        self.assertEqual(index.read_text().count("](issues/ISSUE-"), 60)

    def test_relink_for_log_descends_parent_relative_links_too(self):
        line = "[a](x.md) [b](../tests/t.cjs) [c](https://e.example/x) [d](#h) [e](/abs)"
        self.assertEqual(
            MODULE.relink_for_log(line),
            "[a](../x.md) [b](../../tests/t.cjs) [c](https://e.example/x) [d](#h) [e](/abs)",
        )

    def test_index_just_under_ceiling_is_shortened_to_leave_headroom(self):
        # 2026-09-22: --fix was a no-op on an index 80 bytes from the ceiling,
        # because the budget pass returned early unless the file was already
        # red. It must fire below the ceiling and leave real headroom behind.
        root = self.make_repo()
        index = root / "docs/00_index.md"
        content = index.read_text()
        n = 0
        # grow to just past the target but still under the ceiling: the exact
        # state that used to report "nothing to do"
        while len(content.encode()) <= MODULE.INDEX_TARGET_BYTES:
            content += f"- [ISSUE-{n:04d}](issues/ISSUE-{n:04d}.md) — " + "説明" * 60 + "\n"
            n += 1
        self.assertGreater(len(content.encode()), MODULE.INDEX_TARGET_BYTES)
        self.assertLessEqual(len(content.encode()), MODULE.INDEX_MAX_BYTES)
        index.write_text(content)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        item = report["fixes"]["A2_index_lines_moved"]
        self.assertGreater(item["count"], 0, "budget pass must fire below the ceiling")
        self.assertTrue(item["headroom_target_met"])
        self.assertGreaterEqual(item["headroom_bytes_after"], MODULE.INDEX_HEADROOM_BYTES)
        self.assertLessEqual(index.stat().st_size, MODULE.INDEX_TARGET_BYTES)
        # no row is lost and the full text is recoverable from the log
        self.assertEqual(index.read_text().count("](issues/ISSUE-"), n)
        self.assertIn("説明説明", (root / "docs/log/index-202609.md").read_text())

    def test_shortening_never_drops_a_second_link_from_a_row(self):
        # 2026-09-22, found by running --fix on a real 32,688-byte index: a row
        # whose description pointed at the ADR that recorded its decision lost
        # adrs/ADR-20260916-... when the description was cut. A repo that keeps
        # completed rows (Read Policy) cannot lose a routing target this way.
        line = ("- [ISSUE-x](issues/archive/ISSUE-x.md) — " + "説明" * 80
                + " 人間判断を [ADR-y](adrs/ADR-y.md) に記録した")
        for cap in MODULE.DESC_CAPS:
            short = MODULE.shorten_routing_line(line, cap, "log/index-202609.md")
            self.assertIn("(issues/archive/ISSUE-x.md)", short)
            self.assertIn("(adrs/ADR-y.md)", short, f"second link lost at cap {cap}")
        bullet = MODULE.shorten_bullet(line, "log/index-202609.md")
        self.assertIn("(adrs/ADR-y.md)", bullet)

    def test_real_index_keeps_every_link_through_the_budget_pass(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        content = index.read_text()
        n = 0
        while len(content.encode()) <= MODULE.INDEX_TARGET_BYTES:
            content += (f"- [ISSUE-{n:04d}](issues/ISSUE-{n:04d}.md) — " + "説明" * 60
                        + f" 判断は [ADR-{n:04d}](adrs/ADR-{n:04d}.md) に記録\n")
            n += 1
        index.write_text(content)
        import re as _re
        before = set(_re.findall(r"\]\(([^)\s]+\.md)\)", content))
        MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        after = set(_re.findall(r"\]\(([^)\s]+\.md)\)", index.read_text()))
        self.assertEqual(before - after, set(), "budget pass dropped link targets")

    def test_index_under_target_is_left_untouched(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        before = index.read_text() + "- [ISSUE-a](issues/ISSUE-a.md) — one line\n"
        index.write_text(before)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertEqual(report["fixes"]["A2_index_lines_moved"]["count"], 0)
        self.assertEqual(index.read_text(), before)

    def test_unmeetable_ceiling_is_reported_not_destroyed(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        rows = "".join(
            f"- [ISSUE-2026{n:04d}-very-long-slug-that-keeps-going-and-going-{n}]"
            f"(issues/ISSUE-2026{n:04d}-very-long-slug-that-keeps-going-and-going-{n}.md) — x\n"
            for n in range(400)
        )
        index.write_text(index.read_text() + rows)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        item = report["fixes"]["A2_index_lines_moved"]
        self.assertTrue(item["over_ceiling_after_fix"])
        self.assertIn("close or archive", item["note"])
        self.assertEqual(index.read_text().count("](issues/"), 400)

    def test_reference_written_relative_to_a_package_dir_is_not_dead(self):
        root = self.make_repo()
        (root / "functions/src/domain").mkdir(parents=True)
        (root / "functions/src/domain/events.ts").write_text("")
        (root / "docs/guides/g.md").write_text(
            "---\nupdated_at: 2026-09-01\n---\n# G\n\nSee `domain/events.ts` and `domain/gone.ts`.\n"
        )
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        dead = {i["ref"] for i in report["report"]["R3_dead_references"]["items"]}
        self.assertEqual(dead, {"domain/gone.ts"})


class GitignoredReferenceTest(HygieneFixture):
    """A gitignored path is absent *by design*; a guide naming it is correct, not stale.

    Red first: without the gitignore check both `apps/api/.deploy` (a build output the
    deploy script generates) and `apps/web/.env.local` (a local env file the guide tells
    the reader to create) are reported as dead references, and the noise buries the one
    reference that really is stale.
    """

    def test_gitignored_paths_are_not_dead_but_a_truly_missing_one_still_is(self):
        root = self.make_repo()
        # A directory-only pattern: `git check-ignore apps/api/.deploy` (no trailing
        # slash) MISSES it because the path does not exist, so git cannot tell it is a
        # directory. Only `apps/api/.deploy/` matches. Both forms must be tried.
        (root / ".gitignore").write_text("apps/api/.deploy/\n.env.*\n")
        (root / "docs/guides/g.md").write_text(
            "---\nupdated_at: 2026-09-01\n---\n# G\n\n"
            "Bundle lands in `apps/api/.deploy`; put secrets in `apps/web/.env.local`; "
            "see `docs/never-existed.md`.\n"
        )
        git(root, "add", ".")
        git(root, "commit", "-q", "-m", "ignore")
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        dead = {i["ref"] for i in report["report"]["R3_dead_references"]["items"]}
        self.assertEqual(dead, {"docs/never-existed.md"})

    def test_outside_a_git_repo_the_check_degrades_to_reporting_everything(self):
        root = self.make_repo()
        (root / "docs/guides/g.md").write_text(
            "---\nupdated_at: 2026-09-01\n---\n# G\n\nSee `docs/never-existed.md`.\n"
        )
        import shutil
        shutil.rmtree(root / ".git")
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        dead = {i["ref"] for i in report["report"]["R3_dead_references"]["items"]}
        self.assertEqual(dead, {"docs/never-existed.md"})


class MonorepoScriptsTest(HygieneFixture):
    def test_workspace_package_scripts_count_and_iam_roles_are_not_paths(self):
        root = self.make_repo()
        (root / "apps/api").mkdir(parents=True)
        (root / "apps/api/package.json").write_text(json.dumps({"scripts": {"serve": "x"}}))
        (root / "package.json").write_text(json.dumps({"scripts": {"build": "x"}}))
        (root / "docs/guides/g.md").write_text(
            "---\nupdated_at: 2026-09-01\n---\n# G\n\n`npm run serve`, `npm run nope`, "
            "grant `roles/storage.admin`.\n"
        )
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        dead = {i["ref"] for i in report["report"]["R3_dead_references"]["items"]}
        self.assertEqual(dead, {"npm run nope"})


class TableShapeTest(HygieneFixture):
    def test_padded_separator_collapses_and_self_links_dedupe(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        header = "| " + "ファイル".ljust(700) + " | 状態 |"
        sep = "| " + "-" * 700 + " | --- |"
        rows = "".join(
            f"| [issues/ISSUE-2026{n:04d}-slug-{n}.md](issues/ISSUE-2026{n:04d}-slug-{n}.md) | active |\n"
            for n in range(400)
        )
        index.write_text(index.read_text() + "\n## Active Issues\n\n" + header + "\n" + sep + "\n" + rows)
        self.assertGreater(index.stat().st_size, MODULE.INDEX_MAX_BYTES)
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        text = index.read_text()
        self.assertIn("| --- | --- |", text)
        self.assertNotIn("-" * 200, text)
        self.assertEqual(text.count("| [ISSUE-2026"), 400)
        self.assertLessEqual(index.stat().st_size, MODULE.INDEX_MAX_BYTES)
        # the header row was over-long and is logged; the separator is not
        log = (root / "docs/log/index-202609.md").read_text()
        self.assertIn("ファイル", log)
        self.assertNotIn("-" * 200, log)


class WrappedBulletTest(HygieneFixture):
    def test_wrapped_bullet_is_one_logical_line(self):
        root = self.make_repo()
        index = root / "docs/00_index.md"
        short_wrapped = (
            "- [ISSUE-a](issues/ISSUE-a.md)\n"
            "  — first half of a short description\n"
            "  that wraps once\n"
        )
        long_wrapped = "- [ISSUE-b](issues/ISSUE-b.md)\n" + "".join(
            f"  — continuation {n} " + "詳細" * 40 + "\n" for n in range(6)
        )
        index.write_text(index.read_text() + short_wrapped + long_wrapped)
        MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        text = index.read_text()
        self.assertIn("- [ISSUE-a](issues/ISSUE-a.md) — first half of a short description that wraps once", text)
        b_lines = [l for l in text.splitlines() if "](issues/ISSUE-b.md)" in l]
        self.assertEqual(len(b_lines), 1)
        self.assertIn("— continuation 0", b_lines[0])
        self.assertLessEqual(len(b_lines[0]), MODULE.INDEX_MAX_LINE_CHARS)
        self.assertNotIn("continuation 5", text)
        self.assertIn("continuation 5", (root / "docs/log/index-202609.md").read_text())


class ReportPrivacyTest(HygieneFixture):
    def test_report_carries_no_absolute_path(self):
        root = self.make_repo()
        MODULE.run(root, fix=False, report=True, today=date(2026, 9, 19))
        text = (root / "docs/log/hygiene-20260919.md").read_text()
        self.assertNotIn(str(root), text)
        self.assertNotIn(str(Path.home()), text)
        self.assertNotIn(root.name, text)
        self.assertIn("mode: dry run", text)


class GitIgnoredLargeInputTest(HygieneFixture):
    def test_many_probes_do_not_hang(self):
        root = self.make_repo()
        (root / ".gitignore").write_text("_output/\n")
        refs = [f"src/dir{n}/file{n}.ts" for n in range(400)] + ["_output/index.html"]
        import time
        t = time.time()
        ignored = MODULE.gitignored(root, refs)
        self.assertLess(time.time() - t, 20)
        self.assertEqual(ignored, {"_output/index.html"})


class HandoffFixTest(HygieneFixture):
    def test_updated_at_behind_git_is_synced(self):
        root = self.make_repo()
        path = self.add_issue(root, "ISSUE-20260901-lag", "active", "2026-08-01")
        git(root, "add", ".")
        git(root, "commit", "-q", "-m", "touch", when="2026-09-10T00:00:00")
        report = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertEqual(report["fixes"]["A5_updated_at_synced"]["count"], 1)
        self.assertIn("updated_at: 2026-09-10", path.read_text())
        again = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertEqual(again["fixes"]["A5_updated_at_synced"]["count"], 0)

    def test_finish_language_on_active_issue_is_reported(self):
        root = self.make_repo()
        p = self.add_issue(root, "ISSUE-20260901-said-done", "active", "2026-09-01")
        p.write_text(p.read_text().replace("## Acceptance", "## Current Status\n\n✅ 完了した。残作業なし。\n\n## Acceptance"))
        self.add_issue(root, "ISSUE-20260901-live", "active", "2026-09-01")
        index = root / "docs/00_index.md"
        index.write_text(index.read_text().replace(
            "- [ISSUE-20260901-live](issues/ISSUE-20260901-live.md) — one line",
            "- [ISSUE-20260901-live](issues/ISSUE-20260901-live.md) — 実装済み、あとは merge のみ"))
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        items = {i["file"]: i["where"] for i in report["report"]["R7_says_done_but_active"]["items"]}
        self.assertEqual(set(items), {"docs/issues/ISSUE-20260901-said-done.md", "docs/issues/ISSUE-20260901-live.md"})
        self.assertEqual(items["docs/issues/ISSUE-20260901-said-done.md"], "Current Status")
        self.assertEqual(items["docs/issues/ISSUE-20260901-live.md"], "index row")


class SweepTest(HygieneFixture):
    def test_sweep_writes_checklist_and_apply_archives_only_ticked(self):
        root = self.make_repo()
        old = (date(2026, 9, 19) - timedelta(days=45)).isoformat()
        self.add_issue(root, "ISSUE-20260701-gone", "in_progress", old, branch="feat/gone")
        keep = self.add_issue(root, "ISSUE-20260702-keep", "active", old, branch="feat/keep")
        git(root, "add", ".")
        git(root, "commit", "-q", "-m", "more", when="2026-06-01T00:00:00")
        sweep = MODULE.write_sweep(root, today=date(2026, 9, 19))
        text = sweep.read_text()
        self.assertEqual(sweep.name, "sweep-20260919.md")
        self.assertIn("- [ ] archive `docs/issues/ISSUE-20260701-gone.md`", text)
        self.assertIn("- [ ] archive `docs/issues/ISSUE-20260702-keep.md`", text)
        # human ticks one, writes a reason on the other
        text = text.replace("- [ ] archive `docs/issues/ISSUE-20260701-gone.md`",
                            "- [x] archive `docs/issues/ISSUE-20260701-gone.md`")
        text = text.replace("- [ ] archive `docs/issues/ISSUE-20260702-keep.md`",
                            "- [ ] archive `docs/issues/ISSUE-20260702-keep.md` — keep: still waiting on vendor")
        sweep.write_text(text)
        result = MODULE.apply_sweep(root, sweep, today=date(2026, 9, 19))
        self.assertEqual(result["archived"], ["docs/issues/ISSUE-20260701-gone.md"])
        self.assertEqual(result["kept"], ["docs/issues/ISSUE-20260702-keep.md"])
        self.assertTrue((root / "docs/issues/archive/ISSUE-20260701-gone.md").exists())
        self.assertTrue(keep.exists())
        self.assertIn("status: active", keep.read_text())
        # the decision is recorded in the sweep file itself
        self.assertIn("applied 2026-09-19", sweep.read_text())


class ReviewTest(HygieneFixture):
    """--review re-weights the current-truth layers by how the living docs use them."""

    def build(self, root: Path) -> None:
        g = root / "docs/guides"; s = root / "docs/specs"
        (g / "hot.md").write_text("---\nupdated_at: 2026-09-01\n---\n# Hot\n\nbody\n")
        (g / "cold.md").write_text("---\nupdated_at: 2026-03-01\n---\n# Cold\n\nbody\n")
        (g / "orphan-big.md").write_text(
            "---\nupdated_at: 2026-03-01\n---\n# Big\n\n" +
            "".join(f"## Part {n}\n\n" + "x" * 7000 + "\n\n" for n in range(10)))
        (s / "policy.md").write_text("---\nid: SPEC-policy\nstatus: active\ncreated_at: 2026-01-01\nupdated_at: 2026-01-01\n---\n# P\n")
        self.add_issue(root, "ISSUE-20260901-work", "active", "2026-09-01")
        p = root / "docs/issues/ISSUE-20260901-work.md"
        p.write_text(p.read_text() + "\nSee [hot](../guides/hot.md).\n")
        (s / "cited.md").write_text("---\nid: SPEC-cited\nstatus: active\ncreated_at: 2026-01-01\nupdated_at: 2026-01-01\n---\n# C\n")
        (g / "bare.md").write_text("---\nupdated_at: 2026-03-01\n---\n# Bare\n")
        p.write_text(p.read_text() + "Policy per SPEC-cited.\n")
        index = root / "docs/00_index.md"
        index.write_text(index.read_text() + "\n## Guides\n\n- [hot](guides/hot.md) — x\n- [cold](guides/cold.md) — y\n- docs/guides/bare.md — bare path form\n")
        git(root, "add", ".")
        git(root, "commit", "-q", "-m", "docs", when="2026-03-01T00:00:00")

    def test_review_tiers_and_split_candidates(self):
        root = self.make_repo()
        self.build(root)
        review = MODULE.build_review(root, today=date(2026, 9, 19))
        tiers = {r["file"]: r["tier"] for r in review["docs"]}
        self.assertEqual(tiers["docs/guides/hot.md"], "hot")       # linked from active work
        self.assertEqual(tiers["docs/guides/cold.md"], "warm")     # only the index links it
        self.assertEqual(tiers["docs/guides/orphan-big.md"], "cold")  # no inbound link, old
        self.assertEqual(tiers["docs/specs/policy.md"], "cold")
        self.assertEqual(tiers["docs/specs/cited.md"], "hot")    # cited by id from active work
        self.assertEqual(tiers["docs/guides/bare.md"], "warm")   # bare-path index row
        demote = {d["file"] for d in review["demote_candidates"]}
        self.assertEqual(demote, {"docs/guides/orphan-big.md", "docs/specs/policy.md"})
        split = {d["file"]: d for d in review["split_candidates"]}
        self.assertIn("docs/guides/orphan-big.md", split)
        self.assertEqual(len(split["docs/guides/orphan-big.md"]["sections"]), 10)

    def test_review_file_is_a_checklist_and_records_the_review_date(self):
        root = self.make_repo()
        self.build(root)
        path = MODULE.write_review(root, today=date(2026, 9, 19))
        text = path.read_text()
        self.assertEqual(path.name, "review-20260919.md")
        self.assertIn("- [ ] archive `docs/guides/orphan-big.md`", text)
        self.assertIn("- [ ] split `docs/guides/orphan-big.md`", text)
        self.assertIn("hot: 2", text)
        self.assertEqual(MODULE.last_review_date(root), date(2026, 9, 19))
        # apply_sweep understands the same archive rows
        self.assertIsNone(MODULE.last_review_date(self.make_repo()))

    def test_review_due_is_reported_by_hygiene(self):
        root = self.make_repo()
        self.build(root)
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        self.assertEqual(report["report"]["R8_review_overdue"]["count"], 1)
        MODULE.write_review(root, today=date(2026, 9, 19))
        report = MODULE.run(root, fix=False, report=False, today=date(2026, 9, 19))
        self.assertEqual(report["report"]["R8_review_overdue"]["count"], 0)


class ContextBudgetTest(HygieneFixture):
    def test_report_runs_review_and_counts_context_budget(self):
        root = self.make_repo()
        (root / "docs/guides/big.md").write_text("---\nupdated_at: 2026-09-01\n---\n# B\n\n" + "x" * 40000 + "\n")
        self.add_issue(root, "ISSUE-20260901-w", "active", "2026-09-01", index=False)
        p = root / "docs/issues/ISSUE-20260901-w.md"
        p.write_text(p.read_text() + "\nSee [big](../guides/big.md).\n")
        report = MODULE.run(root, fix=False, report=True, today=date(2026, 9, 19))
        r9 = report["report"]["R9_context_budget"]
        self.assertEqual(r9["items"][0]["oversized_docs"], 1)
        self.assertEqual(r9["items"][0]["hot_docs_kb"], 39)
        self.assertEqual(r9["items"][0]["digest_dropped_entries"], 0)
        self.assertTrue((root / "docs/log/review-20260919.md").exists())
        self.assertEqual(report["review_path"], "docs/log/review-20260919.md")
        self.assertEqual(report["report"]["R8_review_overdue"]["count"], 0)


class KeepCooldownTest(HygieneFixture):
    def test_recently_kept_candidates_are_not_re_listed(self):
        root = self.make_repo()
        old = (date(2026, 9, 19) - timedelta(days=45)).isoformat()
        self.add_issue(root, "ISSUE-20260701-gone", "in_progress", old, branch="feat/gone")
        self.add_issue(root, "ISSUE-20260702-also", "in_progress", old, branch="feat/also")
        git(root, "add", "."); git(root, "commit", "-q", "-m", "m", when="2026-06-01T00:00:00")
        first = MODULE.write_sweep(root, today=date(2026, 9, 19))
        first.write_text(first.read_text().replace(
            "- [ ] archive `docs/issues/ISSUE-20260701-gone.md`",
            "- [ ] archive `docs/issues/ISSUE-20260701-gone.md` — keep: vendor waiting"))
        # the second one was judged with a sub-bullet, the way the first judge wrote it
        first.write_text(first.read_text().replace(
            "- [ ] archive `docs/issues/ISSUE-20260702-also.md`\n",
            "- [ ] archive `docs/issues/ISSUE-20260702-also.md`\n  - keep: still live\n"))
        second = MODULE.write_sweep(root, today=date(2026, 9, 25))
        text = second.read_text()
        self.assertNotIn("ISSUE-20260701-gone", text)
        self.assertNotIn("ISSUE-20260702-also", text)
        self.assertIn("skipped: 2", text)
        later = MODULE.write_sweep(root, today=date(2026, 11, 1))
        self.assertIn("ISSUE-20260701-gone", later.read_text())


class UpdatedAtNoDriftTest(HygieneFixture):
    def test_a_commit_that_only_bumped_updated_at_does_not_move_it_again(self):
        root = self.make_repo()
        path = self.add_issue(root, "ISSUE-20260901-lag", "active", "2026-08-01")
        git(root, "add", "."); git(root, "commit", "-q", "-m", "content", when="2026-09-10T00:00:00")
        MODULE.run(root, fix=True, report=False, today=date(2026, 9, 19))
        self.assertIn("updated_at: 2026-09-10", path.read_text())
        git(root, "add", "."); git(root, "commit", "-q", "-m", "docs(hygiene): sync", when="2026-09-19T00:00:00")
        again = MODULE.run(root, fix=True, report=False, today=date(2026, 9, 25))
        self.assertEqual(again["fixes"]["A5_updated_at_synced"]["count"], 0)
        self.assertIn("updated_at: 2026-09-10", path.read_text())
