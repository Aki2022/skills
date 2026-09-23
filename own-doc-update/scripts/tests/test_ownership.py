"""Issue ownership, routing and orphan prevention (SPEC-doc-governance, Workstream Model).

Fixtures use created_at 2000-01-01 (before any rollout date) and 2099-01-01 (after),
so the rollout gate is exercised without depending on the ROLLOUT_DATE value.
"""
import importlib.util
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))

import ownership  # noqa: E402

_SPEC = importlib.util.spec_from_file_location("validate_repo_docs", SCRIPTS / "validate_repo_docs.py")
VALIDATOR = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(VALIDATOR)

PRE = "2000-01-01"
POST = "2099-01-01"


def make_repo() -> Path:
    root = Path(tempfile.mkdtemp())
    for path in ("docs/adrs", "docs/specs", "docs/issues/archive", "docs/workstreams/archive", "docs/guides"):
        (root / path).mkdir(parents=True, exist_ok=True)
    write_index(root, "")
    return root


def write_index(root: Path, active_issues: str, active_workstreams: str = "") -> None:
    (root / "docs/00_index.md").write_text(
        "---\nupdated_at: 2026-09-23\ncurrent_focus: []\n---\n\n# 00 Index\n\n"
        f"## Active Workstreams\n\n{active_workstreams}\n\n"
        f"## Active Issues\n\n{active_issues}\n\n## Specs\n"
    )


def write_issue(root: Path, issue_id: str, created: str, fm: str = "", title: str = "An issue",
                archived: bool = False) -> Path:
    folder = root / ("docs/issues/archive" if archived else "docs/issues")
    path = folder / f"{issue_id}.md"
    path.write_text(
        f"---\nschema_version: 2\nid: {issue_id}\nstatus: active\ncreated_at: {created}\n"
        f"updated_at: {created}\n{fm}guide_impact: none\nguide_impact_reason: \"test\"\n---\n\n"
        f"# {title}\n\n## Goal\n\n## Acceptance\n\n- verify: machine — true\n\n"
        "## Current Status\n\n## Next Actions\n\n1. next\n"
    )
    return path


def write_ws(root: Path, ws_id: str, created: str, fm: str = "", body: str = "",
             archived: bool = False) -> Path:
    folder = root / ("docs/workstreams/archive" if archived else "docs/workstreams")
    path = folder / f"{ws_id}.md"
    path.write_text(
        f"---\nschema_version: 2\nid: {ws_id}\nstatus: active\ncreated_at: {created}\n"
        f"updated_at: {created}\n{fm}human_boundary_confirmed_at: {created}\nnext_human_gate: g\n---\n\n"
        f"# A workstream\n\n{body}\n"
    )
    return path


def split_block(*issue_ids: str) -> str:
    rows = "\n".join(f"- [{i}](../issues/{i}.md) — t" for i in issue_ids)
    return (
        "## Split Issues\n\n"
        f"{ownership.block_begin('split-issues')}\n{rows}\n{ownership.block_end('split-issues')}\n"
    )


# Minimal fixtures trip unrelated rules (e.g. a workstream without its required sections),
# so every assertion looks only at messages about ownership, routing, priority or due.
OWNERSHIP_WORDS = ("workstream", "Split Issues", "Active Issues", "priority", "due")


def _relevant(messages: list[str], rel: str) -> list[str]:
    prefix = f"{rel}: "
    return [m for m in messages if m.startswith(prefix) and any(w in m[len(prefix):] for w in OWNERSHIP_WORDS)]


def errors_for(root: Path, rel: str) -> list[str]:
    errors, _ = VALIDATOR.validate_repo(root)
    return _relevant(errors, rel)


def warnings_for(root: Path, rel: str) -> list[str]:
    _, warnings = VALIDATOR.validate_repo(root)
    return _relevant(warnings, rel)


class GeneratedBlockTest(unittest.TestCase):
    def test_write_block_inserts_under_heading_matched_by_prefix(self):
        content = "# X\n\n## Active Issues（WS 未所属）\n\n| a | b |\n\n## Specs\n"
        updated = ownership.write_block(content, "active-issues", ["- row"], "Active Issues")
        self.assertIn(ownership.block_begin("active-issues"), updated)
        self.assertLess(updated.index("## Active Issues"), updated.index("- row"))
        self.assertLess(updated.index("- row"), updated.index("## Specs"))
        self.assertIn("| a | b |", updated)

    def test_write_block_is_idempotent(self):
        content = "## Active Issues\n\n"
        once = ownership.write_block(content, "active-issues", ["- row"], "Active Issues")
        twice = ownership.write_block(once, "active-issues", ["- row"], "Active Issues")
        self.assertEqual(once, twice)
        self.assertEqual(ownership.read_block(twice, "active-issues"), ["- row"])

    def test_write_block_adds_missing_heading(self):
        updated = ownership.write_block("# X\n", "split-issues", ["- row"], "Split Issues")
        self.assertIn("## Split Issues", updated)
        self.assertEqual(ownership.read_block(updated, "split-issues"), ["- row"])


class IssueOwnershipValidationTest(unittest.TestCase):
    def test_post_rollout_issue_without_workstream_is_an_error_on_the_issue(self):
        root = make_repo()
        write_issue(root, "ISSUE-20990101-a", POST)
        errs = errors_for(root, "docs/issues/ISSUE-20990101-a.md")
        self.assertTrue(any("workstream" in e for e in errs), errs)

    def test_pre_rollout_issue_without_workstream_is_only_a_warning(self):
        root = make_repo()
        write_issue(root, "ISSUE-20000101-a", PRE)
        rel = "docs/issues/ISSUE-20000101-a.md"
        self.assertFalse(any("workstream" in e for e in errors_for(root, rel)))
        self.assertTrue(any("workstream" in w for w in warnings_for(root, rel)))

    def test_workstream_that_is_not_active_is_an_error_even_before_rollout(self):
        root = make_repo()
        write_ws(root, "WS-20000101-gone", PRE, archived=True)
        write_issue(root, "ISSUE-20000101-a", PRE, fm="workstream: WS-20000101-gone\n")
        errs = errors_for(root, "docs/issues/ISSUE-20000101-a.md")
        self.assertTrue(any("not an active workstream" in e for e in errs), errs)

    def test_owned_issue_missing_from_split_list_is_an_error_on_the_workstream(self):
        root = make_repo()
        write_ws(root, "WS-20000101-w", PRE)
        write_issue(root, "ISSUE-20000101-a", PRE, fm="workstream: WS-20000101-w\n")
        errs = errors_for(root, "docs/workstreams/WS-20000101-w.md")
        self.assertTrue(any("ISSUE-20000101-a" in e for e in errs), errs)

    def test_owned_issue_listed_in_split_list_is_clean(self):
        root = make_repo()
        write_ws(root, "WS-20000101-w", PRE, body=split_block("ISSUE-20000101-a"))
        write_issue(root, "ISSUE-20000101-a", PRE, fm="workstream: WS-20000101-w\n")
        self.assertEqual(errors_for(root, "docs/workstreams/WS-20000101-w.md"), [])
        issue_errs = errors_for(root, "docs/issues/ISSUE-20000101-a.md")
        self.assertFalse(any("workstream" in e for e in issue_errs), issue_errs)

    def test_split_list_row_for_archived_or_foreign_issue_is_an_error_on_the_workstream(self):
        root = make_repo()
        write_ws(root, "WS-20000101-w", PRE, body=split_block("ISSUE-20000101-old", "ISSUE-20000101-other"))
        write_ws(root, "WS-20000101-x", PRE, body=split_block("ISSUE-20000101-other"))
        write_issue(root, "ISSUE-20000101-old", PRE, fm="workstream: WS-20000101-w\n", archived=True)
        write_issue(root, "ISSUE-20000101-other", PRE, fm="workstream: WS-20000101-x\n")
        errs = errors_for(root, "docs/workstreams/WS-20000101-w.md")
        self.assertTrue(any("ISSUE-20000101-old" in e for e in errs), errs)
        self.assertTrue(any("ISSUE-20000101-other" in e for e in errs), errs)

    def test_standalone_after_rollout_needs_priority_and_due(self):
        root = make_repo()
        write_issue(root, "ISSUE-20990101-s", POST, fm="workstream: none\n")
        write_index(root, "- [ISSUE-20990101-s](issues/ISSUE-20990101-s.md) — s")
        errs = errors_for(root, "docs/issues/ISSUE-20990101-s.md")
        self.assertTrue(any("priority" in e for e in errs), errs)
        self.assertTrue(any("due" in e for e in errs), errs)

    def test_due_none_is_accepted_and_bad_values_are_rejected(self):
        root = make_repo()
        write_issue(root, "ISSUE-20990101-s", POST, fm="workstream: none\npriority: high\ndue: none\n")
        write_issue(root, "ISSUE-20990101-t", POST, fm="workstream: none\npriority: urgent\ndue: soon\n")
        write_index(root, "- [ISSUE-20990101-s](issues/ISSUE-20990101-s.md) — s\n"
                          "- [ISSUE-20990101-t](issues/ISSUE-20990101-t.md) — t")
        self.assertEqual(errors_for(root, "docs/issues/ISSUE-20990101-s.md"), [])
        bad = errors_for(root, "docs/issues/ISSUE-20990101-t.md")
        self.assertTrue(any("priority" in e for e in bad), bad)
        self.assertTrue(any("due" in e for e in bad), bad)

    def test_standalone_not_routed_from_active_issues_is_an_error_on_the_issue(self):
        root = make_repo()
        write_issue(root, "ISSUE-20990101-s", POST, fm="workstream: none\npriority: low\ndue: none\n")
        errs = errors_for(root, "docs/issues/ISSUE-20990101-s.md")
        self.assertTrue(any("Active Issues" in e for e in errs), errs)

    def test_owned_issue_listed_in_active_issues_is_an_error_on_the_index(self):
        root = make_repo()
        write_ws(root, "WS-20000101-w", PRE, body=split_block("ISSUE-20000101-a"))
        write_issue(root, "ISSUE-20000101-a", PRE, fm="workstream: WS-20000101-w\n")
        write_index(root, "| [ISSUE-20000101-a](issues/ISSUE-20000101-a.md) | t | High | x |")
        errs = errors_for(root, "docs/00_index.md")
        self.assertTrue(any("ISSUE-20000101-a" in e and "WS-20000101-w" in e for e in errs), errs)

    def test_workstream_priority_and_due_are_gated_by_rollout(self):
        root = make_repo()
        write_ws(root, "WS-20990101-new", POST)
        write_ws(root, "WS-20000101-old", PRE)
        new_errs = errors_for(root, "docs/workstreams/WS-20990101-new.md")
        self.assertTrue(any("priority" in e for e in new_errs), new_errs)
        self.assertFalse(any("priority" in e for e in errors_for(root, "docs/workstreams/WS-20000101-old.md")))
        self.assertTrue(any("priority" in w for w in warnings_for(root, "docs/workstreams/WS-20000101-old.md")))


def run(script: str, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(SCRIPTS / script), *args], capture_output=True, text=True)


ISSUE_BASE = ("--no-guide-reason", "test", "--verify-machine", "true", "--next-action", "first step")
WS_BASE = (
    "--issue", "first", "--scope", "s", "--confirmed-at", "2026-09-23", "--next-human-gate", "g",
    "--autonomous", "a", "--confirm-first", "c", "--verify-machine", "true", "--no-guide-reason", "r",
)


class CreateScriptsTest(unittest.TestCase):
    def test_create_issue_requires_an_ownership_choice(self):
        root = make_repo()
        result = run("create_issue.py", "x", "--repo", str(root), *ISSUE_BASE)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(list((root / "docs/issues").glob("*.md")), [])

    def test_create_issue_for_a_workstream_writes_owner_and_split_row(self):
        root = make_repo()
        write_ws(root, "WS-20000101-w", PRE)
        result = run("create_issue.py", "x", "--date", "20260923", "--repo", str(root),
                     "--workstream", "WS-20000101-w", *ISSUE_BASE)
        self.assertEqual(result.returncode, 0, result.stderr)
        issue = (root / "docs/issues/ISSUE-20260923-x.md").read_text()
        self.assertIn("workstream: WS-20000101-w", issue)
        ws = (root / "docs/workstreams/WS-20000101-w.md").read_text()
        self.assertIn("ISSUE-20260923-x", "\n".join(ownership.read_block(ws, ownership.SPLIT) or []))
        self.assertEqual(errors_for(root, "docs/workstreams/WS-20000101-w.md"), [])
        self.assertEqual(errors_for(root, "docs/00_index.md"), [])

    def test_create_issue_refuses_an_inactive_workstream_and_leaves_nothing(self):
        root = make_repo()
        result = run("create_issue.py", "x", "--repo", str(root), "--workstream", "WS-20000101-nope", *ISSUE_BASE)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(list((root / "docs/issues").glob("*.md")), [])

    def test_standalone_issue_requires_priority_and_due_and_is_routed_from_the_index(self):
        root = make_repo()
        missing = run("create_issue.py", "x", "--repo", str(root), "--standalone", *ISSUE_BASE)
        self.assertNotEqual(missing.returncode, 0)
        result = run("create_issue.py", "x", "--date", "20260923", "--repo", str(root), "--standalone",
                     "--priority", "high", "--due", "2026-10-01", *ISSUE_BASE)
        self.assertEqual(result.returncode, 0, result.stderr)
        issue = (root / "docs/issues/ISSUE-20260923-x.md").read_text()
        for line in ("workstream: none", "priority: high", "due: 2026-10-01"):
            self.assertIn(line, issue)
        rows = ownership.read_block((root / "docs/00_index.md").read_text(), ownership.ACTIVE_ISSUES)
        self.assertTrue(rows and "ISSUE-20260923-x" in rows[0] and "due 2026-10-01" in rows[0], rows)
        self.assertEqual(errors_for(root, "docs/issues/ISSUE-20260923-x.md"), [])

    def test_create_workstream_requires_priority_and_due_and_is_routed_from_the_index(self):
        root = make_repo()
        missing = run("create_workstream.py", "w", "--repo", str(root), *WS_BASE)
        self.assertNotEqual(missing.returncode, 0)
        result = run("create_workstream.py", "w", "--date", "20260923", "--repo", str(root),
                     "--priority", "medium", "--due", "none", *WS_BASE)
        self.assertEqual(result.returncode, 0, result.stderr)
        ws = (root / "docs/workstreams/WS-20260923-w.md").read_text()
        self.assertIn("priority: medium", ws)
        self.assertIn("due: none", ws)
        self.assertIsNotNone(ownership.read_block(ws, ownership.SPLIT))
        rows = ownership.read_block((root / "docs/00_index.md").read_text(), ownership.ACTIVE_WORKSTREAMS)
        self.assertTrue(rows and "WS-20260923-w" in rows[0], rows)


def ready_to_archive(path: Path) -> None:
    text = path.read_text().replace("- [ ]", "- [x]").replace("- status: pending", "- status: complete")
    path.write_text(text.replace("status: active", "status: complete", 1))


class ArchiveScriptsTest(unittest.TestCase):
    def make_ws(self, root: Path, slug: str) -> Path:
        result = run("create_workstream.py", slug, "--date", "20260923", "--repo", str(root),
                     "--priority", "low", "--due", "none", *WS_BASE)
        self.assertEqual(result.returncode, 0, result.stderr)
        return root / f"docs/workstreams/WS-20260923-{slug}.md"

    def test_archive_issue_removes_it_from_the_owning_workstream_list(self):
        root = make_repo()
        ws = self.make_ws(root, "w")
        created = run("create_issue.py", "x", "--date", "20260923", "--repo", str(root),
                      "--workstream", "WS-20260923-w", *ISSUE_BASE)
        self.assertEqual(created.returncode, 0, created.stderr)
        self.assertIn("ISSUE-20260923-x", "\n".join(ownership.read_block(ws.read_text(), ownership.SPLIT)))
        ready_to_archive(root / "docs/issues/ISSUE-20260923-x.md")
        result = run("archive_issue.py", "ISSUE-20260923-x", "--repo", str(root))
        self.assertEqual(result.returncode, 0, result.stderr + result.stdout)
        self.assertEqual(ownership.read_block(ws.read_text(), ownership.SPLIT), [])
        self.assertEqual(errors_for(root, "docs/workstreams/WS-20260923-w.md"), [])

    def test_archive_workstream_refuses_while_an_active_issue_declares_it(self):
        root = make_repo()
        owned = self.make_ws(root, "owned")
        control = self.make_ws(root, "control")
        created = run("create_issue.py", "x", "--date", "20260923", "--repo", str(root),
                      "--workstream", "WS-20260923-owned", *ISSUE_BASE)
        self.assertEqual(created.returncode, 0, created.stderr)
        for ws in (owned, control):
            ready_to_archive(ws)
        refused = run("archive_workstream.py", "WS-20260923-owned", "--repo", str(root))
        self.assertNotEqual(refused.returncode, 0)
        self.assertIn("ISSUE-20260923-x", refused.stderr)
        self.assertTrue(owned.is_file())
        allowed = run("archive_workstream.py", "WS-20260923-control", "--repo", str(root))
        self.assertEqual(allowed.returncode, 0, allowed.stderr + allowed.stdout)


if __name__ == "__main__":
    unittest.main()
