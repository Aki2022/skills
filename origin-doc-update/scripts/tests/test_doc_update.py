import importlib.util
import re
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "validate_repo_docs.py"
SPEC = importlib.util.spec_from_file_location("validate_repo_docs", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class ValidateRepoDocsV2Test(unittest.TestCase):
    def make_repo(self) -> Path:
        root = Path(tempfile.mkdtemp())
        for path in (
            "docs/adrs",
            "docs/specs",
            "docs/issues/archive",
            "docs/workstreams/archive",
            "docs/guides",
        ):
            (root / path).mkdir(parents=True, exist_ok=True)
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-07-19\ncurrent_focus: []\n---\n\n# 00 Index\n"
        )
        return root

    def test_required_guide_impact_needs_a_guide_target(self):
        root = self.make_repo()
        (root / "docs/issues/ISSUE-20260719-example.md").write_text(
            """---
schema_version: 2
id: ISSUE-20260719-example
status: active
created_at: 2026-07-19
updated_at: 2026-07-19
guide_impact: required
guide_impact_reason: ""
related_guides: []
---

# Example
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("guide_impact is required" in error for error in errors))

    def test_no_guide_impact_needs_a_reason(self):
        root = self.make_repo()
        (root / "docs/issues/ISSUE-20260719-example.md").write_text(
            """---
schema_version: 2
id: ISSUE-20260719-example
status: active
created_at: 2026-07-19
updated_at: 2026-07-19
guide_impact: none
guide_impact_reason: ""
related_guides: []
---

# Example
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("guide_impact_reason" in error for error in errors))

    def test_workstream_requires_confirmed_human_boundary(self):
        root = self.make_repo()
        (root / "docs/workstreams/WS-20260719-example.md").write_text(
            """---
schema_version: 2
id: WS-20260719-example
status: active
created_at: 2026-07-19
updated_at: 2026-07-19
human_boundary_confirmed_at: ""
next_human_gate: final-review
---

# Example

## Authorization Envelope

## Issue Queue
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("human_boundary_confirmed_at" in error for error in errors)
        )

    def test_required_guide_must_reference_its_source_issue(self):
        root = self.make_repo()
        (root / "docs/guides/example.md").write_text(
            """---
id: GUIDE-example
updated_at: 2026-07-19
source_issues: []
source_workstreams: []
---

# Example
"""
        )
        (root / "docs/issues/ISSUE-20260719-example.md").write_text(
            """---
schema_version: 2
id: ISSUE-20260719-example
status: active
created_at: 2026-07-19
updated_at: 2026-07-19
branch: ISSUE-20260719-example
guide_impact: required
guide_impact_reason: ""
related_guides: [GUIDE-example]
---

# Example
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("must list source_issues" in error for error in errors))


class FlowStyleFrontMatterTest(unittest.TestCase):
    """A bracket list opened on the line *after* the key used to parse as []."""

    def write(self, body: str) -> Path:
        root = Path(tempfile.mkdtemp())
        path = root / "doc.md"
        path.write_text(body)
        return path

    def test_multiline_flow_list_is_parsed_not_silently_dropped(self):
        path = self.write(
            """---
id: GUIDE-example
source_issues:
  [
    ISSUE-a,
    ISSUE-b,
  ]
---

# Example
"""
        )

        fm = MODULE.parse_front_matter(path)

        self.assertEqual(fm["source_issues"], ["ISSUE-a", "ISSUE-b"])

    def test_single_line_flow_list_under_key_is_parsed(self):
        path = self.write(
            """---
id: GUIDE-example
source_issues:
  [ISSUE-a, ISSUE-b]
---

# Example
"""
        )

        fm = MODULE.parse_front_matter(path)

        self.assertEqual(fm["source_issues"], ["ISSUE-a", "ISSUE-b"])

    def test_flow_style_key_is_recorded_for_reporting(self):
        path = self.write(
            """---
id: GUIDE-example
source_issues:
  [ISSUE-a]
---

# Example
"""
        )

        fm = MODULE.parse_front_matter(path)

        self.assertEqual(MODULE.flow_style_keys(fm), ["source_issues"])

    def test_block_style_records_no_flow_style_key(self):
        path = self.write(
            """---
id: GUIDE-example
source_issues:
  - ISSUE-a
  - ISSUE-b
---

# Example
"""
        )

        fm = MODULE.parse_front_matter(path)

        self.assertEqual(fm["source_issues"], ["ISSUE-a", "ISSUE-b"])
        self.assertEqual(MODULE.flow_style_keys(fm), [])

    def test_same_line_flow_list_is_not_flagged(self):
        path = self.write(
            """---
id: GUIDE-example
source_issues: [ISSUE-a, ISSUE-b]
---

# Example
"""
        )

        fm = MODULE.parse_front_matter(path)

        self.assertEqual(fm["source_issues"], ["ISSUE-a", "ISSUE-b"])
        self.assertEqual(MODULE.flow_style_keys(fm), [])

    def test_empty_block_list_stays_empty(self):
        path = self.write(
            """---
id: GUIDE-example
source_issues:
source_workstreams: []
---

# Example
"""
        )

        fm = MODULE.parse_front_matter(path)

        self.assertEqual(fm["source_issues"], [])
        self.assertEqual(fm["source_workstreams"], [])
        self.assertEqual(MODULE.flow_style_keys(fm), [])


class FlowStyleValidationTest(ValidateRepoDocsV2Test):
    def test_validate_repo_errors_on_flow_style_front_matter(self):
        root = self.make_repo()
        (root / "docs/guides/example.md").write_text(
            """---
id: GUIDE-example
updated_at: 2026-07-19
source_issues:
  [
    ISSUE-20260719-example,
  ]
---

# Example
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any(
                "docs/guides/example.md" in error and "block style" in error
                for error in errors
            ),
            errors,
        )

    def test_flow_style_refs_still_reach_reference_validation(self):
        """The whole point: entries must actually be checked, not skipped."""
        root = self.make_repo()
        (root / "docs/guides/example.md").write_text(
            """---
id: GUIDE-example
updated_at: 2026-07-19
source_issues:
  [
    ISSUE-20260719-does-not-exist,
  ]
---

# Example
"""
        )

        _errors, warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("ISSUE-20260719-does-not-exist" in warning for warning in warnings),
            warnings,
        )


class GeneratedFilesValidateTest(ValidateRepoDocsV2Test):
    """The generators and the validator must agree.

    Nothing used to assert this, which is how issue.template.md came to ship an
    inline `# required | none` comment on guide_impact: parse_front_matter does
    not strip comments, so every generated issue failed validation on creation,
    and the comment on guide_impact_reason made that value non-empty, silently
    satisfying the "reason required when none" check with garbage.
    """

    def run_script(self, name: str, *args: str) -> None:
        script = Path(__file__).resolve().parents[1] / name
        result = subprocess.run(
            [sys.executable, str(script), *args],
            capture_output=True,
            text=True,
        )
        self.assertEqual(result.returncode, 0, f"{name} failed: {result.stderr}")

    def errors_for(self, root: Path, relative: str) -> list[str]:
        errors, _warnings = MODULE.validate_repo(root)
        return [error for error in errors if error.startswith(relative)]

    def test_generated_issue_passes_validation(self):
        root = self.make_repo()
        self.run_script(
            "create_issue.py",
            "example-generated-issue",
            "--no-guide-reason",
            "internal refactor with unchanged behavior",
            "--verify-machine",
            "pytest tests/ exits 0",
            "--repo",
            str(root),
            "--date",
            "20260729",
        )
        target = "docs/issues/ISSUE-20260729-example-generated-issue.md"
        self.assertTrue((root / target).is_file())
        self.assertEqual(self.errors_for(root, target), [])

    def test_generated_adr_passes_validation(self):
        root = self.make_repo()
        self.run_script(
            "create_adr.py",
            "choose-storage-boundary",
            "--scope",
            "development",
            "--status",
            "accepted",
            "--title",
            "Choose Storage Boundary",
            "--repo",
            str(root),
            "--date",
            "20260729",
        )
        target = "docs/adrs/ADR-20260729-choose-storage-boundary.md"
        self.assertTrue((root / target).is_file())
        self.assertEqual(self.errors_for(root, target), [])

    def test_superseded_adr_requires_successor(self):
        root = self.make_repo()
        target = root / "docs/adrs/ADR-20260729-old.md"
        target.write_text(
            """---
id: ADR-20260729-old
status: superseded
scope: spec
created_at: 2026-07-29
updated_at: 2026-07-29
---

# Old decision
"""
        )
        errors, _warnings = MODULE.validate_repo(root)
        self.assertTrue(any("superseded_by is required" in error for error in errors))

    def test_create_issue_requires_a_guide_decision(self):
        """Classification is forced at creation, as it is for workstreams.

        Without it the generated file carries guide_impact: required against an
        empty related_guides, which the validator rejects, so a fresh issue could
        never validate.
        """
        root = self.make_repo()
        script = Path(__file__).resolve().parents[1] / "create_issue.py"
        result = subprocess.run(
            [sys.executable, str(script), "unclassified", "--repo", str(root)],
            capture_output=True,
            text=True,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("--guide", result.stderr)

    def test_generated_issue_names_the_guide_it_must_update(self):
        root = self.make_repo()
        self.run_script(
            "create_issue.py",
            "example-guide-bound-issue",
            "--guide",
            "GUIDE-example",
            "--verify-human",
            "the guide owner reviews the updated section",
            "--repo",
            str(root),
            "--date",
            "20260729",
        )
        text = (root / "docs/issues/ISSUE-20260729-example-guide-bound-issue.md").read_text()
        self.assertIn("guide_impact: required", text)
        self.assertIn("related_guides: [GUIDE-example]", text)

    def test_generated_workstream_passes_validation(self):
        root = self.make_repo()
        self.run_script(
            "create_workstream.py",
            "example-generated-workstream",
            "--issue",
            "first-slice",
            "--scope",
            "Approved scope for the test",
            "--confirmed-at",
            "2026-07-29",
            "--next-human-gate",
            "review the first slice",
            "--autonomous",
            "edit docs and run tests",
            "--confirm-first",
            "external sends",
            "--verify-machine",
            "pytest tests/ exits 0",
            "--no-guide-reason",
            "internal refactor with unchanged behavior",
            "--repo",
            str(root),
            "--date",
            "20260729",
        )
        matches = list((root / "docs/workstreams").glob("*example-generated-workstream*.md"))
        self.assertEqual(len(matches), 1, f"expected one generated workstream, got {matches}")
        target = str(matches[0].relative_to(root))
        self.assertEqual(self.errors_for(root, target), [])

    def test_issue_template_front_matter_carries_no_inline_comments(self):
        template = Path(__file__).resolve().parents[2] / "references/issue.template.md"
        front_matter = template.read_text().split("---", 2)[1]
        offenders = [
            line
            for line in front_matter.splitlines()
            if re.match(r"^[A-Za-z0-9_]+:", line) and "#" in line
        ]
        self.assertEqual(offenders, [], "front matter values must not carry inline comments")



class RecordCompletenessTest(ValidateRepoDocsV2Test):
    """The executors treat a missing record as a gate and stop.

    origin-ws-loop reads runnability per issue and treats a missing record as
    `gated`; origin-goal refuses to start on an unrecorded envelope and cannot
    verify empty acceptance. None of that used to be validated, so a workstream
    could pass clean and still stop every autonomous run at a gate nobody set.
    """

    WS_HEADER = (
        "---\nschema_version: 2\nid: WS-20260817-example\nstatus: active\n"
        "created_at: 2026-08-17\nupdated_at: 2026-08-17\nbranch: \"\"\npr: \"\"\n"
        "human_boundary_confirmed_at: 2026-08-17\nnext_human_gate: none-workstream-complete\n"
        "related_specs: []\nrelated_guides: []\n---\n\n# Example\n"
    )
    ENVELOPE = (
        "\n## Authorization Envelope\n\n"
        "- Approved scope: the example scope\n"
        "- Autonomous actions allowed: edit docs and run tests\n"
        "- Confirm first: external sends\n"
        "- Merge policy: CD (default)\n"
        "- Cost or usage ceiling: none\n"
        "- Out of scope: everything else\n"
        "\n## Human Gates\n\n"
        "- Start gate: confirmed on 2026-08-17\n"
        "- Next gate: none-workstream-complete\n"
        "- Stop conditions: three same-root-cause failures\n"
    )
    QUEUE = (
        "\n## Issue Queue\n\n"
        "| Issue | Status |\n| --- | --- |\n| ISSUE-01-example | pending |\n"
        "\n### ISSUE-01-example\n\n"
        "- status: pending\n"
        "- depends_on: []\n"
        "- runnability: ready\n"
        "- guide_impact: none\n"
        "- related_guides: []\n"
        "- guide_impact_reason: x\n"
        "\n#### Acceptance\n\n"
        "- verify: machine — pytest tests/ exits 0\n"
    )

    def write_ws(self, root, envelope=None, queue=None):
        text = self.WS_HEADER + (envelope or self.ENVELOPE) + (queue or self.QUEUE)
        (root / "docs/workstreams/WS-20260817-example.md").write_text(text)

    def test_a_complete_record_validates_clean(self):
        root = self.make_repo()
        self.write_ws(root)

        errors, _warnings = MODULE.validate_repo(root)

        self.assertEqual(
            [e for e in errors if "WS-20260817-example" in e], [], errors
        )

    def test_an_issue_block_without_runnability_is_an_error(self):
        root = self.make_repo()
        self.write_ws(root, queue=self.QUEUE.replace("- runnability: ready\n", ""))

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("runnability" in e for e in errors),
            "a missing runnability record is what stops every autonomous run",
        )

    def test_gated_runnability_needs_a_reason(self):
        root = self.make_repo()
        self.write_ws(
            root,
            queue=self.QUEUE.replace("- runnability: ready", "- runnability: gated"),
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("runnability" in e for e in errors))

    def test_an_issue_block_without_acceptance_verify_is_an_error(self):
        root = self.make_repo()
        self.write_ws(
            root,
            queue=self.QUEUE.replace(
                "- verify: machine — pytest tests/ exits 0\n", ""
            ),
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("verify" in e for e in errors),
            "empty acceptance makes an agent stall or overclaim",
        )

    def test_an_empty_envelope_bullet_is_an_error(self):
        root = self.make_repo()
        self.write_ws(
            root,
            envelope=self.ENVELOPE.replace(
                "- Autonomous actions allowed: edit docs and run tests",
                "- Autonomous actions allowed:",
            ),
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("Autonomous actions allowed" in e for e in errors))

    def test_an_indented_continuation_satisfies_an_envelope_bullet(self):
        root = self.make_repo()
        self.write_ws(
            root,
            envelope=self.ENVELOPE.replace(
                "- Confirm first: external sends",
                "- Confirm first:\n  - external sends\n  - deploys",
            ),
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertEqual([e for e in errors if "Confirm first" in e], [], errors)

    def test_a_standalone_issue_without_acceptance_verify_is_an_error(self):
        root = self.make_repo()
        (root / "docs/issues/ISSUE-20260817-example.md").write_text(
            """---
schema_version: 2
id: ISSUE-20260817-example
status: active
created_at: 2026-08-17
updated_at: 2026-08-17
branch: ISSUE-20260817-example
guide_impact: none
guide_impact_reason: docs only
related_guides: []
---

# Example

## Goal

## Notes
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("verify" in e for e in errors))

    def test_a_standalone_issue_with_acceptance_verify_is_clean(self):
        root = self.make_repo()
        (root / "docs/issues/ISSUE-20260817-example.md").write_text(
            """---
schema_version: 2
id: ISSUE-20260817-example
status: active
created_at: 2026-08-17
updated_at: 2026-08-17
branch: ISSUE-20260817-example
guide_impact: none
guide_impact_reason: docs only
related_guides: []
---

# Example

## Acceptance

- verify: human-review — the owner reads the filed report
"""
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertEqual(
            [e for e in errors if "ISSUE-20260817-example" in e], [], errors
        )


class SilentPassTest(ValidateRepoDocsV2Test):
    """A check that cannot fail is worse than no check: it reports success."""

    WS_HEADER = '---\nschema_version: 2\nid: WS-20260803-example\nstatus: active\ncreated_at: 2026-08-03\nupdated_at: 2026-08-03\nbranch: ""\npr: ""\nhuman_boundary_confirmed_at: 2026-08-03\nnext_human_gate: none\nrelated_specs: []\nrelated_guides: []\n---\n\n# Example\n\n## Authorization Envelope\n\nx\n\n## Human Gates\n\nx\n\n## Issue Queue\n'

    def write_ws(self, root, body, name="WS-20260803-example.md"):
        (root / "docs/workstreams" / name).write_text(self.WS_HEADER + body)

    def test_a_workstream_without_schema_version_is_not_silently_exempt(self):
        root = self.make_repo()
        (root / "docs/workstreams/WS-20260803-noschema.md").write_text(
            "---\nid: WS-20260803-noschema\nstatus: active\n---\n\n# No envelope, no gates\n"
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("schema_version" in error for error in errors),
            "an unversioned workstream skipped the entire contract and passed",
        )

    def test_an_issue_without_schema_version_is_not_silently_exempt(self):
        root = self.make_repo()
        (root / "docs/issues/ISSUE-20260803-noschema.md").write_text(
            "---\nid: ISSUE-20260803-noschema\nstatus: active\n---\n\n# No guide impact\n"
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("schema_version" in error for error in errors))

    def test_an_issue_in_the_queue_table_needs_a_block(self):
        root = self.make_repo()
        self.write_ws(
            root,
            """
| Issue | Status |
| --- | --- |
| ISSUE-01-done | complete |
| ISSUE-02-still-open | pending |

### ISSUE-01-done

- status: complete
- depends_on: []
- guide_impact: none
- related_guides: []
- guide_impact_reason: x
""",
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("ISSUE-02-still-open" in error for error in errors),
            "an issue listed in the queue table with no block was invisible to every gate",
        )

    def test_a_block_missing_from_the_queue_table_is_reported(self):
        root = self.make_repo()
        self.write_ws(
            root,
            """
| Issue | Status |
| --- | --- |
| ISSUE-01-done | complete |

### ISSUE-01-done

- status: complete
- depends_on: []
- guide_impact: none
- related_guides: []
- guide_impact_reason: x

### ISSUE-02-orphan

- status: pending
- depends_on: []
- guide_impact: none
- related_guides: []
- guide_impact_reason: x
""",
        )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(any("ISSUE-02-orphan" in error for error in errors))

    def test_a_duplicate_id_is_reported_rather_than_silently_shadowed(self):
        root = self.make_repo()
        for directory in ("docs/guides", "docs/guides/archive"):
            (root / directory).mkdir(parents=True, exist_ok=True)
            (root / directory / "GUIDE-dup.md").write_text(
                "---\nid: GUIDE-dup\nstatus: active\ncreated_at: 2026-08-03\n"
                "updated_at: 2026-08-03\nsource_workstreams: []\nsource_issues: []\n---\n\n# Dup\n"
            )

        errors, _warnings = MODULE.validate_repo(root)

        self.assertTrue(
            any("GUIDE-dup" in error and "duplicate" in error.lower() for error in errors)
        )


class ArchiveGateTest(ValidateRepoDocsV2Test):
    """An archive gate that cannot see unfinished work reports success on it.

    Every fixture here must be valid apart from the one thing under test —
    otherwise the script exits non-zero for an unrelated reason and the test
    passes without exercising the gate at all.
    """

    ARCHIVE_WS = Path(__file__).resolve().parents[1] / "archive_workstream.py"
    ARCHIVE_ISSUE = Path(__file__).resolve().parents[1] / "archive_issue.py"

    WS_BODY = """
# Example

## Authorization Envelope

- Autonomous actions allowed: edit docs
- Confirm first: external sends

## Human Gates

x

## Issue Queue

| Issue | Status |
| --- | --- |
| ISSUE-01-a | complete |

### ISSUE-01-a

- status: complete
- depends_on: []
- runnability: ready
- guide_impact: none
- related_guides: []
- guide_impact_reason: x

#### Acceptance

- verify: machine — the check exits 0

## Completion

"""

    def run_script(self, script, *args):
        return subprocess.run(
            [sys.executable, str(script), *args], capture_output=True, text=True
        )

    def write_ws(self, root, completion, ws_id="WS-20260803-boxes", status_line="status: active\n"):
        (root / "docs/workstreams" / f"{ws_id}.md").write_text(
            "---\nschema_version: 2\n"
            f"id: {ws_id}\n" + status_line +
            "created_at: 2026-08-03\nupdated_at: 2026-08-03\n"
            'branch: ""\npr: ""\n'
            "human_boundary_confirmed_at: 2026-08-03\nnext_human_gate: none\n"
            "related_specs: []\nrelated_guides: []\n---\n"
            + self.WS_BODY
            + completion
        )
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-08-03\ncurrent_focus: x\n---\n\n# 00 Index\n\n"
            f"- [{ws_id}](workstreams/{ws_id}.md)\n"
        )

    def assert_fixture_is_otherwise_valid(self, root):
        errors, _warnings = MODULE.validate_repo(root)
        self.assertEqual(
            errors, [], "fixture is invalid for an unrelated reason; the test would pass vacuously"
        )

    def test_unchecked_boxes_in_other_spellings_still_block_archiving(self):
        for label, box in (
            ("canonical", "- [ ] item"),
            ("no inner space", "- [] item"),
            ("indented", "  - [ ] item"),
            ("asterisk bullet", "* [ ] item"),
            ("no trailing text", "- [ ]"),
        ):
            with self.subTest(label):
                root = self.make_repo()
                self.write_ws(root, "- [x] done\n" + box + "\n")
                self.assert_fixture_is_otherwise_valid(root)

                result = self.run_script(
                    self.ARCHIVE_WS, "WS-20260803-boxes", "--repo", str(root)
                )

                self.assertNotEqual(
                    result.returncode, 0, f"{label} archived with work unchecked"
                )
                self.assertIn(
                    "checklist",
                    (result.stdout + result.stderr).lower(),
                    f"{label} was blocked, but not by the checklist gate",
                )

    def test_a_fully_checked_workstream_still_archives(self):
        root = self.make_repo()
        self.write_ws(root, "- [x] done\n- [X] also done\n")
        self.assert_fixture_is_otherwise_valid(root)

        result = self.run_script(
            self.ARCHIVE_WS, "WS-20260803-boxes", "--repo", str(root)
        )

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_an_unstarted_issue_does_not_archive(self):
        root = self.make_repo()
        (root / "docs/issues/ISSUE-20260803-unstarted.md").write_text(
            "---\nschema_version: 2\nid: ISSUE-20260803-unstarted\nstatus: active\n"
            "created_at: 2026-08-03\nupdated_at: 2026-08-03\nbranch: main\n"
            "guide_impact: none\nguide_impact_reason: x\nrelated_guides: []\n---\n\n"
            "# Unstarted\n\n## Acceptance\n\n- verify: machine — the check exits 0\n\n"
            "## Completion\n\n- [ ] not done yet\n"
        )
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-08-03\ncurrent_focus: x\n---\n\n# 00 Index\n\n"
            "- [ISSUE-20260803-unstarted](issues/ISSUE-20260803-unstarted.md)\n"
        )
        self.assert_fixture_is_otherwise_valid(root)

        result = self.run_script(
            self.ARCHIVE_ISSUE, "ISSUE-20260803-unstarted.md", "--repo", str(root)
        )

        self.assertNotEqual(result.returncode, 0, "an unstarted issue archived cleanly")
        self.assertIn("checklist", (result.stdout + result.stderr).lower())

    def test_the_self_referential_archive_box_does_not_block_archiving(self):
        """The "Workspace archived" box is the action this script performs.

        archive_issue.py already ticks its equivalent box on the script's behalf;
        archive_workstream.py did not, so every straight run failed on a box the
        caller could only satisfy by lying before the fact.
        """
        root = self.make_repo()
        self.write_ws(
            root,
            "- [x] Every issue meets its acceptance criteria\n"
            "- [ ] Workstream archived when complete\n",
        )
        self.assert_fixture_is_otherwise_valid(root)

        result = self.run_script(
            self.ARCHIVE_WS, "WS-20260803-boxes", "--repo", str(root)
        )

        self.assertEqual(
            result.returncode,
            0,
            f"self-referential box blocked archiving: {result.stdout}{result.stderr}",
        )

    def test_archiving_reports_a_status_it_could_not_set(self):
        root = self.make_repo()
        self.write_ws(root, "- [x] done\n", status_line="")
        # No `status:` line at all: update_scalar substitutes nothing, so the file
        # would move into archive/ having never been marked archived.
        result = self.run_script(
            self.ARCHIVE_WS, "WS-20260803-boxes", "--repo", str(root)
        )

        self.assertNotEqual(
            result.returncode, 0, "archived without ever marking the file archived"
        )
        self.assertIn("status", (result.stdout + result.stderr).lower())


class ValidatorReportsItsTargetTest(unittest.TestCase):
    SCRIPT = Path(__file__).resolve().parents[1] / "validate_repo_docs.py"

    def test_the_resolved_repository_is_printed(self):
        root = Path(tempfile.mkdtemp())
        for path in (
            "docs/adrs",
            "docs/specs",
            "docs/issues/archive",
            "docs/workstreams/archive",
            "docs/guides",
        ):
            (root / path).mkdir(parents=True, exist_ok=True)
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-08-03\ncurrent_focus: x\n---\n\n# 00 Index\n"
        )

        result = subprocess.run(
            [sys.executable, str(self.SCRIPT), str(root)],
            capture_output=True,
            text=True,
        )

        self.assertIn(
            str(root.resolve()),
            result.stdout,
            "the validator never said which repository it read",
        )

if __name__ == "__main__":
    unittest.main()
