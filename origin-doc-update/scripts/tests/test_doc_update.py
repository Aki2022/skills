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


if __name__ == "__main__":
    unittest.main()
