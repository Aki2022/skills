import importlib.util
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


if __name__ == "__main__":
    unittest.main()
