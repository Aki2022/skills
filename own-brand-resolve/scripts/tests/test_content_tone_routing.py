"""Content Tone routes to the outline tone, not to design/DESIGN.md.

SPEC-deliverable-production-layers requirement 8 (biz_ops): the repository-wide voice is retired; the tone of a
deliverable belongs to its outline (infoarch). A reference DESIGN.md's Content Tone is handed to own-infoarch-author
as material for the outline tone; the repository's DESIGN.md gets no Content Tone section.
"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
MAPPING = ROOT / "own-brand-resolve/plugins/awesome-design-md-jp-import/mapping.md"
TYPOGRAPHY = ROOT / "own-web-build/validators/common/typography.md"


def content_tone_row():
    rows = [line for line in MAPPING.read_text(encoding="utf-8").splitlines()
            if re.match(r"\|\s*Content Tone", line)]
    assert len(rows) == 1, rows
    return [cell.strip() for cell in rows[0].strip("|").split("|")]


class ContentToneRouting(unittest.TestCase):
    def test_content_tone_goes_to_the_outline_tone_as_material(self):
        _, copy_to, _ = content_tone_row()
        self.assertIn("own-infoarch-author", copy_to)
        self.assertIn("outline", copy_to)
        self.assertIn("tone", copy_to)
        self.assertNotIn("声", copy_to)

    def test_content_tone_is_not_written_into_design_md(self):
        _, copy_to, _ = content_tone_row()
        self.assertNotIn("DESIGN.md", copy_to)
        self.assertNotIn("散文に", copy_to)

    def test_web_validator_points_at_the_outline_tone_not_a_design_md_section(self):
        text = TYPOGRAPHY.read_text(encoding="utf-8")
        self.assertNotIn("DESIGN.md の Content Tone", text)
        self.assertIn("outline", text)
        self.assertNotIn("「声」", text)


if __name__ == "__main__":
    unittest.main()
