"""Content Tone routes to the IA voice, not to design/DESIGN.md.

SPEC-information-architecture-standardization v5 (Q10, biz_ops): the voice canon is
the IA "声" atom. A reference DESIGN.md's Content Tone is handed to own-infoarch-author
as proposed material for the voice; the repository's DESIGN.md gets no Content Tone
section.
"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
MAPPING = ROOT / "own-brand-resolve/plugins/awesome-design-md-jp-import/mapping.md"
TYPOGRAPHY = ROOT / "own-web-design/validators/common/typography.md"


def content_tone_row():
    rows = [line for line in MAPPING.read_text(encoding="utf-8").splitlines()
            if re.match(r"\|\s*Content Tone", line)]
    assert len(rows) == 1, rows
    return [cell.strip() for cell in rows[0].strip("|").split("|")]


class ContentToneRouting(unittest.TestCase):
    def test_content_tone_goes_to_the_ia_voice_as_proposed_material(self):
        _, copy_to, _ = content_tone_row()
        self.assertIn("design/info_arch", copy_to)
        self.assertIn("声", copy_to)
        self.assertIn("proposed", copy_to)

    def test_content_tone_is_not_written_into_design_md(self):
        _, copy_to, _ = content_tone_row()
        self.assertNotIn("DESIGN.md", copy_to)
        self.assertNotIn("散文に", copy_to)

    def test_web_validator_points_at_the_ia_voice_not_a_design_md_section(self):
        text = TYPOGRAPHY.read_text(encoding="utf-8")
        self.assertNotIn("DESIGN.md の Content Tone", text)


if __name__ == "__main__":
    unittest.main()
