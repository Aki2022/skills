"""The direct route: someone calls own-website-audit without the entry skill. The gate holds here too, and the skill
owns no measuring script (the measurement is own-web-build's)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-web-build/scripts/tests'))
from reviewer_gate_testkit import SKILLS, assert_route  # noqa: E402

AUDIT = SKILLS / 'own-website-audit'


def test_the_direct_route_starts_no_reviewer_when_nothing_was_measured_or_the_run_failed_or_there_is_no_profile(tmp_path):
    assert_route(tmp_path, 'direct')


def test_the_audit_skill_tells_the_agent_to_pass_the_gate_before_the_reviewer():
    text = (AUDIT / 'SKILL.md').read_text()
    assert '/own-web-build/scripts/reviewer_gate.py' in text and '--route direct' in text
    assert '/own-web-build/scripts/measure_render.mjs' in text and 'measure_lp' not in text
    for name in ('SKILL.md', 'references/marketing-rubric.md'):
        assert 'measure_lp' not in (AUDIT / name).read_text(), name


def test_the_audit_skill_has_no_measuring_script_and_no_threshold_table():
    scripts = sorted(p.name for p in (AUDIT / 'scripts').iterdir() if p.is_file())
    assert scripts == [], f'own-website-audit/scripts holds files: {scripts}'
    assert not (AUDIT / 'references/aesthetic-thresholds.md').exists()
    assert (AUDIT / 'references/marketing-rubric.md').is_file(), 'the marketing rubric stays: it is this skill'
