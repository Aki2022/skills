"""The entry route: own-deliverable-check hands a landing-type target to the audit only through the reviewer gate."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / 'own-web-build/scripts/tests'))
from reviewer_gate_testkit import SKILLS, assert_route  # noqa: E402


def test_the_entry_route_starts_no_reviewer_when_nothing_was_measured_or_the_run_failed_or_there_is_no_profile(tmp_path):
    assert_route(tmp_path, 'entry')


def test_the_entry_skill_tells_the_agent_to_pass_the_gate_and_not_to_start_the_reviewer_otherwise():
    text = (SKILLS / 'own-deliverable-check/SKILL.md').read_text()
    assert '/own-web-build/scripts/reviewer_gate.py' in text and '--route entry' in text
    assert 'measure_render.mjs' in text and 'measure_lp' not in text
    assert 'measurement-profiles.json' in text
