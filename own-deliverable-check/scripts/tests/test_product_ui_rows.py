"""The coverage table has the structure, skeleton and state rows for a product UI, and they name a check that exists."""
import re
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
TEXT = (SKILLS / 'own-deliverable-check/SKILL.md').read_text(encoding='utf-8')


def row(label):
    hits = [ln for ln in TEXT.splitlines() if ln.startswith('|') and label in ln.split('|')[1]]
    assert len(hits) == 1, (label, hits)
    return [c.strip() for c in hits[0].strip('|').split('|')]


def test_structure_skeleton_and_state_rows_run_the_screen_check_for_a_product_ui():
    for label in ('構造', '骨格', '状態'):
        cells = row(label)
        assert 'audit_screen.py' in cells[2] or '同スクリプト' in cells[2], (label, cells)
    assert 'audit_screen.py' in row('構造')[2]


def test_the_check_the_table_names_exists_and_the_skill_that_owns_it_says_how_to_run_it():
    assert (SKILLS / 'own-web-build/scripts/audit_screen.py').is_file() and (SKILLS / 'own-web-build/scripts/measure_screen.mjs').is_file()
    assert 'audit_screen.py' in (SKILLS / 'own-web-build/SKILL.md').read_text(encoding='utf-8')


def test_a_product_ui_without_a_blueprint_is_reported_as_missing_not_skipped_silently():
    assert re.search(r'blueprint が無い成果物には、この 3 行は「欠落', TEXT)
