"""The coverage table checks a deliverable's official names, fixed values and expiry against docs/dictionary/.

Without the row, "check this deliverable" never runs the dictionary check, and a wrong value or an expired
campaign ships while the router reports green.
"""
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
TEXT = (SKILLS / 'own-deliverable-check/SKILL.md').read_text(encoding='utf-8')
CLI = SKILLS / 'own-doc-update/scripts/dictionary.py'


def row(label):
    hits = [ln for ln in TEXT.splitlines() if ln.startswith('|') and label in ln.split('|')[1]]
    assert len(hits) == 1, (label, hits)
    return [c.strip() for c in hits[0].strip('|').split('|')]


def test_dictionary_row_runs_the_html_check_for_every_html_deliverable():
    cells = row('正式名称・固定値')
    lp, product_ui, docs = cells[1], cells[2], cells[3]
    assert 'dictionary.py check-html' in lp
    assert product_ui in ('同左',) or 'dictionary.py check-html' in product_ui
    assert 'HTML' in docs and ('同左' in docs or 'dictionary.py check-html' in docs)


def test_the_named_check_exists_and_nothing_checked_is_not_a_pass():
    assert CLI.is_file()
    r = subprocess.run([sys.executable, str(CLI), 'check-html', str(SKILLS / 'nonexistent-repo')], capture_output=True, text=True)
    assert r.returncode == 2 and 'not checked' in r.stdout
    # The router must say so rather than drop the row.
    assert 'docs/dictionary/' in TEXT and '確かめていない' in TEXT.split('正式名称・固定値', 1)[1]
