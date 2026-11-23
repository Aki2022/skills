"""The coverage table has the information-architecture rows (layers requirement 7), and a tree without the IA is reported.

Without these rows "check this deliverable" never runs the IA consistency check or the meaning judgment, and the
router stays silent about it: exactly the silence the router exists to remove.
"""
import re
import sys
from pathlib import Path

import pytest

SKILLS = Path(__file__).resolve().parents[3]
TEXT = (SKILLS / 'own-deliverable-check/SKILL.md').read_text(encoding='utf-8')
sys.path.insert(0, str(SKILLS / 'own-design-author/scripts'))
import layer_paths  # noqa: E402  (the layer rule: where the positioning and the infoarch live)

CONSISTENCY = '情報設計の整合'
SEMANTIC = '意味の判定'


def row(text, label):
    hits = [ln for ln in text.splitlines() if ln.startswith('|') and label in ln.split('|')[1]]
    assert len(hits) == 1, (label, hits)
    return [c.strip() for c in hits[0].strip('|').split('|')]


def assert_ia_rows(text):
    """Everything the rows must say; the tests below call it on the real file and on a file with a row removed."""
    consistency = row(text, CONSISTENCY)
    # LP and document name both checks; the product UI cell says "same as the LP, plus blueprint"
    for cell in (consistency[1], consistency[3]):
        assert 'infoarch-author check' in cell and 'positioning-author check' in cell, consistency
    assert '同左' in consistency[2], consistency
    assert 'infoarch-author blueprint' in consistency[2], consistency  # the product UI's blueprint
    semantic = row(text, SEMANTIC)
    for cell in semantic[1:]:
        assert 'infoarch-author judge' in cell or '同左' in cell, semantic
    assert 'infoarch-author judge' in semantic[1] and 'infoarch-author evaluate' in semantic[1], semantic
    # one kind of judgment, a warning only: it never decides pass or fail and never stops a shipment
    assert 'ia-value-fit' in semantic[1] and '警告だけ' in semantic[1] and '止めない' in semantic[1], semantic


def test_the_real_table_has_both_ia_rows():
    assert_ia_rows(TEXT)


@pytest.mark.parametrize('label', [CONSISTENCY, SEMANTIC])
def test_removing_a_row_turns_the_check_red(label):
    kept = '\n'.join(ln for ln in TEXT.splitlines() if not (ln.startswith('|') and label in ln.split('|')[1]))
    with pytest.raises(AssertionError):
        assert_ia_rows(kept)


def test_lp_and_product_ui_both_get_the_consistency_check():
    cells = row(TEXT, CONSISTENCY)
    lp, product_ui = cells[1], cells[2]
    assert 'positioning-author check' in lp and 'infoarch-author check' in lp
    assert '同左' in product_ui and 'infoarch-author blueprint' in product_ui and 'infoarch-author blueprint' not in lp


def test_every_command_the_rows_name_exists():
    infoarch = (SKILLS / 'own-infoarch-author/scripts/infoarch_author.py').read_text(encoding='utf-8')
    positioning = (SKILLS / 'own-positioning-author/scripts/positioning_author.py').read_text(encoding='utf-8')
    for sub in ('check', 'blueprint', 'judge', 'evaluate'):
        assert f"add_parser('{sub}'" in infoarch, sub
    assert "add_parser('check'" in positioning
    assert (SKILLS / 'own-infoarch-author/scripts/infoarch-author').is_file()
    assert (SKILLS / 'own-positioning-author/scripts/positioning-author').is_file()


def test_a_tree_without_the_ia_is_reported_as_not_run_never_as_a_pass_or_a_failure():
    """`check` on a tree with no IA exits 2 ("it is not there"), which is not a verdict on the deliverable."""
    section = TEXT.split('情報設計の 2 行は', 1)[1].split('### 3.', 1)[0]
    assert layer_paths.rel('positioning_file') in section and layer_paths.rel('infoarch') in section, section
    assert '当てなかった' in section and '欠落' in section
    assert re.search(r'失格にも合格にも数えない', section), section
    # the report contract has to carry them too
    report = TEXT.split('### 4.', 1)[1]
    assert '情報設計' in report and ('欠落' in report), report


def test_the_semantic_judgment_runs_after_the_deterministic_checks_and_is_reported_as_unmeasured():
    section = TEXT.split('情報設計の 2 行は', 1)[1].split('### 3.', 1)[0]
    assert 'judge' in section and '決定論の検査が緑になった後に走らせる' in section, section
    assert '未カバー' in section and 'screened' in section and '精度は未測定' in section, section
