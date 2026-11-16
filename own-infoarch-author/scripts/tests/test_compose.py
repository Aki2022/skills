"""From positioning to composition (v5 requirements 12, 15, 19 on the v6 positioning; ISSUE-14 replaces the outline).

Steps 2-4 choose positioning atoms, select the messages that stand on them, and record the choice and its rationale
in the outline (field `strategy` until ISSUE-14). Whether a message really serves its positioning is a Jev judgment
(ia-strategy-fit until ISSUE-16) whose input carries the positioning path hash; it is advisory until measured (D8).
"""
import copy
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
import judge  # noqa: E402
from test_check import DECK, OUTLINE, SSOT, check, cli, codes, write_repo  # noqa: E402


def model(ssot=SSOT):
    return ia_check.Model(root=Path('.'), ssot=ssot)


def test_compose_selects_messages_standing_on_the_chosen_positioning():
    assert ia_check.compose(model(), 'aud-0001', ['val-0001']) == ['msg-0001', 'msg-0002', 'msg-0003']
    # another segment and value select different messages from the same SSOT
    assert ia_check.compose(model(), 'aud-0002', ['val-0002']) == ['msg-0002', 'msg-0003']


def test_compose_cli(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('compose', '--root', str(tmp_path), '--audience', 'aud-0002', '--positioning', 'val-0002')
    assert rc == 0 and out['messages'] == ['msg-0002', 'msg-0003'] and out['segment'] == 'seg-0002', out


def test_outline_must_record_its_strategy_choice(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    del outline['strategy']
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_SCHEMA' in codes(report), report['findings']


def test_chosen_atom_of_another_segment_fails(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['strategy']['choose'] = ['val-0002']
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_COMPOSITION_POSITIONING_INVALID' in codes(report), report['findings']


def test_message_off_the_chosen_positioning_warns(tmp_path):
    outline = copy.deepcopy(OUTLINE)
    outline['strategy']['choose'] = ['val-0003']  # the parity value: no LP message serves it
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert 'IA_COMPOSITION_OFF_POSITIONING' in codes(report, 'warn'), report['findings']


def test_dropping_the_top_priority_needs_a_recorded_reason(tmp_path):  # requirement 12
    outline = copy.deepcopy(OUTLINE)
    outline['pages'][0]['sections'] = [s for s in outline['pages'][0]['sections'] if s.get('message') != 'msg-0001']
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert 'IA_COMPOSITION_DEVIATION_UNEXPLAINED' in codes(report, 'warn'), report['findings']
    outline['strategy']['deviation'] = 'A/B の比べる案：1 位を差し替えた案'
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert 'IA_COMPOSITION_DEVIATION_UNEXPLAINED' not in codes(report, 'warn'), report['findings']


def fit_rows(ssot=SSOT):
    jl = judge._jl()
    return {r['id']: r for r in judge.outline_candidates('lp-municipal', OUTLINE, ssot, jl.load_policy(), jl)
            if r['input']['kind'] == 'ia-strategy-fit'}


def test_fit_candidates_carry_the_positioning_path_hash():
    rows = fit_rows()
    assert set(rows) == {'ia-strategy-fit:lp-municipal/top/1', 'ia-strategy-fit:lp-municipal/top/2', 'ia-strategy-fit:lp-municipal/top/3'}
    ctx = rows['ia-strategy-fit:lp-municipal/top/1']['input']['context']
    assert ctx['message']['id'] == 'msg-0001' and len(ctx['path_hash']) == 64
    ids = {a['id'] for a in ctx['positioning']}
    assert ids >= {'val-0001', 'str-0001', 'seg-0001', 'satisfies:val-0001:pan-0001', 'produces:str-0001:val-0001'}
    assert 'val-0003' not in ids and not any('seg-0002' in i or 'job-0002' in i for i in ids)


def test_positioning_change_invalidates_only_affected_fit_judgments():  # no reuse after a positioning change
    base = fit_rows()
    changed = copy.deepcopy(SSOT)
    changed['strengths']['str-0001']['text'] = '別の強み'
    after = fit_rows(changed)
    assert all(after[k]['judgment_hash'] != base[k]['judgment_hash'] for k in base)  # every LP message stands on val-0001
    unrelated = copy.deepcopy(SSOT)
    unrelated['values']['val-0002']['text'] += '（改）'
    same = fit_rows(unrelated)
    assert all(same[k]['judgment_hash'] == base[k]['judgment_hash'] for k in base)


def test_mapping_records_the_positioning_choice(tmp_path):
    write_repo(tmp_path)
    rc, table = cli('mapping', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert rc == 0 and table['strategy']['choose'] == ['val-0001'], table
