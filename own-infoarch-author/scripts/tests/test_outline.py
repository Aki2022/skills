"""outline (SPEC-information-architecture-standardization v6, requirements 6-9 and 13; Quality 1 and 3).

An outline stands on one positioning (its target's segment), keeps the constants and the objective fixed, and
varies messages, sections and copy. Requirement 7 disqualifies it; ids come from content only; two sets differ by
elements, and a difference no element owns disqualifies the variant.
"""
import copy
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
import outline as ia_outline  # noqa: E402
import positioning  # noqa: E402
from test_check import DECK, OUTLINE, SKILL, SSOT, check, cli, codes, confirm, dump, write_repo  # noqa: E402

TABLE = json.loads((SKILL / 'fixtures/outline-red-cases.json').read_text(encoding='utf-8'))
CONDITIONS = {c['id']: c for c in TABLE['conditions']}
CASES, DIFF_CASES = TABLE['cases'], TABLE['diff_cases']


def apply(ops, outline=OUTLINE):
    outline = copy.deepcopy(outline)
    unapproved = []
    for op in ops:
        if op['op'] == 'unapprove':
            unapproved.append(op['key'])
            continue
        *parent, last = op['path']
        node = outline
        for k in parent:
            node = node[k]
        if op['op'] == 'set':
            node[last] = op['value']
        elif op['op'] == 'delete':
            del node[last]
        else:
            raise KeyError(op['op'])
    return outline, tuple(unapproved)


def write_case(root, case):
    outline, unapproved = apply(case['ops'])
    ssot = copy.deepcopy(SSOT)
    for op in case.get('ssot_ops', []):
        if op['op'] != 'empty_segment':
            raise KeyError(op['op'])
        ssot['segments'][op['id']] = {'state': 'empty', 'provenance': 'proposal'}
        ssot['ledger']['next']['seg'] = int(op['id'].split('-')[1]) + 1
    write_repo(root, ssot, outlines=(outline, DECK), confirmed=False)
    confirm(root, ssot, skip=unapproved)


# ---- the table (Quality 1) ----------------------------------------------------------------------------------------

def test_every_condition_has_a_red_case_and_every_case_names_a_condition():
    covered = {c['condition'] for c in CASES + DIFF_CASES}
    assert sorted(set(CONDITIONS) - covered) == []  # 対応の無い失格条件 0 件
    assert sorted(covered - set(CONDITIONS)) == []
    ids = [c['id'] for c in CASES + DIFF_CASES]
    assert len(set(ids)) == len(ids)
    assert len([c for c in CONDITIONS.values() if c['requirement'] == '7']) == 16  # spec v6 requirement 7, split by code
    assert [c['id'] for c in CONDITIONS.values() if c['requirement'] == '9'] == ['9-unassigned-difference']


def test_every_outline_code_is_raised_by_the_checker():
    source = ''.join((SKILL / f'scripts/{n}').read_text(encoding='utf-8') for n in ('outline.py', 'check.py'))
    for c in CONDITIONS.values():
        if c['code'] != 'IA_SCHEMA':
            assert f"'{c['code']}'" in source, c


@pytest.mark.parametrize('case', CASES, ids=[c['id'] for c in CASES])
def test_red_case_fails_with_its_condition_code(case, tmp_path):
    write_case(tmp_path, case)
    rc, report = check(tmp_path)
    assert rc == 2, report['findings']
    assert CONDITIONS[case['condition']]['code'] in codes(report), codes(report)


@pytest.mark.parametrize('case', DIFF_CASES, ids=[c['id'] for c in DIFF_CASES])
def test_unassigned_difference_disqualifies_the_variant(case, tmp_path):
    variant = copy.deepcopy(OUTLINE)
    variant['id'], variant['outputs'] = 'lp-municipal-b', ['site/b.html']
    write_repo(tmp_path / 'a', outlines=(OUTLINE, DECK))
    ssot_b = copy.deepcopy(SSOT)
    for fid, value in case.get('fact_b', {}).items():
        ssot_b['facts'][fid]['value'] = value
    write_repo(tmp_path / 'b', ssot_b, outlines=(variant, DECK))
    extra = []
    if 'dom_a' in case:
        for name in ('a', 'b'):
            (tmp_path / f'dom-{name}.json').write_text(json.dumps(case[f'dom_{name}'], ensure_ascii=False), encoding='utf-8')
        extra = ['--dom-a', str(tmp_path / 'dom-a.json'), '--dom-b', str(tmp_path / 'dom-b.json')]
    rc, out = cli('diff', '--root', str(tmp_path / 'a'), '--root-b', str(tmp_path / 'b'), '--a', 'lp-municipal',
                  '--b', 'lp-municipal-b', *extra)
    assert rc == 2 and CONDITIONS[case['condition']]['code'] in [f['code'] for f in out['findings']], out
    assert out['changes'] == []  # nothing in the sets changed: the difference belongs to no element


def test_unknown_operation_is_rejected():
    with pytest.raises(KeyError):
        apply([{'op': 'nope', 'path': ['id']}])


# ---- what passes, warns and is reported (requirement 7) -------------------------------------------------------------

def test_base_outlines_pass(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
    assert not [f for f in report['findings'] if f['code'].startswith('IA_OUTLINE_')], report['findings']


def test_unplaced_support_message_is_a_warning(tmp_path):
    outline, _ = apply([{'op': 'set', 'path': ['messages', 'msg-0004'], 'value': {'value': 'val-0003', 'rank': 'support', 'type': 'normal'}}])
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 1 and 'IA_OUTLINE_MESSAGE_UNPLACED' in codes(report, 'warn'), report['findings']


def test_declared_no_cta_passes_and_is_always_reported(tmp_path):
    outline, _ = apply([{'op': 'set', 'path': ['no_cta'], 'value': True},
                        {'op': 'delete', 'path': ['messages', 'msg-0003']},
                        {'op': 'set', 'path': ['sections', 'sec-0003', 'places'], 'value': ['msg-0001']},
                        {'op': 'delete', 'path': ['copies', 'cpy-0005']}])
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
    hits = [f for f in report['findings'] if f['code'] == 'IA_OUTLINE_NO_CTA_INTENDED']
    assert len(hits) == 1 and 'CTA なし（意図）' in hits[0]['message'] and 'lp-municipal' in hits[0]['message']


def test_unapproved_value_only_warns_in_a_draft(tmp_path):
    write_repo(tmp_path, confirmed=False)
    confirm(tmp_path, SSOT, skip=('val-0002',))
    rc, report = check(tmp_path)
    assert 'IA_OUTLINE_VALUE_UNAPPROVED' in codes(report, 'warn') and 'IA_OUTLINE_VALUE_UNAPPROVED' not in codes(report)


# ---- ids (requirement 8, Quality 3) -----------------------------------------------------------------------------------

def model(ssot=SSOT, outlines=(OUTLINE, DECK), skip=()):
    m = ia_check.Model(root=Path('.'), ssot=ssot)
    for o in outlines:
        m.outlines[o['id']] = (Path(f'design/info_arch/outlines/{o["id"]}.yaml'), o)
    rows = {k: v for k, v in positioning.approval_targets(m).items() if k[len(positioning.APPROVAL_PREFIX):] not in skip}
    return m, positioning.approved_keys(m, rows)


def ids_of(oid='lp-municipal', **kw):
    m, approved = model(**kw)
    return ia_outline.set_ids(m, oid, approved)


def test_ids_are_reproducible_and_shaped():
    a, b = ids_of(), ids_of()
    assert a == b
    assert a['outline'].startswith('out-') and a['positioning'].startswith('pos-')
    assert a['set'] == f'{a["positioning"]}+{a["outline"]}'


def test_section_order_moves_outline_and_set_not_positioning():
    swapped = copy.deepcopy(OUTLINE)
    swapped['sections']['sec-0001']['order'], swapped['sections']['sec-0002']['order'] = 2, 1
    a, b = ids_of(), ids_of(outlines=(swapped, DECK))
    assert b['outline'] != a['outline'] and b['set'] != a['set'] and b['positioning'] == a['positioning']


def test_copy_order_and_no_cta_are_part_of_the_outline_id():
    swapped = copy.deepcopy(OUTLINE)
    swapped['copies']['cpy-0001']['order'], swapped['copies']['cpy-0002']['order'] = 2, 1
    assert ids_of(outlines=(swapped, DECK))['outline'] != ids_of()['outline']


def test_value_text_moves_positioning_and_set_not_outline():
    changed = copy.deepcopy(SSOT)
    changed['values']['val-0001']['text'] = '運用の手間がなくなる'
    a, b = ids_of(), ids_of(ssot=changed)  # re-approved as the owner would
    assert b['outline'] == a['outline'] and b['positioning'] != a['positioning'] and b['set'] != a['set']


def test_file_order_and_placement_do_not_move_the_outline_id():
    shuffled = copy.deepcopy(OUTLINE)
    for key in ('messages', 'sections', 'copies'):
        shuffled[key] = dict(reversed(list(shuffled[key].items())))
    shuffled = dict(reversed(list(shuffled.items())))
    moved = copy.deepcopy(OUTLINE)
    moved['outputs'] = ['site/lp/index.html']  # where it is written is not what it says
    assert ids_of(outlines=(shuffled, DECK))['outline'] == ids_of()['outline']
    assert ia_outline.outline_id(moved) == ia_outline.outline_id(OUTLINE)


def test_deliverable_id_does_not_move_the_outline_id():
    renamed = copy.deepcopy(OUTLINE)
    renamed['id'] = 'lp-municipal-b'
    assert ia_outline.outline_id(renamed) == ia_outline.outline_id(OUTLINE)


def test_unapproved_positioning_has_no_set_id():
    out = ids_of(skip=('val-0001', 'val-0003', 'job-0001', 'pan-0001', 'gan-0001', 'gan-0002', 'seg-0001',
                       'alt-0001', 'alt-0002', 'cat-0001', 'str-0001', 'prf-0001', 'prf-0003'))
    assert out['positioning'] is None and out['set'] is None and out['outline'].startswith('out-')


def test_set_cli_reports_ids_and_elements(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('set', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert rc == 0, out
    assert out['set'] == f'{out["positioning"]}+{out["outline"]}'
    assert sorted(out['elements']['copies']) == ['cpy-0001', 'cpy-0002', 'cpy-0003', 'cpy-0004', 'cpy-0005']
    assert all(len(h) == 64 for h in out['elements']['messages'].values())


# ---- diff (requirement 9) ---------------------------------------------------------------------------------------------

def diff(a, b, ssot_b=SSOT, skip_b=()):
    ma, appr_a = model(outlines=(a,))
    mb, appr_b = model(ssot=ssot_b, outlines=(b,), skip=skip_b)
    return ia_outline.diff(ma, a['id'], appr_a, mb, b['id'], appr_b)


def test_diff_lists_only_the_changed_elements():
    b = copy.deepcopy(OUTLINE)
    b['sections']['sec-0001']['order'], b['sections']['sec-0002']['order'] = 2, 1
    b['copies']['cpy-0002']['text'] = '運用の手間が 3 分の 1 になります'
    out = diff(OUTLINE, b)
    assert out['changes'] == [
        {'layer': 'outline', 'element': 'copies/cpy-0002', 'change': 'changed', 'attributes': ['text']},
        {'layer': 'outline', 'element': 'sections/sec-0001', 'change': 'changed', 'attributes': ['order']},
        {'layer': 'outline', 'element': 'sections/sec-0002', 'change': 'changed', 'attributes': ['order']}]
    assert out['changed_layers'] == ['outline'] and out['same_experiment'] is True and out['findings'] == []


def test_diff_of_identical_sets_is_empty():
    out = diff(OUTLINE, copy.deepcopy(OUTLINE))
    assert out['changes'] == [] and out['changed_layers'] == [] and out['ids']['a'] == out['ids']['b']


def test_diff_names_added_and_removed_elements_and_the_lead_change():
    b = copy.deepcopy(OUTLINE)
    b['messages']['msg-0001']['rank'], b['messages']['msg-0002']['rank'] = 'support', 'lead'
    del b['copies']['cpy-0004']
    b['copies']['cpy-0009'] = {'section': 'sec-0003', 'order': 1, 'message': 'msg-0001', 'text': '保守の要らない運用を'}
    out = diff(OUTLINE, b)
    assert {(c['element'], c['change']) for c in out['changes']} == {
        ('messages/msg-0001', 'changed'), ('messages/msg-0002', 'changed'),
        ('copies/cpy-0004', 'removed'), ('copies/cpy-0009', 'added')}


def test_positioning_and_outline_may_change_together():
    ssot_b = copy.deepcopy(SSOT)
    [edge] = [e for e in ssot_b['edges'] if e['kind'] == 'fulfills' and e['from'] == 'alt-0002' and e['to'] == 'pan-0001']
    edge['degree'] = 'none'
    b = copy.deepcopy(OUTLINE)
    b['copies']['cpy-0001']['text'] = '運用、重くないですか'
    out = diff(OUTLINE, b, ssot_b=ssot_b)
    assert out['changed_layers'] == ['outline', 'positioning']
    assert {'layer': 'positioning', 'element': 'fulfills:alt-0002:pan-0001', 'change': 'changed', 'attributes': ['degree']} in out['changes']


def test_different_constants_or_objective_is_another_experiment():
    for path, value in ((('tone',), 'くだけた調子'), (('touchpoint',), 'ad'), (('target', 'condition'), '別の条件'),
                        (('objective', 'action'), '問い合わせ')):
        b = copy.deepcopy(OUTLINE)
        node = b
        for k in path[:-1]:
            node = node[k]
        node[path[-1]] = value
        out = diff(OUTLINE, b)
        assert out['same_experiment'] is False and out['findings'] == [], (path, out)
        assert out['experiment_differs'] == ['/'.join(path)]


def test_diff_cli_passes_on_a_clean_variant(tmp_path):
    variant = copy.deepcopy(OUTLINE)
    variant['id'], variant['outputs'] = 'lp-municipal-b', ['site/b.html']
    variant['copies']['cpy-0002']['text'] = '運用の手間が 3 分の 1 になります'
    write_repo(tmp_path, outlines=(OUTLINE, variant, DECK))
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 0, out
    assert [c['element'] for c in out['changes']] == ['copies/cpy-0002'] and out['ids']['a']['positioning'] == out['ids']['b']['positioning']
    assert 'dom' in ' '.join(out['not_checked'])  # without rendered reports, text outside anchors is not compared


# ---- external review of 2026-11-17 --------------------------------------------------------------------------------

def test_unapproved_evidence_of_a_shipped_value_stops_shipping(tmp_path):
    """Evidence a shipped message inherits is something the outline stands on: it needs the owner's approval."""
    write_repo(tmp_path, confirmed=False)
    confirm(tmp_path, SSOT, skip=('supports:prf-0001:val-0001',))
    rc, report = check(tmp_path)
    hits = [f for f in report['findings'] if f['code'] == 'IA_POSITIONING_UNAPPROVED' and f['level'] == 'fail']
    assert rc == 2 and hits and 'supports:prf-0001:val-0001' in hits[0]['detail']['atoms'], report['findings']


def test_unapproved_evidence_does_not_count_as_support_for_the_lead(tmp_path):
    write_repo(tmp_path, confirmed=False)
    confirm(tmp_path, SSOT, skip=('prf-0002', 'supports:prf-0002:val-0002'))
    rc, report = check(tmp_path)
    assert 'IA_LEAD_UNSUPPORTED' in codes(report, 'warn'), report['findings']


def test_expired_evidence_of_a_gap_arrow_stops_a_shipped_outline(tmp_path):
    """fulfills arrows decide the gaps a value fills: evidence behind them is evidence the deliverable stands on."""
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'supports' and e['from'] == 'prf-0003'
                                                      and e['to'] == 'str-0001')]
    ssot['proofs']['prf-0003']['as_of'] = '2024-01-01'  # supports only fulfills:alt-0001:pan-0001 now
    write_repo(tmp_path, ssot)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_EXPIRED_SHIPPED' in codes(report), report['findings']


def test_compared_alternatives_are_approval_targets(tmp_path):
    outline, _ = apply([{'op': 'set', 'path': ['messages', 'msg-0001', 'compares'], 'value': ['alt-0002']}])
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, out = cli('targets', '--root', str(tmp_path))
    assert 'ia:positioning:alt-0002' in out['must_approval_keys'], out


def test_unanchored_text_is_compared_regardless_of_output_order():
    a = {'unanchored': [{'viewport': 'mobile', 'file': 'x.html', 'text': '一'}, {'viewport': 'mobile', 'file': 'y.html', 'text': '二'}]}
    b = {'unanchored': [{'viewport': 'mobile', 'file': 'y.html', 'text': '二'}, {'viewport': 'mobile', 'file': 'x.html', 'text': '一'}]}
    assert ia_outline._unanchored(a) == ia_outline._unanchored(b)
