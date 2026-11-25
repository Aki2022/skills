"""script, 台本 (SPEC-information-architecture-standardization, requirements 6-9 and 13; Quality 1 and 3).

A script stands on one positioning (its target's segment), keeps the constants and the objective fixed, and
varies messages, blocks (each with one content function) and copy. Requirement 7 disqualifies it; ids come from content only; two sets differ by
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
from test_check import DECK, DICTIONARY, OUTLINE, SKILL, SSOT, check, cli, codes, confirm, dump, write_repo  # noqa: E402
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)

TABLE = json.loads((SKILL / 'fixtures/script-red-cases.json').read_text(encoding='utf-8'))
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
    base = ({**copy.deepcopy(OUTLINE), 'id': 'lp-base', 'outputs': ['site/base.html']},) if case.get('base') else ()
    write_repo(root, ssot, outlines=(outline, DECK, *base), confirmed=False)
    confirm(root, ssot, skip=unapproved)


# ---- the table (Quality 1) ----------------------------------------------------------------------------------------

def test_every_condition_has_a_red_case_and_every_case_names_a_condition():
    covered = {c['condition'] for c in CASES + DIFF_CASES}
    assert sorted(set(CONDITIONS) - covered) == []  # 対応の無い失格条件 0 件
    assert sorted(covered - set(CONDITIONS)) == []
    ids = [c['id'] for c in CASES + DIFF_CASES]
    assert len(set(ids)) == len(ids)
    assert len([c for c in CONDITIONS.values() if c['requirement'] == '7']) == 22  # spec requirement 7 (2026-11-25, the script), split by code
    assert [c['id'] for c in CONDITIONS.values() if c['requirement'] == '9'] == [
        '9-unassigned-difference', '9-variant-other-experiment', '9-undeclared-difference', '9-override-unchanged']


def test_every_outline_code_is_raised_by_the_checker():
    source = ''.join((SKILL / f'scripts/{n}').read_text(encoding='utf-8') for n in ('outline.py', 'check.py'))
    for c in CONDITIONS.values():
        if c['code'] != 'IA_SCHEMA':
            assert f"'{c['code']}'" in source, c


@pytest.mark.parametrize('case', CASES, ids=[c['id'] for c in CASES])
def test_red_case_fails_with_its_condition_code(case, tmp_path):
    write_case(tmp_path, case)
    rc, report = check(tmp_path)
    if CONDITIONS[case['condition']].get('level') == 'warn':  # a warning condition: reported, never a failure
        assert rc == 1 and not codes(report), report['findings']
        assert CONDITIONS[case['condition']]['code'] in codes(report, 'warn'), report['findings']
        return
    assert rc == 2, report['findings']
    assert CONDITIONS[case['condition']]['code'] in codes(report), codes(report)


@pytest.mark.parametrize('case', DIFF_CASES, ids=[c['id'] for c in DIFF_CASES])
def test_unassigned_difference_disqualifies_the_variant(case, tmp_path):
    variant = copy.deepcopy(OUTLINE)
    variant['id'], variant['outputs'] = 'lp-municipal-b', ['site/b.html']
    write_repo(tmp_path / 'a', outlines=(OUTLINE, DECK))
    dictionary_b = copy.deepcopy(DICTIONARY)
    for fid, value in case.get('fact_b', {}).items():
        next(e for e in dictionary_b['entries'] if e['id'] == fid)['text'] = value
    write_repo(tmp_path / 'b', outlines=(variant, DECK), dictionary=dictionary_b)
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
                        {'op': 'set', 'path': ['blocks', 'blk-0003', 'places'], 'value': []},
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


def test_blocks_of_any_function_and_any_count_pass(tmp_path):
    """Requirement 6: blocks of every function, the same function more than once, copy citing evidence or the dictionary."""
    outline, _ = apply([
        {'op': 'set', 'path': ['blocks', 'blk-0004'], 'value': {'order': 4, 'function': 'identify'}},
        {'op': 'set', 'path': ['blocks', 'blk-0005'], 'value': {'order': 5, 'function': 'fact'}},
        {'op': 'set', 'path': ['blocks', 'blk-0006'], 'value': {'order': 6, 'function': 'evidence'}},
        {'op': 'set', 'path': ['blocks', 'blk-0007'], 'value': {'order': 7, 'function': 'offer'}},
        {'op': 'set', 'path': ['blocks', 'blk-0008'], 'value': {'order': 8, 'function': 'source'}},
        {'op': 'set', 'path': ['blocks', 'blk-0009'], 'value': {'order': 9, 'function': 'evidence'}},
        {'op': 'set', 'path': ['copies', 'cpy-0006'], 'value': {'block': 'blk-0004', 'order': 1, 'text': '{{value-phone}} 自治体向けサーバー'}},
        {'op': 'set', 'path': ['copies', 'cpy-0007'], 'value': {'block': 'blk-0005', 'order': 1, 'text': '導入は{{value-customer-count}}'}},
        {'op': 'set', 'path': ['copies', 'cpy-0008'], 'value': {'block': 'blk-0006', 'order': 1, 'refs': ['prf-0001'], 'text': '導入した自治体の数を台帳で数えています'}},
        {'op': 'set', 'path': ['copies', 'cpy-0009'], 'value': {'block': 'blk-0007', 'order': 1, 'text': '月額で使えます'}},
        {'op': 'set', 'path': ['copies', 'cpy-0010'], 'value': {'block': 'blk-0008', 'order': 1, 'text': '運営：自治体サーバー事業部'}},
        {'op': 'set', 'path': ['copies', 'cpy-0011'], 'value': {'block': 'blk-0009', 'order': 1, 'refs': ['prf-0001', 'prf-0003'], 'text': '稼働記録と導入台帳から'}}])
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']


def test_proposal_is_a_warning_while_producing(tmp_path):
    """Requirement 7 / SPEC-deliverable-production-layers requirement 17: a copy with no reference may stay as a
    proposal; production only warns (the ship gate refuses it, see test_ship)."""
    outline, _ = apply([
        {'op': 'set', 'path': ['blocks', 'blk-0004'], 'value': {'order': 4, 'function': 'fact'}},
        {'op': 'set', 'path': ['copies', 'cpy-0006'], 'value': {'block': 'blk-0004', 'order': 1, 'proposal': True,
                                                               'text': '自治体の 8 割が更新に悩んでいます'}}])
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, report = check(tmp_path)
    assert rc == 1 and 'IA_SCRIPT_PROPOSAL' in codes(report, 'warn') and not codes(report), report['findings']
    assert ia_outline.proposals(outline) == ['cpy-0006'] and ia_outline.proposals(OUTLINE) == []


def test_proposal_is_part_of_the_script_id():
    outline, _ = apply([{'op': 'delete', 'path': ['copies', 'cpy-0003', 'message']},
                        {'op': 'set', 'path': ['copies', 'cpy-0003', 'proposal'], 'value': True}])
    cited, _ = apply([{'op': 'set', 'path': ['copies', 'cpy-0004', 'refs'], 'value': ['prf-0001']}])
    base = ia_outline.outline_id(OUTLINE)
    assert len({base, ia_outline.outline_id(outline), ia_outline.outline_id(cited)}) == 3


# ---- ids (requirement 8, Quality 3) -----------------------------------------------------------------------------------

def model(ssot=SSOT, outlines=(OUTLINE, DECK), skip=()):
    m = ia_check.Model(root=Path('.'), ssot=ssot)
    for o in outlines:
        m.outlines[o['id']] = (Path(layer_paths.rel('infoarch', 'scripts', f'{o["id"]}.yaml')), o)
    rows = {k: v for k, v in positioning.approval_targets(m).items() if k[len(positioning.APPROVAL_PREFIX):] not in skip}
    return m, positioning.approved_keys(m, rows)


def ids_of(oid='lp-municipal', **kw):
    m, approved = model(**kw)
    return ia_outline.set_ids(m, oid, approved)


def test_ids_are_reproducible_and_shaped():
    a, b = ids_of(), ids_of()
    assert a == b
    assert a['outline'].startswith('spt-') and a['positioning'].startswith('pos-')
    assert sorted(a['paths']) == ['val-0001', 'val-0003'] and all(p.startswith('pth-') for p in a['paths'].values())
    assert a['set'] == a['outline'] + ''.join('+' + a['paths'][v] for v in sorted(a['paths']))  # no positioning id


def test_block_order_moves_script_and_set_not_positioning():
    swapped = copy.deepcopy(OUTLINE)
    swapped['blocks']['blk-0001']['order'], swapped['blocks']['blk-0002']['order'] = 2, 1
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
    for key in ('messages', 'blocks', 'copies'):
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


def test_set_id_hashes_content_whether_or_not_the_paths_are_approved():
    out = ids_of(skip=('val-0001', 'val-0003'))
    assert out['positioning'] is None and out['set'] == ids_of()['set']  # approval is the ship gate's, not the id's


def test_set_cli_reports_ids_and_elements(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('set', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert rc == 0, out
    assert out['set'] == out['outline'] + ''.join('+' + out['paths'][v] for v in sorted(out['paths']))
    assert sorted(out['elements']['copies']) == ['cpy-0001', 'cpy-0002', 'cpy-0003', 'cpy-0004', 'cpy-0005']
    assert all(len(h) == 64 for h in out['elements']['messages'].values())


# ---- diff (requirement 9) ---------------------------------------------------------------------------------------------

def diff(a, b, ssot_b=SSOT, skip_b=()):
    ma, appr_a = model(outlines=(a,))
    mb, appr_b = model(ssot=ssot_b, outlines=(b,), skip=skip_b)
    return ia_outline.diff(ma, a['id'], appr_a, mb, b['id'], appr_b)


def variant(b, *overrides, base='lp-municipal'):
    """Declare `b` a variant of `base` that overrides `overrides` (requirement 6, last bullet)."""
    b['variant_of'], b['overrides'] = base, list(overrides)
    return b


def test_diff_lists_only_the_changed_elements():
    b = copy.deepcopy(OUTLINE)
    b['blocks']['blk-0001']['order'], b['blocks']['blk-0002']['order'] = 2, 1
    b['copies']['cpy-0002']['text'] = '運用の手間が 3 分の 1 になります'
    out = diff(OUTLINE, b)
    assert out['changes'] == [
        {'layer': 'outline', 'element': 'blocks/blk-0001', 'change': 'changed', 'attributes': ['order']},
        {'layer': 'outline', 'element': 'blocks/blk-0002', 'change': 'changed', 'attributes': ['order']},
        {'layer': 'outline', 'element': 'copies/cpy-0002', 'change': 'changed', 'attributes': ['text']}]
    assert out['changed_layers'] == ['outline'] and out['same_experiment'] is True
    assert [f['code'] for f in out['findings']] == ['IA_SCRIPT_UNDECLARED_DIFFERENCE'] and out['confounded'] is True
    declared = diff(OUTLINE, variant(b, 'blocks/blk-0001', 'blocks/blk-0002', 'copies/cpy-0002'))
    assert declared['findings'] == [] and declared['confounded'] is False and declared['exit_code'] == 0


def test_diff_of_identical_sets_is_empty():
    out = diff(OUTLINE, copy.deepcopy(OUTLINE))
    assert out['changes'] == [] and out['changed_layers'] == [] and out['ids']['a'] == out['ids']['b']


def test_diff_names_added_and_removed_elements_and_the_lead_change():
    b = copy.deepcopy(OUTLINE)
    b['messages']['msg-0001']['rank'], b['messages']['msg-0002']['rank'] = 'support', 'lead'
    del b['copies']['cpy-0004']
    b['copies']['cpy-0009'] = {'block': 'blk-0003', 'order': 1, 'text': '保守の要らない運用を'}
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
    assert {'layer': 'positioning', 'element': 'fulfills:alt-0002:pan-0001', 'change': 'changed', 'attributes': ['degree'],
            'path': 'val-0001'} in out['changes']


def test_different_constants_or_objective_is_another_experiment():
    for path, value in ((('tone',), 'くだけた調子'), (('touchpoint',), 'ad'), (('target', 'condition'), '別の条件'),
                        (('objective', 'action'), '問い合わせ')):
        b = copy.deepcopy(OUTLINE)
        node = b
        for k in path[:-1]:
            node = node[k]
        node[path[-1]] = value
        out = diff(OUTLINE, b)
        assert out['same_experiment'] is False and out['findings'] == [], (path, out)  # another experiment, not confounding
        assert out['experiment_differs'] == ['/'.join(path)]


def test_diff_cli_passes_on_a_clean_variant(tmp_path):
    b = copy.deepcopy(OUTLINE)
    b['id'], b['outputs'] = 'lp-municipal-b', ['site/b.html']
    b['copies']['cpy-0002']['text'] = '運用の手間が 3 分の 1 になります'
    write_repo(tmp_path, outlines=(OUTLINE, variant(b, 'copies/cpy-0002'), DECK))
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 0, out
    assert out['confounded'] is False and out['differs_in'] == 'script'
    assert any(n.startswith(ia_outline.PLAN_NOT_CHECKED) for n in out['not_checked']) and out['plan'] is None  # no plans
    assert [c['element'] for c in out['changes']] == ['copies/cpy-0002'] and out['ids']['a']['positioning'] == out['ids']['b']['positioning']
    assert 'dom' in ' '.join(out['not_checked'])  # without rendered reports, text outside anchors is not compared


def test_diff_cli_reports_an_undeclared_difference_as_confounded(tmp_path):
    """Requirement 9 / production layers 18: a difference not declared as an override warns, and the comparison is
    labeled 「交絡あり」 with the undeclared keys; a declared key that does not differ also warns."""
    b = copy.deepcopy(OUTLINE)
    b['id'], b['outputs'] = 'lp-municipal-b', ['site/b.html']
    b['copies']['cpy-0002']['text'] = '運用の手間が 3 分の 1 になります'
    b['blocks']['blk-0001']['order'], b['blocks']['blk-0002']['order'] = 2, 1
    write_repo(tmp_path, outlines=(OUTLINE, variant(b, 'copies/cpy-0002', 'copies/cpy-0001'), DECK))
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 1, out
    assert out['confounded'] is True and out['confounded_label'] == '交絡あり'
    assert out['undeclared'] == ['blocks/blk-0001', 'blocks/blk-0002'] and out['unchanged_overrides'] == ['copies/cpy-0001']
    assert {f['code'] for f in out['findings']} == {'IA_SCRIPT_UNDECLARED_DIFFERENCE', 'IA_SCRIPT_OVERRIDE_UNCHANGED'}
    assert all(f['level'] == 'warn' for f in out['findings']) and out['differs_in'] == 'script'
    rc, report = check(tmp_path)  # the same comparison runs inside check, for each variant against its base
    assert rc == 1 and {'IA_SCRIPT_UNDECLARED_DIFFERENCE', 'IA_SCRIPT_OVERRIDE_UNCHANGED'} <= set(codes(report, 'warn'))
    assert not codes(report), report['findings']


def test_declaring_a_variant_moves_neither_the_script_id_nor_the_set_id():
    """Quality 3: `variant_of` and `overrides` say how a file relates to another, not what it says (like channel and
    outputs): the script id and the set id come from the content only."""
    declared = variant(copy.deepcopy(OUTLINE), 'copies/cpy-0002')
    assert ia_outline.outline_id(declared) == ia_outline.outline_id(OUTLINE)
    declared['id'] = 'lp-municipal-v'
    a, b = ids_of(), ids_of(outlines=(declared, DECK), oid='lp-municipal-v')
    assert (a['outline'], a['set']) == (b['outline'], b['set'])


# ---- external review of 2026-11-17 --------------------------------------------------------------------------------

def test_changed_evidence_of_a_shipped_value_stops_shipping(tmp_path):
    """Evidence a shipped message inherits is on its path: changing it withdraws the path's approval."""
    write_repo(tmp_path)
    changed = copy.deepcopy(SSOT)
    changed['proofs']['prf-0001']['text'] = '導入自治体の数（改）'
    dump(layer_paths.path(tmp_path, 'positioning_file'), changed)
    rc, report = check(tmp_path)
    hits = [f for f in report['findings'] if f['code'] == 'IA_POSITIONING_UNAPPROVED' and f['level'] == 'fail']
    assert rc == 2 and hits and hits[0]['detail']['values'] == ['val-0001'], report['findings']


def test_lead_without_evidence_on_its_path_warns(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'supports' and e['from'] == 'prf-0002')]
    write_repo(tmp_path, ssot)
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


def test_compared_alternatives_are_approved_with_the_path(tmp_path):
    outline, _ = apply([{'op': 'set', 'path': ['messages', 'msg-0001', 'compares'], 'value': ['alt-0002']}])
    write_repo(tmp_path, outlines=(outline, DECK))
    rc, out = cli('targets', '--root', str(tmp_path))
    assert out['must_approval_keys'] == ['ia:path:val-0001', 'ia:path:val-0003'], out
    m = ia_check.Model(root=Path('.'), ssot=SSOT)
    assert 'alt-0002' in positioning.paths(m)['val-0001'].roles  # every alternative of the segment is on a gap path


def test_unanchored_text_is_compared_regardless_of_output_order():
    a = {'unanchored': [{'viewport': 'mobile', 'file': 'x.html', 'text': '一'}, {'viewport': 'mobile', 'file': 'y.html', 'text': '二'}]}
    b = {'unanchored': [{'viewport': 'mobile', 'file': 'y.html', 'text': '二'}, {'viewport': 'mobile', 'file': 'x.html', 'text': '一'}]}
    assert ia_outline._unanchored(a) == ia_outline._unanchored(b)


# ---- the plan side of the comparison (production layers requirement 18) ----------------------------------------------

PLAN_A = {'schema_version': 1, 'deliverable': 'lp-municipal', 'script': 'lp-municipal',
          'blocks': {'blk-0001': {'pattern': 'split-2'}, 'blk-0002': {'pattern': 'stats-row'},
                     'blk-0003': {'pattern': 'band'}},
          'comp': {'reviewed': True}}


def write_plan(root, plan):
    dump(layer_paths.path(root, 'build', plan['deliverable'], 'plan.yaml'), plan)


def plan_b(*overrides, change=None):
    b = copy.deepcopy(PLAN_A)
    b.update(deliverable='lp-municipal-b', script='lp-municipal-b', variant_of='lp-municipal', overrides=list(overrides))
    if change:
        change(b)
    return b


def _cover_first(b):
    b['blocks']['blk-0001'] = {'pattern': 'cover', 'details': '絵の上に見出しを重ねる'}


def script_b(*overrides, copy_changed=True):
    b = copy.deepcopy(OUTLINE)
    b['id'], b['outputs'] = 'lp-municipal-b', ['site/b.html']
    if copy_changed:
        b['copies']['cpy-0002']['text'] = '運用の手間が 3 分の 1 になります'
    return variant(b, *overrides)


def test_diff_compares_the_plans_when_each_deliverable_has_one(tmp_path):
    write_repo(tmp_path, outlines=(OUTLINE, script_b('copies/cpy-0002'), DECK))
    write_plan(tmp_path, PLAN_A)
    write_plan(tmp_path, plan_b('blocks/blk-0001', change=_cover_first))
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 0, out
    assert out['differs_in'] == 'both' and 'plan' in out['changed_layers']
    assert not any(n.startswith(ia_outline.PLAN_NOT_CHECKED) for n in out['not_checked'])
    assert {'layer': 'plan', 'element': 'blocks/blk-0001', 'change': 'changed', 'attributes': ['details', 'pattern']} in out['changes']
    assert out['plan']['a'] == 'lp-municipal' and out['plan']['b'] == 'lp-municipal-b' and out['plan']['undeclared'] == []


def test_diff_says_plan_when_only_the_plan_differs(tmp_path):
    write_repo(tmp_path, outlines=(OUTLINE, script_b(copy_changed=False), DECK))
    write_plan(tmp_path, PLAN_A)
    write_plan(tmp_path, plan_b('blocks/blk-0001', change=_cover_first))
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 0, out
    assert out['differs_in'] == 'plan' and out['confounded'] is False


def test_diff_says_script_when_the_plans_are_the_same(tmp_path):
    write_repo(tmp_path, outlines=(OUTLINE, script_b('copies/cpy-0002'), DECK))
    write_plan(tmp_path, PLAN_A)
    write_plan(tmp_path, plan_b())
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 0, out
    assert out['differs_in'] == 'script' and out['plan']['changed'] == []


def test_an_undeclared_plan_difference_is_confounded(tmp_path):
    write_repo(tmp_path, outlines=(OUTLINE, script_b('copies/cpy-0002'), DECK))
    write_plan(tmp_path, PLAN_A)
    write_plan(tmp_path, plan_b(change=_cover_first))
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 1, out
    assert out['confounded'] is True and out['confounded_label'] == '交絡あり'
    assert out['plan']['undeclared'] == ['blocks/blk-0001']
    assert [f['code'] for f in out['findings']] == ['PLAN_UNDECLARED_DIFFERENCE']


def test_without_a_plan_on_one_side_the_plan_is_reported_not_checked(tmp_path):
    write_repo(tmp_path, outlines=(OUTLINE, script_b('copies/cpy-0002'), DECK))
    write_plan(tmp_path, PLAN_A)
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert rc == 0, out
    assert out['differs_in'] == 'script' and out['plan'] is None
    assert any(n.startswith(ia_outline.PLAN_NOT_CHECKED) for n in out['not_checked'])


def test_two_plans_for_one_script_are_not_guessed_between(tmp_path):
    write_repo(tmp_path, outlines=(OUTLINE, script_b('copies/cpy-0002'), DECK))
    write_plan(tmp_path, PLAN_A)
    other = copy.deepcopy(PLAN_A)
    other['deliverable'] = 'lp-municipal-print'
    write_plan(tmp_path, other)
    write_plan(tmp_path, plan_b())
    rc, out = cli('diff', '--root', str(tmp_path), '--a', 'lp-municipal', '--b', 'lp-municipal-b')
    assert out['plan'] is None and any(n.startswith(ia_outline.PLAN_NOT_CHECKED) for n in out['not_checked'])
