"""Shared fixtures of the positioning layer's tests: the positioning SSOT (schema 5) and the owner's approval rows.

The infoarch tests import this module too (the outlines stand on this positioning), so the fixture is written once.
"""
import copy
import sys
from pathlib import Path

import yaml

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))
import positioning  # noqa: E402
import positioning_model as pm  # noqa: E402

def _a(**kw):
    return {'state': 'proposed', 'provenance': 'owner', **kw}


def _e(kind, frm, to, **kw):
    return {'kind': kind, 'from': frm, 'to': to, 'state': 'proposed', 'provenance': 'owner', **kw}


# Positioning (SPEC-information-architecture-standardization v6, requirement 2): atoms belong to one segment each;
# arrows are listed once, with their kind and direction. seg-0001 has a gap (pan-0001, gan-0001: no alternative fully
# meets them) and a point of parity (gan-0002: every alternative of the member category meets it). seg-0002 reuses
# alt-0001 as a copy (alt-0003, copied_from) instead of sharing it.
EDGES = [
    _e('has_job', 'seg-0001', 'job-0001'),
    _e('need', 'job-0001', 'pan-0001'), _e('need', 'job-0001', 'gan-0001'), _e('need', 'job-0001', 'gan-0002'),
    _e('recalls', 'job-0001', 'cat-0001'), _e('includes', 'cat-0001', 'alt-0001'),
    _e('member_of', 'seg-0001', 'cat-0001'),
    _e('fulfills', 'alt-0001', 'pan-0001', degree='none'), _e('fulfills', 'alt-0002', 'pan-0001', degree='partial'),
    _e('fulfills', 'alt-0001', 'gan-0001', degree='none'), _e('fulfills', 'alt-0002', 'gan-0001', degree='partial'),
    _e('fulfills', 'alt-0001', 'gan-0002', degree='full'), _e('fulfills', 'alt-0002', 'gan-0002', degree='full'),
    _e('lacks', 'str-0001', 'alt-0001'), _e('produces', 'str-0001', 'val-0001'),
    _e('satisfies', 'val-0001', 'pan-0001'), _e('satisfies', 'val-0001', 'gan-0001'),
    _e('satisfies', 'val-0003', 'gan-0002'),
    _e('supports', 'prf-0003', 'str-0001'), _e('supports', 'prf-0003', 'fulfills:alt-0001:pan-0001'),
    _e('supports', 'prf-0001', 'val-0001'), _e('supports', 'prf-0002', 'val-0002'),
    _e('has_job', 'seg-0002', 'job-0002'), _e('need', 'job-0002', 'pan-0002'),
    _e('recalls', 'job-0002', 'cat-0002'), _e('includes', 'cat-0002', 'alt-0003'),
    _e('member_of', 'seg-0002', 'cat-0002'),
    _e('fulfills', 'alt-0003', 'pan-0002', degree='none'),
    _e('lacks', 'str-0002', 'alt-0003'), _e('produces', 'str-0002', 'val-0002'),
    _e('satisfies', 'val-0002', 'pan-0002'),
]

SSOT = {
    'schema_version': 5,
    'ledger': {'next': {'prf': 4,
                        'seg': 3, 'job': 3, 'pan': 3, 'gan': 3, 'alt': 4, 'cat': 3, 'str': 3, 'val': 4},
               'tombstones': {}},
    'segments': {
        'seg-0001': _a(qualifier='自治体の情報部門', situation='既存システムの運用が重い'),
        'seg-0002': _a(qualifier='中小企業の経営者', situation='人手が足りない')},
    'jobs': {
        'job-0001': _a(segment='seg-0001', text='既存システムを無理なく運用する'),
        'job-0002': _a(segment='seg-0002', text='少ない人数で業務を回す')},
    'pains': {
        'pan-0001': _a(segment='seg-0001', text='運用が重い'),
        'pan-0002': _a(segment='seg-0002', text='人手が足りない')},
    'gains': {
        'gan-0001': _a(segment='seg-0001', text='職員が本来の業務に時間を使える'),
        'gan-0002': _a(segment='seg-0001', text='止まらずに動き続ける')},
    'alternatives': {
        'alt-0001': _a(segment='seg-0001', text='今のシステムを使い続ける'),
        'alt-0002': _a(segment='seg-0001', text='保守を外部に委託する'),
        'alt-0003': _a(segment='seg-0002', text='今のシステムを使い続ける', copied_from='alt-0001')},
    'categories': {
        'cat-0001': _a(segment='seg-0001', text='自治体向け Web サーバー'),
        'cat-0002': _a(segment='seg-0002', text='中小企業向けの業務システム')},
    'strengths': {
        'str-0001': _a(segment='seg-0001', text='自動更新で保守が要らない'),
        'str-0002': _a(segment='seg-0002', text='設定なしで使い始められる')},
    'values': {
        'val-0001': _a(segment='seg-0001', kind='difference', text='運用の手間が減る'),
        'val-0002': _a(segment='seg-0002', kind='difference', text='少人数でも回る'),
        'val-0003': _a(segment='seg-0001', kind='parity', text='止まらずに動く')},
    'edges': EDGES,
    # positioning evidence: belongs to a segment and supports atoms and arrows through `supports` edges
    'proofs': {'prf-0001': _a(segment='seg-0001', text='導入自治体の数', source='導入台帳', as_of='2026-10-01'),
               'prf-0002': _a(segment='seg-0002', text='導入前後の運用時間の比較', source='導入事例の計測', as_of='2026-10-01'),
               'prf-0003': _a(segment='seg-0001', text='自動更新の稼働記録', source='運用ログ', as_of='2026-10-01')},
}


def dump(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding='utf-8')


def confirm(root, ssot, skip=()):
    """Record the owner's approval of every proposed positioning path (as the owner would with design-author
    approve). `skip` names value ids left unapproved."""
    m = pm.Model(root=root, ssot=ssot)
    rows = [{'key': key, 'value_hash': digest, 'approver': 'fixture-owner', 'date': '2026-11-16'}
            for key, digest in sorted(positioning.approval_targets(m).items())
            if key[len(positioning.APPROVAL_PREFIX):] not in skip]
    path = root / 'design/approvals.yaml'
    kept = yaml.safe_load(path.read_text(encoding='utf-8')) if path.is_file() else []
    kept = [r for r in (kept or []) if not str(r.get('key', '')).startswith(positioning.APPROVAL_PREFIX)]
    dump(path, kept + rows)  # keep the design rows (ds:*) the owner already recorded


def _key(edge):
    return f'{edge["kind"]}:{edge["from"]}:{edge["to"]}'


def apply(ops, ssot=SSOT):
    """Apply a red case's declarative operations to a copy of the base fixture. -> (ssot, keys left unapproved)."""
    ssot = copy.deepcopy(ssot)
    blocks = {p: b for b, p in positioning.ATOM_BLOCKS.items()} | {'prf': 'proofs'}
    unapproved = []
    for op in ops:
        kind = op['op']
        if kind == 'set_atom':
            ssot[blocks[op['id'].split('-')[0]]][op['id']] = op['atom']
        elif kind == 'set_field':
            ssot[blocks[op['id'].split('-')[0]]][op['id']][op['field']] = op['value']
        elif kind == 'set_next':
            ssot['ledger']['next'][op['prefix']] = op['value']
        elif kind == 'add_edge':
            ssot['edges'].append({'state': 'proposed', 'provenance': 'owner', **op['edge']})
        elif kind == 'remove_edge':
            before = len(ssot['edges'])
            ssot['edges'] = [e for e in ssot['edges'] if _key(e) != op['key']]
            assert len(ssot['edges']) == before - 1, op
        elif kind in ('set_field_edge', 'remove_field_edge'):
            [edge] = [e for e in ssot['edges'] if _key(e) == op['key']]
            if kind == 'set_field_edge':
                edge[op['field']] = op['value']
            else:
                del edge[op['field']]
        elif kind == 'set_block':
            ssot[op['block']] = op['value']
        elif kind == 'remove_atom':
            del ssot[blocks[op['id'].split('-')[0]]][op['id']]
            ssot['edges'] = [e for e in ssot['edges'] if op['id'] not in (e['from'], e['to'])]
        elif kind == 'unapprove':
            unapproved.append(op['key'])
        else:
            raise KeyError(kind)
    return ssot, tuple(unapproved)
