"""The release approval: the owner's last gate (SPEC-deliverable-production-layers requirements 17 and 19).

One ledger row per shipped script, key `ia:release:<script id>`. Its hash binds three things by the bytes of their
files: the script, the layout plan (`ux/build/<script id>/plan.yaml`, found by its `script` field) and every file the
script lists in `outputs`. One character changed in any of them after the approval, and the hash no longer matches: the
approval is off. The comp images are not in git and not part of it (requirement 22), so changing one changes nothing.

A script that cannot be bound cannot be approved: a proposal mark left in a copy (requirement 17: the check builds no
approval request), no plan or more than one, an output that is not a file. `targets` then carries no row to approve and
the ship gate stops (IA_RELEASE_UNAPPROVED).

States: `approved` · `unapproved` (bindable, no row) · `changed` (a row exists, its hash is not the current one) ·
`blocked` (cannot be bound; `problems` says why) · `not-shipped` (a draft channel: nothing to approve).
"""
from __future__ import annotations

import hashlib
from pathlib import Path

import layer_paths  # on sys.path through common
import outline as ol
import positioning as pos

PREFIX = pos.RELEASE_PREFIX


def key(oid):
    return PREFIX + oid


def _bytes_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _plans(root, oid):
    """[plan.yaml path] of every readable plan under the build entry whose `script` is `oid` (build's own reader)."""
    plan_check = ol._plan_module()
    return [layer_paths.path(root, 'build', deliverable, 'plan.yaml') for deliverable in sorted(plan_check.find_plans(root, oid))]


def digest(root, oid, script_rel, o):
    """-> (hash or None, problems). None when the script cannot be bound (see the module doc)."""
    root = Path(root)
    problems = []
    left = ol.proposals(o)
    if left:
        problems.append(f'提案の印が残る文案 {left} がある（承認依頼を作らない。証拠を足すか文を消す）')
    script_path = root / script_rel
    files = {'script': {}, 'plan': {}, 'outputs': {}}
    if script_path.is_file() and not script_path.is_symlink():
        files['script'][str(script_rel)] = _bytes_hash(script_path)
    else:
        problems.append(f'台本 {script_rel} が通常ファイルでない')
    try:
        plans = _plans(root, oid)
    except RuntimeError as exc:
        plans, problems = [], problems + [f'組み方の計画を読めない（{exc}）']
    if len(plans) != 1:
        problems.append(f'台本ごとに組み方の計画がちょうど 1 つ要る（{oid} に {len(plans)} 個）')
    for plan in plans[:1]:
        files['plan'][plan.relative_to(root).as_posix()] = _bytes_hash(plan)
    outputs = o.get('outputs') or []
    if not outputs:
        problems.append('台本に成果物（outputs）が無い')
    for out in outputs:
        path = root / out
        if path.is_file() and not path.is_symlink():
            files['outputs'][out] = _bytes_hash(path)
        else:
            problems.append(f'成果物 {out} が通常ファイルとして無い')
    return (None if problems else pos.sha256(files)), problems


def state(root, oid, script_rel, o, rows, shipped):
    """-> {'script', 'key', 'shipped', 'hash', 'state', 'problems'} for one script against the ledger's `rows`."""
    out = {'script': oid, 'key': key(oid), 'shipped': shipped, 'hash': None, 'state': 'not-shipped', 'problems': []}
    if not shipped:
        return out
    out['hash'], out['problems'] = digest(root, oid, script_rel, o)
    if out['hash'] is None:
        out['state'] = 'blocked'
    elif rows.get(out['key']) == out['hash']:
        out['state'] = 'approved'
    else:
        out['state'] = 'changed' if out['key'] in rows else 'unapproved'
    return out


def states(m, rows):
    """{script id: state} of every script of the model."""
    result = {}
    for oid, (rel, o) in sorted(m.outlines.items()):
        ch = m.channels.get(o['channel'])
        result[oid] = state(m.root, oid, rel, o, rows, bool(ch and ch[1]['ship']))
    return result


def findings(m, rows):
    """[(code, message, path)]: IA_RELEASE_UNAPPROVED for every shipped script whose release approval is not current."""
    out = []
    for oid, s in states(m, rows).items():
        if s['state'] in ('approved', 'not-shipped'):
            continue
        why = {'unapproved': 'オーナーの仕上がりの承認が無い', 'changed': '承認の後に台本・計画・成果物のどれかが変わった（承認が外れた）',
               'blocked': '承認の対象を作れない：' + '；'.join(s['problems'])}[s['state']]
        out.append(('IA_RELEASE_UNAPPROVED', f'{oid}: {why}（鍵 {s["key"]}）', str(m.outlines[oid][0])))
    return out


def targets(m, rows):
    """{approval key: hash} for `targets`: the rows the owner can approve now. A key also stays when the ledger already
    holds a row for it (the design layer refuses a ledger row that names no target); then a script that can no longer be
    bound gets a hash no file has, so the row reads as not matching and never as approved."""
    out = {}
    for oid, s in states(m, rows).items():
        if not s['shipped']:
            continue
        if s['hash'] is not None:
            out[s['key']] = s['hash']
        elif s['key'] in rows:
            out[s['key']] = pos.sha256({'not-bindable': oid, 'problems': s['problems']})
    return out
