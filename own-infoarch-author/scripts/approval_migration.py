"""One-time replacement of per-element positioning approvals by per-path rows (requirement 17).

SPEC-information-architecture-standardization requirement 17 and ADR-20261118-ia-approval-per-value-path (biz_ops):
a path row is made only when every atom and arrow on the path has an old `ia:positioning:<key>` row whose hash
matches its current content, so the new row never approves more than the owner did. The owner's words are not
stored in the ledger; they are read from the `Owner-Approval:` trailers of the commits that wrote it, and quoted
as the basis of the migration. A path with any member missing (or no owner words) stays unapproved. After the
replacement the old rows are removed: the ledger keeps them in its git history, and reverting the migration
commit restores them.
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import tempfile
from pathlib import Path

import positioning as pos

LEDGER = 'design/approvals.yaml'
TRAILER = re.compile(r'^Owner-Approval:\s*「(?P<words>.+)」\s*\((?P<date>\d{4}-\d{2}-\d{2})\)\s*$')


def owner_words(root):
    """[(words, date)] of every Owner-Approval trailer on the commits that touched the ledger, oldest first."""
    run = subprocess.run(['git', '-C', str(root), 'log', '--reverse', '--format=%B%x00', '--', LEDGER],
                         capture_output=True, text=True)
    if run.returncode:
        return []
    out = []
    for body in run.stdout.split('\0'):
        for line in body.splitlines():
            hit = TRAILER.match(line.strip())
            if hit and (hit['words'], hit['date']) not in out:
                out.append((hit['words'], hit['date']))
    return out


def plan(m, rows, words):
    """-> (new rows, migrated report, unmigrated report). `rows` is the verified ledger list."""
    legacy = {r['key'][len(pos.LEGACY_PREFIX):]: r for r in rows if r['key'].startswith(pos.LEGACY_PREFIX)}
    A, E = pos.atoms(m), pos.edges(m)
    new_rows, migrated, unmigrated = [], [], []
    for vid, p in pos.paths(m).items():
        missing = [k for k in p.members
                   if (legacy.get(k) or {}).get('value_hash') != pos.element_hash(k, A.get(k) or E[k])]
        if missing:
            unmigrated.append({'value': vid, 'reason': 'not-every-member-approved', 'missing': missing})
            continue
        sources = [legacy[k] for k in p.members]
        approvers = sorted({r['approver'] for r in sources})
        dates = sorted({r['date'] for r in sources})
        quoted = [{'words': w, 'date': d} for w, d in words if d in dates]
        if len(approvers) != 1:
            unmigrated.append({'value': vid, 'reason': 'several-approvers', 'approvers': approvers})
            continue
        if not quoted or {q['date'] for q in quoted} != set(dates):
            unmigrated.append({'value': vid, 'reason': 'no-owner-words', 'dates': dates})
            continue
        row = {'key': pos.PATH_PREFIX + vid, 'value_hash': p.hash, 'approver': approvers[0], 'date': dates[-1]}
        new_rows.append(row)
        migrated.append({**row, 'from': [pos.LEGACY_PREFIX + k for k in p.members], 'owner_words': quoted})
    return new_rows, migrated, unmigrated


def migrate(root, m, rows, write=False):
    """Plan, and with `write` replace, the ledger. `rows`: the ledger list already verified by own-design-author."""
    legacy = [r for r in rows if r['key'].startswith(pos.LEGACY_PREFIX)]
    words = owner_words(root)
    new_rows, migrated, unmigrated = plan(m, rows, words)
    clash = sorted(r['key'] for r in new_rows if any(o['key'] == r['key'] for o in rows))
    kept = [r for r in rows if not r['key'].startswith(pos.LEGACY_PREFIX)]
    quoted = sorted({(q['date'], q['words']) for r in migrated for q in r['owner_words']},
                    key=lambda dw: next(i for i, w in enumerate(words) if (w[1], w[0]) == dw))
    result = {'legacy_rows': len(legacy), 'migrated': migrated, 'unmigrated': unmigrated, 'kept_rows': len(kept),
              'commit_trailers': [f'Owner-Approval-Migrated: 「{w}」 ({d})' for d, w in quoted],
              'revert': f'git revert <this commit>（{LEDGER} の旧い行が戻る）', 'written': False}
    if clash:
        result.update(exit_code=2, error=f'経路の行がすでにある {clash}（移行は一度だけ）')
        return result
    if not legacy:
        result.update(exit_code=2, error='原子と矢印ごとの旧い行が無い（移すものが無い）')
        return result
    if not migrated:
        result.update(exit_code=2, error='経路の行を 1 つも作れない（unmigrated の理由を見る）')
        return result
    result['exit_code'] = 0
    if write:
        _write(Path(root) / LEDGER, kept + new_rows)
        result['written'] = True
    return result


def _write(target, rows):
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=target.parent, delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(rows, stream, ensure_ascii=False, indent=2, allow_nan=False)
            stream.write('\n')
            stream.flush()
            os.fsync(stream.fileno())
        temporary.replace(target)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()
