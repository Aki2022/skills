"""IA ship gate and maintenance (ISSUE-05): the shipped SHA, never the working tree, is what gets checked.

- gate: on a clean commit, run the deterministic `check`, `dom` on every shipped HTML output, and `judge`
  (read-only against the shared judgment ledger). A shipped outline's uncovered judgment stops the gate;
  a score at or above theta goes to a human (warning), because the emphasis gate is still undecided.
- channel: `ship_command` is the formal entry. It runs only after a green gate on the same commit.
- pre-push: the same gate on the predicted merge of each pushed commit into the channel's `merge_base`.
- doctor: a configuration diagnostic. A shipping channel without ship_command or merge_base, a missing or
  invalid marker, and a missing scanner-preserving dispatcher or IA extension are failures.
- maintain: the IA trigger of own-session-close Step 2.5. One line; skips when design/info_arch/ is absent.

Git plumbing is reused from own-design-author's ship gate (hook-free environment, clean-commit binding).
Nothing here calls a model or installs a dependency.
"""
from __future__ import annotations

from contextlib import contextmanager
import datetime as dt
import json
import os
from pathlib import Path
import shlex
import subprocess
import tempfile

import check
import dom
import judge
from common import INFO_ARCH_DIR, SKILL, YamlError, load_ref, load_yaml

EXTENSION = 'pre-push.d/61-infoarch-ship'
MARKER = Path('.vibe-guard/infoarch-ship.enabled')
HTML_SUFFIXES = ('.html', '.htm')


def _design_ship():
    judge._jl()  # puts own-design-author/scripts on sys.path (the judgment ledger is shared there too)
    import ship
    return ship


def _git(root, *args, input=None):
    return _design_ship().git(root, *args, input=input)


def _clean_commit(root):
    return _design_ship().clean_commit(root)


def _stage(name, report, **extra):
    data = report.to_dict()
    return {'stage': name, 'exit_code': data['exit_code'], 'counts': data['counts'], 'targets': data['targets'],
            'findings': data['findings'], **extra}


def _messages(stage, level):
    return [f['code'] + ': ' + f['message'] for f in stage['findings'] if f['level'] == level]


def shipped_outlines(m):
    return sorted(oid for oid, (_, o) in m.outlines.items()
                  if (ch := m.channels.get(o['channel'])) and ch[1]['ship'])


def _outline_of(judgment_id):
    ident = judgment_id.split(':', 1)[1]
    return ident.split('/', 1)[0].split('@', 1)[0]


def gate(root, date, *, source=None, predicted=False):
    """-> (result, model). Check, render and judge one committed tree; never deploys."""
    root = Path(root).resolve()
    commit, tree = _clean_commit(root)
    result = {'exit_code': 0, 'commit': commit, 'tree': tree, 'source_commit': source or commit,
              'predicted_merge': predicted, 'evaluated_on': date, 'stages': [], 'errors': [], 'warnings': [],
              'shipped_outlines': [], 'deployment': 'not-requested'}

    def finish(error=None):
        if error:
            result['errors'].append(error)
        for stage in result['stages']:
            result['errors'] += _messages(stage, 'fail')
            result['warnings'] += _messages(stage, 'warn')
        result['exit_code'] = 2 if result['errors'] else max([s['exit_code'] for s in result['stages']], default=0)
        if _clean_commit(root) != (commit, tree):
            raise ValueError('source changed during the IA ship gate')
        return result, m

    report, m = check.check(root, date)
    result['stages'].append(_stage('check', report))
    if report.exit_code == 2:
        return finish('IA check failed')
    if not any(row['ship'] for _, row in m.channels.values()):
        return finish('no ship:true IA channel declared; nothing ships through the IA gate')
    shipped = shipped_outlines(m)
    result['shipped_outlines'] = shipped
    if not shipped:
        return finish('zero shipped IA outlines (an empty gate is not a pass)')

    with tempfile.TemporaryDirectory(prefix='infoarch-ship-') as temporary:
        dom_paths, rendered, skipped, findings, screens = [], [], [], [], 0
        worst = 0
        for oid in shipped:
            outputs = m.outlines[oid][1]['outputs']
            if not any(Path(o).suffix.lower() in HTML_SUFFIXES for o in outputs):
                skipped.append({'outline': oid, 'reason': 'no HTML output; the rendered DOM check covers HTML only'})
                continue
            rep = dom.run(root, oid, date)
            data = rep.to_dict()
            data['first_screens'] = getattr(rep, 'first_screens', [])
            path = Path(temporary) / f'dom-{oid}.json'
            path.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
            dom_paths.append(str(path))
            rendered.append(oid)
            findings += data['findings']
            screens += len(data['first_screens'])
            worst = max(worst, data['exit_code'])
        counts = {k: sum(f['level'] == k for f in findings) for k in ('fail', 'warn', 'info')}
        result['stages'].append({'stage': 'dom', 'exit_code': worst, 'counts': counts, 'targets': {'outlines': len(rendered)},
                                 'rendered': rendered, 'skipped': skipped, 'first_screens': screens, 'findings': findings})
        if worst == 2:
            return finish('IA rendered DOM check failed')
        report, _, _ = judge.run(root, date, dom_paths)

    blocking = [u for u in report.uncovered if _outline_of(u['id']) in shipped]
    if blocking:
        report.add('fail', 'IA_SHIP_UNCOVERED',
                   f'出荷される成果物の判定 {len(blocking)} 件が台帳に無い（Jev で取得して記録してから出荷する）',
                   ids=sorted(u['id'] for u in blocking))
    result['stages'].append(_stage('judge', report, uncovered=report.uncovered, screened=len(report.screened)))
    return finish('IA judgment gate failed' if report.exit_code == 2 else None)


def ship_channel(root, name, date):
    root = Path(root).resolve()
    _, m = check.check(root, date)
    row = (m.channels.get(name) or (None, {}))[1]
    if not row.get('ship'):
        raise ValueError('unknown or nonshipping IA channel: ' + name)
    result, m = gate(root, date)
    if result['exit_code'] == 2:
        return result
    row = m.channels[name][1]
    env = _design_ship().hook_free_env()
    env.update(INFOARCH_SHIP_COMMIT=result['commit'], INFOARCH_SHIP_TREE=result['tree'], INFOARCH_SHIP_CHANNEL=name)
    run = subprocess.run(shlex.split(row['ship_command']), cwd=root, env=env)
    if run.returncode:
        raise ValueError('configured ship_command failed (exit ' + str(run.returncode) + ')')
    if _clean_commit(root) != (result['commit'], result['tree']):
        raise ValueError('ship_command changed the checked source')
    result['deployment'] = 'command-completed'
    result['channel'] = name
    return result


# ---- predicted merge ------------------------------------------------------------------------

def _channels_at(root, commit):
    """Shipping channels as committed in `commit` (the pushed tree decides its own merge base)."""
    names = _git(root, 'ls-tree', '--name-only', commit, (INFO_ARCH_DIR / 'channels').as_posix() + '/').splitlines()
    rows = {}
    for path in names:
        if not path.endswith('.yaml'):
            continue
        try:
            data = load_yaml_text(_git_raw(root, 'show', commit + ':' + path))
        except YamlError as exc:
            raise ValueError(f'{path}: unreadable channel in the pushed commit') from exc
        if isinstance(data, dict) and data.get('ship') is True:
            rows[Path(path).stem] = data
    return rows


def _git_raw(root, *args):
    run = subprocess.run(['git', '-C', str(root), *args], env=_design_ship().hook_free_env(), capture_output=True)
    if run.returncode:
        raise ValueError('git ' + args[0] + ' failed')
    return run.stdout.decode('utf-8')


def load_yaml_text(text):
    from common import safe_load_text
    return safe_load_text(text)


@contextmanager
def predicted_checkout(root, base_ref, source):
    base = _git(root, 'rev-parse', '--verify', base_ref + '^{commit}')
    # A conflict or a missing merge base makes git fail, and that stops the push.
    tree = _git(root, 'merge-tree', '--write-tree', base, source).splitlines()[0]
    # The synthetic commit is never referenced or pushed; it takes the pushing user's configured identity,
    # and a checkout with no identity fails here, which stops the push.
    synthetic = _git(root, '-c', 'user.name=IA Gate', 'commit-tree', tree, '-p', source, '-p', base,
                     input='IA predicted merge validation\n')
    with tempfile.TemporaryDirectory(prefix='infoarch-ship-') as temporary:
        snapshot = Path(temporary) / 'checkout'
        _git(root, 'worktree', 'add', '--detach', str(snapshot), synthetic)
        try:
            yield snapshot
        finally:
            _git(root, 'worktree', 'remove', '--force', str(snapshot))


def pre_push(root, refs_file, date):
    root = Path(root).resolve()
    if refs_file is None:
        raise ValueError('pre-push requires the preserved refs file')
    records = Path(refs_file).read_text().splitlines()
    if not records:
        raise ValueError('zero push refs')
    sources = []
    for row in records:
        parts = row.split()
        if len(parts) != 4:
            raise ValueError('invalid push ref record')
        sha = parts[1]
        if set(sha) == {'0'}:
            continue
        if _git(root, 'rev-parse', '--verify', sha + '^{commit}') != sha:
            raise ValueError('push source must be a commit')
        if sha not in sources:
            sources.append(sha)
    # The snapshot has no node_modules: rendering resolves Playwright from the pushing checkout unless told otherwise.
    os.environ.setdefault('INFOARCH_PLAYWRIGHT_DIR', str(root))
    checks = []
    for source in sources:
        channels = _channels_at(root, source)
        if not channels:
            raise ValueError('no ship:true IA channel in the pushed commit; the enabled IA gate cannot pass it')
        missing = sorted(name for name, row in channels.items() if not isinstance(row.get('merge_base'), str) or not row['merge_base'])
        if missing:
            raise ValueError('ship:true IA channel without merge_base (no predicted merge): ' + ', '.join(missing))
        for base_ref in sorted({row['merge_base'] for row in channels.values()}):
            with predicted_checkout(root, base_ref, source) as snapshot:
                predicted = _channels_at(snapshot, 'HEAD')
                if set(predicted) != set(channels) or any(predicted[n].get('merge_base') != channels[n].get('merge_base') for n in channels):
                    raise ValueError('predicted tree changed IA ship coverage or merge base')
                result, _ = gate(snapshot, date, source=source, predicted=True)
                result['merge_base'] = base_ref
                checks.append(result)
                if result['exit_code'] == 2:
                    return {'exit_code': 2, 'errors': ['predicted merge IA ship gate failed'], 'checks': checks}
    # The pre-push dispatcher stops on any nonzero exit: warnings travel in the body, only failures exit 2.
    return {'exit_code': 0, 'checks': checks, 'warnings': [w for c in checks for w in c['warnings']],
            'skipped': [] if sources else ['deleted refs only; no shipment']}


# ---- doctor ------------------------------------------------------------------------------------

def doctor(root):
    root = Path(root).resolve()
    errors, shipping = [], []
    base = root / INFO_ARCH_DIR / 'channels'
    if not base.is_dir():
        errors.append('design/info_arch/channels/ absent; no IA ship channel to protect')
    for path in sorted(base.glob('*.yaml')) if base.is_dir() else []:
        try:
            row = load_yaml(path)
        except YamlError:
            errors.append(f'{path.stem}: unreadable channel')
            continue
        if not isinstance(row, dict) or row.get('ship') is not True:
            continue
        shipping.append(path.stem)
        if not isinstance(row.get('ship_command'), str) or not row['ship_command'].strip():
            errors.append(f'{path.stem}: ship:true channel has no ship_command (no formal shipping entry)')
        if not isinstance(row.get('merge_base'), str) or not row['merge_base'].strip():
            errors.append(f'{path.stem}: ship:true channel has no merge_base (pre-push cannot build the predicted merge)')
    if base.is_dir() and not shipping:
        errors.append('no ship:true IA channel declared')
    marker = root / MARKER
    if marker.is_symlink() or not marker.is_file() or marker.read_text().strip() != '1':
        errors.append('IA ship marker missing or invalid (' + MARKER.as_posix() + ')')
    try:
        hook = Path(_git(root, 'rev-parse', '--git-path', 'hooks/pre-push'))
        if not hook.is_absolute():
            hook = root / hook
        extension = hook.parent / EXTENSION
        if not hook.is_file() or not os.access(hook, os.X_OK):
            raise ValueError('pre-push hook missing')
        text = hook.read_text()
        if 'pre-push.d' not in text or '"$ext"' not in text or 'scan-push' not in text:
            raise ValueError('scanner-preserving pre-push dispatcher missing')
        if (not extension.is_file() or not os.access(extension, os.X_OK)
                or extension.read_bytes() != (SKILL / 'hooks' / EXTENSION).read_bytes()):
            raise ValueError('current IA pre-push extension missing (' + EXTENSION + ')')
    except (OSError, ValueError) as exc:
        errors.append(str(exc))
    return {'exit_code': 2 if errors else 0, 'errors': errors, 'channels': shipping,
            'production_protected': False, 'scope': 'configuration-diagnostic'}


# ---- maintenance (own-session-close Step 2.5) ------------------------------------------------

def _dirty(root):
    rows = _git_raw(root, 'status', '--porcelain=v1', '-z', '--untracked-files=all').split('\0')
    paths, i = [], 0
    while i < len(rows):
        row = rows[i]
        i += 1
        if not row:
            continue
        paths.append(row[3:])
        if row[:1] in ('R', 'C') or row[1:2] in ('R', 'C'):
            paths.append(rows[i])
            i += 1
    return paths


def _atoms(ssot):
    if not isinstance(ssot, dict):
        return {}
    return {aid: atom for block in list(check.BLOCKS) + ['entities', 'voices']
            for aid, atom in (ssot.get(block) or {}).items()}


def maintain(root, date, within_days=None):
    root = Path(root).resolve()
    result = {'exit_code': 0, 'status': 'skipped', 'errors': [], 'warnings': [], 'due': [], 'changed_paths': [],
              'impacted_outlines': [], 'uncovered_judgments': 0}
    if not (root / INFO_ARCH_DIR).is_dir():
        return {'exit_code': 0, 'status': 'skipped', 'reason': 'design/info_arch/ absent'}
    today = dt.date.fromisoformat(date)
    expiry = load_ref('expiry.json')
    within = expiry['due_within_days'] if within_days is None else within_days
    if type(within) is not int or within < 0:
        raise ValueError('due window must be a nonnegative integer')
    report, m = check.check(root, date)
    result.update(status='checked', errors=[f.code + ': ' + f.message for f in report.findings if f.level == 'fail'],
                  warnings=[f.code + ': ' + f.message for f in report.findings if f.level == 'warn'])
    if m.ssot is None:
        result['exit_code'] = 2
        return result
    shipped = set(shipped_outlines(m))
    refs = {oid: {x for _, x in check.outline_refs(o)} for oid, (_, o) in m.outlines.items()}
    days = {**expiry['expiry_days'], **((m.ssot.get('policy') or {}).get('expiry_days') or {})}
    for block, kind in (('proofs', 'proof'), ('facts', 'fact')):
        for aid, atom in sorted(m.ssot[block].items()):
            exp = check._expiry(atom, days[kind]) if atom['state'] == 'proposed' else None
            if exp is None or not today <= exp <= today + dt.timedelta(days=within):
                continue
            users = sorted(oid for oid, ids in refs.items() if aid in ids)
            result['due'].append({'id': aid, 'expires': exp.isoformat(), 'outlines': users, 'shipped': bool(shipped & set(users))})

    # What changed since the shipping merge base (committed) and in the working tree (uncommitted).
    prefix = INFO_ARCH_DIR.as_posix() + '/'
    changed, old_ref = set(p for p in _dirty(root) if p.startswith(prefix)), 'HEAD'
    bases = sorted({row.get('merge_base') for _, row in m.channels.values() if row['ship'] and row.get('merge_base')})
    for ref in bases:
        try:
            fork = _git(root, 'merge-base', ref, 'HEAD')
        except ValueError:
            result['warnings'].append(f'merge_base {ref} is not resolvable here; committed IA changes are not listed')
            continue
        changed |= set(p for p in _git(root, 'diff', '--name-only', fork, 'HEAD', '--', prefix).splitlines() if p)
        old_ref = fork
    result['changed_paths'] = sorted(changed)
    impacted = set()
    for path in changed:
        stem = Path(path).stem
        if path.startswith(prefix + 'outlines/') and stem in m.outlines:
            impacted.add(stem)
        elif path == check.SSOT_PATH.as_posix():
            try:
                before = load_yaml_text(_git_raw(root, 'show', old_ref + ':' + path))
            except (ValueError, YamlError):
                before = None
            old, new = _atoms(before), _atoms(m.ssot)
            moved = {aid for aid in set(old) | set(new) if old.get(aid) != new.get(aid)}
            if before is None or any(aid.startswith(('ent-', 'vce-')) for aid in moved):
                impacted |= set(m.outlines)  # names and voice apply to every outline
            impacted |= {oid for oid, ids in refs.items() if ids & moved}
    result['impacted_outlines'] = sorted(impacted)

    if report.exit_code < 2:
        jrep, _, _ = judge.run(root, date)
        result['uncovered_judgments'] = len(jrep.uncovered)
        for f in jrep.findings:
            if f.code == 'IA_JUDGE_EMPHASIS_MISSING':
                continue  # the ship gate renders first screens; maintenance stays browser-free
            if f.level in ('fail', 'warn'):
                result['errors' if f.level == 'fail' else 'warnings'].append(f.code + ': ' + f.message)
    if result['due']:
        result['warnings'].append(f'{len(result["due"])} 件の証拠・固定値の期限が {within} 日以内に来る')
    result['exit_code'] = 2 if result['errors'] else (1 if result['warnings'] else 0)
    return result
