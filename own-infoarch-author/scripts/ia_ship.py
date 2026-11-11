"""IA ship gate and maintenance (ISSUE-05): the shipped SHA, never the working tree, is what gets checked.

- gate: one commit, read from git objects through a detached worktree of it (tracked files only). `tree` refuses
  symlinked IA inputs and shipped outputs that are not regular files in the commit (HTML or not); then the
  deterministic `check` (every ship:true channel needs at least one outline), `dom` on every shipped HTML output,
  and `judge` (read-only against the shared judgment ledger). A shipped outline's uncovered judgment stops the
  gate; a score at or above theta goes to a human (warning), because the emphasis gate is still undecided.
- channel: `ship_command` is the formal entry. It runs only after a green gate on the same commit.
- pre-push: every pushed commit itself, its predicted merge into each channel `merge_base`, and its predicted
  merge into the commit this same push sets that base branch to. Deleted refs ship nothing.
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
import sys
import tempfile

import check
import dom
import judge
from common import INFO_ARCH_DIR, SKILL, YamlError, load_ref, load_yaml

EXTENSION = 'pre-push.d/61-infoarch-ship'
MARKER = Path('.vibe-guard/infoarch-ship.enabled')
HTML_SUFFIXES = ('.html', '.htm')
FILE_MODES = ('100644', '100755')
SYMLINK_MODE = '120000'
TREE_MODE = '040000'


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


# ---- what the commit holds (git objects, not the working tree) ------------------------------------

def _kind(root, commit, path):
    """'file' | 'symlink' | 'tree' | 'other' | 'missing' for `path` in `commit`, read from git objects.

    Every parent is walked, so a path under a symlinked directory reads as 'symlink', never as the link's target.
    """
    parts, mode = path.split('/'), None
    for i, name in enumerate(parts):
        parent = '/'.join(parts[:i])
        try:
            rows = _git_raw(root, 'ls-tree', '-z', commit + ':' + parent if parent else commit)
        except ValueError:
            return 'missing'
        mode = next((meta.split()[0] for meta, _, entry in (row.partition('\t') for row in rows.split('\0'))
                     if entry == name and meta), None)
        if mode is None:
            return 'missing'
        if mode == SYMLINK_MODE:
            return 'symlink'
        if i < len(parts) - 1 and mode != TREE_MODE:
            return 'missing'
    return {TREE_MODE: 'tree', SYMLINK_MODE: 'symlink'}.get(mode, 'file' if mode in FILE_MODES else 'other')


def _input_findings(root, commit):
    """The IA inputs the gate reads must be committed as files: a symlink would make the check read elsewhere."""
    out = []
    base = INFO_ARCH_DIR.as_posix()
    kind = _kind(root, commit, base)
    if kind == 'symlink':
        out.append(('IA_SHIP_SYMLINK', f'{base} はシンボリックリンク（出荷する木の中身として読めない）', base))
    elif kind == 'tree':
        for row in _git_raw(root, 'ls-tree', '-r', '-z', commit + ':' + base).split('\0'):
            meta, _, entry = row.partition('\t')
            if meta and meta.split()[0] == SYMLINK_MODE:
                out.append(('IA_SHIP_SYMLINK', f'{base}/{entry} はシンボリックリンク（出荷する木の中身として読めない）', f'{base}/{entry}'))
    if _kind(root, commit, judge.LEDGER.as_posix()) == 'symlink':
        out.append(('IA_SHIP_SYMLINK', f'{judge.LEDGER.as_posix()} はシンボリックリンク', judge.LEDGER.as_posix()))
    return out


def _output_findings(root, commit, m, shipped):
    """Every output of a shipped outline, HTML or not, must be a regular file in the shipped commit."""
    out = []
    for oid in shipped:
        for path in m.outlines[oid][1]['outputs']:
            kind = _kind(root, commit, path)
            if kind == 'symlink':
                out.append(('IA_SHIP_SYMLINK', f'{oid}: 成果物 {path} はシンボリックリンク（出荷する木の中身として読めない）', path))
            elif kind != 'file':
                out.append(('IA_SHIP_OUTPUT_MISSING', f'{oid}: 成果物 {path} が出荷するコミットに通常ファイルとして無い', path))
    return out


def _tree_stage(stage, rows):
    stage['findings'] += [{'level': 'fail', 'code': code, 'message': msg, 'path': path} for code, msg, path in rows]
    stage['counts'] = {k: sum(f['level'] == k for f in stage['findings']) for k in ('fail', 'warn', 'info')}
    stage['exit_code'] = 2 if stage['counts']['fail'] else 0


@contextmanager
def checkout(root, commit):
    """A detached worktree of `commit`: tracked content only, so ignored files and later edits are never read."""
    with tempfile.TemporaryDirectory(prefix='infoarch-ship-') as temporary:
        snapshot = Path(temporary) / 'checkout'
        _git(root, 'worktree', 'add', '--detach', str(snapshot), commit)
        try:
            yield snapshot.resolve()
        finally:
            _git(root, 'worktree', 'remove', '--force', str(snapshot))


# ---- gate ----------------------------------------------------------------------------------------

def gate(root, date, *, commit=None, source=None, predicted=False):
    """-> (result, model). Check, render and judge one commit read from git objects; never deploys.

    Without `commit` the clean HEAD of `root` is gated (and must still be the same commit afterwards).
    """
    root = Path(root).resolve()
    bound = commit is None
    if bound:
        commit, tree = _clean_commit(root)
    else:
        commit = _git(root, 'rev-parse', '--verify', commit + '^{commit}')
        tree = _git(root, 'rev-parse', commit + '^{tree}')
    # The snapshot has no node_modules: rendering resolves Playwright from this checkout unless told otherwise.
    os.environ.setdefault('INFOARCH_PLAYWRIGHT_DIR', str(root))
    with checkout(root, commit) as snapshot:
        result, m = _evaluate(snapshot, commit, tree, date, source, predicted)
    if bound and _clean_commit(root) != (commit, tree):
        raise ValueError('source changed during the IA ship gate')
    return result, m


def _evaluate(root, commit, tree, date, source, predicted):
    result = {'exit_code': 0, 'commit': commit, 'tree': tree, 'source_commit': source or commit,
              'predicted_merge': predicted, 'evaluated_on': date, 'stages': [], 'errors': [], 'warnings': [],
              'shipped_outlines': [], 'deployment': 'not-requested'}
    m = None
    # Inputs are read before check (so check never reads through a link); outputs are added once outlines are known.
    files = {'stage': 'tree', 'exit_code': 0, 'counts': {}, 'targets': {'commit': commit}, 'findings': []}
    result['stages'].append(files)
    _tree_stage(files, _input_findings(root, commit))

    def finish(error=None):
        if error:
            result['errors'].append(error)
        for stage in result['stages']:
            result['errors'] += _messages(stage, 'fail')
            result['warnings'] += _messages(stage, 'warn')
        result['exit_code'] = 2 if result['errors'] else max([s['exit_code'] for s in result['stages']], default=0)
        if _clean_commit(root) != (commit, tree):
            raise ValueError('the IA ship gate changed the checked snapshot')
        return result, m

    if files['exit_code'] == 2:
        return finish('IA ship inputs are not committed files')
    report, m = check.check(root, date)
    result['stages'].append(_stage('check', report))
    if report.exit_code == 2:
        return finish('IA check failed')
    channels = sorted(name for name, (_, row) in m.channels.items() if row['ship'])
    if not channels:
        return finish('no ship:true IA channel declared; nothing ships through the IA gate')
    shipped = shipped_outlines(m)
    result['shipped_outlines'] = shipped
    per_channel = {name: [oid for oid in shipped if m.outlines[oid][1]['channel'] == name] for name in channels}
    result['channels'] = per_channel
    empty = [name for name, oids in per_channel.items() if not oids]
    if empty:
        return finish('ship:true IA channel with zero outlines (an empty gate is not a pass): ' + ', '.join(empty))
    _tree_stage(files, _output_findings(root, commit, m, shipped))
    files['targets'] = {'commit': commit, 'outputs': sum(len(m.outlines[oid][1]['outputs']) for oid in shipped)}
    if files['exit_code'] == 2:
        return finish('shipped IA outputs are not committed files')

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


def _is_zero(sha):
    return set(sha) == {'0'}


def _commit_or_none(root, sha):
    try:
        return _git(root, 'rev-parse', '--verify', '--quiet', sha + '^{commit}')
    except ValueError:
        return None


def predicted_commit(root, base, source):
    """A synthetic merge of `source` into `base`; never referenced or pushed."""
    # A conflict or a missing merge base makes git fail, and that stops the push.
    tree = _git(root, 'merge-tree', '--write-tree', base, source).splitlines()[0]
    # It takes the pushing user's configured email; a checkout with no identity fails here, which stops the push.
    return _git(root, '-c', 'user.name=IA Gate', 'commit-tree', tree, '-p', source, '-p', base,
                input='IA predicted merge validation\n')


def _marker_enabled(root, rev):
    """Whether `rev` commits a valid marker. A committed marker that is not a one-line "1" file stops the push."""
    kind = _kind(root, rev, MARKER.as_posix())
    if kind == 'missing':
        return False
    if kind != 'file' or _git_raw(root, 'show', rev + ':' + MARKER.as_posix()).strip() != '1':
        raise ValueError('invalid IA ship marker committed in ' + rev[:12] + '; shipment stopped')
    return True


def _bases(root, base_ref, pushed):
    """The merge base as it is now, plus what this same push sets the branch behind it to."""
    bases = [_git(root, 'rev-parse', '--verify', base_ref + '^{commit}')]
    try:
        full = _git(root, 'rev-parse', '--symbolic-full-name', base_ref)
    except ValueError:
        full = ''
    if full.startswith('refs/remotes/') and full.count('/') >= 3:
        full = 'refs/heads/' + full.split('/', 3)[3]
    if full.startswith('refs/heads/') and pushed.get(full) and pushed[full] not in bases:
        bases.append(pushed[full])
    return bases


def pre_push(root, refs_file, date):
    root = Path(root).resolve()
    if refs_file is None:
        raise ValueError('pre-push requires the preserved refs file')
    records = Path(refs_file).read_text().splitlines()
    if not records:
        raise ValueError('zero push refs')
    updates, pushed = [], {}
    for row in records:
        parts = row.split()
        if len(parts) != 4:
            raise ValueError('invalid push ref record')
        _, sha, remote_ref, remote_sha = parts
        if _is_zero(sha):
            continue  # a deleted ref ships nothing
        if _git(root, 'rev-parse', '--verify', sha + '^{commit}') != sha:
            raise ValueError('push source must be a commit')
        updates.append((sha, remote_ref, remote_sha))
        pushed[remote_ref] = sha
    # The extension decides whether this repository's gate is on; once called, every pushed commit is gated.
    checks, warnings, seen = [], [], set()

    def run(target, source, base_ref=None, base=None):
        tree = _git(root, 'rev-parse', target + '^{tree}')
        if tree in seen:
            return None  # the same tree was already gated in this push
        seen.add(tree)
        result, _ = gate(root, date, commit=target, source=source, predicted=base is not None)
        if base is not None:
            result.update(merge_base=base_ref, merge_base_commit=base)
        checks.append(result)
        return result if result['exit_code'] == 2 else None

    for sha, remote_ref, remote_sha in updates:
        channels = _channels_at(root, sha)
        if not _marker_enabled(root, sha):
            # Removing the committed marker does not switch the gate off for the push that removes it: say so.
            before = [] if _is_zero(remote_sha) else [_commit_or_none(root, remote_sha)]
            before += [_commit_or_none(root, row['merge_base']) for row in channels.values() if isinstance(row.get('merge_base'), str)]
            if any(rev and _marker_enabled(root, rev) for rev in before):
                note = (f'{remote_ref}: this push removes {MARKER.as_posix()}; the IA ship gate still checks it '
                        'while the remote branch or the merge base keeps the marker')
                warnings.append(note)
                print('warning: ' + note, file=sys.stderr)
        if not channels:
            raise ValueError('no ship:true IA channel in the pushed commit; the enabled IA gate cannot pass it')
        missing = sorted(name for name, row in channels.items() if not isinstance(row.get('merge_base'), str) or not row['merge_base'])
        if missing:
            raise ValueError('ship:true IA channel without merge_base (no predicted merge): ' + ', '.join(missing))
        # The pushed commit itself: an old red commit pushed to a new branch can be green once merged.
        failed = run(sha, sha)
        for base_ref in sorted({row['merge_base'] for row in channels.values()}):
            if failed:
                break
            for base in _bases(root, base_ref, pushed):
                synthetic = predicted_commit(root, base, sha)
                predicted = _channels_at(root, synthetic)
                if set(predicted) != set(channels) or any(predicted[n].get('merge_base') != channels[n].get('merge_base') for n in channels):
                    raise ValueError('predicted tree changed IA ship coverage or merge base')
                failed = run(synthetic, sha, base_ref, base)
                if failed:
                    break
        if failed:
            what = 'predicted merge IA ship gate failed' if failed['predicted_merge'] else 'pushed commit IA ship gate failed'
            return {'exit_code': 2, 'errors': [what], 'checks': checks, 'warnings': warnings}
    # The pre-push dispatcher stops on any nonzero exit: warnings travel in the body, only failures exit 2.
    return {'exit_code': 0, 'checks': checks, 'warnings': warnings + [w for c in checks for w in c['warnings']],
            'skipped': [] if updates else ['deleted refs only; no shipment']}


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
