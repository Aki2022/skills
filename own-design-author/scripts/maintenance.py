"""Regenerate changed/due design inputs with offline L0; never record approvals."""
from datetime import date, timedelta
import fcntl
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import subprocess

import layer_paths


def _git(root, *args):
    run = subprocess.run(['git', '-C', str(root), *args], capture_output=True)
    if run.returncode:
        raise ValueError('maintenance requires a readable git commit/history')
    return run.stdout


def validate_maintenance(config, known_ids):
    value = config.get('maintenance', {})
    if not isinstance(value, dict) or set(value) - {'within_days', 'review_due_dates'}:
        raise ValueError('invalid maintenance configuration')
    days = value.get('within_days', 7)
    dates = value.get('review_due_dates', {})
    if type(days) is not int or days < 0 or not isinstance(dates, dict):
        raise ValueError('invalid maintenance review window/dates')
    for ident, deadline in dates.items():
        if ident not in known_ids or not isinstance(deadline, str):
            raise ValueError('unknown maintenance review id/date')
        if date.fromisoformat(deadline).isoformat() != deadline:
            raise ValueError('maintenance review date must be ISO YYYY-MM-DD')
    return days, dates


def _versions(config, catalogs):
    from design_author import SKILL
    # Hash the implementations and policy tables, not just self-declared version labels.
    paths = [p for folder in ('scripts', 'references') for p in (SKILL / folder).rglob('*')
             if p.is_file() and 'tests' not in p.parts and '__pycache__' not in p.parts]
    hashes = {'skill/' + p.relative_to(SKILL).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
              for p in sorted(paths)}
    for name in (config['catalog_version'] + '.json', 'id-ledger.json'):
        path = Path(catalogs) / name
        hashes['catalog/' + name] = hashlib.sha256(path.read_bytes()).hexdigest()
    import layout_patterns  # the vocabulary is resolved into the manifest: a new catalog re-runs L0
    hashes['catalog/' + layout_patterns.CATALOG_FILE.name] = hashlib.sha256(layout_patterns.CATALOG_FILE.read_bytes()).hexdigest()
    return hashlib.sha256(json.dumps(hashes, sort_keys=True).encode()).hexdigest()


def _dirty_paths(root):
    records = _git(root, 'status', '--porcelain=v1', '-z', '--untracked-files=all').split(b'\0')
    paths, index = [], 0
    while index < len(records):
        row = records[index]
        index += 1
        if not row:
            continue
        paths.append(row[3:].decode())
        if row[:1] in (b'R', b'C') or row[1:2] in (b'R', b'C'):
            paths.append(records[index].decode())
            index += 1
    return paths


def _relevant(paths):
    generated = layer_paths.rel('resolved') + '/'
    return sorted({p for p in paths if not p.startswith(generated) and p != layer_paths.rel('assets_index')})


def _leaf(key):
    return PurePosixPath(layer_paths.rel(key)).name


def _inputs(layers):
    """Every file under the layers root (the ledger and the judgment record included), as the L0 manifest hashes them."""
    values = {}
    for path in sorted(layers.rglob('*')):
        relative = path.relative_to(layers)
        if _leaf('resolved') in relative.parts or path.name in (_leaf('assets_index'), 'migration.json') \
                or not layer_paths.design_input(relative.as_posix()):  # build records are not design inputs
            continue
        if path.is_symlink() or not path.resolve().is_relative_to(layers):
            raise ValueError(f"maintenance input is a symlink or escapes {layer_paths.rel('layers')}/")
        if path.is_file():
            values[relative.as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    return values


def maintain_design(root, *, catalog_dir=None, today=None, force=False):
    from design_author import CATALOG, check_design, design_dir, design_rel, git_root, load_catalog, read_data, write_data
    root = git_root(root).resolve()
    result = {'exit_code': 0, 'status': 'skipped', 'triggers': [], 'changed_paths': [],
              'due_items': [], 'invalidated_approvals': [], 'approval_candidates': [],
              'judgment_candidates': [], 'errors': []}
    try:
        design = design_dir(root)  # the layout guard: a repository this version does not read is a failure, never a skip
    except layer_paths.LayerLayoutError as error:
        return {**result, 'exit_code': 2, 'status': 'failed', 'errors': [str(error)]}
    if not design.exists():
        return {**result, 'reason': f'{design_rel()}/ absent'}
    config_path = layer_paths.path(root, 'design_config')
    catalog_dir = Path(catalog_dir or CATALOG)
    today = today or date.today()
    try:
        if design.is_symlink() or not design.is_dir():
            raise ValueError(f'{design_rel()}/ must be a real directory')
        config = read_data(config_path)
        if not isinstance(config, dict):
            raise ValueError('design configuration must be an object')
        catalog = load_catalog(config.get('catalog_version'), catalog_dir)
        days, deadlines = validate_maintenance(config, {x['id'] for x in catalog['items']} |
                                               set(config.get('retired_items', {})))
        head = _git(root, 'rev-parse', '--verify', 'HEAD^{commit}').decode().strip()
        versions = _versions(config, catalog_dir)
        resolved = layer_paths.path(root, 'resolved')
        if resolved.is_symlink():
            raise ValueError('maintenance output directory is a symlink')
        resolved.mkdir(exist_ok=True)
        lock_path = resolved / 'maintenance.lock'
        if lock_path.is_symlink():
            raise ValueError('maintenance lock is a symlink')
        with lock_path.open('a') as lock:
            try:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                raise ValueError('maintenance writer already active') from None
            stamp_path = resolved / 'maintenance-stamp.json'
            stamp = None
            if stamp_path.exists() or stamp_path.is_symlink():
                try:
                    if stamp_path.is_symlink():
                        raise ValueError('symlink stamp')
                    stamp = read_data(stamp_path)
                    if (set(stamp) != {'schema_version', 'commit', 'versions', 'checked_at'} or
                            stamp['schema_version'] != 1 or
                            not re.fullmatch('[a-f0-9]{40,64}', stamp['commit']) or
                            not re.fullmatch('[a-f0-9]{64}', stamp['versions'])):
                        raise ValueError('invalid stamp')
                    date.fromisoformat(stamp['checked_at'])
                    _git(root, 'cat-file', '-e', stamp['commit'] + '^{commit}')
                except (ValueError, KeyError, TypeError, OSError):
                    stamp = None
                    result['triggers'].append('invalid-stamp')
            else:
                result['triggers'].append('first-run')
            dirty = _relevant(_dirty_paths(root))
            committed = []
            if stamp:
                committed = _relevant(p.decode() for p in _git(root, 'diff', '--name-only', '-z',
                    '--diff-filter=ACDMRT', stamp['commit'], head).split(b'\0') if p)
                if committed:
                    result['triggers'].append('commit-diff')
                if stamp['versions'] != versions:
                    result['triggers'].append('versions-changed')
            if dirty:
                result['triggers'].append('working-tree')
            if stamp and any(not (resolved / (medium + '.json')).is_file()
                             for medium in config.get('media', [])):
                result['triggers'].append('resolved-missing')
            result['changed_paths'] = sorted(set(dirty + committed))
            result['due_items'] = [{'id': ident, 'review_by': deadline}
                for ident, deadline in sorted(deadlines.items())
                if date.fromisoformat(deadline) <= today + timedelta(days=days)]
            if result['due_items']:
                result['triggers'].append('review-due')
            if force:
                result['triggers'].append('forced')
            if not result['triggers']:
                return {**result, 'reason': 'no changes, due reviews, or version changes'}
            ledger_path = layer_paths.path(root, 'ledger')
            before_approvals = ledger_path.read_bytes()
            layers = layer_paths.path(root, 'layers')
            before_inputs = _inputs(layers)
            checked = check_design(root, catalog_dir=catalog_dir)
            result.update(status='checked', exit_code=checked['exit_code'], errors=checked['errors'],
                          warnings=checked['warnings'], counts=checked['counts'], report=checked['report'])
            targets = checked['approval_targets']
            approvals = read_data(ledger_path)
            if isinstance(approvals, list):
                result['invalidated_approvals'] = sorted({row['key'] for row in approvals
                    if isinstance(row, dict) and isinstance(row.get('key'), str)
                    and targets.get(row['key']) != row.get('value_hash')})
                approved = {row['key']: row.get('value_hash') for row in approvals
                            if isinstance(row, dict) and isinstance(row.get('key'), str)}
                result['approval_candidates'] = sorted(key for key in checked['approval_required']
                    if approved.get(key) != targets.get(key))
            result['judgment_candidates'] = [row['judgment_hash'] for row in checked['judgment_candidates']]
            if (ledger_path.read_bytes() != before_approvals or
                    _inputs(layers) != before_inputs or
                    _git(root, 'rev-parse', 'HEAD').decode().strip() != head or
                    _versions(config, catalog_dir) != versions):
                raise ValueError('maintenance inputs changed during validation; rerun')
            if checked['exit_code'] != 2:
                write_data(stamp_path, {'schema_version': 1, 'commit': head, 'versions': versions,
                                       'checked_at': today.isoformat()})
            return result
    except (ValueError, OSError, KeyError, TypeError) as error:
        return {**result, 'exit_code': 2, 'status': 'failed', 'errors': [str(error)]}


def maintain_all(root, catalog_dir=None, today=None, force=False):
    """own-session-close Step 2.5 in one entry (ADR-20261115-design-ia-two-skills-shared-base, biz_ops).

    The design stage may return early when nothing changed; the IA stage still runs every time, because evidence
    and facts fall due by date alone. Without the infoarch layer the IA stage is a one-line skip.
    """
    import subprocess
    import sys
    result = maintain_design(root, catalog_dir=catalog_dir, today=today, force=force)
    base = Path(root).resolve()
    try:
        has_infoarch = layer_paths.has_ia(base)
    except layer_paths.LayerLayoutError as error:  # a repository this version does not read is a failure, never a skip
        ia = {'exit_code': 2, 'status': 'failed', 'errors': [str(error)]}
    else:
        if not has_infoarch:
            ia = {'exit_code': 0, 'status': 'skipped', 'reason': f"{layer_paths.rel('infoarch')}/ and {layer_paths.rel('positioning_dir')}/ absent"}
        else:
            cli = Path(__file__).resolve().parents[2] / 'own-infoarch-author/scripts/infoarch_author.py'
            args = [sys.executable, str(cli), 'maintain', '--root', str(base)]
            if today:
                args += ['--date', today.isoformat()]
            run = subprocess.run(args, capture_output=True, text=True, timeout=600)
            try:
                ia = json.loads(run.stdout)
                if not isinstance(ia, dict) or ia.get('exit_code') != run.returncode:
                    raise ValueError
            except ValueError:
                ia = {'exit_code': 2, 'status': 'failed', 'errors': ['infoarch-author maintain gave no valid report']}
    return {**result, 'ia': ia, 'exit_code': max(result['exit_code'], ia['exit_code'])}
