"""Regenerate changed/due design inputs with offline L0; never record approvals."""
from datetime import date, timedelta
import fcntl
import hashlib
import json
from pathlib import Path
import re
import subprocess


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
    return sorted({p for p in paths if not p.startswith('design/.resolved/') and p != 'design/assets.json'})


def _inputs(design):
    values = {}
    for path in sorted(design.rglob('*')):
        relative = path.relative_to(design)
        if '.resolved' in relative.parts or path.name in ('assets.json', 'migration.json'):
            continue
        if path.is_symlink() or not path.resolve().is_relative_to(design):
            raise ValueError('maintenance input is a symlink or escapes design/')
        if path.is_file():
            values[relative.as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    return values


def maintain_design(root, *, catalog_dir=None, today=None, force=False):
    from design_author import CATALOG, check_design, git_root, load_catalog, read_data, write_data
    root = git_root(root).resolve()
    design = root / 'design'
    result = {'exit_code': 0, 'status': 'skipped', 'triggers': [], 'changed_paths': [],
              'due_items': [], 'invalidated_approvals': [], 'approval_candidates': [],
              'judgment_candidates': [], 'errors': []}
    if not design.exists():
        return {**result, 'reason': 'design/ absent'}
    catalog_dir = Path(catalog_dir or CATALOG)
    today = today or date.today()
    try:
        if design.is_symlink() or not design.is_dir():
            raise ValueError('design/ must be a real directory')
        config = read_data(design / 'design.yaml')
        if not isinstance(config, dict):
            raise ValueError('design configuration must be an object')
        catalog = load_catalog(config.get('catalog_version'), catalog_dir)
        days, deadlines = validate_maintenance(config, {x['id'] for x in catalog['items']} |
                                               set(config.get('retired_items', {})))
        head = _git(root, 'rev-parse', '--verify', 'HEAD^{commit}').decode().strip()
        versions = _versions(config, catalog_dir)
        resolved = design / '.resolved'
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
            before_approvals = (design / 'approvals.yaml').read_bytes()
            before_inputs = _inputs(design)
            checked = check_design(root, catalog_dir=catalog_dir)
            result.update(status='checked', exit_code=checked['exit_code'], errors=checked['errors'],
                          warnings=checked['warnings'], counts=checked['counts'], report=checked['report'])
            targets = checked['approval_targets']
            approvals = read_data(design / 'approvals.yaml')
            if isinstance(approvals, list):
                result['invalidated_approvals'] = sorted({row['key'] for row in approvals
                    if isinstance(row, dict) and isinstance(row.get('key'), str)
                    and targets.get(row['key']) != row.get('value_hash')})
                approved = {row['key']: row.get('value_hash') for row in approvals
                            if isinstance(row, dict) and isinstance(row.get('key'), str)}
                result['approval_candidates'] = sorted(key for key in checked['approval_required']
                    if approved.get(key) != targets.get(key))
            result['judgment_candidates'] = [row['judgment_hash'] for row in checked['judgment_candidates']]
            if ((design / 'approvals.yaml').read_bytes() != before_approvals or
                    _inputs(design) != before_inputs or
                    _git(root, 'rev-parse', 'HEAD').decode().strip() != head or
                    _versions(config, catalog_dir) != versions):
                raise ValueError('maintenance inputs changed during validation; rerun')
            if checked['exit_code'] != 2:
                write_data(stamp_path, {'schema_version': 1, 'commit': head, 'versions': versions,
                                       'checked_at': today.isoformat()})
            return result
    except (ValueError, OSError, KeyError, TypeError) as error:
        return {**result, 'exit_code': 2, 'status': 'failed', 'errors': [str(error)]}
