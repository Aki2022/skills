"""Schema-2 bridge to the existing deterministic IA checker; IA has no approvals."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

VERSION = '0.2.0'
CONTRACT = 'ia-git-history-v2'
CATEGORIES = ('executed', 'skipped', 'missing', 'uncovered', 'screened')


def _unique(pairs):
    out = {}
    for key, value in pairs:
        if key in out:
            raise ValueError('duplicate IA response key')
        out[key] = value
    return out


def resolve(*, root, approvals, profiles):
    root = Path(root).resolve(strict=True)
    git = subprocess.run(['git', '-C', str(root), 'rev-parse', '--show-toplevel'], capture_output=True, text=True)
    if git.returncode or Path(git.stdout.strip()).resolve() != root:
        raise ValueError('IA root must be the git repository root')
    base = root / 'design/info_arch'
    # Check every component and input before handing paths to the upstream loader.
    for path in [root / 'design', base, *sorted(base.rglob('*'))]:
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError('IA input symlink or escape')
        if not path.is_file() and not path.is_dir():
            raise ValueError('IA input is missing or not a regular file/directory')
    inputs = {p.relative_to(root).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
              for p in sorted(base.rglob('*')) if p.is_file()}
    cli = Path(__file__).resolve().parent / 'infoarch_author.py'
    evaluated_on = dt.date.today().isoformat()
    run = subprocess.run([sys.executable, str(cli), 'check', '--root', str(root),
                          '--date', evaluated_on], cwd=root,
                         capture_output=True, text=True, timeout=60)
    data = json.loads(run.stdout, object_pairs_hook=_unique)
    expected = {'tool', 'version', 'level', 'evaluated_on', 'exit_code', 'environment',
                'counts', 'targets', 'uncovered', 'findings'}
    if (not isinstance(data, dict) or set(data) != expected or data['tool'] != 'own-infoarch-author'
            or data['version'] != VERSION or data['level'] != 'check'
            or type(data['exit_code']) is not int or run.returncode != data['exit_code']
            or run.returncode not in (0, 1, 2) or data['evaluated_on'] != evaluated_on
            or not isinstance(data['environment'], dict) or set(data['environment']) != {'python', 'platform'}
            or any(not isinstance(value, str) for value in data['environment'].values()) or data['uncovered'] != []):
        raise ValueError('invalid upstream IA check response')
    targets = data['targets']
    if not isinstance(targets, dict) or set(targets) != {'channels', 'outlines', 'sections'} or any(
            type(n) is not int or n < 0 for n in targets.values()):
        # A malformed SSOT can return no counts. Fail closed with its actual diagnostics.
        if run.returncode != 2 or targets != {}:
            raise ValueError('invalid upstream IA target counts')
        targets = {'channels': 0, 'outlines': 0, 'sections': 0}
    counts = data['counts']
    findings = data['findings']
    if not isinstance(findings, list) or any(not isinstance(f, dict) or set(f) != {
            'level', 'code', 'message', 'path', 'line', 'detail'} or f['level'] not in ('fail', 'warn', 'info')
            or any(not isinstance(f[k], str) for k in ('code', 'message', 'path'))
            or type(f['line']) is not int or not isinstance(f['detail'], dict) for f in findings):
        raise ValueError('invalid upstream IA findings')
    if counts != {level: sum(f['level'] == level for f in findings) for level in ('fail', 'warn', 'info')}:
        raise ValueError('inconsistent upstream IA finding counts')
    if data['exit_code'] != max(({'info': 0, 'warn': 1, 'fail': 2}[f['level']] for f in findings), default=0):
        raise ValueError('inconsistent upstream IA exit code')
    errors = [f['code'] + ': ' + f['message'] for f in findings if f['level'] == 'fail']
    warnings = [f['code'] + ': ' + f['message'] for f in findings if f['level'] == 'warn']
    if any(n == 0 for n in targets.values()):
        errors.append('zero IA check targets')
    now = {p.relative_to(root).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
           for p in sorted(base.rglob('*')) if p.is_file() and not p.is_symlink()}
    if now != inputs:
        raise ValueError('IA inputs changed during check')
    head = subprocess.run(['git', '-C', str(root), 'rev-parse', '--verify', 'HEAD'], capture_output=True, text=True)
    commit = head.stdout.strip() if head.returncode == 0 else None
    if commit is not None and not re.fullmatch(r'[0-9a-f]{40}|[0-9a-f]{64}', commit):
        raise ValueError('invalid IA git commit')
    governance = {'contract': CONTRACT, 'method': 'git-history', 'schema_version': 2,
                  'commit': commit, 'input_hashes': inputs}
    report = {key: [] for key in CATEGORIES}
    report['executed'].append({'id': 'ia-deterministic', 'exit_code': 2 if errors else run.returncode,
                               'targets': targets['sections'], 'counts': targets})
    report['missing'].append({'id': 'ia-semantic-judgments', 'reason': 'IA semantic judgment integration is not implemented; deterministic IA check does not judge meaning'})
    report['skipped'].append({'id': 'ia-dom', 'reason': 'L0 does not run the implemented rendered IA check; run infoarch-author dom separately'})
    return {'contract': CONTRACT, 'governance': governance, 'targets': targets,
            'media': {medium: {'governance': governance, 'targets': targets} for medium in profiles},
            'approval_hashes': {}, 'must_approval_keys': [], 'errors': errors, 'warnings': warnings, 'report': report}
