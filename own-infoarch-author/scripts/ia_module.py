"""Bridge to the deterministic IA checker; the owner's approval of each value path is an approval target."""
from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

# This bridge is imported by own-design-author's L0 and has no sibling imports of its own; the layer rule is design's.
_DESIGN_SCRIPTS = Path(__file__).resolve().parents[2] / 'own-design-author' / 'scripts'
if not (_DESIGN_SCRIPTS / 'layer_paths.py').is_file():
    raise RuntimeError('own-design-author/scripts/layer_paths.py が無い（layer の置き場所の規則は design が持つ）')
if str(_DESIGN_SCRIPTS) not in sys.path:
    sys.path.append(str(_DESIGN_SCRIPTS))
import layer_paths  # noqa: E402

VERSION = '0.4.0'
CONTRACT = 'ia-git-history-v4'
# ia:item:<item id>: one row per product item of a blueprint (the item's own hash, SPEC-product-ui-infoarch requirement 3).
# ia:path:<value id>: one row per value, its hash binding the whole path (ADR-20261118-ia-approval-per-value-path).
# Per-element keys (ia:positioning:...) are not targets: a ledger row with one is not counted as an approval.
KEY = re.compile(r'ia:path:val-[0-9]{4,}|ia:item:(?:sto|pri|con|ojc|ntr)-[0-9]{4,}')  # a value's path; a product item (blueprint)
# The targets a check counted: the outline side (a deliverable) and the blueprint side (a product UI). The five keys are
# fixed; design_modules._git_contract validates the same five, and a resolved manifest records them in media.*.targets.
OUTLINE_TARGETS = ('channels', 'outlines', 'sections')
BLUEPRINT_TARGETS = ('screens', 'items')
TARGET_KEYS = OUTLINE_TARGETS + BLUEPRINT_TARGETS
CATEGORIES = ('executed', 'skipped', 'missing', 'uncovered', 'screened')
# The fixed values and official names the checker compares against are the docs dictionary's, outside ux/. They are
# inputs too: one file per domain, `docs/dictionary/<domain>.json` (own-doc-update/scripts/dictionary.py reads exactly
# `DICTIONARY_DIR/*.json`; a test pins the two together). Its hashes sit beside the ux/infoarch/ and ux/positioning/ ones in
# governance.input_hashes, so a changed word changes them (design L0's consumers compare them against the tree).
DICTIONARY_DIR = 'docs/dictionary'


def zero_targets(targets):
    """True when the check counted nothing it can stand on. A side is all or nothing: a side with some counts and a zero
    (channels with no outline, screens with no item) is empty, and so is a check where both sides are zero. One whole
    side is enough, so a product UI with no outline is a target and a deliverable with no blueprint stays one."""
    sides = [[targets[k] for k in OUTLINE_TARGETS], [targets[k] for k in BLUEPRINT_TARGETS]]
    return any(any(side) and not all(side) for side in sides) or not any(all(side) for side in sides)


def _unique(pairs):
    out = {}
    for key, value in pairs:
        if key in out:
            raise ValueError('duplicate IA response key')
        out[key] = value
    return out


def dictionary_inputs(root):
    """{repo-relative path: sha256} of `docs/dictionary/*.json`; empty when the repository has no dictionary."""
    folder = root / DICTIONARY_DIR
    if not folder.exists() and not folder.is_symlink():
        return {}
    for path in [root / 'docs', folder]:
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError('IA input symlink or escape')
    if not folder.is_dir():
        raise ValueError('IA input is missing or not a regular file/directory')
    inputs = {}
    for path in sorted(folder.glob('*.json')):
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError('IA input symlink or escape')
        if path.is_file():
            inputs[path.relative_to(root).as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    return inputs


def resolve(*, root, approvals, profiles):
    root = Path(root).resolve(strict=True)
    git = subprocess.run(['git', '-C', str(root), 'rev-parse', '--show-toplevel'], capture_output=True, text=True)
    if git.returncode or Path(git.stdout.strip()).resolve() != root:
        raise ValueError('IA root must be the git repository root')
    base = layer_paths.path(root, 'infoarch')  # a repository that moved stops here (LAYER_LAYOUT_FOREIGN_ROOT)
    # Check every component and input before handing paths to the upstream loader.
    positioning = layer_paths.path(root, 'positioning_dir')
    # Either IA input tree may exist alone (a positioning with no infoarch yet is a real state, reported by check as
    # red); only a repository with neither stops here as "missing".
    trees = [t for t in (base, positioning) if t.exists() or t.is_symlink()] or [base]  # a dangling link is not absence
    for path in [layer_paths.path(root, 'layers'), *trees, *[p for tree in trees for p in sorted(tree.rglob('*'))]]:
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError('IA input symlink or escape')
        if not path.is_file() and not path.is_dir():
            raise ValueError('IA input is missing or not a regular file/directory')
    inputs = {p.relative_to(root).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
              for tree in trees for p in sorted(tree.rglob('*')) if p.is_file()}
    inputs.update(dictionary_inputs(root))
    cli = Path(__file__).resolve().parent / 'infoarch_author.py'
    evaluated_on = dt.date.today().isoformat()
    run = subprocess.run([sys.executable, str(cli), 'check', '--root', str(root),
                          '--date', evaluated_on], cwd=root,
                         capture_output=True, text=True, timeout=60)
    data = json.loads(run.stdout, object_pairs_hook=_unique)
    expected = {'tool', 'version', 'level', 'evaluated_on', 'exit_code', 'environment',
                'counts', 'targets', 'uncovered', 'findings', 'executed', 'skipped', 'missing', 'screened'}
    if (not isinstance(data, dict) or set(data) != expected or data['tool'] != 'own-infoarch-author'
            or data['version'] != VERSION or data['level'] != 'check'
            or type(data['exit_code']) is not int or run.returncode != data['exit_code']
            or run.returncode not in (0, 1, 2) or data['evaluated_on'] != evaluated_on
            or not isinstance(data['environment'], dict) or set(data['environment']) != {'python', 'platform'}
            or any(not isinstance(value, str) for value in data['environment'].values()) or data['uncovered'] != []
            or any(not isinstance(data[key], list) for key in ('executed', 'skipped', 'missing', 'screened'))):
        raise ValueError('invalid upstream IA check response')
    targets = data['targets']
    if not isinstance(targets, dict) or set(targets) != set(TARGET_KEYS) or any(
            type(n) is not int or n < 0 for n in targets.values()):
        # A malformed SSOT can return no counts. Fail closed with its actual diagnostics.
        if run.returncode != 2 or targets != {}:
            raise ValueError('invalid upstream IA target counts')
        targets = dict.fromkeys(TARGET_KEYS, 0)
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
    if zero_targets(targets):
        errors.append('zero IA check targets')
    # Owner approvals per value path (SPEC-information-architecture-standardization requirement 18).
    targets_run = subprocess.run([sys.executable, str(cli), 'targets', '--root', str(root)], cwd=root,
                                 capture_output=True, text=True, timeout=60)
    approval_hashes, must = {}, []
    if targets_run.returncode == 0:
        t = json.loads(targets_run.stdout, object_pairs_hook=_unique)
        if (not isinstance(t, dict) or set(t) != {'approval_hashes', 'must_approval_keys'}
                or not isinstance(t['approval_hashes'], dict) or not isinstance(t['must_approval_keys'], list)
                or any(not isinstance(k, str) or not KEY.fullmatch(k) or not isinstance(v, str)
                       or not re.fullmatch(r'[0-9a-f]{64}', v) for k, v in t['approval_hashes'].items())
                or any(k not in t['approval_hashes'] for k in t['must_approval_keys'])):
            raise ValueError('invalid upstream IA approval targets')
        approval_hashes, must = t['approval_hashes'], t['must_approval_keys']
    elif run.returncode != 2:
        raise ValueError('IA approval targets failed on a loadable SSOT')
    now = {p.relative_to(root).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
           for tree in trees for p in sorted(tree.rglob('*')) if p.is_file() and not p.is_symlink()}
    now.update(dictionary_inputs(root))
    if now != inputs:
        raise ValueError('IA inputs changed during check')
    head = subprocess.run(['git', '-C', str(root), 'rev-parse', '--verify', 'HEAD'], capture_output=True, text=True)
    commit = head.stdout.strip() if head.returncode == 0 else None
    if commit is not None and not re.fullmatch(r'[0-9a-f]{40}|[0-9a-f]{64}', commit):
        raise ValueError('invalid IA git commit')
    governance = {'contract': CONTRACT, 'method': 'git-history', 'schema_version': 4,
                  'commit': commit, 'input_hashes': inputs}
    report = {key: [] for key in CATEGORIES}
    report['executed'].append({'id': 'ia-deterministic', 'exit_code': 2 if errors else run.returncode,
                               'targets': targets['sections'] + targets['items'], 'counts': targets})
    report['missing'].append({'id': 'ia-semantic-judgments', 'reason': 'design L0 does not read IA judgments; the IA ship gate (infoarch-author ship) reads them from the shared ledger'})
    report['skipped'].append({'id': 'ia-dom', 'reason': 'design L0 does not render; the IA ship gate (infoarch-author ship) runs dom on shipped HTML outputs'})
    return {'contract': CONTRACT, 'governance': governance, 'targets': targets,
            'media': {medium: {'governance': governance, 'targets': targets} for medium in profiles},
            'approval_hashes': approval_hashes, 'must_approval_keys': must, 'errors': errors, 'warnings': warnings,
            'report': report}
