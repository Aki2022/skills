"""Fixtures for the release approval: a layout plan, and the owner's ledger row for exactly the committed content.

The hash is computed here from the bytes of the three things the spec names, not through release.py, so a test that
approves with it checks the real binding rather than the module against itself."""
import hashlib
from pathlib import Path
import sys

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import layer_paths  # noqa: E402
import positioning as pos  # noqa: E402

OID = 'deck-smb'
KEY = 'ia:release:' + OID


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


OUTPUTS = {OID: ('decks/smb.pptx',)}


def plan_of(root, oid=OID):
    return layer_paths.path(root, 'build', oid, 'plan.yaml')


def script_of(root, oid=OID):
    return layer_paths.path(root, 'infoarch', 'scripts', f'{oid}.yaml')


def expected_hash(root, oid=OID, outputs=None):
    outputs = outputs or OUTPUTS[oid]
    return pos.sha256({'script': {layer_paths.rel('infoarch', 'scripts', f'{oid}.yaml'): sha(script_of(root, oid))},
                       'plan': {layer_paths.rel('build', oid, 'plan.yaml'): sha(plan_of(root, oid))},
                       'outputs': {o: sha(Path(root) / o) for o in outputs}})


def write_plan(root, oid=OID):
    script = yaml.safe_load(script_of(root, oid).read_text(encoding='utf-8'))
    plan = {'schema_version': 1, 'deliverable': oid, 'script': oid, 'blocks': {b: {'pattern': 'band'} for b in script['blocks']},
            'comp': {'reviewed': True}}
    plan_of(root, oid).parent.mkdir(parents=True, exist_ok=True)
    plan_of(root, oid).write_text(yaml.safe_dump(plan, allow_unicode=True, sort_keys=False), encoding='utf-8')


def ledger_rows(root):
    path = layer_paths.path(root, 'ledger')
    return (yaml.safe_load(path.read_text(encoding='utf-8')) if path.is_file() else None) or []


def record(root, digest, oid=OID):
    rows = [r for r in ledger_rows(root) if r['key'] != 'ia:release:' + oid]
    rows.append({'key': 'ia:release:' + oid, 'value_hash': digest, 'approver': 'fixture-owner', 'date': '2026-11-27'})
    layer_paths.path(root, 'ledger').write_text(yaml.safe_dump(rows, allow_unicode=True, sort_keys=False), encoding='utf-8')


def approve_release(root, oid=OID, outputs=None):
    """The owner approves what is on disk now: (re)write the plan, then record the hash of script + plan + files."""
    write_plan(root, oid)
    record(root, expected_hash(root, oid, outputs), oid)


def unapprove(root):
    """Back to a script with no plan and no release row."""
    plan_of(root).unlink()
    rows = [r for r in ledger_rows(root) if r['key'] != KEY]
    layer_paths.path(root, 'ledger').write_text(yaml.safe_dump(rows, allow_unicode=True, sort_keys=False), encoding='utf-8')
