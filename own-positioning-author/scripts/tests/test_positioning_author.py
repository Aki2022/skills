"""own-positioning-author on its own: the standalone check (every requirement-5 red case), schema 4 trees, the CLI."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

from ia_fixtures import SSOT, apply, confirm, dump

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
TABLE = json.loads((SKILL / 'fixtures/positioning-red-cases.json').read_text(encoding='utf-8'))
CONDITIONS = {c['id']: c for c in TABLE['conditions']}
# Requirement 15 (a shipped outline points at an unapproved path) needs an outline: it is the infoarch layer's check.
NEEDS_OUTLINE = {'15-unapproved-shipped'}
STANDALONE = [c for c in TABLE['cases'] if c['condition'] not in NEEDS_OUTLINE]


def cli(*args):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'positioning_author.py'), *args], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


def write(root, ssot=SSOT, confirmed=True, skip=()):
    dump(root / 'design/info_arch/info_arch.yaml', ssot)
    if confirmed:
        confirm(root, ssot, skip=skip)


def fails(report):
    return [f['code'] for f in report['findings'] if f['level'] == 'fail']


def test_green_fixture_passes_and_counts_real_targets(tmp_path):
    write(tmp_path)
    rc, report = cli('check', '--root', str(tmp_path), '--strict')
    assert rc == 0, report['findings']
    assert report['tool'] == 'own-positioning-author'
    assert report['targets'] == {'values': 3, 'approved': 3}


def test_the_standalone_table_is_not_empty_and_names_only_known_conditions():
    assert len(STANDALONE) >= 26  # 28 cases, only the 2 that need an outline are left out
    assert {c['condition'] for c in STANDALONE} <= set(CONDITIONS)
    assert {c['condition'] for c in TABLE['cases']} - {c['condition'] for c in STANDALONE} == NEEDS_OUTLINE


@pytest.mark.parametrize('case', STANDALONE, ids=[c['id'] for c in STANDALONE])
def test_every_positioning_red_case_is_red_without_any_outline(case, tmp_path):
    ssot, unapproved = apply(case['ops'])
    write(tmp_path, ssot, skip=unapproved)
    rc, report = cli('check', '--root', str(tmp_path), '--strict')
    assert rc == 2, report['findings']
    assert CONDITIONS[case['condition']]['code'] in fails(report), fails(report)


@pytest.mark.parametrize('case', [c for c in TABLE['cases'] if c['condition'] in NEEDS_OUTLINE], ids=lambda c: c['id'])
def test_the_outline_use_conditions_are_not_raised_here(case, tmp_path):
    """Whether an unapproved path stops anything depends on which outline points at it: infoarch's check says so."""
    ssot, unapproved = apply(case['ops'])
    write(tmp_path, ssot, skip=unapproved)
    rc, report = cli('check', '--root', str(tmp_path))
    assert 'IA_POSITIONING_UNAPPROVED' not in fails(report) and rc != 2, report['findings']


def test_without_strict_a_disqualified_path_nobody_points_at_only_warns(tmp_path):
    case = next(c for c in STANDALONE if c['id'] == 'strength-uncompared')
    ssot, unapproved = apply(case['ops'])
    write(tmp_path, ssot, skip=unapproved)
    rc, report = cli('check', '--root', str(tmp_path))
    assert rc == 1 and CONDITIONS[case['condition']]['code'] in [f['code'] for f in report['findings'] if f['level'] == 'warn']


def test_no_values_is_not_a_pass(tmp_path):
    ssot = copy.deepcopy(SSOT)
    for block in ('segments', 'jobs', 'pains', 'gains', 'alternatives', 'categories', 'strengths', 'values', 'proofs'):
        ssot[block] = {}
    ssot['edges'] = []
    write(tmp_path, ssot, confirmed=False)
    rc, report = cli('check', '--root', str(tmp_path))
    assert rc == 2 and 'IA_NO_VALUES' in fails(report)


def test_missing_ssot_is_exit_2(tmp_path):
    rc, report = cli('check', '--root', str(tmp_path))
    assert rc == 2 and 'IA_MISSING' in fails(report)


# ---- schema 4: the format before the dictionary, read with a warning --------------------------------------------

V4 = {**copy.deepcopy(SSOT), 'schema_version': 4,
      'facts': {'fct-0001': {'state': 'proposed', 'provenance': 'owner', 'label': '導入社数', 'value': '100社以上', 'as_of': '2026-10-01'}},
      'entities': {'ent-0001': {'state': 'proposed', 'provenance': 'owner', 'type': 'product', 'name': 'Webサーバー'}},
      'voices': {'vce-0001': {'state': 'proposed', 'provenance': 'owner', 'traits': [
          {'trait': t, 'is': 'できることを具体的に言う', 'is_not': '大げさな約束はしない'} for t in ('誠実', '前向き', 'わかりやすい')]}}}
V4['ledger'] = {'next': {**SSOT['ledger']['next'], 'fct': 2, 'ent': 2, 'vce': 2}, 'tombstones': {}}


def test_a_schema_4_tree_is_read_with_a_warning_and_its_voice_is_dropped(tmp_path):
    write(tmp_path, V4)
    rc, report = cli('check', '--root', str(tmp_path), '--strict')
    assert rc == 1 and not fails(report), report['findings']
    assert [f['code'] for f in report['findings'] if f['level'] == 'warn'] == ['IA_SSOT_LEGACY_V4']


def test_the_lift_moves_facts_and_entities_out_and_keeps_every_positioning_id():
    import positioning_model as pm
    lifted, terms = pm.lift_v4(V4)
    assert lifted['schema_version'] == 5 and not {'facts', 'entities', 'voices'} & set(lifted)
    assert set(lifted['ledger']['next']) == set(SSOT['ledger']['next'])
    assert set(terms) == {'facts', 'entities'} and 'fct-0001' in terms['facts']
    assert {b: lifted[b] for b in pm.BLOCKS} == {b: V4[b] for b in pm.BLOCKS}


def test_a_broken_schema_4_tree_is_still_a_failure(tmp_path):
    broken = copy.deepcopy(V4)
    broken['values']['val-0001']['state'] = 'approved'
    write(tmp_path, broken, confirmed=False)
    rc, report = cli('check', '--root', str(tmp_path))
    assert rc == 2 and 'IA_SCHEMA' in fails(report)


def test_a_schema_5_tree_that_still_carries_a_voice_is_refused(tmp_path):
    bad = {**copy.deepcopy(SSOT), 'voices': V4['voices']}
    write(tmp_path, bad, confirmed=False)
    rc, report = cli('check', '--root', str(tmp_path))
    assert rc == 2 and 'IA_SCHEMA' in fails(report)


def test_other_schema_versions_are_refused(tmp_path):
    write(tmp_path, {**copy.deepcopy(SSOT), 'schema_version': 3}, confirmed=False)
    rc, report = cli('check', '--root', str(tmp_path))
    assert rc == 2 and 'IA_SCHEMA' in fails(report)


# ---- the other commands ------------------------------------------------------------------------------------------

def test_targets_are_one_hash_per_value_path(tmp_path):
    write(tmp_path)
    rc, out = cli('targets', '--root', str(tmp_path))
    assert rc == 0 and sorted(out['approval_hashes']) == ['ia:path:val-0001', 'ia:path:val-0002', 'ia:path:val-0003']
    assert all(len(h) == 64 for h in out['approval_hashes'].values())


def test_positioning_lists_paths_and_what_is_approved(tmp_path):
    write(tmp_path, skip=('val-0003',))
    rc, out = cli('positioning', '--root', str(tmp_path))
    assert rc == 0
    assert {v: p['approved'] for v, p in out['paths'].items()} == {'val-0001': True, 'val-0002': True, 'val-0003': False}


# ---- the voice is gone (SPEC-deliverable-production-layers requirement 8) -----------------------------------------

# The only files that may still name a voice: the schema 4 format that is read with a warning, the migrations out of
# it, and the tests and red cases that prove each of those. Anything else is a voice that came back.
VOICE_RECORDS = {
    'own-positioning-author/references/schema/info_arch.v4.schema.json',
    'own-positioning-author/scripts/positioning_model.py',
    'own-positioning-author/scripts/tests/test_positioning_author.py',
    'own-infoarch-author/scripts/migrate_v6.py',
    'own-infoarch-author/scripts/migrate_dictionary.py',
    'own-infoarch-author/scripts/tests/test_migrate_v6.py',
    'own-infoarch-author/scripts/tests/test_migrate_dictionary.py',
    'own-infoarch-author/scripts/tests/test_check.py',
    'own-infoarch-author/fixtures/red-cases.json',
    'own-infoarch-author/references/contract.md',
}


def test_no_skill_names_a_voice_outside_the_migration_records():
    root = SKILL.parent
    hits = set()
    for path in root.rglob('*'):
        parts = path.relative_to(root).parts
        if (not path.is_file() or parts[0] not in ('own-positioning-author', 'own-infoarch-author', 'own-design-author',
                                                   'own-brand-resolve', 'own-web-build', 'own-deliverable-check')
                or '__pycache__' in parts or '.pytest_cache' in parts or path.suffix in ('.pyc', '.png', '.woff2')):
            continue
        try:
            text = path.read_text(encoding='utf-8')
        except (UnicodeDecodeError, OSError):
            continue
        if 'vce-' in text or 'voices' in text:
            hits.add('/'.join(parts))
    assert hits <= VOICE_RECORDS, sorted(hits - VOICE_RECORDS)
    assert VOICE_RECORDS - hits == set(), 'an allowed record no longer names a voice: shrink the allowlist'
