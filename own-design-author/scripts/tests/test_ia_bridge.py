"""Actual IA checker bridge: git history, IA positioning approvals as approval targets, stale consumers."""
import copy
import importlib.util
import json
import os
from pathlib import Path
import runpy
import shutil
import subprocess
import sys

import pytest

from test_core import SKILL, approve, check, project, write
from design_modules import resolve_ia, validate_module_state
from reporting import empty_report

IA = SKILL.parent / 'own-infoarch-author'
FIXTURE = runpy.run_path(str(IA / 'scripts/tests/test_check.py'))


def attach(project):
    root, _ = project
    # Fixed artificial dates stay future-proof without weakening the expiry test.
    ssot = copy.deepcopy(FIXTURE['SSOT'])
    ssot['proofs']['prf-0001']['as_of'] = '2099-01-01'
    dictionary = copy.deepcopy(FIXTURE['DICTIONARY'])
    dictionary['entries'][0]['expires'] = '2099-01-01'  # a dated value far in the future: the expiry test is the one below
    FIXTURE['write_repo'](root, ssot, dictionary=dictionary)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['modules'] = {'ia': '0.4.0'}
    write(root / 'design/design.yaml', config)
    return root


def test_schema2_actual_check_has_counts_and_preserves_ds_approvals(project):
    approve(project)
    before = check(project, ship=True)
    root = attach(project)
    result = check(project, ship=True)
    assert result['exit_code'] == 0, result['errors']
    # The design targets stay exactly as they were; one IA row per value path is added beside them.
    ds = {k: v for k, v in result['approval_targets'].items() if k.startswith('ds:')}
    assert ds == before['approval_targets']
    ia = sorted(k for k in result['approval_targets'] if k.startswith('ia:'))
    assert ia == ['ia:path:val-0001', 'ia:path:val-0002', 'ia:path:val-0003']
    assert [k for k in result['approval_required'] if k.startswith('ds:')] == before['approval_required']
    assert 'ia:path:val-0001' in result['approval_required']
    resolved = result['manifest']['web']['modules']['ia']
    assert resolved['version'] == '0.4.0'
    assert resolved['resolved']['governance']['contract'] == 'ia-git-history-v4'
    assert resolved['resolved']['governance']['schema_version'] == 4
    assert resolved['resolved']['targets'] == {'channels': 2, 'outlines': 2, 'sections': 5}
    assert resolved['resolved']['governance']['method'] == 'git-history'
    assert resolved['resolved']['governance']['input_hashes']
    assert any(row['id'] == 'ia-deterministic' and row['targets'] == 5 for row in result['report']['executed'])
    assert any(row['id'] == 'ia-semantic-judgments' for row in result['report']['missing'])
    assert any(row['id'] == 'ia-dom' for row in result['report']['skipped'])
    validate_module_state(result['manifest']['web']['modules'])


@pytest.mark.parametrize('broken', ['ref', 'expiry', 'zero', 'corrupt', 'symlink'])
def test_actual_ia_bad_inputs_make_l0_red(project, broken):
    root = attach(project)
    base = root / 'design/info_arch'
    if broken == 'ref':
        data = copy.deepcopy(FIXTURE['OUTLINE'])
        data['sections']['sec-0002']['places'] = ['msg-9999']
        FIXTURE['dump'](base / 'outlines/lp-municipal.yaml', data)
    elif broken == 'expiry':
        dictionary = copy.deepcopy(FIXTURE['DICTIONARY'])
        dictionary['entries'][0]['expires'] = '2000-01-01'  # an expired value in the docs dictionary
        FIXTURE['write_dictionary'](root, dictionary)
    elif broken == 'zero':
        for path in (base / 'outlines').iterdir():
            path.unlink()
    elif broken == 'corrupt':
        (base / 'info_arch.yaml').write_text('schema_version: 2\nschema_version: 2\n')
    else:
        path = base / 'info_arch.yaml'
        outside = root.parent / 'outside.yaml'
        outside.write_bytes(path.read_bytes())
        path.unlink()
        path.symlink_to(outside)
    result = check(project)
    assert result['exit_code'] == 2, result
    assert result['errors']
    if broken in ('ref', 'expiry', 'zero'):
        expected = {'ref': 'IA_OUTLINE_REF_BROKEN', 'expiry': 'IA_EXPIRED_SHIPPED', 'zero': 'zero IA check targets'}[broken]
        assert any(expected in error for error in result['errors']), result['errors']


@pytest.mark.parametrize('version', [None, '0.1.0', '0.3.0'])
def test_missing_or_wrong_pin_stays_red(project, version):
    root = attach(project)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['modules'] = {} if version is None else {'ia': version}
    write(root / 'design/design.yaml', config)
    assert check(project)['exit_code'] == 2


def test_python_and_javascript_reject_changed_checker_and_schema(project, tmp_path):
    approve(project)
    root = attach(project)
    installed = tmp_path / 'installed'
    shutil.copytree(SKILL, installed / SKILL.name, ignore=shutil.ignore_patterns('__pycache__'))
    shutil.copytree(IA, installed / IA.name, ignore=shutil.ignore_patterns('__pycache__'))
    # the checker is not only own-infoarch-author: it imports the positioning layer and reads the docs dictionary
    for sibling in ('own-positioning-author', 'own-doc-update'):
        shutil.copytree(SKILL.parent / sibling, installed / sibling, ignore=shutil.ignore_patterns('__pycache__', 'tests'))
    command = [sys.executable, str(installed / SKILL.name / 'scripts/design_author.py'),
               'check', '--repo', str(root), '--catalog-dir', str(project[1])]
    run = subprocess.run(command, capture_output=True, text=True)
    assert run.returncode == 0, run.stdout + run.stderr
    manifest = root / 'design/.resolved/web.json'
    python = [sys.executable, str(installed / SKILL.name / 'scripts/resolved_contract.py'), str(manifest)]
    js = 'import {loadManifest} from ' + json.dumps((installed / SKILL.name / 'scripts/resolved_contract.mjs').as_uri()) + ';loadManifest(process.argv.at(-1));'
    node = [os.environ.get('DESIGN_TEST_NODE', 'node'), '--input-type=module', '-e', js, str(manifest)]
    for consumer in (python, node):
        good = subprocess.run(consumer, capture_output=True, text=True)
        assert good.returncode == 0, good.stdout + good.stderr
    for relative in ('own-infoarch-author/scripts/check.py', 'own-infoarch-author/references/schema/outline.v3.schema.json',
                     'own-positioning-author/scripts/positioning.py', 'own-positioning-author/references/schema/info_arch.v5.schema.json',
                     'own-doc-update/scripts/dictionary.py'):
        path = installed / relative
        saved = path.read_bytes()
        path.write_bytes(saved + b'\n')
        for consumer in (python, node):
            assert subprocess.run(consumer, capture_output=True).returncode != 0
        path.write_bytes(saved)
    # Prose is not code: editing a reference document leaves the resolved result valid (ADR-20261115, biz_ops),
    # while a JSON reference the checker reads still binds.
    for relative in ('references/contract.md', 'references/media/lp.md'):
        path = installed / IA.name / relative
        path.write_bytes(path.read_bytes() + '\n追記した説明。\n'.encode())
        for consumer in (python, node):
            kept = subprocess.run(consumer, capture_output=True, text=True)
            assert kept.returncode == 0, (relative, kept.stdout + kept.stderr)
    hidden = installed / IA.name / 'references/notes.md'  # a directory named like prose still binds its code
    hidden.mkdir()
    (hidden / 'extra.py').write_text('x = 1\n')
    for consumer in (python, node):
        assert subprocess.run(consumer, capture_output=True).returncode != 0
    shutil.rmtree(hidden)
    path = installed / IA.name / 'references/expiry.json'
    saved = path.read_bytes()
    path.write_bytes(saved + b'\n')
    for consumer in (python, node):
        assert subprocess.run(consumer, capture_output=True).returncode != 0
    path.write_bytes(saved)
    # A symlinked trusted dependency is rejected by both consumers, even if bytes match.
    path = installed / IA.name / 'scripts/check.py'
    original = path.read_bytes()
    outside = tmp_path / 'checker.py'
    outside.write_bytes(original)
    path.unlink()
    path.symlink_to(outside)
    for consumer in (python, node):
        assert subprocess.run(consumer, capture_output=True).returncode != 0


def test_legacy_empty_targets_remain_rejected(project, tmp_path, monkeypatch):
    import design_modules
    copied = tmp_path / 'installed/own-design-author/scripts/design_modules.py'
    copied.parent.mkdir(parents=True)
    copied.write_bytes(Path(design_modules.__file__).read_bytes())
    adapter = copied.parents[2] / 'own-infoarch-author/scripts/ia_module.py'
    adapter.parent.mkdir(parents=True)
    adapter.write_text("VERSION='legacy-v1'\ndef resolve(**kw):\n return {'media':{k:{} for k in kw['profiles']},'approval_hashes':{},'must_approval_keys':[],'errors':[],'warnings':[],'report':{k:[] for k in ('executed','skipped','missing','uncovered','screened')}}\n")
    spec = importlib.util.spec_from_file_location('fixture_design_modules', copied)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    with pytest.raises(ValueError, match='zero IA approval targets'):
        module.resolve_ia(project[0], {'modules': {'ia': 'legacy-v1'}}, {}, {'web': {}}, empty_report())


def test_absent_fixed_adapter_is_red(project, tmp_path):
    import design_modules
    copied = tmp_path / 'missing/own-design-author/scripts/design_modules.py'
    copied.parent.mkdir(parents=True)
    copied.write_bytes(Path(design_modules.__file__).read_bytes())
    spec = importlib.util.spec_from_file_location('missing_design_modules', copied)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    report = empty_report()
    with pytest.raises(ValueError, match='IA module is unavailable'):
        module.resolve_ia(project[0], {'modules': {'ia': '0.4.0'}}, {}, {'web': {}}, report)
    assert report['missing'][0]['id'] == 'ia'


def test_unapproved_positioning_can_be_prepared_and_recorded_by_the_owner(project):
    """An unapproved positioning is an approval gap, not a structural error: preparing the bundle must not be blocked
    by the very approval it prepares (WS-20261114 I6, biz_ops)."""
    from owner_approval import approve_request, prepare
    approve(project)
    root = attach(project)
    import yaml
    ledger = root / 'design/approvals.yaml'
    rows = [r for r in yaml.safe_load(ledger.read_text()) if not r['key'].startswith('ia:')]
    ledger.write_text(yaml.safe_dump(rows, allow_unicode=True, sort_keys=False))  # the owner has approved nothing yet
    blocked = check(project, ship=True)
    assert any(e.startswith('IA_POSITIONING_UNAPPROVED') for e in blocked['errors']), blocked['errors']
    request = prepare(root, project[1], all_must=True)
    keys = {r['key'] for r in request['rows']}
    assert 'ia:path:val-0001' in keys and 'ia:path:val-0003' in keys
    approve_request(root, project[1], request, 'fixture-owner', owner_words='fixture: approve',
                    approved_on=request['rows'][0]['date'])
    assert not any(e.startswith('IA_POSITIONING_UNAPPROVED') for e in check(project, ship=True)['errors'])


def test_changing_one_positioning_element_withdraws_only_the_paths_it_is_on(project):
    """Per-path approval through the design bridge; ds: approvals and the other paths' rows stay."""
    approve(project)
    before = check(project, ship=True)
    root = attach(project)
    approved = check(project, ship=True)
    assert approved['exit_code'] == 0, approved['errors']
    ssot = copy.deepcopy(FIXTURE['SSOT'])
    ssot['proofs']['prf-0001']['as_of'] = '2099-01-01'
    ssot['strengths']['str-0001']['text'] = '自動更新で保守がほぼ要らない'
    FIXTURE['dump'](root / 'design/info_arch/info_arch.yaml', ssot)
    after = check(project, ship=True)
    moved = sorted(k for k, v in approved['approval_targets'].items() if after['approval_targets'].get(k) != v)
    assert moved == ['ia:path:val-0001'], moved  # str-0001 produces val-0001 only
    assert {k: v for k, v in after['approval_targets'].items() if k.startswith('ds:')} == before['approval_targets']
    assert any(e.startswith('IA_POSITIONING_UNAPPROVED') and 'val-0001' in e for e in after['errors']), after['errors']


def test_legacy_per_element_rows_are_refused_until_migrated(project):
    """An ia:positioning row is no target any more: the design check stops on it instead of reading it (fail closed),
    and migrate-approvals is the way out. The ds: approvals are untouched by either."""
    import yaml
    approve(project)
    root = attach(project)
    ledger = root / 'design/approvals.yaml'
    rows = yaml.safe_load(ledger.read_text())
    rows.append({'key': 'ia:positioning:val-0001', 'value_hash': 'b' * 64, 'approver': 'fixture-owner', 'date': '2026-11-16'})
    ledger.write_text(yaml.safe_dump(rows, allow_unicode=True, sort_keys=False))
    result = check(project, ship=True)
    assert result['exit_code'] == 2 and result['errors'] == ['approval target absent from catalog/tokens'], result['errors']


# ---- the docs dictionary is an input of L0 (WS-20261120, ISSUE-20261121-design-author-input-hashes-skip-dictionary) ----
# The fixed values and official names the IA checker compares against live in docs/dictionary/<domain>.json, outside
# design/. A word changed there must change the L0 input hashes, or a resolved manifest built before the change
# keeps passing every consumer's stale check.

def dictionary_path(root):
    return root / 'docs/dictionary' / f'{FIXTURE["DICTIONARY"]["domain"]}.json'


def ia_input_hashes(result):
    return result['manifest']['web']['modules']['ia']['resolved']['governance']['input_hashes']


def edit_dictionary_word(root):
    path = dictionary_path(root)
    data = json.loads(path.read_text(encoding='utf-8'))
    data['entries'][0]['meaning'] += '（言い換えた）'
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def consumers(manifest):
    python = [sys.executable, str(SKILL / 'scripts/resolved_contract.py'), str(manifest)]
    js = 'import {loadManifest} from ' + json.dumps((SKILL / 'scripts/resolved_contract.mjs').as_uri()) + ';loadManifest(process.argv.at(-1));'
    node = [os.environ.get('DESIGN_TEST_NODE', 'node'), '--input-type=module', '-e', js, str(manifest)]
    return python, node


def test_dictionary_word_change_changes_l0_input_hashes(project):
    approve(project)
    root = attach(project)
    before = ia_input_hashes(check(project, ship=True))
    rel = dictionary_path(root).relative_to(root).as_posix()
    assert rel in before, sorted(before)
    assert any(key.startswith('design/info_arch/') for key in before)
    assert check(project, ship=True)['exit_code'] == 0
    assert ia_input_hashes(check(project, ship=True)) == before  # same inputs, same hashes
    edit_dictionary_word(root)
    result = check(project, ship=True)
    assert result['exit_code'] == 0, result['errors']
    after = ia_input_hashes(result)
    assert after[rel] != before[rel]
    assert {k: v for k, v in after.items() if k != rel} == {k: v for k, v in before.items() if k != rel}


def test_resolved_manifest_is_stale_for_both_consumers_after_a_dictionary_change(project):
    approve(project)
    root = attach(project)
    command = [sys.executable, str(SKILL / 'scripts/design_author.py'), 'check', '--repo', str(root), '--catalog-dir', str(project[1])]
    run = subprocess.run(command, capture_output=True, text=True)
    assert run.returncode == 0, run.stdout + run.stderr
    manifest = root / 'design/.resolved/web.json'
    for consumer in consumers(manifest):
        good = subprocess.run(consumer, capture_output=True, text=True)
        assert good.returncode == 0, good.stdout + good.stderr
    path = dictionary_path(root)
    saved = path.read_bytes()
    edit_dictionary_word(root)
    for consumer in consumers(manifest):
        stale = subprocess.run(consumer, capture_output=True, text=True)
        assert stale.returncode != 0, stale.stdout + stale.stderr
    from resolved_contract import load_manifest
    with pytest.raises(ValueError, match='stale manifest'):
        load_manifest(manifest)
    path.write_bytes(saved)  # put the word back: the manifest is valid again, so the stale check is not a one-way latch
    for consumer in consumers(manifest):
        assert subprocess.run(consumer, capture_output=True, text=True).returncode == 0
    # a dictionary file that is added or removed after L0 is a change too
    extra = root / 'docs/dictionary/other.json'
    extra.write_text(json.dumps({'schema_version': 1, 'domain': 'other', 'summary': 's', 'entries': []}), encoding='utf-8')
    for consumer in consumers(manifest):
        assert subprocess.run(consumer, capture_output=True, text=True).returncode != 0
    extra.unlink()
    path.unlink()
    for consumer in consumers(manifest):
        assert subprocess.run(consumer, capture_output=True, text=True).returncode != 0


def test_repository_without_a_dictionary_has_no_dictionary_inputs_and_stays_valid(project):
    approve(project)
    root = project[0]
    ssot = copy.deepcopy(FIXTURE['SSOT'])
    ssot['proofs']['prf-0001']['as_of'] = '2099-01-01'
    FIXTURE['write_repo'](root, ssot, dictionary=None)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['modules'] = {'ia': '0.4.0'}
    write(root / 'design/design.yaml', config)
    assert not (root / 'docs/dictionary').exists()
    command = [sys.executable, str(SKILL / 'scripts/design_author.py'), 'check', '--repo', str(root), '--catalog-dir', str(project[1])]
    run = subprocess.run(command, capture_output=True, text=True)
    data = json.loads(run.stdout)
    manifest = root / 'design/.resolved/web.json'
    if manifest.exists():  # an IA tree that needs a dictionary may be red; what matters is no dictionary key appears
        assert not any(k.startswith('docs/') for k in ia_input_hashes({'manifest': {'web': json.loads(manifest.read_text())}}))
        for consumer in consumers(manifest):
            assert subprocess.run(consumer, capture_output=True, text=True).returncode == 0
    else:
        assert data['exit_code'] == 2, data


@pytest.mark.parametrize('key', ['docs/dictionary/sub/x.json', 'docs/dictionary/../x.json', 'docs/other.json',
                                 'docs/dictionary/x.txt', 'docs/dictionary/', '/docs/dictionary/x.json', 'design/info_arch/../../x'])
def test_governance_input_hash_keys_stay_confined_to_ia_inputs_and_the_dictionary(project, key):
    import design_modules
    root = attach(project)
    resolved = resolve_ia(root, json.loads((root / 'design/design.yaml').read_text()), {}, {'web': {}}, empty_report())
    resolved.pop('_implementation_sha256')
    design_modules._git_contract(resolved)  # the real result passes

    def with_key(name):
        row = copy.deepcopy(resolved)
        row['governance']['input_hashes'][name] = '0' * 64
        row['media'] = {m: {'governance': row['governance'], 'targets': row['targets']} for m in row['media']}
        return row

    design_modules._git_contract(with_key('docs/dictionary/another-domain.json'))  # a dictionary file is an input
    with pytest.raises(ValueError, match='invalid IA input hash'):
        design_modules._git_contract(with_key(key))


def test_dictionary_folder_is_named_the_same_by_the_reader_the_bridge_and_the_consumers():
    import design_modules
    loaded = {}
    for name, path in (('dictionary', SKILL.parent / 'own-doc-update/scripts/dictionary.py'), ('ia_module', IA / 'scripts/ia_module.py')):
        spec = importlib.util.spec_from_file_location('pin_' + name, path)
        loaded[name] = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(loaded[name])
    assert loaded['dictionary'].DICTIONARY_DIR.as_posix() == loaded['ia_module'].DICTIONARY_DIR == design_modules.IA_DICTIONARY_DIR
    assert "['docs', 'dictionary']" in (SKILL / 'scripts/resolved_contract.mjs').read_text()
