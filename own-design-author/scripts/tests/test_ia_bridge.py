"""Actual schema-2 IA checker bridge: git history, no IA approval rows, stale consumers."""
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
    ssot['facts']['fct-0001']['as_of'] = '2099-01-01'
    FIXTURE['write_repo'](root, ssot)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['modules'] = {'ia': '0.2.0'}
    write(root / 'design/design.yaml', config)
    return root


def test_schema2_actual_check_has_counts_and_preserves_ds_approvals(project):
    approve(project)
    before = check(project, ship=True)
    root = attach(project)
    result = check(project, ship=True)
    assert result['exit_code'] == 0, result['errors']
    assert result['approval_targets'] == before['approval_targets']
    assert result['approval_required'] == before['approval_required']
    resolved = result['manifest']['web']['modules']['ia']
    assert resolved['version'] == '0.2.0'
    assert resolved['resolved']['targets'] == {'channels': 2, 'outlines': 2, 'sections': 6}
    assert resolved['resolved']['governance']['method'] == 'git-history'
    assert resolved['resolved']['governance']['input_hashes']
    assert any(row['id'] == 'ia-deterministic' and row['targets'] == 6 for row in result['report']['executed'])
    assert any(row['id'] == 'ia-semantic-judgments' for row in result['report']['missing'])
    assert any(row['id'] == 'ia-dom' for row in result['report']['skipped'])
    validate_module_state(result['manifest']['web']['modules'])


@pytest.mark.parametrize('broken', ['ref', 'expiry', 'zero', 'corrupt', 'symlink'])
def test_actual_ia_bad_inputs_make_l0_red(project, broken):
    root = attach(project)
    base = root / 'design/info_arch'
    if broken == 'ref':
        data = copy.deepcopy(FIXTURE['OUTLINE'])
        data['pages'][0]['sections'][1]['message'] = 'msg-9999'
        FIXTURE['dump'](base / 'outlines/lp-municipal.yaml', data)
    elif broken == 'expiry':
        data = copy.deepcopy(FIXTURE['SSOT'])
        data['facts']['fct-0001']['expires'] = '2000-01-01'
        FIXTURE['dump'](base / 'info_arch.yaml', data)
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


@pytest.mark.parametrize('version', [None, '0.1.0'])
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
    for relative in ('scripts/check.py', 'references/schema/info_arch.v2.schema.json'):
        path = installed / IA.name / relative
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
        module.resolve_ia(project[0], {'modules': {'ia': '0.2.0'}}, {}, {'web': {}}, report)
    assert report['missing'][0]['id'] == 'ia'
