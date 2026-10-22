"""Behavioral L0/CLI tests; fixture catalogs are synthetic, never upstream classifications."""
import copy
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from design_author import check_design, init_design, migrate_design, value_hash
from dtcg import resolve_tokens
from classify import load_policy, classify_response

SKILL = Path(__file__).resolve().parents[2]


def write(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n')


def item(ident='dads-000001', grade='must'):
    row = {'id': ident, 'match_key': ident, 'document': 'fixture', 'title': ident,
           'headings': [ident], 'text': 'A synthetic fixture policy.', 'source_url': 'https://example.org/',
           'evidence_types': ['token', 'doc', 'rule']}
    policy = load_policy()
    choice = 'global1-lp0-product0-presentation0' if grade == 'must' else 'global0-lp0-product0-presentation0'
    answer = {'type':'choice', 'choice':choice, 'confidence':1.0,
              'probabilities':{k:float(k == choice) for k in policy['outcomes']}}
    row['classification'] = classify_response(row,answer,policy['model'],policy)
    return row


@pytest.fixture
def project(tmp_path):
    root = tmp_path / 'repo'
    root.mkdir()
    subprocess.run(['git', 'init', '-q', str(root)], check=True)
    catalogs = tmp_path / 'catalogs'
    write(catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': [item()]})
    write(catalogs / 'id-ledger.json', {'schema_version': 1, 'next_id': 2,
        'entries': [{'id': 'dads-000001', 'match_key': 'dads-000001', 'retired': False}]})
    init_design(root, 'product-ui', 'test-v1', catalog_dir=catalogs)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['items'] = {'dads-000001': {'state': 'defined', 'evidence': ['token:color.primary']}}
    write(root / 'design/design.yaml', config)
    write(root / 'design/tokens/brand.tokens.json', {'color': {'$type': 'color', 'primary': {
        '$value': {'colorSpace': 'srgb', 'components': [0, 0, 0], 'hex': '#000000'},
        '$extensions': {'org.own.design': {'origin': 'proposal'}}}}})
    return root, catalogs


def check(project, **kw):
    root, catalogs = project
    return check_design(root, catalog_dir=catalogs, **kw)


def approve(project):
    root, _ = project
    result = check(project)
    value = result['manifest']['web']['tokens']['color.primary']['value_hash']
    # Human-authored fixture ledger, not a production approval-writing operation.
    write(root / 'design/approvals.yaml', [{'token_path': 'color.primary', 'value_hash': value,
                                          'approver': 'fixture-owner', 'date': '2026-10-21'}])


def test_proposed_warning_and_ship_rejection(project):
    result = check(project)
    assert result['exit_code'] == 1
    assert result['counts']['proposed'] == 1
    assert check(project, ship=True)['exit_code'] == 2
    assert not (project[0] / 'design/.resolved/web.json').exists()


def test_approved_green_then_changed_value_is_red(project):
    approve(project)
    assert check(project, ship=True)['exit_code'] == 0
    p = project[0] / 'design/tokens/brand.tokens.json'
    data = json.loads(p.read_text())
    data['color']['primary']['$value']['components'] = [1, 1, 1]
    data['color']['primary']['$value']['hex'] = '#ffffff'
    write(p, data)
    result = check(project)
    assert result['exit_code'] == 2
    assert any('approval' in e for e in result['errors'])


@pytest.mark.parametrize('case', ['must-empty', 'approval-mismatch', 'relaxed-minimum', 'zero-target', 'unknown-id'])
def test_five_required_red_cases(project, case):
    root, catalogs = project
    config = json.loads((root / 'design/design.yaml').read_text())
    if case == 'must-empty':
        config['items'] = {}
    elif case == 'approval-mismatch':
        write(root / 'design/approvals.yaml', [{'token_path': 'color.primary', 'value_hash': 'bad',
                                               'approver': 'fixture-owner', 'date': '2026-10-21'}])
    elif case == 'relaxed-minimum':
        write(root / 'design/profiles/web.json', {'media': 'web', 'minimums': {'text_contrast': 1}})
    elif case == 'zero-target':
        write(catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': []})
    else:
        config['items']['unknown-id'] = {'state': 'n-a', 'reason': 'not applicable'}
    write(root / 'design/design.yaml', config)
    result = check(project)
    assert result['exit_code'] == 2, result
    assert result['errors']
    run = subprocess.run([sys.executable, str(SKILL / 'scripts/design_author.py'), 'check',
        '--repo', str(root), '--catalog-dir', str(catalogs)], cwd=root, capture_output=True, text=True)
    assert run.returncode == 2, run.stdout + run.stderr
    assert json.loads(run.stdout)['exit_code'] == 2


def test_cli_real_git_root_and_five_cases(project):
    root, catalogs = project
    cmd = [sys.executable, str(SKILL / 'scripts/design_author.py'), 'check', '--repo', str(root),
           '--catalog-dir', str(catalogs)]
    run = subprocess.run(cmd, cwd=root, capture_output=True, text=True)
    assert run.returncode == 1, run.stderr + run.stdout
    data = json.loads(run.stdout)
    assert data['counts']['items'] == 1
    assert (root / 'design/.resolved/web.json').exists()
    (root / 'design/design.yaml').write_text('{"catalog_version":"missing"}')
    run = subprocess.run(cmd, cwd=root, capture_output=True, text=True)
    assert run.returncode == 2
    assert not (root / 'design/.resolved/web.json').exists()


def test_unknown_version_unclassified_and_unknown_preset_fail(project):
    root, catalogs = project
    p = catalogs / 'test-v1.json'
    data = json.loads(p.read_text())
    data['items'][0]['classification'] = None
    write(p, data)
    result = check(project)
    assert result['exit_code'] == 2
    assert result['counts']['unclassified'] == 1
    with pytest.raises(ValueError):
        init_design(root.parent / 'other', 'unknown', 'test-v1', catalog_dir=catalogs)


def test_defined_requires_real_evidence_and_doc_body(project):
    root, _ = project
    p = root / 'design/design.yaml'
    config = json.loads(p.read_text())
    config['items']['dads-000001']['evidence'] = ['doc:DESIGN.md#missing']
    write(p, config)
    assert check(project)['exit_code'] == 2
    config['items']['dads-000001']['evidence'] = ['doc:DESIGN.md#dads-000001']
    write(p, config)
    (root / 'design/DESIGN.md').write_text('# Brand\n\n## dads-000001\n\n')
    assert check(project)['exit_code'] == 2
    (root / 'design/DESIGN.md').write_text('# Brand\n\n## dads-000001\nUse token:color.primary for primary actions.\n')
    assert check(project)['exit_code'] == 1


def test_must_na_needs_reason_and_hash_bound_approval(project):
    root, _ = project
    p = root / 'design/design.yaml'
    config = json.loads(p.read_text())
    config['items']['dads-000001'] = {'state': 'n-a', 'reason': 'No interactive surface.'}
    write(p, config)
    assert check(project)['exit_code'] == 2
    write(root / 'design/approvals.yaml', [{'item_id': 'dads-000001',
        'value_hash': value_hash({'state': 'n-a', 'reason': 'No interactive surface.'}),
        'approver': 'fixture-owner', 'date': '2026-10-21'}])
    assert check(project)['exit_code'] == 1  # remaining token is still proposed
    config['items']['dads-000001']['reason'] = 'Changed reason'
    write(p, config)
    assert check(project)['exit_code'] == 2


def test_origin_reference_requires_notice(project):
    root, _ = project
    p = root / 'design/tokens/brand.tokens.json'
    data = json.loads(p.read_text())
    data['color']['primary']['$extensions']['org.own.design']['origin'] = 'reference:example'
    write(p, data)
    assert check(project)['exit_code'] == 2
    write(root / 'design/NOTICE', {'sources': [{'id': 'example', 'url': 'https://example.org/',
        'commit': 'v1', 'license': 'MIT', 'copyright': 'Example contributors', 'edited': True}]})
    assert check(project)['exit_code'] == 1


def test_rule_uncovered_is_visible_and_ship_fails(project):
    root, _ = project
    config = json.loads((root / 'design/design.yaml').read_text())
    config['items']['dads-000001']['evidence'] = ['rule:contrast-text']
    write(root / 'design/design.yaml', config)
    write(root / 'design/checks/web.yaml', {'rules': [{'id': 'contrast-text', 'kind': 'contrast',
        'targets': ['body'], 'minimum': 4.5}]})
    assert check(project)['counts']['uncovered'] == 1
    approve(project)
    assert check(project, ship=True)['exit_code'] == 2
    write(root / 'design/checks/web.yaml', {'rules': [{'id': 'evil', 'command': 'echo bypass'}]})
    assert check(project)['exit_code'] == 2


def test_duplicate_yaml_key_path_escape_and_duplicate_approval_fail(project):
    root, _ = project
    original = (root / 'design/design.yaml').read_text()
    (root / 'design/design.yaml').write_text('catalog_version: test-v1\ncatalog_version: test-v1\n')
    assert check(project)['exit_code'] == 2
    (root / 'design/design.yaml').write_text(original)
    config = json.loads(original)
    config['items']['dads-000001']['evidence'] = ['doc:../../secret#entry']
    write(root / 'design/design.yaml', config)
    assert check(project)['exit_code'] == 2
    write(root / 'design/design.yaml', json.loads(original))
    approve(project)
    rows = json.loads((root / 'design/approvals.yaml').read_text())
    write(root / 'design/approvals.yaml', rows * 2)
    assert check(project)['exit_code'] == 2


def test_migrate_only_adds_deferred_and_keeps_old_until_command(project):
    root, catalogs = project
    old = (root / 'design/design.yaml').read_text()
    write(catalogs / 'test-v2.json', {'schema_version': 1, 'version': 'test-v2',
        'items': [item(), item('dads-000002')]})
    ledger = json.loads((catalogs / 'id-ledger.json').read_text())
    ledger['entries'].append({'id': 'dads-000002', 'match_key': 'dads-000002', 'retired': False})
    ledger['next_id'] = 3
    write(catalogs / 'id-ledger.json', ledger)
    assert check(project)['exit_code'] == 1
    assert (root / 'design/design.yaml').read_text() == old
    result = migrate_design(root, 'test-v2', catalog_dir=catalogs)
    assert result['added'] == ['dads-000002']
    assert check(project)['exit_code'] == 2
    assert json.loads((root / 'design/design.yaml').read_text())['items']['dads-000001']['state'] == 'defined'
    with pytest.raises(ValueError):
        init_design(root, 'lp', 'test-v2', catalog_dir=catalogs)


def test_asset_manifest_has_content_hash_and_no_absolute_paths(project):
    root, _ = project
    (root / 'design/assets').mkdir(exist_ok=True)
    (root / 'design/assets/logo.svg').write_text('<svg/>')
    assert check(project)['exit_code'] == 1
    assets = json.loads((root / 'design/assets.json').read_text())
    assert assets['assets'][0]['path'] == 'assets/logo.svg'
    assert len(assets['assets'][0]['sha256']) == 64


def test_dtcg_alias_units_css_and_schema_failure():
    tokens = {'space': {'$type': 'dimension', 'base': {'$value': {'value': 1.5, 'unit': 'rem'}},
                      'alias': {'$value': '{space.base}'}}}
    result = resolve_tokens([tokens])
    assert result['space.alias']['css'] == '1.5rem'
    assert result['space.alias']['value'] == result['space.base']['value']
    with pytest.raises(ValueError):
        resolve_tokens([{'bad': {'$type': 'dimension', '$value': {'value': 'wrong', 'unit': 'px'}}}])
    with pytest.raises(ValueError):
        resolve_tokens([{'a': {'$type': 'number', '$value': '{b}'}, 'b': {'$value': '{a}'}}])
    with pytest.raises(ValueError):
        resolve_tokens([{'a': {'$type': 'color', '$value': {'colorSpace': 'srgb', 'components': [0, 0, 0]}}}])


def test_alias_hash_changes_with_target_and_css_names_do_not_collide():
    data = {'color': {'base': {'$type': 'number', '$value': 1},
                      'alias': {'$type': 'number', '$value': '{color.base}'}}}
    first = resolve_tokens([data])
    data['color']['base']['$value'] = 2
    second = resolve_tokens([data])
    assert first['color.alias']['value_hash'] != second['color.alias']['value_hash']
    with pytest.raises(ValueError):
        resolve_tokens([{'a-b': {'$type': 'number', '$value': 1}, 'a': {'b': {'$type': 'number', '$value': 2}}}])


def test_nice_na_approval_is_still_hash_bound(project):
    root, catalogs = project
    catalog = json.loads((catalogs / 'test-v1.json').read_text())
    catalog['items'][0] = item(grade='nice-to-have')
    write(catalogs / 'test-v1.json', catalog)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['items']['dads-000001'] = {'state':'n-a','reason':'No such surface.'}
    write(root / 'design/design.yaml',config)
    write(root / 'design/approvals.yaml',[{'item_id':'dads-000001','value_hash':'wrong',
        'approver':'fixture-owner','date':'2026-10-21'}])
    assert check(project)['exit_code'] == 2


def test_fake_markdown_anchor_inside_code_fence_is_not_evidence(project):
    root, _ = project
    config = json.loads((root / 'design/design.yaml').read_text())
    config['items']['dads-000001']['evidence'] = ['doc:DESIGN.md#dads-000001']
    write(root / 'design/design.yaml',config)
    (root / 'design/DESIGN.md').write_text('# Brand\n```md\n## dads-000001\nUse token:color.primary.\n```\n')
    assert check(project)['exit_code'] == 2


def test_srgb_fallback_cannot_silently_disagree_with_components():
    with pytest.raises(ValueError):
        resolve_tokens([{'color':{'$type':'color','$value':{'colorSpace':'srgb','components':[0,0,0],'hex':'#ffffff'}}}])


def test_catalog_classification_stale_input_is_rejected(project):
    _, catalogs = project
    p = catalogs / 'test-v1.json'
    data = json.loads(p.read_text())
    data['items'][0]['text'] += ' Materially changed.'
    write(p, data)
    assert check(project)['exit_code'] == 2


def test_generic_font_fallback_remains_a_css_keyword():
    resolved = resolve_tokens([{'font':{'$type':'fontFamily','$value':['Noto Sans JP','sans-serif']}}])
    assert resolved['font']['css'] == '"Noto Sans JP", sans-serif'


def test_removed_unresolved_must_remains_a_gate_after_migration(project):
    root,catalogs=project
    old=json.loads((catalogs/'test-v1.json').read_text())
    old['items'].append(item('dads-000002'))
    write(catalogs/'test-v1.json',old)
    ledger=json.loads((catalogs/'id-ledger.json').read_text())
    ledger['entries'].append({'id':'dads-000002','match_key':'dads-000002','retired':True})
    ledger['next_id']=3
    write(catalogs/'id-ledger.json',ledger)
    assert check(project)['exit_code']==2
    write(catalogs/'test-v2.json',{'schema_version':1,'version':'test-v2','items':[item()]})
    migrate_design(root,'test-v2',catalog_dir=catalogs)
    assert check(project)['exit_code']==2
    config=json.loads((root/'design/design.yaml').read_text())
    assert config['retired_items']['dads-000002']['state']=='deferred'
    write(catalogs/'test-v3.json',{'schema_version':1,'version':'test-v3','items':old['items']})
    migrate_design(root,'test-v3',catalog_dir=catalogs)
    restored=json.loads((root/'design/design.yaml').read_text())
    assert not restored.get('retired_items')
    assert restored['items']['dads-000002']['state']=='deferred'


def test_failed_check_invalidates_asset_index_and_unlinks_derived_symlink(project):
    root,_=project
    assert check(project)['exit_code']==1
    assert (root/'design/assets.json').exists()
    resolved=root/'design/.resolved'
    old=root/'previous-resolved'
    resolved.rename(old)
    resolved.symlink_to(old,target_is_directory=True)
    result=check(project)
    assert result['exit_code']==2
    assert not (root/'design/assets.json').exists()
    assert not resolved.exists() and not resolved.is_symlink()
    assert (old/'web.json').is_file()  # No deletion through a link to another directory.


def test_export_failure_leaves_no_partial_generated_contract(project,monkeypatch):
    root,_=project
    original=Path.write_text
    def fail_css(path,*args,**kwargs):
        if path.suffix=='.css':
            raise OSError('fixture CSS export failure')
        return original(path,*args,**kwargs)
    monkeypatch.setattr(Path,'write_text',fail_css)
    assert check(project)['exit_code']==2
    assert not (root/'design/assets.json').exists()
    assert not (root/'design/.resolved/web.json').exists()


def test_fresh_init_reports_empty_must_even_before_first_token(project):
    root,catalogs=project
    import shutil
    shutil.rmtree(root/'design')
    init_design(root,'product-ui','test-v1',catalog_dir=catalogs)
    result=check(project)
    assert result['exit_code']==2
    assert result['counts']['deferred']==1
    assert any('must item is empty' in e for e in result['errors'])
