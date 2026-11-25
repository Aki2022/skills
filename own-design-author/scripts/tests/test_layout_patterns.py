"""The layout pattern vocabulary (SPEC-design-system-standardization requirement 15) and the comp size (requirement 16).

- The skill holds a versioned catalog of media-neutral layout patterns; a repository's design.yaml holds only a diff
  (`layout_patterns: {disabled: [...], added: {...}}`). An added pattern follows the item lifecycle: its approval is
  the owner's, recorded in the ledger and bound by hash (never a state the file declares); proposed warns, and fails
  under --ship.
- The comp size per media comes from the skill's default; a repository profile may only override it. It is resolved
  into the manifest and never moves an approval (it is not a brand value).
- The build records (ux/build/: the layout plan and the git-ignored comp images) are not design inputs: writing them
  never makes the manifest stale (SPEC-deliverable-production-layers requirement 22).
"""
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import layer_paths as lp  # noqa: E402
import layout_patterns  # noqa: E402
from design_author import check_design, init_design  # noqa: E402
from classify import load_policy, classify_response  # noqa: E402

SKILL = Path(__file__).resolve().parents[2]
FUNCTIONS = {'identify', 'claim', 'fact', 'evidence', 'offer', 'action', 'source'}


def write(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n')


def _item(ident='dads-000001'):
    row = {'id': ident, 'match_key': ident, 'document': 'fixture', 'title': ident, 'headings': [ident],
           'text': 'A synthetic fixture policy.', 'source_url': 'https://example.org/',
           'evidence_types': ['token', 'doc', 'rule']}
    policy = load_policy()
    choice = 'global1-lp0-product0-presentation0'
    answer = {'type': 'choice', 'choice': choice, 'confidence': 1.0,
              'probabilities': {k: float(k == choice) for k in policy['outcomes']}}
    row['classification'] = classify_response(row, answer, policy['model'], policy)
    return row


@pytest.fixture
def project(tmp_path):
    root = tmp_path / 'repo'
    root.mkdir()
    subprocess.run(['git', 'init', '-q', str(root)], check=True)
    catalogs = tmp_path / 'catalogs'
    write(catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': [_item()]})
    write(catalogs / 'id-ledger.json', {'schema_version': 1, 'next_id': 2,
        'entries': [{'id': 'dads-000001', 'match_key': 'dads-000001', 'retired': False}]})
    init_design(root, 'lp', 'test-v1', catalog_dir=catalogs)
    config = json.loads(lp.path(root, 'design_config').read_text())
    config['items'] = {'dads-000001': {'state': 'defined', 'evidence': ['token:color.primary']}}
    write(lp.path(root, 'design_config'), config)
    write(lp.path(root, 'design_dir', 'tokens/brand.tokens.json'), {'color': {'$type': 'color', 'primary': {
        '$value': {'colorSpace': 'srgb', 'components': [0, 0, 0], 'hex': '#000000'},
        '$extensions': {'org.own.design': {'origin': 'proposal'}}}}})
    return root, catalogs


def check(project, **kw):
    root, catalogs = project
    return check_design(root, catalog_dir=catalogs, **kw)


def set_patterns(project, value):
    root, _ = project
    config = json.loads(lp.path(root, 'design_config').read_text())
    config['layout_patterns'] = value
    write(lp.path(root, 'design_config'), config)


def approve(project, *keys):
    root, _ = project
    targets = check(project)['approval_targets']
    rows = [{'key': 'ds:web:token:color.primary', 'value_hash': targets['ds:web:token:color.primary']}]
    rows += [{'key': k, 'value_hash': targets[k]} for k in keys]
    write(lp.path(root, 'ledger'), [{**r, 'approver': 'fixture-owner', 'date': '2026-11-25'} for r in rows])


def added(**extra):
    row = {'description': '左に大きな数字、右に短い説明を置く', 'columns': 2, 'band': False, 'images': ['none'],
           'separators': ['rule'], 'narrow': 'stack-1', 'functions': ['fact']}
    row.update(extra)
    return row


def codes(result):
    return ' '.join(result['errors'] + result['warnings'])


# --- the catalog ---------------------------------------------------------------------------------------------------

def test_the_catalog_is_versioned_and_every_pattern_has_every_field():
    catalog = layout_patterns.load_catalog()
    assert catalog['schema_version'] == 1 and isinstance(catalog['version'], str) and catalog['version']
    patterns = catalog['patterns']
    assert 10 <= len(patterns) <= 16
    for name, row in patterns.items():
        assert layout_patterns.validate_pattern(name, row) == [], name
        assert set(row['functions']) <= FUNCTIONS, name
        assert row['narrow'] in patterns, name  # the narrow default is itself a name of the vocabulary
    covered = {f for row in patterns.values() for f in row['functions']}
    assert covered == FUNCTIONS  # an LP's blocks always have a suited pattern


def test_the_catalog_names_are_media_neutral():
    for name in layout_patterns.load_catalog()['patterns']:
        for word in ('web', 'html', 'pptx', 'slide', 'px', 'mobile', 'desktop', 'flex', 'grid-template'):
            assert word not in name, name


def test_validate_pattern_names_each_broken_field():
    assert layout_patterns.validate_pattern('x', {**added(), 'columns': 0})
    assert layout_patterns.validate_pattern('x', {**added(), 'band': 'yes'})
    assert layout_patterns.validate_pattern('x', {**added(), 'functions': ['hero']})
    assert layout_patterns.validate_pattern('x', {k: v for k, v in added().items() if k != 'narrow'})
    assert layout_patterns.validate_pattern('x', {**added(), 'extra': 1})
    assert layout_patterns.validate_pattern('Bad Name', added())


# --- resolution into the manifest ----------------------------------------------------------------------------------

def test_without_a_diff_the_manifest_holds_the_catalog_vocabulary(project):
    result = check(project)
    assert result['exit_code'] == 1, result['errors']
    resolved = result['manifest']['web']['layout_patterns']
    catalog = layout_patterns.load_catalog()
    assert resolved['catalog_version'] == catalog['version']
    assert set(resolved['patterns']) == set(catalog['patterns'])
    assert all(row['origin'] == 'catalog' for row in resolved['patterns'].values())


def test_disabling_a_known_pattern_removes_it(project):
    set_patterns(project, {'disabled': ['quotes']})
    result = check(project)
    assert result['exit_code'] == 1, result['errors']
    assert 'quotes' not in result['manifest']['web']['layout_patterns']['patterns']


def test_disabling_an_unknown_name_fails(project):
    set_patterns(project, {'disabled': ['no-such-pattern']})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_LAYOUT_PATTERN_UNKNOWN_DISABLED' in codes(result)
    assert not lp.path(project[0], 'resolved', 'web.json').exists()


def test_an_unknown_key_in_the_diff_fails(project):
    set_patterns(project, {'removed': ['quotes']})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_LAYOUT_PATTERNS_INVALID' in codes(result)


def test_an_added_pattern_missing_a_field_fails(project):
    row = added()
    del row['functions']
    set_patterns(project, {'added': {'number-split': row}})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_LAYOUT_PATTERN_FIELDS' in codes(result)


def test_an_added_pattern_may_not_declare_its_own_state(project):
    """Approval comes from the ledger, never from a flag the file sets (the AI never sets approved)."""
    set_patterns(project, {'added': {'number-split': added(state='approved')}})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_LAYOUT_PATTERN_FIELDS' in codes(result)


def test_an_added_name_that_the_catalog_already_has_fails(project):
    set_patterns(project, {'added': {'stack-1': added()}})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_LAYOUT_PATTERN_NAME_TAKEN' in codes(result)


def test_a_narrow_default_outside_the_vocabulary_fails(project):
    set_patterns(project, {'disabled': ['stack-1'], 'added': {'number-split': added()}})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_LAYOUT_PATTERN_NARROW_UNKNOWN' in codes(result)


def test_a_proposed_added_pattern_warns_and_fails_the_ship_gate_until_the_owner_approves(project):
    set_patterns(project, {'added': {'number-split': added()}})
    result = check(project)
    assert result['exit_code'] == 1 and 'DS_LAYOUT_PATTERN_PROPOSED' in codes(result)
    row = result['manifest']['web']['layout_patterns']['patterns']['number-split']
    assert row['origin'] == 'repo' and row['approval_state'] == 'proposed'
    key = 'ds:web:pattern:number-split'
    assert key in result['approval_targets']
    shipped = check(project, ship=True)
    assert shipped['exit_code'] == 2 and 'DS_LAYOUT_PATTERN_PROPOSED' in codes(shipped)
    approve(project, key)
    result = check(project, ship=True)
    assert result['exit_code'] == 0, result['errors'] + result['warnings']
    assert result['manifest']['web']['layout_patterns']['patterns']['number-split']['approval_state'] == 'approved'


def test_changing_an_approved_added_pattern_drops_the_approval(project):
    set_patterns(project, {'added': {'number-split': added()}})
    approve(project, 'ds:web:pattern:number-split')
    assert check(project, ship=True)['exit_code'] == 0
    set_patterns(project, {'added': {'number-split': added(columns=3)}})
    result = check(project)
    assert result['exit_code'] == 2 and 'approval hash mismatch' in codes(result)


# --- the comp size --------------------------------------------------------------------------------------------------

def test_the_comp_size_defaults_from_the_skill(project):
    result = check(project)
    defaults = json.loads((SKILL / 'references/comp-sizes.json').read_text())
    assert result['manifest']['web']['comp'] == {**defaults['web'], 'source': 'skill'}
    assert defaults['web']['width'] == 1280


def test_a_profile_may_override_the_comp_size_without_moving_any_approval(project):
    approve(project)
    before = check(project, ship=True)
    assert before['exit_code'] == 0
    root, _ = project
    write(lp.path(root, 'design_dir', 'profiles/web.json'), {'schema_version': 1, 'media': 'web', 'minimums': {},
                                                             'comp': {'width': 1440}})
    after = check(project, ship=True)
    assert after['exit_code'] == 0, after['errors']
    assert after['manifest']['web']['comp'] == {'width': 1440, 'source': 'profile'}
    assert after['approval_targets'] == before['approval_targets']


@pytest.mark.parametrize('comp', [{'width': 0}, {'width': '1280'}, {'width': True}, {'height': 900}, [], {'width': 1.5}])
def test_an_invalid_comp_size_fails(project, comp):
    root, _ = project
    write(lp.path(root, 'design_dir', 'profiles/web.json'), {'schema_version': 1, 'media': 'web', 'minimums': {},
                                                             'comp': comp})
    result = check(project)
    assert result['exit_code'] == 2 and 'DS_COMP_SIZE_INVALID' in codes(result)


# --- build records are not design inputs ----------------------------------------------------------------------------

def _write_build_records(root):
    plan = lp.path(root, 'build', 'lp-sample', 'plan.yaml')
    plan.parent.mkdir(parents=True, exist_ok=True)
    plan.write_text('schema_version: 1\n')
    comp = lp.path(root, 'build', 'lp-sample', 'comp', 'comp.png')
    comp.parent.mkdir(parents=True, exist_ok=True)
    comp.write_bytes(b'\x89PNG fixture')


def test_build_records_are_not_hashed_into_the_manifest(project):
    root, _ = project
    _write_build_records(root)
    result = check(project)
    assert result['exit_code'] == 1, result['errors']
    build = lp.rel('build')[len(lp.rel('layers')) + 1:]
    assert not [k for k in result['manifest']['web']['input_hashes'] if k.split('/')[0] == build]


def test_writing_a_plan_or_a_comp_after_l0_leaves_the_manifest_fresh(project):
    from resolved_contract import load_manifest
    root, _ = project
    assert check(project)['exit_code'] == 1
    manifest = lp.path(root, 'resolved', 'web.json')
    load_manifest(manifest)
    _write_build_records(root)
    load_manifest(manifest)  # still fresh
    lp.path(root, 'design_dir', 'DESIGN.md').write_text('# changed\n')
    with pytest.raises(ValueError, match='stale'):
        load_manifest(manifest)  # the guard still sees a real design input change


def test_the_node_manifest_reader_also_ignores_build_records(project):
    root, _ = project
    assert check(project)['exit_code'] == 1
    _write_build_records(root)
    manifest = lp.path(root, 'resolved', 'web.json')
    code = ('import {loadManifest} from "' + (SKILL / 'scripts/resolved_contract.mjs').as_uri() + '";'
            'try{loadManifest(process.argv[1]);console.log("fresh")}catch(e){console.log(e.message)}')
    run = subprocess.run(['node', '--input-type=module', '-e', code, str(manifest)], capture_output=True, text=True)
    assert run.returncode == 0, run.stderr
    assert run.stdout.strip() == 'fresh'


def test_maintenance_inputs_ignore_build_records(project):
    from maintenance import _inputs
    root, _ = project
    layers = lp.path(root, 'layers')
    before = _inputs(layers)
    _write_build_records(root)
    assert _inputs(layers) == before


def test_a_new_pattern_catalog_triggers_maintenance(project, monkeypatch):
    """The vocabulary is an input of the resolved manifest: a new catalog version must re-run L0 at maintenance."""
    import maintenance
    root, catalogs = project
    config = json.loads(lp.path(root, 'design_config').read_text())
    before = maintenance._versions(config, catalogs)
    moved = Path(catalogs) / 'layout-patterns-moved.json'
    data = layout_patterns.load_catalog()
    data['version'] = 'next'
    write(moved, data)
    monkeypatch.setattr(layout_patterns, 'CATALOG_FILE', moved)
    assert maintenance._versions(config, catalogs) != before
