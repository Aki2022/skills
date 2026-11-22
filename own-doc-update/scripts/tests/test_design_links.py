"""The docs gate checks references only; design semantics belong to the runner."""
import importlib.util
import json
from pathlib import Path
import sys

import pytest

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))
spec = importlib.util.spec_from_file_location('design_docs_validator', SCRIPTS / 'validate_repo_docs.py')
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)


@pytest.fixture
def project(tmp_path):
    (tmp_path / 'design').mkdir()
    (tmp_path / 'docs/specs').mkdir(parents=True)
    (tmp_path / 'docs/specs/visual.md').write_text('# Policy\n\n[Master](../../design/design.yaml)\n')
    (tmp_path / 'design/design.yaml').write_text(json.dumps({'spec': 'docs/specs/visual.md'}))
    return tmp_path


def findings(root):
    errors = []
    validator.validate_design_references(root, errors)
    return errors


def test_reciprocal_json_references_pass_without_evaluating_design(project):
    # No tokens, items, approvals or catalog: this check must not run L0.
    assert findings(project) == []


def test_yaml_comments_and_multiple_specs(project):
    (project / 'docs/specs/content.md').write_text('# Content\n\n[Master](../../design/)\n')
    (project / 'design/design.yaml').write_text('spec:\n  - docs/specs/visual.md\n  - docs/specs/content.md # policy\n')
    assert findings(project) == []


@pytest.mark.parametrize('value', [None, '', [], True, ['docs/specs/visual.md'] * 2])
def test_missing_or_invalid_spec_declaration_is_red(project, value):
    (project / 'design/design.yaml').write_text(json.dumps({'spec': value}))
    assert any(error.startswith('design/design.yaml:') for error in findings(project))


def test_missing_spec_file_is_red(project):
    (project / 'docs/specs/visual.md').unlink()
    assert findings(project)


def test_missing_backlink_is_on_the_changed_spec(project):
    (project / 'docs/specs/visual.md').write_text('# Policy\n')
    assert any(error.startswith('docs/specs/visual.md:') for error in findings(project))


def test_unlisted_spec_backlink_is_on_design_configuration(project):
    (project / 'docs/specs/content.md').write_text('# Content\n\n[Master](../../design/)\n')
    assert any('does not list docs/specs/content.md' in error for error in findings(project))


def test_link_to_nonexistent_design_is_red(project):
    (project / 'design/design.yaml').unlink()
    assert findings(project)


@pytest.mark.parametrize('target', ['../outside.md', '/tmp/outside.md', 'docs/specs/archive/old.md', 'https://example.org/policy'])
def test_spec_path_must_be_inside_active_repository_specs(project, target):
    (project / 'design/design.yaml').write_text(json.dumps({'spec': target}))
    assert findings(project)


def test_escaped_symlink_cannot_satisfy_spec_reference(project, tmp_path):
    outside = tmp_path / 'outside.md'
    outside.write_text('# Not a project policy\n')
    path = project / 'docs/specs/visual.md'
    path.unlink()
    path.symlink_to(outside)
    assert findings(project)


def test_code_spans_do_not_count_as_backlinks(project):
    (project / 'docs/specs/visual.md').write_text('`[example](../../design/)`\n')
    assert findings(project)


def test_duplicate_yaml_spec_key_is_red(project):
    (project / 'design/design.yaml').write_text('spec: docs/specs/visual.md\nspec: docs/specs/visual.md\n')
    assert findings(project)


def test_unrelated_repo_needs_no_design_directory(project):
    import shutil
    shutil.rmtree(project / 'design')
    (project / 'docs/specs/visual.md').write_text('# Ordinary policy\n')
    assert findings(project) == []


def test_main_validator_calls_reference_check(project):
    (project / 'docs/00_index.md').write_text('---\nupdated_at: 2026-10-23\n---\n# Index\n')
    (project / 'docs/specs/visual.md').write_text('# Policy\n')
    errors, _ = validator.validate_repo(project)
    assert any('design back-reference' in error for error in errors)


def test_either_side_of_a_broken_pair_is_reported_to_staged_file_filter(project):
    (project / 'docs/specs/visual.md').write_text('# Policy\n')
    errors = findings(project)
    assert any(error.startswith('design/design.yaml:') for error in errors)
    assert any(error.startswith('docs/specs/visual.md:') for error in errors)
    (project / 'docs/specs/visual.md').write_text('[Master](../../design/)\n')
    (project / 'docs/specs/content.md').write_text('[Master](../../design/)\n')
    errors = findings(project)
    assert any(error.startswith('design/design.yaml:') for error in errors)
    assert any(error.startswith('docs/specs/content.md:') for error in errors)
