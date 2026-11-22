"""The before/after comparison of a repository's check result, one record at a time.

A comparison that looks at nothing must not pass: zero compared records, fewer than the stated count, a missing
record, a changed hash or a changed verdict are each a failure."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import layer_snapshot as ls  # noqa: E402
from layer_testkit import SKILLS, git_init, run  # noqa: E402

H1, H2, H3 = 'a' * 64, 'b' * 64, 'c' * 64


def snap():
    return {
        'schema_version': 1,
        'ledger': {'ds:web:item:dads-000001': {'value_hash': H1, 'current_hash': H1, 'valid': True},
                   'ia:path:val-0001': {'value_hash': H2, 'current_hash': H3, 'valid': False}},
        'targets': {'ds:web:item:dads-000001': H1, 'ia:path:val-0001': H3, 'ds:web:item:dads-000002': H2},
        'rules': ['R-a', 'R-b'],
        'verdict': {'exit_code': 1, 'stages': {'L0': {'exit_code': 1, 'targets': 3}, 'ia': {'exit_code': 0, 'targets': 2}},
                    'errors': [], 'warnings': ['<design>/x awaiting approval: 2']},
    }


def test_identical_snapshots_compare_equal_and_count_what_they_compared():
    result = ls.compare(snap(), snap())
    assert result['differences'] == []
    assert result['compared'] == {'ledger': 2, 'targets': 3, 'rules': 2, 'stages': 2}
    assert ls.exit_code(result, minimum=1) == 0


def test_an_empty_comparison_is_not_a_pass():
    empty = {'schema_version': 1, 'ledger': {}, 'targets': {}, 'rules': [],
             'verdict': {'exit_code': 0, 'stages': {}, 'errors': [], 'warnings': []}}
    result = ls.compare(empty, copy.deepcopy(empty))
    assert result['differences'] == []
    assert ls.exit_code(result, minimum=1) == 2


def test_fewer_records_than_stated_is_not_a_pass():
    result = ls.compare(snap(), snap())
    assert ls.exit_code(result, minimum=1) == 0
    assert ls.exit_code(result, expected={'targets': 4}) == 2


@pytest.mark.parametrize('field,change', [
    ('approval hash', lambda s: s['ledger']['ds:web:item:dads-000001'].update(value_hash=H3)),
    ('approval validity', lambda s: s['ledger']['ia:path:val-0001'].update(valid=True)),
    ('target set', lambda s: s['targets'].pop('ds:web:item:dads-000002')),
    ('target hash', lambda s: s['targets'].update({'ds:web:item:dads-000002': H1})),
    ('rule id', lambda s: s['rules'].remove('R-b')),
    ('verdict', lambda s: s['verdict'].update(exit_code=2)),
    ('stage', lambda s: s['verdict']['stages']['ia'].update(exit_code=1)),
    ('warning', lambda s: s['verdict']['warnings'].append('new')),
])
def test_each_kind_of_difference_is_found(field, change):
    after = snap()
    change(after)
    result = ls.compare(snap(), after)
    assert result['differences'], field
    assert ls.exit_code(result, minimum=1) == 1


def test_normalisation_makes_two_layouts_comparable():
    pre = {'entries': {'layers': 'design', 'design': 'design', 'ledger': 'design/approvals.yaml'}}
    post = {'entries': {'layers': 'ux', 'design': 'ux/design', 'ledger': 'ux/approvals.yaml'}}
    assert ls.normalise('design/approvals.yaml missing and design/x', pre) == '<ledger> missing and <design>/x'
    assert ls.normalise('ux/approvals.yaml missing and ux/design/x', post) == '<ledger> missing and <design>/x'


def test_snapshot_of_a_fresh_design_is_stable_and_nonempty(tmp_path):
    catalogs = tmp_path / 'catalogs'
    catalogs.mkdir()
    item = {'id': 'dads-000001', 'match_key': 'dads-000001', 'document': 'fixture', 'title': 'x', 'headings': ['x'],
            'text': 'A synthetic fixture policy.', 'source_url': 'https://example.org/', 'evidence_types': ['token']}
    (catalogs / 'test-v1.json').write_text(json.dumps({'schema_version': 1, 'version': 'test-v1', 'items': [item]}))
    (catalogs / 'id-ledger.json').write_text(json.dumps({'schema_version': 1, 'next_id': 2, 'entries': [
        {'id': 'dads-000001', 'match_key': 'dads-000001', 'retired': False}]}))
    repo = git_init(tmp_path / 'repo')
    subprocess.run(['git', '-C', str(repo), 'remote', 'add', 'origin', 'https://example.org/fixture.git'], check=True)
    init = run([SKILLS / 'own-design-author/scripts/design-author', 'init', '--repo', repo, '--preset', 'lp',
                '--catalog-version', 'test-v1', '--catalog-dir', catalogs])
    assert init.returncode == 0, init.stdout + init.stderr
    out = []
    for n in (1, 2):
        dest = tmp_path / f'snap{n}.json'
        r = run([sys.executable, SKILLS / 'own-design-author/scripts/layer_snapshot.py', 'snapshot', '--repo', repo,
                 '--catalog-dir', catalogs, '--out', dest])
        assert r.returncode == 0, r.stdout + r.stderr
        out.append(json.loads(dest.read_text()))
    assert out[0] == out[1]
    assert out[0]['verdict']['stages'], 'a snapshot that ran no stage compares nothing'
    cmp = run([sys.executable, SKILLS / 'own-design-author/scripts/layer_snapshot.py', 'compare',
               tmp_path / 'snap1.json', tmp_path / 'snap2.json', '--min', '0', '--expect', 'stages=1'])
    assert cmp.returncode == 0, cmp.stdout + cmp.stderr
