"""The owner decides, the AI records (ADR-20261118-ai-records-owner-approvals, biz_ops).

Recording is refused without the owner's own approval words and the date they were given, and the date must be the
one in the reviewed request. The value hashes are recomputed right before writing. Positioning is approved before any
outline exists, so a positioning-only IA (no outlines yet) can be prepared and recorded, also in a repository with
no design/design.yaml at all.
"""
import copy
import json
import runpy
import subprocess
import sys

import pytest

from test_core import SKILL, approve, check, project, write
from owner_approval import approve_request, prepare

IA = SKILL.parent / 'own-infoarch-author'
FIXTURE = runpy.run_path(str(IA / 'scripts/tests/test_check.py'))
WORDS = 'positioningは承認する'
DAY = '2026-11-18'


def ssot():
    data = copy.deepcopy(FIXTURE['SSOT'])
    data['proofs']['prf-0001']['as_of'] = '2099-01-01'
    data['facts']['fct-0001']['as_of'] = '2099-01-01'
    return data


def ia_only(tmp_path):
    """A repository with design/info_arch only: no design.yaml, no outlines, nothing approved yet."""
    root = tmp_path / 'ia-only'
    root.mkdir()
    subprocess.run(['git', 'init', '-q', str(root)], check=True)
    subprocess.run(['git', '-C', str(root), 'remote', 'add', 'origin', 'https://example.org/ia-only.git'], check=True)
    FIXTURE['write_repo'](root, ssot(), outlines=(), confirmed=False)
    return root


def ia_keys(root):
    sys.path.insert(0, str(IA / 'scripts'))
    import check as ia_check
    import positioning
    m = ia_check.Model(root=root, ssot=ssot())
    return sorted(positioning.approval_targets(m))


def ledger(root):
    path = root / 'design/approvals.yaml'
    return json.loads(path.read_text()) if path.is_file() else None


def cli(root, *args, catalogs=None):
    extra = ['--catalog-dir', str(catalogs)] if catalogs else []
    return subprocess.run([sys.executable, str(SKILL / 'scripts/design_author.py'), 'approve', '--repo', str(root),
                           *extra, *args], capture_output=True, text=True)


# ---- the owner's words are required ----------------------------------------------------------------------------

@pytest.mark.parametrize('words,day', [(None, DAY), ('', DAY), ('   ', DAY), (WORDS, None), (WORDS, 'yesterday')])
def test_recording_without_the_owners_words_and_date_is_refused(project, words, day):
    root, catalogs = project
    approve(project)
    before = ledger(root)
    request = prepare(root, catalogs, all_must=True, when=DAY)
    with pytest.raises(ValueError, match="owner's approval"):
        approve_request(root, catalogs, request, 'owner', owner_words=words, approved_on=day)
    assert ledger(root) == before


def test_the_request_date_must_be_the_day_the_owner_approved(project):
    root, catalogs = project
    request = prepare(root, catalogs, all_must=True, when='2026-11-17')
    with pytest.raises(ValueError, match='approval date'):
        approve_request(root, catalogs, request, 'owner', owner_words=WORDS, approved_on=DAY)


def test_recording_returns_the_quote_for_the_commit_and_the_pr(project):
    root, catalogs = project
    request = prepare(root, catalogs, all_must=True, when=DAY)
    out = approve_request(root, catalogs, request, 'owner', owner_words=WORDS, approved_on=DAY)
    assert out['recorded_rows'] == 1
    assert out['owner_approval'] == {'words': WORDS, 'date': DAY}
    assert WORDS in out['commit_trailer'] and DAY in out['commit_trailer']


def test_cli_refuses_recording_without_owner_words(project, tmp_path):
    root, catalogs = project
    request = prepare(root, catalogs, all_must=True, when=DAY)
    path = tmp_path / 'request.json'
    write(path, request)
    before = ledger(root)
    refused = cli(root, '--request', str(path), '--approver', 'owner', catalogs=catalogs)
    assert refused.returncode != 0 and ledger(root) == before
    done = cli(root, '--request', str(path), '--approver', 'owner', '--owner-words', WORDS, '--approved-on', DAY,
               catalogs=catalogs)
    assert done.returncode in (0, 1, 2) and json.loads(done.stdout)['owner_approval'] == {'words': WORDS, 'date': DAY}
    assert len(ledger(root)) == 1


# ---- positioning before outlines -------------------------------------------------------------------------------

def test_positioning_without_outlines_can_be_prepared_in_a_design_repository(project):
    approve(project)
    root, catalogs = project
    FIXTURE['write_repo'](root, ssot(), outlines=(), confirmed=False)
    config = json.loads((root / 'design/design.yaml').read_text())
    config['modules'] = {'ia': '0.4.0'}
    write(root / 'design/design.yaml', config)
    assert any(e.startswith('IA_NO_OUTLINES') for e in check(project)['errors'])
    request = prepare(root, catalogs, targets=ia_keys(root), when=DAY)
    assert {r['key'] for r in request['rows']} == set(ia_keys(root))


def test_ia_only_repository_prepares_and_records_without_design_yaml(tmp_path):
    root = ia_only(tmp_path)
    keys = ia_keys(root)
    request = prepare(root, None, targets=keys, when=DAY)
    assert len(request['rows']) == len(keys) and not (root / 'design/design.yaml').exists()
    out = approve_request(root, None, request, 'owner', owner_words=WORDS, approved_on=DAY)
    assert out['recorded_rows'] == len(keys)
    assert {r['key'] for r in ledger(root)} == set(keys)
    report = json.loads(subprocess.run([sys.executable, str(IA / 'scripts/infoarch_author.py'), 'check',
                                        '--root', str(root)], capture_output=True, text=True).stdout)
    assert [f['code'] for f in report['findings'] if f['level'] == 'fail'] == ['IA_NO_OUTLINES'], report['findings']


def test_ia_only_change_after_review_is_refused_before_write(tmp_path):
    root = ia_only(tmp_path)
    request = prepare(root, None, targets=ia_keys(root), when=DAY)
    changed = ssot()
    changed['strengths']['str-0001']['text'] = '自動更新で保守がほぼ要らない'
    FIXTURE['dump'](root / 'design/info_arch/info_arch.yaml', changed)
    with pytest.raises(ValueError, match='reviewed inputs changed'):
        approve_request(root, None, request, 'owner', owner_words=WORDS, approved_on=DAY)
    assert ledger(root) is None


def test_ia_only_refuses_a_structural_error(tmp_path):
    root = ia_only(tmp_path)
    (root / 'design/info_arch/info_arch.yaml').write_text('schema_version: [', encoding='utf-8')
    with pytest.raises(ValueError, match='repair structural'):
        prepare(root, None, targets=['ia:positioning:val-0001'], when=DAY)


# ---- approving a positioning in parts --------------------------------------------------------------------------

def test_a_partly_approved_positioning_does_not_block_approving_the_rest(tmp_path):
    """After the first part is recorded the positioning is incomplete (no alternatives yet). Incomplete is fixed by
    approving the rest, so it must not refuse preparing and recording that rest (officeohana pilot, 2026-11-18)."""
    root = ia_only(tmp_path)
    keys = ia_keys(root)
    later = [k for k in keys if 'alt-' in k]
    first = [k for k in keys if k not in later]
    assert later and first
    approve_request(root, None, prepare(root, None, targets=first, when=DAY), 'owner', owner_words=WORDS, approved_on=DAY)
    rest = prepare(root, None, targets=later, when=DAY)
    out = approve_request(root, None, rest, 'owner', owner_words='代替の案は承認する', approved_on=DAY)
    assert out['recorded_rows'] == len(later) and {r['key'] for r in ledger(root)} == set(keys)
    assert [e for e in out['errors'] if not e.startswith(('IA_NO_OUTLINES', 'zero IA'))] == [], out['errors']


def test_a_broken_arrow_still_refuses_prepare(tmp_path):
    root = ia_only(tmp_path)
    broken = ssot()
    broken['edges'].append(copy.deepcopy(broken['edges'][0]))  # the same arrow twice is damage, not an approval gap
    FIXTURE['dump'](root / 'design/info_arch/info_arch.yaml', broken)
    with pytest.raises(ValueError, match='repair structural'):
        prepare(root, None, targets=ia_keys(root)[:1], when=DAY)
