"""Reviews the owner accepted (officeohana pilot, 2026-11-18).

A score at or above theta goes to a human as IA_JUDGE_REVIEW. When the owner reads the copy and accepts it, the AI
records that decision with the owner's words (ADR-20261118-ai-records-owner-approvals, biz_ops) in
ux/infoarch/accepted_reviews.yaml (the layout rule's `accepted_reviews`), bound to the judgment hash. An accepted review is no longer a warning; it is
reported as accepted. Changing the copy (or its context) changes the hash, so the record stops applying and the review
comes back. The shared judgment ledger keeps its shape.
"""
import copy
import json

import pytest

from test_check import DATE, OUTLINE, dump
from test_judge import candidates, cli, jl, repo, write_ledger  # noqa: F401  (repo is a fixture)
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)

FLAGGED = 'ia-value-fit:lp-municipal/sec-0001/cpy-0002'
OTHER = 'ia-value-fit:lp-municipal/sec-0001/cpy-0001'
WORDS = '認める'
DAY = '2026-11-18'
FILE = layer_paths.rel('accepted_reviews')


@pytest.fixture
def flagged(repo):
    _, _, rows = candidates(repo)
    theta = jl.load_policy()['theta']
    write_ledger(repo, [{'candidate': r, 'score': theta if r['id'] in (FLAGGED, OTHER) else 0.0} for r in rows])
    return repo


def judge(root):
    return cli('judge', '--root', str(root), '--date', DATE)


def accept(root, *ids, words=WORDS, day=DAY, write=True):
    args = ['accept-reviews', '--root', str(root), '--date', DATE]
    for i in ids:
        args += ['--id', i]
    if words is not None:
        args += ['--owner-words', words]
    if day is not None:
        args += ['--approved-on', day]
    if write:
        args.append('--write')
    return cli(*args)


def codes(out):
    return [f['code'] for f in out['findings']]


def reviews(out):
    return sorted(f['detail']['id'] for f in out['findings'] if f['code'] == 'IA_JUDGE_REVIEW')


def test_an_accepted_review_is_no_longer_a_warning(flagged):
    root = flagged
    code, out = accept(root, FLAGGED)
    assert code == 0, out
    assert out['recorded'] == 1 and out['owner_approval'] == {'words': WORDS, 'date': DAY}
    assert WORDS in out['commit_trailer'] and DAY in out['commit_trailer']
    code, out = judge(root)
    assert reviews(out) == [OTHER]
    accepted = [f for f in out['findings'] if f['code'] == 'IA_JUDGE_ACCEPTED']
    assert len(accepted) == 1 and accepted[0]['level'] == 'info' and accepted[0]['detail']['ids'] == [FLAGGED]


def test_accepting_every_review_turns_judge_green(flagged):
    root = flagged
    assert accept(root, FLAGGED, OTHER)[0] == 0
    code, out = judge(root)
    assert code == 0 and 'IA_JUDGE_REVIEW' not in codes(out), out['findings']


def test_changed_copy_brings_the_review_back(flagged):
    root = flagged
    assert accept(root, FLAGGED)[0] == 0
    changed = copy.deepcopy(OUTLINE)
    changed['copies']['cpy-0002']['text'] = '運用の手間がかなり減ります'
    dump(layer_paths.path(root, 'infoarch', 'outlines', 'lp-municipal.yaml'), changed)
    _, _, rows = candidates(root)
    theta = jl.load_policy()['theta']
    write_ledger(root, [{'candidate': r, 'score': theta if r['id'] in (FLAGGED, OTHER) else 0.0} for r in rows])
    code, out = judge(root)
    assert reviews(out) == sorted([FLAGGED, OTHER])
    stale = [f for f in out['findings'] if f['code'] == 'IA_JUDGE_ACCEPTED_STALE']
    assert len(stale) == 1 and stale[0]['level'] == 'info'


@pytest.mark.parametrize('words,day', [(None, DAY), ('  ', DAY), (WORDS, None), (WORDS, 'yesterday')])
def test_recording_without_the_owners_words_and_date_is_refused(flagged, words, day):
    root = flagged
    code, _ = accept(root, FLAGGED, words=words, day=day)
    assert code == 2 and not (root / FILE).exists()


def test_only_a_current_review_can_be_accepted(flagged):
    root = flagged
    not_flagged = 'ia-value-fit:deck-smb/sec-0001/cpy-0001'
    for ident in (not_flagged, 'ia-value-fit:nowhere/sec-0001/cpy-0001'):
        code, out = accept(root, ident)
        assert code == 2 and not (root / FILE).exists(), out


def test_preview_writes_nothing(flagged):
    root = flagged
    code, out = accept(root, FLAGGED, write=False)
    assert code == 0 and out['status'] == 'preview' and out['recorded'] == 0
    assert not (root / FILE).exists()


def test_recording_twice_keeps_one_row(flagged):
    root = flagged
    assert accept(root, FLAGGED)[0] == 0
    assert accept(root, FLAGGED)[0] == 0
    rows = json.loads((root / FILE).read_text(encoding='utf-8'))['rows']
    assert [r['id'] for r in rows] == [FLAGGED]


@pytest.mark.parametrize('text', ['{', '{"schema_version": 1, "rows": [{"judgment_hash": "x"}]}',
                                  '{"schema_version": 1, "rows": [], "extra": 1}'])
def test_a_broken_record_fails_closed(flagged, text):
    root = flagged
    (root / FILE).write_text(text, encoding='utf-8')
    code, out = judge(root)
    assert code == 2 and 'IA_JUDGE_ACCEPTED_FILE' in codes(out)
    assert reviews(out) == sorted([FLAGGED, OTHER])  # a broken record never hides a review


def test_a_symlinked_record_fails_closed(flagged, tmp_path):
    root = flagged
    target = tmp_path / 'elsewhere.yaml'
    target.write_text('{"schema_version": 1, "rows": []}', encoding='utf-8')
    (root / FILE).symlink_to(target)
    code, out = judge(root)
    assert code == 2 and 'IA_JUDGE_ACCEPTED_FILE' in codes(out)
