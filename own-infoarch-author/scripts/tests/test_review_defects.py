"""Defects found by the external review of 2026-11-15 (biz_ops docs/log/design-ia-boundary-review-20261115.md).

1. Reverse lookup must follow the positioning transitively (v6 requirement 14).
2. The mapping table's version hash must move when a positioning element the deliverable stands on changes.
3. A replacement written from one deliverable's run must not drop the valid IA judgments of other deliverables.
5. The owner-approval bundle must be readable as content, and refuse a stale request (v6 requirement 3).
"""
import copy
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check as ia_check  # noqa: E402
import judge  # noqa: E402
from test_check import DECK, OUTLINE, SSOT, cli, dump, write_repo  # noqa: E402


def test_impact_follows_the_positioning(tmp_path):
    write_repo(tmp_path)
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'str-0001')
    assert rc == 0 and out['outlines'] == ['lp-municipal'], out  # str-0001 backs val-0001 on the LP's path
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'val-0002')
    assert out['outlines'] == ['deck-smb'], out
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'seg-0001')
    assert out['outlines'] == ['lp-municipal'], out


def test_mapping_version_moves_with_a_positioning_dependency(tmp_path):
    write_repo(tmp_path)
    _, before = cli('mapping', '--root', str(tmp_path), '--outline', 'lp-municipal')
    changed = copy.deepcopy(SSOT)
    changed['strengths']['str-0001']['text'] = '別の強み'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    _, after = cli('mapping', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert after['version_hash'] != before['version_hash']
    unrelated = copy.deepcopy(SSOT)
    unrelated['values']['val-0002']['text'] += '（改）'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', unrelated)
    _, same = cli('mapping', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert same['version_hash'] == before['version_hash']


def test_replacement_keeps_valid_judgments_of_other_deliverables():
    jl = judge._jl()
    policy = jl.load_policy()
    lp = judge.outline_candidates('lp-municipal', OUTLINE, SSOT, policy, jl)
    deck = judge.outline_candidates('deck-smb', DECK, SSOT, policy, jl)
    emphasis = jl.candidate('ia-emphasis', 'deck-smb/decks/smb.html@mobile', '見出し', {'outline': 'deck-smb'}, policy=policy)
    ledger = {'schema_version': jl.CURRENT_SCHEME, 'policy': policy, 'holdout': {'source': 'unmeasured'},
              'rows': [{'candidate': r, 'score': 0.1} for r in lp + deck + [emphasis]]}
    # this run rendered only the LP: its candidates are the outline-level ones of both, without the deck's emphasis
    rows = lp + deck
    answers = {'status': 'model-answers-unlabeled', 'policy': policy,
               'rows': [{'candidate': lp[0], 'score': 0.2, 'raw_answer': {'type': 'noul', 'noul': 0.2}}]}
    out, summary = judge.replacement(ledger, rows, answers, policy, jl, outlines={'lp-municipal', 'deck-smb'})
    ids = {r['candidate']['id'] for r in out['rows']}
    assert emphasis['id'] in ids  # another deliverable's first screen, not rendered in this run, stays
    assert {r['id'] for r in deck} <= ids  # still current and valid
    assert summary['dropped_ia_rows'] == 0


def test_replacement_drops_superseded_and_orphaned_judgments():
    jl = judge._jl()
    policy = jl.load_policy()
    lp = judge.outline_candidates('lp-municipal', OUTLINE, SSOT, policy, jl)
    changed = copy.deepcopy(OUTLINE)
    changed['pages'][0]['sections'][1]['copy'] = '運用の手間がかなり減ります'
    now = judge.outline_candidates('lp-municipal', changed, SSOT, policy, jl)
    gone = jl.candidate('ia-emphasis', 'retired/x.html@mobile', '見出し', {'outline': 'retired'}, policy=policy)
    ledger = {'schema_version': jl.CURRENT_SCHEME, 'policy': policy, 'holdout': {'source': 'unmeasured'},
              'rows': [{'candidate': r, 'score': 0.1} for r in lp + [gone]]}
    answers = {'status': 'model-answers-unlabeled', 'policy': policy, 'rows': []}
    out, summary = judge.replacement(ledger, now, answers, policy, jl, outlines={'lp-municipal'})
    ids = {r['candidate']['judgment_hash'] for r in out['rows']}
    stale = {r['judgment_hash'] for r in lp} - {r['judgment_hash'] for r in now}
    assert stale and not (stale & ids) and gone['judgment_hash'] not in ids
    assert summary['dropped_ia_rows'] == len(stale) + 1


def test_review_bundle_shows_content_and_refuses_a_stale_request(tmp_path):
    write_repo(tmp_path, confirmed=False)
    m = ia_check.Model(root=tmp_path, ssot=SSOT)
    hashes, _ = ia_check.positioning_targets(m)
    request = {'schema_version': 2, 'approval_scheme': 'hash-jcs-v1', 'repo_id': 'x',
               'rows': [{'key': k, 'value_hash': v, 'date': '2026-11-15'} for k, v in sorted(hashes.items())]}
    req = tmp_path / 'request.json'
    req.write_text(json.dumps(request), encoding='utf-8')
    rc, out = cli('review', '--root', str(tmp_path), '--request', str(req))
    assert rc == 0, out
    text = out['markdown']
    assert '自動更新で保守が要らない' in text and 'str-0001' in text and 'str-0001 → alt-0001：代替に無い' in text
    assert 'alt-0001 → pan-0001：満たす度合い（none）' in text and '写しの出どころ：alt-0001' in text
    changed = copy.deepcopy(SSOT)
    changed['strengths']['str-0001']['text'] = '別の強み'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, out = cli('review', '--root', str(tmp_path), '--request', str(req))
    assert rc == 2 and out['stale'] == ['ia:positioning:str-0001'], out  # its arrows keep their approval
