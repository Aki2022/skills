"""Meaning judgment and the answer key (SPEC-information-architecture-standardization v6, requirements 11 and 12).

Jev is asked one question only: does this copy convey the value its message points to (ia-value-fit). The answer
key pairs every copy with its own value (expected to convey) and re-pairs it with the other values of the same
positioning (or, with a single value, another positioning's value) as candidates that should not. The owner's
labels decide; the machine only proposes. Agreement is measured against the shared ledger, and when the owner's
labels hold no conveying or no non-conveying pair, the result is "unmeasurable", never a number.
All tests use a fake Jev: ledger rows with chosen scores. No model is called.
"""
import copy
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import answer_key  # noqa: E402
import check as ia_check  # noqa: E402
import judge  # noqa: E402
import positioning  # noqa: E402
from test_check import DECK, OUTLINE, SSOT_VIEW as SSOT, cli, write_repo  # noqa: E402

DESIGN = Path(__file__).resolve().parents[3] / 'own-design-author/scripts'
sys.path.insert(0, str(DESIGN))
import judgment_ledger as jl  # noqa: E402
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)


def model(ssot=SSOT, outlines=(OUTLINE, DECK)):
    m = ia_check.Model(root=Path('.'), ssot=ssot)
    for o in outlines:
        m.outlines[o['id']] = (Path(layer_paths.rel('infoarch', 'outlines', f'{o["id"]}.yaml')), o)
    return m, positioning.approved_keys(m, positioning.approval_targets(m))


def items(ssot=SSOT, outlines=(OUTLINE, DECK)):
    m, approved = model(ssot, outlines)
    return answer_key.build(m, approved, jl.load_policy(), jl)


# ---- the question Jev is asked (requirement 11) ----------------------------------------------------------------

def test_value_fit_is_copy_times_one_value_and_replaces_strategy_fit():
    rows = judge.outline_candidates('lp-municipal', OUTLINE, SSOT, jl.load_policy(), jl)
    kinds = {r['input']['kind'] for r in rows}
    assert 'ia-value-fit' in kinds and 'ia-strategy-fit' not in kinds
    fit = {r['id']: r for r in rows if r['input']['kind'] == 'ia-value-fit'}
    # trust (facts, proof) and act (the call to action) sections do not carry a value: they are not asked (D32).
    assert sorted(fit) == ['ia-value-fit:lp-municipal/sec-0001/cpy-0001', 'ia-value-fit:lp-municipal/sec-0001/cpy-0002']
    ctx = fit['ia-value-fit:lp-municipal/sec-0001/cpy-0002']['input']['context']
    assert set(ctx) == {'outline', 'message', 'value'}


def test_trust_and_act_copy_is_not_asked_whether_it_conveys_a_value():
    """Pilot and I16: every fact listing and call-to-action copy was labeled "does not convey" by the owner, because
    conveying a value is not its job. The value judgment skips those sections; the name and claim judgments stay."""
    rows = judge.outline_candidates('lp-municipal', OUTLINE, SSOT, jl.load_policy(), jl)
    roles = {sid: sec['role'] for sid, sec in OUTLINE['sections'].items()}
    by_kind = {}
    for r in rows:
        by_kind.setdefault(r['input']['kind'], set()).add(r['id'].split(':', 1)[1])
    skipped = {f'lp-municipal/{c["section"]}/{cid}' for cid, c in OUTLINE['copies'].items()
               if roles[c['section']] in ('trust', 'act')}
    assert skipped and not skipped & by_kind['ia-value-fit']
    assert skipped <= by_kind['ia-unregistered-claim']
    assert not any(i['id'].split('@')[0] in skipped for i in items())


def test_value_fit_kind_is_in_the_shared_policy():
    assert 'ia-value-fit' in jl.KINDS and jl.load_policy()['prompts']['ia-value-fit'].strip()


@pytest.mark.parametrize('change', ['copy', 'value'])
def test_judgment_input_binds_the_copy_and_the_value_content(change):
    """Changing either side gives a new judgment hash, so an old judgment is never reused."""
    def key(ssot, outline):
        rows = judge.outline_candidates('lp-municipal', outline, ssot, jl.load_policy(), jl)
        return {r['id']: r['judgment_hash'] for r in rows if r['input']['kind'] == 'ia-value-fit'}
    ssot, outline = copy.deepcopy(SSOT), copy.deepcopy(OUTLINE)
    if change == 'copy':
        outline['copies']['cpy-0002']['text'] = '運用の手間がかなり減ります'
    else:
        ssot['values']['val-0001']['text'] = '運用がかなり楽になる'
    before, after = key(SSOT, OUTLINE), key(ssot, outline)
    moved = sorted(k for k in before if before[k] != after[k])
    # cpy-0001 and cpy-0002 both point to val-0001; trust and act copy are not asked (D32)
    expect = ['ia-value-fit:lp-municipal/sec-0001/cpy-0002'] if change == 'copy' else \
        ['ia-value-fit:lp-municipal/sec-0001/cpy-0001', 'ia-value-fit:lp-municipal/sec-0001/cpy-0002']
    assert moved == expect, moved


# ---- the answer key (requirement 12) -----------------------------------------------------------------------------

def test_answer_key_pairs_each_copy_with_its_value_and_re_pairs_it_within_the_positioning():
    out = {i['id']: i for i in items()}
    own = out['lp-municipal/sec-0001/cpy-0002@val-0001']
    assert own['kind'] == 'own' and own['value']['id'] == 'val-0001'
    swap = out['lp-municipal/sec-0001/cpy-0002@val-0003']  # the other value of seg-0001
    assert swap['kind'] == 'swap' and swap['candidate']['input']['kind'] == 'ia-value-fit'
    assert swap['candidate']['input']['text'] == own['candidate']['input']['text']  # same copy, another value


def test_a_single_value_positioning_is_re_paired_with_another_positioning():
    out = {i['id']: i for i in items()}
    deck = [i for i in out.values() if i['id'].startswith('deck-smb/sec-0001/cpy-0001@')]
    assert {i['value']['id']: i['kind'] for i in deck} == {'val-0002': 'own', 'val-0001': 'swap'}


def test_answer_key_is_reproducible_and_ordered():
    a, b = items(), items()
    assert [i['id'] for i in a] == [i['id'] for i in b] == sorted(i['id'] for i in a)
    assert answer_key.key_hash(a) == answer_key.key_hash(b)


def test_page_lists_every_pair_and_works_offline():
    its = items()
    html = answer_key.page(its)
    assert all(i['id'] in html for i in its)
    assert '<script src' not in html and 'http' not in html.split('</style>')[-1].lower().replace('http-equiv', '')
    assert answer_key.key_hash(its) in html  # the labels file is bound to this key


PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')
CLICK = r"""
const { chromium } = require(require('path').join(process.env.PW_DIR, 'node_modules', 'playwright'));
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + process.argv[2]);
  const rows = await p.$$('tr[data-id]');
  const picks = {};
  for (let n = 0; n < rows.length; n++) {
    const v = n % 3 === 0 ? 'true' : (n % 3 === 1 ? 'false' : 'null');
    await rows[n].$eval('input[value="' + v + '"]', e => e.click());
    picks[await rows[n].getAttribute('data-id')] = v === 'null' ? null : v === 'true';
  }
  const shown = await p.$eval('#labels-json', e => e.value);
  const notice = await p.isVisible('#no-script');
  console.log(JSON.stringify({ picks, shown, notice }));
  await b.close();
})();
"""
NO_SCRIPT = r"""
const { chromium } = require(require('path').join(process.env.PW_DIR, 'node_modules', 'playwright'));
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ javaScriptEnabled: false });
  await p.goto('file://' + process.argv[2]);
  console.log(JSON.stringify({ notice: await p.isVisible('#no-script'),
                               text: await p.textContent('#no-script') }));
  await b.close();
})();
"""


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser test not run (skipped, not passed)')
def test_a_viewer_without_scripts_is_told_to_open_the_page_in_a_browser(tmp_path):
    """The owner's file viewer ran no script, so save and copy silently did nothing (pilot, 2026-11-19)."""
    page = tmp_path / 'page.html'
    page.write_text(answer_key.page(items()), encoding='utf-8')
    script = tmp_path / 'noscript.cjs'
    script.write_text(NO_SCRIPT, encoding='utf-8')
    proc = subprocess.run(['node', str(script), str(page)], capture_output=True, text=True, check=True,
                          env={**os.environ, 'PW_DIR': PW_DIR})
    out = json.loads(proc.stdout)
    assert out['notice'] and 'ブラウザ' in out['text']


def test_page_shows_the_labels_as_copyable_json():
    """The download alone left no file on the owner's Mac (I16): the page also shows labels.json to copy."""
    html = answer_key.page(items())
    assert 'id="labels-json"' in html and 'readonly' in html and 'id="copy"' in html


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser test not run (skipped, not passed)')
def test_copyable_json_follows_the_choices_and_measures(tmp_path):
    its = items()
    page = tmp_path / 'page.html'
    page.write_text(answer_key.page(its), encoding='utf-8')
    script = tmp_path / 'click.cjs'
    script.write_text(CLICK, encoding='utf-8')
    proc = subprocess.run(['node', str(script), str(page)], capture_output=True, text=True, check=True,
                          env={**os.environ, 'PW_DIR': PW_DIR})
    out = json.loads(proc.stdout)
    shown = json.loads(out['shown'])
    assert shown == {'answer_key': answer_key.key_hash(its), 'labels': out['picks']}
    assert out['notice'] is False  # scripts run, so the no-script notice is hidden
    theta = jl.load_policy()['theta']
    result = answer_key.measure(its, shown, ledger(its, {i['id']: 0.0 for i in its}), theta)
    assert result['status'] == 'measured'


# ---- measuring (fake Jev) ----------------------------------------------------------------------------------------

def ledger(its, scores):
    return {'rows': [{'candidate': i['candidate'], 'score': scores[i['id']]} for i in its if i['id'] in scores]}


def labels(its, conveys):
    return {'answer_key': answer_key.key_hash(its), 'labels': conveys}


def test_agreement_misses_and_false_detections_are_counted():
    its = items()
    theta = jl.load_policy()['theta']
    own = [i['id'] for i in its if i['kind'] == 'own'][:3]
    swap = [i['id'] for i in its if i['kind'] == 'swap'][:3]
    owner = {own[0]: True, own[1]: True, own[2]: True, swap[0]: False, swap[1]: False, swap[2]: None}
    jev = {own[0]: 0.0, own[1]: theta, own[2]: 0.1, swap[0]: 0.95, swap[1]: 0.0, swap[2]: 0.9}
    out = answer_key.measure(its, labels(its, owner), ledger(its, jev), theta)
    assert out['status'] == 'measured'
    assert out['counts'] == {'labeled': 5, 'conveys': 3, 'does_not_convey': 2, 'unknown': 1, 'uncovered': 0}
    assert out['agreement'] == pytest.approx(3 / 5)
    assert out['misses'] == [swap[1]]          # owner: does not convey; Jev did not flag it
    assert out['false_detections'] == [own[1]]  # owner: conveys; Jev flagged it


@pytest.mark.parametrize('side', [True, False])
def test_zero_on_either_side_is_unmeasurable_not_a_number(side):
    its = items()
    picked = [i['id'] for i in its if (i['kind'] == 'own') == side][:3]
    out = answer_key.measure(its, labels(its, {k: side for k in picked}),
                             ledger(its, {k: 0.0 for k in picked}), jl.load_policy()['theta'])
    assert out['status'] == 'unmeasurable' and 'agreement' not in out and out['reason'], out


def test_uncovered_pairs_are_reported_and_not_counted():
    its = items()
    own, swap = [i['id'] for i in its if i['kind'] == 'own'][:2], [i['id'] for i in its if i['kind'] == 'swap'][:2]
    owner = {own[0]: True, own[1]: True, swap[0]: False, swap[1]: False}
    out = answer_key.measure(its, labels(its, owner), ledger(its, {own[0]: 0.0, swap[0]: 0.9}), jl.load_policy()['theta'])
    assert out['counts']['uncovered'] == 2 and sorted(out['uncovered']) == sorted([own[1], swap[1]])
    assert out['agreement'] == 1.0


def test_labels_for_another_key_are_refused():
    its = items()
    stale = {'answer_key': '0' * 64, 'labels': {its[0]['id']: True}}
    with pytest.raises(ValueError, match='answer key'):
        answer_key.measure(its, stale, ledger(its, {}), 0.5)
    with pytest.raises(ValueError, match='unknown'):
        answer_key.measure(its, labels(its, {'nope@val-0001': True}), ledger(its, {}), 0.5)


def test_cli_writes_the_key_and_measures_from_the_shared_ledger(tmp_path):
    write_repo(tmp_path)
    out = tmp_path / 'key'
    rc, made = cli('answer-key', '--root', str(tmp_path), '--out', str(out))
    assert rc == 0 and (out / 'answer-key-page.html').is_file() and (out / 'candidates.json').is_file(), made
    cands = json.loads((out / 'candidates.json').read_text(encoding='utf-8'))['judgment_candidates']
    key = json.loads((out / 'answer-key.json').read_text(encoding='utf-8'))
    rc, again = cli('answer-key', '--root', str(tmp_path), '--out', str(out))
    assert rc == 2  # never overwrites
    its = key['items']
    own = [i for i in its if i['kind'] == 'own'][0]['id']
    swap = [i for i in its if i['kind'] == 'swap'][0]['id']
    (tmp_path / 'labels.json').write_text(json.dumps({'answer_key': key['answer_key'], 'labels': {own: True, swap: False}}),
                                          encoding='utf-8')
    policy = jl.load_policy()
    by_id = {c['input']['catalog_id']: c for c in cands}
    rows = [{'candidate': by_id[own], 'score': 0.0}, {'candidate': by_id[swap], 'score': 0.99}]
    layer_paths.path(tmp_path, 'judgments').write_text(json.dumps({'schema_version': 3, 'policy': policy, 'rows': rows,
                                                                'holdout': {'source': 'unmeasured'}}), encoding='utf-8')
    rc, measured = cli('answer-key', '--root', str(tmp_path), '--measure', str(tmp_path / 'labels.json'),
                       '--key', str(out / 'answer-key.json'))
    assert rc == 0 and measured['status'] == 'measured' and measured['agreement'] == 1.0, measured
