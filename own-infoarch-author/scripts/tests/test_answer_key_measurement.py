"""WS-20261123 ISSUE-02: the answer key measures accuracy through the real recording path (requirement 12).

own pairs (a copy with its own value) are the shipping question, so Jev's answer is read from the shared ledger
(`ux/judgments.json`). swap pairs (the same copy with another value) exist only to measure, so their answers are
recorded in a measurement ledger next to the answer key (`measurement-judgments.json`), which the ship gate never
reads. Accuracy is measured from the two together. Jev is replaced by a fake transport; no model is called.
"""
import hashlib
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import answer_key  # noqa: E402
import judge  # noqa: E402
from test_check import DATE, cli, write_repo  # noqa: E402
from test_evaluate import responder, usage_ledger  # noqa: E402

DESIGN = Path(__file__).resolve().parents[3] / 'own-design-author/scripts'
sys.path.insert(0, str(DESIGN))
import judgment_ledger as jl  # noqa: E402
import layer_paths  # noqa: E402

SCRIPTS = Path(__file__).resolve().parents[1]
MEASUREMENT = 'measurement-judgments.json'


@pytest.fixture
def made(tmp_path, monkeypatch):
    """A repository with its answer key written by the real CLI -> (root, key dir, key file contents)."""
    monkeypatch.setenv('TYPESAFE_API_KEY', 'test-key')
    root = tmp_path / 'repo'
    root.mkdir()
    write_repo(root)
    out = root / 'key'
    rc, made_ = cli('answer-key', '--root', str(root), '--out', str(out))
    assert rc == 0, made_
    return root, out, json.loads((out / 'answer-key.json').read_text(encoding='utf-8'))


def write_shared_ledger(root, rows):
    ledger = {'schema_version': jl.CURRENT_SCHEME, 'policy': jl.load_policy(), 'rows': rows, 'holdout': {'source': 'unmeasured'}}
    path = layer_paths.path(root, 'judgments')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False), encoding='utf-8')


def judged_candidates(root):
    """What `judge` asks (the shipping question): the candidates a Jev run would answer into the shared ledger."""
    rc, out = cli('judge', '--root', str(root), '--date', DATE, '--include-candidates')
    return out['judgment_candidates']


def write_labels(path, key, own_value=True, swap_value=False):
    labels = {i['id']: (own_value if i['kind'] == 'own' else swap_value) for i in key['items']}
    path.write_text(json.dumps({'answer_key': key['answer_key'], 'labels': labels}), encoding='utf-8')
    return path


def measure(root, out, labels):
    return cli('answer-key', '--root', str(root), '--measure', str(labels), '--key', str(out / 'answer-key.json'))


def record(root, out, tmp_path, *, limit=500, execute=True, score=0.95, calls=None):
    """The recording tool, driven the way production drives it, with a fake Jev behind it."""
    items = answer_key.load_items(*_model(root), out / 'answer-key.json')
    return answer_key.record(items, out, limit, execute=execute, _usage_path=_usage(tmp_path),
                             _transport=responder(calls if calls is not None else [], score=score))


def _usage(tmp_path):
    """One usage ledger per test, shared by every recording run (the cumulative ceiling is the real one's job)."""
    path = tmp_path / 'usage/classification-usage.json'
    return path if path.exists() else usage_ledger(tmp_path)


def _model(root):
    """The same loading `infoarch-author answer-key` does: the model, the approved keys, the shared policy."""
    import check
    import positioning
    report = check.Report('answer-key', DATE)
    m = check.load(root, report)
    approved = positioning.approved_keys(m, check.read_approvals(root, report))
    jlmod = judge._jl()
    return m, approved, jlmod.load_policy(), jlmod


def test_own_pairs_ask_the_shipping_question_and_swap_pairs_do_not(made):
    root, out, key = made
    shipping = {r['judgment_hash'] for r in judged_candidates(root)}
    own = {i['judgment_hash'] for i in key['items'] if i['kind'] == 'own'}
    swap = {i['judgment_hash'] for i in key['items'] if i['kind'] == 'swap'}
    assert own and swap
    assert own <= shipping, 'an own pair must be the very candidate the ship gate asks, so one answer serves both'
    assert not swap & shipping, 'a swap pair must never look like a shipping judgment'


def test_without_the_recording_tool_the_swap_side_is_missing_and_nothing_is_measured(made, tmp_path):
    root, out, key = made
    theta = jl.load_policy()['theta']
    write_shared_ledger(root, [{'candidate': c, 'score': 0.0} for c in judged_candidates(root)])
    labels = write_labels(tmp_path / 'labels.json', key)
    rc, result = measure(root, out, labels)
    assert result['status'] == 'unmeasurable', result
    assert result['counts']['uncovered'] == sum(i['kind'] == 'swap' for i in key['items'])
    assert 'agreement' not in result


def test_the_real_recording_path_makes_accuracy_measurable(made, tmp_path):
    root, out, key = made
    write_shared_ledger(root, [{'candidate': c, 'score': 0.0} for c in judged_candidates(root)])  # own: conveys
    calls = []
    result = record(root, out, tmp_path, calls=calls)
    swaps = sum(i['kind'] == 'swap' for i in key['items'])
    assert result['status'] == 'answered' and sum(calls) == swaps, result  # only swap pairs are asked
    assert (out / MEASUREMENT).is_file()
    rc, measured = measure(root, out, write_labels(tmp_path / 'labels.json', key))
    assert rc == 0 and measured['status'] == 'measured' and measured['agreement'] == 1.0, measured
    assert measured['counts']['uncovered'] == 0


def test_a_swap_answer_never_enters_the_shared_ledger_or_the_ship_gate(made, tmp_path):
    root, out, key = made
    write_shared_ledger(root, [{'candidate': c, 'score': 0.0} for c in judged_candidates(root)])
    path = layer_paths.path(root, 'judgments')
    before = hashlib.sha256(path.read_bytes()).hexdigest()
    record(root, out, tmp_path)
    assert hashlib.sha256(path.read_bytes()).hexdigest() == before  # the recorder never touches ux/judgments.json
    swap = {i['judgment_hash'] for i in key['items'] if i['kind'] == 'swap'}
    shared = {r['candidate']['judgment_hash'] for r in json.loads(path.read_text(encoding='utf-8'))['rows']}
    assert not swap & shared
    rc, judged = cli('judge', '--root', str(root), '--date', DATE)
    assert rc == 0 and judged['uncovered'] == [], judged  # judge is the same with or without the measurement ledger
    assert len(judged['screened']) == len(shared)


def test_no_script_but_the_answer_key_ones_names_the_measurement_ledger():
    named = [p.name for p in SCRIPTS.glob('*.py') if MEASUREMENT in p.read_text(encoding='utf-8')]
    assert named and set(named) <= {'answer_key.py', 'infoarch_author.py'}, named


def test_the_recording_tool_previews_without_execute_and_respects_the_limit(made, tmp_path):
    root, out, key = made
    calls = []
    preview = record(root, out, tmp_path, execute=False, calls=calls)
    assert preview['status'] == 'dry-run' and calls == [] and not (out / MEASUREMENT).exists()
    swaps = sum(i['kind'] == 'swap' for i in key['items'])
    assert preview['selected'] == swaps
    first = record(root, out, tmp_path, limit=1, calls=calls)
    assert sum(calls) == 1 and first['answered'] == 1
    second = record(root, out, tmp_path, limit=500, calls=calls)
    assert sum(calls) == swaps  # the second run continues from where the first stopped
    assert len(json.loads((out / MEASUREMENT).read_text(encoding='utf-8'))['rows']) == swaps


def test_only_swap_pairs_not_yet_answered_are_asked(made, tmp_path):
    root, out, key = made
    calls = []
    record(root, out, tmp_path, calls=calls)
    again = record(root, out, tmp_path, calls=calls)
    assert again['status'] == 'already-complete' and sum(calls) == sum(i['kind'] == 'swap' for i in key['items'])


def test_a_damaged_or_symlinked_measurement_ledger_is_refused_not_ignored(made, tmp_path):
    root, out, key = made
    write_shared_ledger(root, [{'candidate': c, 'score': 0.0} for c in judged_candidates(root)])
    labels = write_labels(tmp_path / 'labels.json', key)
    (out / MEASUREMENT).write_text('{"schema_version": 1}', encoding='utf-8')
    rc, result = measure(root, out, labels)
    assert rc == 2, result
    (out / MEASUREMENT).unlink()
    target = tmp_path / 'elsewhere.json'
    target.write_text('{}', encoding='utf-8')
    (out / MEASUREMENT).symlink_to(target)
    rc, result = measure(root, out, labels)
    assert rc == 2, result


def test_cli_record_previews_and_never_calls_without_execute(made, tmp_path):
    root, out, key = made
    rc, result = cli('answer-key', '--root', str(root), '--key', str(out / 'answer-key.json'), '--record', '--limit', '5')
    assert rc == 0 and result['status'] == 'dry-run' and result['selected'] == min(5, sum(i['kind'] == 'swap' for i in key['items']))
    assert not (out / MEASUREMENT).exists()
