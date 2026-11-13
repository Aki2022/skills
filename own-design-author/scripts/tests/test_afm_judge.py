"""AFM stage: adopted items that code cannot measure are judged against each artifact by the on-device model."""
import json
from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from afm_judge import judge, AfmUnavailable

STUB = '''#!/usr/bin/env python3
import json, os, re, sys
calls = os.environ['STUB_CALLS']
with open(calls, 'a') as f: f.write(' '.join(sys.argv[1:2]) + '\\n')
if sys.argv[1] == 'models':
    # mirrors the real CLI: exit 1 because the cloud model is unavailable, on-device model marked ready
    mark = '\u2717' if os.environ.get('STUB_NO_MODEL') else '\u2713'
    print('  Apple Foundation Models\\n  ' + mark + ' system (AFM 3 Core)\\n  \u2717 pcc  (unavailable)')
    raise SystemExit(1)
texts = [sys.argv[i + 1] for i, a in enumerate(sys.argv) if a == '--text']
ids = re.findall(r'^\\[(dads-[0-9]+)\\]', texts[0], re.M)
mode = os.environ.get('STUB_MODE', 'pass')
if mode == 'garbage': print('not json'); raise SystemExit(0)
if mode == 'drop': ids = ids[1:]
bad = set(os.environ.get('STUB_VIOLATE', '').split(','))
print(json.dumps({'verdicts': [{'id': i, 'applies': True, 'violates': i in bad, 'reason': 'stub'} for i in ids]}))
'''


@pytest.fixture
def stub(tmp_path, monkeypatch):
    path = tmp_path / 'fm'
    path.write_text(STUB)
    path.chmod(0o755)
    calls = tmp_path / 'calls.log'
    monkeypatch.setenv('STUB_CALLS', str(calls))
    return path, calls, tmp_path / 'cache'


CANDIDATES = [{'id': f'dads-00000{i}', 'rule': f'rule {i}', 'evidence': f'evidence {i}'} for i in range(1, 4)]


def artifacts(tmp_path):
    image = tmp_path / 'shot.png'
    image.write_bytes(b'png-bytes')
    return [{'id': 'L2:DESIGN_WEB_URL:1280:light', 'kind': 'image', 'path': str(image)},
            {'id': 'html:site/index.html', 'kind': 'text', 'text': '<main>hello</main>'}]


def calls_of(calls, verb='respond'):
    return [line for line in calls.read_text().splitlines() if line == verb] if calls.exists() else []


def test_every_item_is_judged_on_every_artifact_and_cached(stub, tmp_path):
    command, calls, cache = stub
    result = judge(CANDIDATES, artifacts(tmp_path), command=str(command), cache_dir=cache)
    assert result['errors'] == [] and result['uncovered'] == [] and result['warnings'] == []
    assert len(result['screened']) == 6
    first = len(calls_of(calls))
    assert first == 2  # one batch per artifact
    again = judge(CANDIDATES, artifacts(tmp_path), command=str(command), cache_dir=cache)
    assert len(calls_of(calls)) == first and again['cached'] == 6


def test_violation_is_reported_as_advisory_warning(stub, tmp_path, monkeypatch):
    command, _, cache = stub
    monkeypatch.setenv('STUB_VIOLATE', 'dads-000002')
    result = judge(CANDIDATES, artifacts(tmp_path)[:1], command=str(command), cache_dir=cache)
    assert result['errors'] == []
    assert result['warnings'] == ['AFM advisory: dads-000002 on L2:DESIGN_WEB_URL:1280:light: stub']


def test_changed_artifact_is_judged_again(stub, tmp_path):
    command, calls, cache = stub
    rows = artifacts(tmp_path)[:1]
    judge(CANDIDATES, rows, command=str(command), cache_dir=cache)
    Path(rows[0]['path']).write_bytes(b'changed')
    judge(CANDIDATES, rows, command=str(command), cache_dir=cache)
    assert len(calls_of(calls)) == 2


@pytest.mark.parametrize('mode', ['garbage', 'drop'])
def test_missing_or_invalid_verdicts_leave_items_uncovered(stub, tmp_path, monkeypatch, mode):
    command, _, cache = stub
    monkeypatch.setenv('STUB_MODE', mode)
    result = judge(CANDIDATES, artifacts(tmp_path)[:1], command=str(command), cache_dir=cache)
    assert result['uncovered'], result
    assert all(row['artifact'] == 'L2:DESIGN_WEB_URL:1280:light' for row in result['uncovered'])


def test_missing_model_command_stops(tmp_path):
    with pytest.raises(AfmUnavailable):
        judge(CANDIDATES, artifacts(tmp_path), command=str(tmp_path / 'absent'), cache_dir=tmp_path / 'cache')


def test_model_marked_unavailable_stops(stub, tmp_path, monkeypatch):
    command, _, cache = stub
    monkeypatch.setenv('STUB_NO_MODEL', '1')
    with pytest.raises(AfmUnavailable):
        judge(CANDIDATES, artifacts(tmp_path), command=str(command), cache_dir=cache)


def test_no_artifacts_is_not_a_pass(stub):
    command, _, cache = stub
    result = judge(CANDIDATES, [], command=str(command), cache_dir=cache)
    assert result['uncovered'] and result['errors'] == ['AFM has zero artifacts to judge']
