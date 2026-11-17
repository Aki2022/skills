"""AFM stage: adopted items that code cannot measure are judged against each artifact by the on-device model."""
import json
from pathlib import Path
import shutil
import struct
import sys
import zlib

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
if mode == 'overflow':
    # the real CLI refuses an image whose tokens plus the prompt exceed the window: model it by image height
    image = sys.argv[sys.argv.index('--image') + 1] if '--image' in sys.argv else None
    if image:
        import struct
        with open(image, 'rb') as f: head = f.read(24)
        if struct.unpack('>I', head[20:24])[0] > int(os.environ.get('STUB_MAX_HEIGHT', '100')):
            sys.stderr.write("Error: The session's transcript exceeded the model's context size.\\n"); raise SystemExit(1)
        with open(os.environ['STUB_CALLS'], 'a') as f: f.write('height ' + str(struct.unpack('>I', head[20:24])[0]) + '\\n')
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


def png(path, width, height):
    """A real, tiny grey PNG so the shrink step (macOS sips) can read it."""
    raw = b''.join(b'\x00' + b'\x80' * width for _ in range(height))
    chunk = lambda kind, data: (struct.pack('>I', len(data)) + kind + data
                                + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff))
    path.write_bytes(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 0, 0, 0, 0))
                     + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))
    return path


@pytest.mark.skipif(shutil.which('sips') is None, reason='image shrinking uses macOS sips, as the fm CLI does')
def test_image_too_large_for_the_window_is_judged_on_a_smaller_copy(stub, tmp_path, monkeypatch):
    # A tall page plus a long rule overflowed the 4,096-token window and every pair stayed unjudged.
    command, calls, cache = stub
    monkeypatch.setenv('STUB_MODE', 'overflow')
    image = png(tmp_path / 'tall.png', 40, 300)
    before = image.read_bytes()
    result = judge(CANDIDATES, [{'id': 'L2:DESIGN_WEB_URL:1280:light', 'kind': 'image', 'path': str(image)}],
                   command=str(command), cache_dir=cache)
    assert result['uncovered'] == [] and len(result['screened']) == 3, result
    assert result['shrunk'] == 3
    heights = [int(line.split()[1]) for line in calls.read_text().splitlines() if line.startswith('height ')]
    assert heights and max(heights) <= 100
    assert image.read_bytes() == before  # the L2 evidence image itself is never rewritten


@pytest.mark.skipif(shutil.which('sips') is None, reason='image shrinking uses macOS sips, as the fm CLI does')
def test_image_that_never_fits_stays_uncovered(stub, tmp_path, monkeypatch):
    command, _, cache = stub
    monkeypatch.setenv('STUB_MODE', 'overflow')
    monkeypatch.setenv('STUB_MAX_HEIGHT', '1')
    image = png(tmp_path / 'tall.png', 40, 300)
    result = judge(CANDIDATES, [{'id': 'L2:DESIGN_WEB_URL:1280:light', 'kind': 'image', 'path': str(image)}],
                   command=str(command), cache_dir=cache)
    assert len(result['uncovered']) == 3 and result['screened'] == []
