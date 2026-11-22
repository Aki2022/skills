"""Shared cases for the independent-reviewer gate, used by both routes (the entry skill and the direct audit skill).

A launch is counted through a channel the gate does not own: the command the gate may start appends a line to a counter
file, and the test reads that file. The gate's own ledger must agree with it."""
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
GATE_REL = 'own-web-build/scripts/reviewer_gate.py'
GATE = SKILLS / GATE_REL
# situation -> the reason the gate must give. Each reason is its own condition: taking one out must change this table's outcome,
# even where another condition would still have stopped the run (overlap must not hide a removed condition).
STOP_CASES = {'zero-measured': 'zero-measured', 'execution-failed': 'execution-failed', 'execution-no-report': 'report-missing',
              'unreadable-report': 'report-unreadable', 'no-profile': 'no-profile', 'unknown-profile': 'unknown-profile',
              'failures-remain': 'failures-remain'}


def normal_report():
    return {'url': 'http://127.0.0.1/', 'profile': 'lp', 'state': None, 'skipped': [], 'applied': ['contrast'],
            'measured': {'runs': 2, 'targets': 41}, 'failures': [], 'warnings': [], 'uncovered': [],
            'runs': [{'scheme': 'light', 'width': 375}, {'scheme': 'dark', 'width': 375}]}


def make_report(case):
    """The report a measurement left for each situation; None means it left no report at all."""
    report = normal_report()
    if case == 'normal':
        return report
    if case == 'zero-measured':
        report['measured'] = {'runs': 2, 'targets': 0}
    elif case == 'execution-failed':
        report['execution_error'] = 'render-unstable'
        report['failures'] = [{'id': 'render-execution-failed', 'stage': 'collect', 'error': 'render-unstable'}]
        report['measured'] = {'runs': 0, 'targets': 0}
    elif case == 'execution-no-report':
        return None
    elif case == 'unreadable-report':
        return '{ this is not json'
    elif case == 'no-profile':
        del report['profile']
    elif case == 'unknown-profile':
        report['profile'] = 'print'
    elif case == 'failures-remain':
        report['failures'] = [{'id': 'contrast', 'at': 'light/375px'}]
    else:
        raise ValueError(case)
    return report


def launcher(tmp_path):
    """A command standing in for 'start the independent reviewer': it appends one line to a counter file."""
    counter = tmp_path / 'reviewer-launches.txt'
    script = tmp_path / 'fake_reviewer.py'
    script.write_text('import sys\nopen(sys.argv[1], "a").write("started\\n")\n')
    return counter, [sys.executable, str(script), str(counter)]


def launches(counter):
    return len(counter.read_text().splitlines()) if counter.exists() else 0


def run_route(tmp_path, route, case, gate=GATE):
    """Run one route on one situation. Returns (exit code, launches seen by the counter, ledger rows)."""
    tmp_path.mkdir(parents=True, exist_ok=True)
    report_path = tmp_path / 'measurement.json'
    report = make_report(case)
    if report is not None:
        report_path.write_text(report if isinstance(report, str) else json.dumps(report))
    counter, command = launcher(tmp_path)
    ledger = tmp_path / 'ledger.jsonl'
    result = subprocess.run([sys.executable, str(gate), '--route', route, '--report', str(report_path),
                             '--ledger', str(ledger), '--launch', *command], capture_output=True, text=True)
    rows = [json.loads(line) for line in ledger.read_text().splitlines()] if ledger.exists() else []
    return result, launches(counter), rows


def mutated_gate(tmp_path):
    """The gate with its stopping conditions taken out: it always lets the reviewer start."""
    source = GATE.read_text()
    patched, count = re.subn(r'def decide\(', 'def decide_original(', source, count=1)
    assert count == 1, 'the gate has no decide() to take the conditions out of'
    marker = "if __name__ == '__main__':"
    assert marker in patched, 'the gate has no entry point to put the replacement before'
    patched = patched.replace(marker, 'def decide(report, report_exists=True):\n    return "go", "mutated"\n\n\n' + marker)
    path = tmp_path / 'mutated'
    (path / 'own-web-build/scripts').mkdir(parents=True)
    (path / 'own-design-author').mkdir()
    shutil.copytree(SKILLS / 'own-design-author/scripts', path / 'own-design-author/scripts',
                    ignore=shutil.ignore_patterns('tests', '__pycache__'))
    shutil.copytree(SKILLS / 'own-design-author/references', path / 'own-design-author/references')
    target = path / 'own-web-build/scripts/reviewer_gate.py'
    target.write_text(patched)
    return target


def assert_route(tmp_path, route):
    """The contract of one route: every stopping situation starts the reviewer 0 times and says why; a normal
    measurement starts it at least once; the ledger agrees with the counter; taking the conditions out turns it red."""
    for case, reason in STOP_CASES.items():
        result, started, rows = run_route(tmp_path / ('stop-' + case), route, case)
        assert started == 0, f'{route}/{case}: the reviewer was started {started} time(s)\n{result.stdout}{result.stderr}'
        assert result.returncode == 3, f'{route}/{case}: expected exit 3, got {result.returncode}\n{result.stderr}'
        assert [r['decision'] for r in rows] == ['stop'] and rows[0]['route'] == route, rows
        assert rows[0]['reason'] == reason, f'{route}/{case}: stopped for {rows[0]["reason"]!r}, expected {reason!r}'
        assert rows[0]['launched'] == 0
    result, started, rows = run_route(tmp_path / 'normal', route, 'normal')
    assert result.returncode == 0, result.stdout + result.stderr
    assert started >= 1, f'{route}/normal: the control did not start the reviewer'
    assert [r['decision'] for r in rows] == ['go'] and rows[0]['launched'] == started
    # With the conditions taken out, the same observation sees the reviewer start in every stopping situation.
    gate = mutated_gate(tmp_path / 'mut')
    leaked = {case: run_route(tmp_path / ('mut-' + case), route, case, gate=gate)[1] for case in STOP_CASES}
    assert all(n >= 1 for n in leaked.values()), f'the test does not see a launch when the gate is removed: {leaked}'
