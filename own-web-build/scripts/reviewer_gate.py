#!/usr/bin/env python3
"""The gate in front of the independent reviewer (marketing scoring, or any reviewer that reads a rendered page).

A reviewer costs tens of thousands of tokens per head, and a page that did not render, was not measured, or still fails
the measurement only produces display-defect findings that bury what the reviewer was meant to judge. So the reviewer
may start only when the measurement after rendering ran, measured something, knows its profile and is green.

Both ways into the audit pass through here, and the way in is recorded:
  --route entry   the entry skill (own-deliverable-check) hands the target on
  --route direct  the audit skill (own-website-audit) is called by itself

    reviewer_gate.py --route entry|direct --report <measurement.json> --ledger <ledger.jsonl> [--launch <command...>]

exit 0 = go (the reviewer may start), exit 3 = stop (it may not), exit 2 = the gate itself was called wrongly.
Every decision is appended to the ledger with its reason and how many reviewers it started. With --launch the gate
starts the given command only on go, which is how a test counts launches through a channel the gate does not own.

Stops, each with its own reason:
  report-missing / report-unreadable   no measurement result exists (the run failed before writing one)
  execution-failed                     the run stopped part-way (a render-execution-failed failure or execution_error)
  no-profile / unknown-profile         the measurement did not say which profile judged the page, or an unknown one
  zero-measured                        nothing was actually measured (0 runs or 0 targets)
  failures-remain                      the measurement still has failures: fix the page first, then measure again
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'own-design-author/scripts'))
import measurement_profiles  # noqa: E402

ROUTES = ('entry', 'direct')


def decide(report, report_exists=True):
    """('go', reason) or ('stop', reason). `report` is the parsed measurement, or None when it could not be read."""
    if not report_exists:
        return 'stop', 'report-missing'
    if not isinstance(report, dict):
        return 'stop', 'report-unreadable'
    failures = report.get('failures') if isinstance(report.get('failures'), list) else None
    if report.get('execution_error') or any(isinstance(f, dict) and f.get('id') == 'render-execution-failed'
                                            for f in (failures or [])):
        return 'stop', 'execution-failed'
    profile = report.get('profile')
    if profile is None:
        return 'stop', 'no-profile'
    if profile not in measurement_profiles.profile_ids():
        return 'stop', 'unknown-profile'
    measured = report.get('measured')
    if (not isinstance(measured, dict) or type(measured.get('runs')) is not int or type(measured.get('targets')) is not int
            or measured['runs'] <= 0 or measured['targets'] <= 0):
        return 'stop', 'zero-measured'
    if failures is None or failures:
        return 'stop', 'failures-remain'
    return 'go', 'measured-green'


def main(argv):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--route', required=True, choices=ROUTES)
    parser.add_argument('--report', required=True, type=Path)
    parser.add_argument('--ledger', required=True, type=Path)
    parser.add_argument('--launch', nargs=argparse.REMAINDER, default=[])
    args = parser.parse_args(argv[1:])
    report, exists = None, args.report.is_file()
    if exists:
        try:
            report = json.loads(args.report.read_text(encoding='utf-8'))
        except (OSError, ValueError):
            report = None
    decision, reason = decide(report, exists)
    launched, launch_exit = 0, None
    if decision == 'go' and args.launch:
        launch_exit = subprocess.run(args.launch).returncode
        launched = 1 if launch_exit == 0 else 0
    row = {'route': args.route, 'decision': decision, 'reason': reason, 'launched': launched}
    if launch_exit is not None:
        row['launch_exit'] = launch_exit
    args.ledger.parent.mkdir(parents=True, exist_ok=True)
    with args.ledger.open('a', encoding='utf-8') as handle:
        handle.write(json.dumps(row, ensure_ascii=False) + '\n')
    print(json.dumps(row, ensure_ascii=False))
    return 0 if decision == 'go' else 3


if __name__ == '__main__':
    sys.exit(main(sys.argv))
