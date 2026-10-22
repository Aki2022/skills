#!/usr/bin/env python3
"""Run the same L0 --ship for hooks and mise; warning-only results allow shipment."""
import argparse
import json
import os
from pathlib import Path
import subprocess


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', default='.')
    parser.add_argument('--catalog-dir', type=Path)
    args = parser.parse_args()
    command = os.environ.get('DESIGN_AUTHOR_COMMAND') or str(Path(__file__).with_name('design-author'))
    argv = [command, 'check', '--ship', '--repo', args.repo]
    if args.catalog_dir:
        argv += ['--catalog-dir', str(args.catalog_dir)]
    try:
        run = subprocess.run(argv, capture_output=True, text=True)
        result = json.loads(run.stdout)
        if run.returncode not in (0, 1, 2) or result.get('exit_code') != run.returncode:
            raise ValueError('invalid L0 result')
    except (OSError, ValueError):
        print(json.dumps({'exit_code': 2, 'errors': ['L0 unavailable or invalid; shipment stopped']}))
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 2 if run.returncode == 2 else 0


if __name__ == '__main__':
    raise SystemExit(main())
