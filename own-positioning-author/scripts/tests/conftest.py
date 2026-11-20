"""Test import paths: this skill's scripts and tests."""
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
for path in (SCRIPTS, SCRIPTS / 'tests'):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))
