"""Test import paths: this skill's scripts, and its siblings' (the positioning layer and the docs dictionary reader)."""
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
ROOT = SCRIPTS.parents[1]
for path in (ROOT / 'own-doc-update/scripts', ROOT / 'own-design-author/scripts', ROOT / 'own-positioning-author/scripts',
             ROOT / 'own-positioning-author/scripts/tests', SCRIPTS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))
