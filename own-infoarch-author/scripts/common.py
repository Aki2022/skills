"""Shared loading for own-infoarch-author.

The YAML reader, the finding and report shapes and the schema validator are the positioning layer's (one reader for
the two layers: own-positioning-author/scripts/ia_common.py). This module adds what is this skill's own: its
references, and the sys.path entries that let it import its siblings (the positioning layer, and the docs
dictionary reader that own-doc-update owns). A missing sibling is an error, never a silent fallback.
"""
from __future__ import annotations

import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
REFS = SKILL / 'references'
SIBLING_SCRIPTS = {name: SKILL.parent / name / 'scripts' for name in ('own-positioning-author', 'own-doc-update')}

for _name, _path in SIBLING_SCRIPTS.items():
    if not _path.is_dir():
        raise RuntimeError(f'{_name}/scripts が無い（own-infoarch-author は positioning と辞書を兄弟 skill から読む）')
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

import ia_common  # noqa: E402
from ia_common import (  # noqa: E402,F401
    INFO_ARCH_DIR, LEVELS, VERSION, DependencyMissing, Finding, Report, YamlError, load_yaml, safe_load_text)


def load_ref(name):
    return ia_common.load_ref(name, REFS)


def validator_for(schema_name):
    """A schema of this skill (channel, outline) is read from here; the SSOT's schemas are the positioning layer's."""
    if (REFS / 'schema' / schema_name).is_file():
        return ia_common.validator_for(schema_name, REFS)
    return ia_common.validator_for(schema_name)
