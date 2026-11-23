"""Shared loading, findings and exit codes for the positioning and infoarch layers.

own-infoarch-author imports this module from its sibling skill: one YAML reader, one finding/report shape, one id
and ledger vocabulary, so the two layers cannot read the same file two ways.
"""
from __future__ import annotations

import json
import platform
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
REFS = SKILL / 'references'

# Where the layers live is own-design-author's one rule (references/layer-paths.json, read by layer_paths.py). A missing
# sibling is an error, never a silent fallback to a spelled path.
_DESIGN_SCRIPTS = SKILL.parent / 'own-design-author' / 'scripts'
if not (_DESIGN_SCRIPTS / 'layer_paths.py').is_file():
    raise RuntimeError('own-design-author/scripts/layer_paths.py が無い（layer の置き場所の規則は design が持つ）')
if str(_DESIGN_SCRIPTS) not in sys.path:
    sys.path.append(str(_DESIGN_SCRIPTS))
import layer_paths  # noqa: E402

VERSION = '0.4.0'
INFO_ARCH_DIR = Path(layer_paths.rel('infoarch'))  # the repository-relative name only; a road that touches a repo uses layer_paths.path
LEVELS = {'fail': 2, 'warn': 1, 'info': 0}


class DependencyMissing(RuntimeError):
    pass


def load_ref(name, refs=None):
    return json.loads(((refs or REFS) / name).read_text(encoding='utf-8'))


class YamlError(ValueError):
    pass


def _loader():
    try:
        import yaml
    except ImportError as exc:  # same runtime contract as own-design-author: never auto-install
        raise DependencyMissing('PyYAML is required; set DESIGN_AUTHOR_PYTHON to an interpreter that has it') from exc

    class StrictLoader(yaml.SafeLoader):
        """Rejects duplicate keys (PyYAML keeps the last one silently) and keeps dates as strings."""

    StrictLoader.yaml_implicit_resolvers = {
        k: [(tag, rx) for tag, rx in v if tag != 'tag:yaml.org,2002:timestamp']
        for k, v in yaml.SafeLoader.yaml_implicit_resolvers.items()}

    def construct_mapping(loader, node, deep=False):
        seen = set()
        for key_node, _ in node.value:
            key = loader.construct_object(key_node, deep=deep)
            if key in seen:
                raise yaml.constructor.ConstructorError(None, None, f'duplicate key {key!r}', key_node.start_mark)
            seen.add(key)
        return yaml.SafeLoader.construct_mapping(loader, node, deep)

    StrictLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, construct_mapping)
    return yaml, StrictLoader


def safe_load_text(text):
    yaml, loader = _loader()
    try:
        return yaml.load(text, Loader=loader)  # noqa: S506 - StrictLoader derives from SafeLoader
    except yaml.YAMLError as exc:
        raise YamlError(str(exc).replace('\n', ' ')) from exc


def load_yaml(path):
    try:
        text = Path(path).read_bytes().decode('utf-8')
    except UnicodeDecodeError as exc:
        raise YamlError(f'not UTF-8: {exc}') from exc
    return safe_load_text(text)


def validator_for(schema_name, refs=None):
    try:
        import jsonschema
    except ImportError as exc:
        raise DependencyMissing('jsonschema is required; set DESIGN_AUTHOR_PYTHON to an interpreter that has it') from exc
    schema = load_ref(f'schema/{schema_name}', refs)
    cls = jsonschema.validators.validator_for(schema)
    cls.check_schema(schema)
    return cls(schema)


@dataclass
class Finding:
    level: str
    code: str
    message: str
    path: str = ''
    line: int = 0
    detail: dict = field(default_factory=dict)


# The five categories of a report (SPEC-information-architecture-standardization, Quality 2). What each one says:
#   executed  a check that ran: {id, targets}
#   skipped   a check this command does not run, because another command owns it: {id, reason}
#   missing   an input that was needed and is absent, so nothing could be checked: {id, reason}
#   uncovered a target that exists and was not looked at (a judgment not in the ledger, a non-HTML output): {id|path, reason}
#   screened  a judgment that exists in the ledger and was read (never a guarantee that the text is fine)
# `check`, `blueprint` and the positioning layer's `check` carry all five keys, empty or not (`categorized`). judge,
# evaluate and ship keep the shape they had: they own `uncovered`, `screened` and `missing` in their own terms.
CATEGORIES = ('executed', 'skipped', 'missing', 'uncovered', 'screened')

# A failure that says an input is absent names it in `missing` (the id is the input, not the code).
MISSING_INPUT = {
    'IA_MISSING': ('ia-positioning', 'positioning の SSOT ファイルが無い（置き場所は layer の規則が決める）'),
    'IA_NO_VALUES': ('ia-values', 'positioning に提供価値が 0 件'),
    'IA_NO_CHANNELS': ('ia-channels', 'channels/*.yaml が 0 件'),
    'IA_NO_OUTLINES': ('ia-outlines', 'outlines/*.yaml が 0 件'),
    'IA_BLUEPRINT_NO_SCREENS': ('ia-blueprint-screens', 'blueprint/screens/*.yaml が 0 件'),
    'IA_BLUEPRINT_PRODUCT_MISSING': ('ia-blueprint-product', 'blueprint/product.yaml が無い'),
}


@dataclass
class Report:
    level: str
    evaluated_on: str
    findings: list = field(default_factory=list)
    targets: dict = field(default_factory=dict)
    uncovered: list = field(default_factory=list)
    tool: str = 'own-infoarch-author'
    categorized: bool = False
    executed: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    missing: list = field(default_factory=list)
    screened: list = field(default_factory=list)

    def add(self, level, code, message, path='', line=0, **detail):
        if level not in LEVELS:
            raise ValueError(level)
        self.findings.append(Finding(level, code, message, str(path), line, detail))

    def codes(self, level=None):
        return [f.code for f in self.findings if level is None or f.level == level]

    @property
    def exit_code(self):
        return max([LEVELS[f.level] for f in self.findings], default=0)

    def categories(self):
        """-> the five categories. `missing` is what was named plus the inputs that a failure says are absent."""
        missing = list(self.missing)
        named = {m['id'] for m in missing}
        for f in self.findings:
            entry = MISSING_INPUT.get(f.code)
            if f.level == 'fail' and entry and entry[0] not in named:
                named.add(entry[0])
                missing.append({'id': entry[0], 'reason': entry[1], 'code': f.code})
        return {'executed': self.executed, 'skipped': self.skipped, 'missing': missing,
                'uncovered': self.uncovered, 'screened': self.screened}

    def to_dict(self):
        counts = {k: sum(1 for f in self.findings if f.level == k) for k in LEVELS}
        out = {
            'tool': self.tool, 'version': VERSION, 'level': self.level,
            'evaluated_on': self.evaluated_on, 'exit_code': self.exit_code,
            'environment': {'python': platform.python_version(), 'platform': sys.platform},
            'counts': counts, 'targets': self.targets, 'uncovered': self.uncovered,
            'findings': [asdict(f) for f in self.findings],
        }
        if self.categorized:
            out.update(self.categories())
        return out
