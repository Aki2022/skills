"""Shared loading, findings and exit codes for own-infoarch-author."""
from __future__ import annotations

import json
import platform
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
REFS = SKILL / 'references'
VERSION = '0.2.0'
INFO_ARCH_DIR = Path('design/info_arch')
LEVELS = {'fail': 2, 'warn': 1, 'info': 0}


class DependencyMissing(RuntimeError):
    pass


def load_ref(name):
    return json.loads((REFS / name).read_text(encoding='utf-8'))


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


def validator_for(schema_name):
    try:
        import jsonschema
    except ImportError as exc:
        raise DependencyMissing('jsonschema is required; set DESIGN_AUTHOR_PYTHON to an interpreter that has it') from exc
    schema = load_ref(f'schema/{schema_name}')
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


@dataclass
class Report:
    level: str
    evaluated_on: str
    findings: list = field(default_factory=list)
    targets: dict = field(default_factory=dict)
    uncovered: list = field(default_factory=list)

    def add(self, level, code, message, path='', line=0, **detail):
        if level not in LEVELS:
            raise ValueError(level)
        self.findings.append(Finding(level, code, message, str(path), line, detail))

    def codes(self, level=None):
        return [f.code for f in self.findings if level is None or f.level == level]

    @property
    def exit_code(self):
        return max([LEVELS[f.level] for f in self.findings], default=0)

    def to_dict(self):
        counts = {k: sum(1 for f in self.findings if f.level == k) for k in LEVELS}
        return {
            'tool': 'own-infoarch-author', 'version': VERSION, 'level': self.level,
            'evaluated_on': self.evaluated_on, 'exit_code': self.exit_code,
            'environment': {'python': platform.python_version(), 'platform': sys.platform},
            'counts': counts, 'targets': self.targets, 'uncovered': self.uncovered,
            'findings': [asdict(f) for f in self.findings],
        }
