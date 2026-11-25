"""The layout pattern vocabulary (SPEC-design-system-standardization requirement 15).

The skill holds a versioned catalog (catalog/layout-patterns.json) of media-neutral patterns a layout plan names per
block (own-web-build). A repository's design.yaml holds only a diff:

    layout_patterns:
      disabled: [<catalog names the repository does not use>]
      added: {<name>: {description, columns, band, images, separators, narrow, functions}}

An added pattern follows the item lifecycle: it is proposed until the owner approves it, and the approval is a ledger row
(`ds:<media>:pattern:<name>`) bound by hash to the pattern's definition. The file never declares its own state (the AI never
sets approved). L0 resolves the vocabulary into the manifest, so build reads one place.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from dtcg import value_hash

CATALOG_FILE = Path(__file__).resolve().parents[1] / 'catalog' / 'layout-patterns.json'
FIELDS = ('description', 'columns', 'band', 'images', 'separators', 'narrow', 'functions')
FUNCTIONS = ('identify', 'claim', 'fact', 'evidence', 'offer', 'action', 'source')  # docs dictionary term-function-*
IMAGES = ('none', 'left', 'right', 'top', 'background', 'inline', 'avatar', 'icon', 'logo')
SEPARATORS = ('none', 'rule', 'space', 'tint', 'number')
NAME = re.compile(r'^[a-z][a-z0-9]*(-[a-z0-9]+)*$')
DIFF_KEYS = {'disabled', 'added'}


def validate_pattern(name, row):
    """The problems of one pattern definition (empty when it is whole). Unknown fields are problems too."""
    problems = []
    if not isinstance(name, str) or not NAME.match(name):
        problems.append(f'name {name!r} is not lower-case kebab-case')
    if not isinstance(row, dict):
        return problems + ['definition is not an object']
    missing = [k for k in FIELDS if k not in row]
    extra = sorted(set(row) - set(FIELDS))
    if missing:
        problems.append(f'missing {missing}')
    if extra:
        problems.append(f'unknown fields {extra}')
    if 'description' in row and (not isinstance(row['description'], str) or not row['description'].strip()):
        problems.append('description must be non-empty text')
    if 'columns' in row:
        c = row['columns']
        if not (c == 'n' or (isinstance(c, int) and not isinstance(c, bool) and c >= 1)):
            problems.append('columns must be an integer >= 1 or "n"')
    if 'band' in row and not isinstance(row['band'], bool):
        problems.append('band must be true or false')
    for key, allowed in (('images', IMAGES), ('separators', SEPARATORS), ('functions', FUNCTIONS)):
        if key in row:
            value = row[key]
            if not isinstance(value, list) or not value or len(set(map(str, value))) != len(value) \
                    or any(v not in allowed for v in value):
                problems.append(f'{key} must be a non-empty list of distinct values from {list(allowed)}')
    if 'narrow' in row and (not isinstance(row['narrow'], str) or not NAME.match(row['narrow'])):
        problems.append('narrow must be a pattern name')
    return problems


def load_catalog(path=CATALOG_FILE):
    data = json.loads(Path(path).read_text(encoding='utf-8'))
    if data.get('schema_version') != 1 or not isinstance(data.get('version'), str) or not data['version'] \
            or not isinstance(data.get('patterns'), dict) or not data['patterns']:
        raise ValueError('layout pattern catalog is malformed')
    for name, row in data['patterns'].items():
        problems = validate_pattern(name, row)
        if problems:
            raise ValueError(f'layout pattern catalog: {name}: {problems}')
    return data


def definition_hash(name, row):
    """What an owner's approval of an added pattern binds: its name and its whole definition."""
    return value_hash({'layout_pattern': name, 'definition': row})


def resolve(diff, approvals, media, *, catalog=None, ship=False):
    """The resolved vocabulary of a repository, its findings and the approval targets of its added patterns.

    -> {'catalog_version', 'patterns': {name: {...definition, origin, approval_state?}}, 'errors', 'warnings',
        'targets': {ledger key: hash}}"""
    catalog = catalog or load_catalog()
    out = {'catalog_version': catalog['version'], 'patterns': {}, 'errors': [], 'warnings': [], 'targets': {}}
    errors, warnings = out['errors'], out['warnings']
    diff = {} if diff is None else diff
    if not isinstance(diff, dict) or set(diff) - DIFF_KEYS:
        errors.append(f'DS_LAYOUT_PATTERNS_INVALID: layout_patterns は disabled と added だけを持つ（{sorted(set(diff) - DIFF_KEYS) if isinstance(diff, dict) else "object でない"}）')
        return out
    disabled = diff.get('disabled', [])
    added = diff.get('added', {})
    if not isinstance(disabled, list) or any(not isinstance(n, str) for n in disabled) or len(set(disabled)) != len(disabled):
        errors.append('DS_LAYOUT_PATTERNS_INVALID: disabled は重複の無い名前の配列')
        disabled = []
    if not isinstance(added, dict):
        errors.append('DS_LAYOUT_PATTERNS_INVALID: added は名前を鍵にした object')
        added = {}
    unknown = sorted(n for n in disabled if n not in catalog['patterns'])
    if unknown:
        errors.append(f'DS_LAYOUT_PATTERN_UNKNOWN_DISABLED: カタログ {catalog["version"]} に無い型を disabled にしている {unknown}')
    patterns = {name: {**row, 'origin': 'catalog'} for name, row in catalog['patterns'].items() if name not in disabled}
    for name, row in added.items():
        if name in catalog['patterns']:
            errors.append(f'DS_LAYOUT_PATTERN_NAME_TAKEN: 足した型 {name} はカタログにある名前（使わないなら disabled、変えるなら別の名前で足す）')
            continue
        problems = validate_pattern(name, row)
        if problems:
            hint = '（state は書かない。承認は台帳が持つ）' if isinstance(row, dict) and 'state' in row else ''
            errors.append(f'DS_LAYOUT_PATTERN_FIELDS: 足した型 {name}: {"; ".join(problems)}{hint}')
            continue
        digest = definition_hash(name, row)
        keys = [f'ds:{medium}:pattern:{name}' for medium in media]
        state = 'approved'
        for key in keys:
            out['targets'][key] = digest
            have = approvals.get(key)
            if have is not None and have != digest:
                errors.append(f'approval hash mismatch: pattern:{name}')
            if have != digest:
                state = 'proposed'
        patterns[name] = {**row, 'origin': 'repo', 'approval_state': state}
        if state != 'approved':
            message = f'DS_LAYOUT_PATTERN_PROPOSED: 足した型 {name} はオーナーの承認が無い（{keys[0]}）'
            (errors if ship else warnings).append(message + ('（出荷では失格）' if ship else ''))
    for name, row in patterns.items():
        if row['narrow'] not in patterns:
            errors.append(f'DS_LAYOUT_PATTERN_NARROW_UNKNOWN: 型 {name} の狭い画面の既定 {row["narrow"]} が語彙に無い')
    out['patterns'] = patterns
    return out
