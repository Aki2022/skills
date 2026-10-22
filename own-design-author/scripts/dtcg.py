"""One offline DTCG 2025.10 resolver; validators consume its manifest, never tokens."""
import copy
import hashlib
import json
import math
from pathlib import Path
import re

EXTENSION = 'org.own.design'
SCHEMA_PATH = Path(__file__).resolve().parents[1] / 'references/schema/dtcg-2025.10.json'


def value_hash(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                    separators=(',', ':'), allow_nan=False).encode()).hexdigest()


def _validate(document):
    try:
        from jsonschema import Draft7Validator
        from referencing import Registry, Resource
    except ImportError as exc:
        raise ValueError('DTCG schema validation requires an existing Python environment with jsonschema') from exc
    schema = json.loads(SCHEMA_PATH.read_text())
    resources = [(schema['$id'], Resource.from_contents(schema))]
    resources += [(v['$id'], Resource.from_contents(v)) for v in schema['definitions'].values()
                  if isinstance(v, dict) and '$id' in v]
    # The official download embeds all referenced schemas. No network retrieval at runtime.
    registry = Registry().with_resources(resources)
    errors = sorted(Draft7Validator(schema, registry=registry).iter_errors(document), key=lambda e: str(e.path))
    if errors:
        error = errors[0]
        raise ValueError(f'DTCG schema {".".join(map(str,error.path))}: {error.message[:180]}')


def _num(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError('non-finite or non-numeric token value')
    return f'{value:g}'


def css_value(kind, value):
    if kind == 'color':
        if not isinstance(value, dict) or not re.fullmatch(r'#[0-9a-fA-F]{6}', value.get('hex', '')):
            raise ValueError('color requires a six-digit hex fallback')
        alpha = value.get('alpha', 1)
        if value.get('colorSpace') != 'srgb':
            raise ValueError('web CSS currently supports srgb colors; convert other spaces explicitly')
        components = value.get('components', [])
        if len(components) != 3 or any(isinstance(c, bool) or not isinstance(c, (int, float)) for c in components):
            raise ValueError('srgb components must be three numbers')
        rgb_bytes = [int(value['hex'][i:i+2], 16) for i in (1, 3, 5)]
        if any(abs(component * 255 - byte) > 0.501 for component, byte in zip(components, rgb_bytes)):
            raise ValueError('srgb components disagree with hex fallback')
        if alpha == 1:
            return value['hex'].lower()
        raw = value['hex'][1:]
        rgb = [str(int(raw[i:i+2], 16)) for i in (0, 2, 4)]
        return f'rgb({" ".join(rgb)} / {_num(alpha)})'
    if kind in ('dimension', 'duration'):
        return _num(value['value']) + value['unit']
    if kind in ('number', 'fontWeight'):
        return _num(value) if isinstance(value, (int, float)) else str(value)
    if kind == 'fontFamily':
        generic = {'serif','sans-serif','monospace','cursive','fantasy','system-ui','ui-serif',
                   'ui-sans-serif','ui-monospace','ui-rounded','emoji','math','fangsong'}
        return ', '.join(s if s in generic else json.dumps(s, ensure_ascii=False)
                         for s in (value if isinstance(value, list) else [value]))
    if kind == 'cubicBezier':
        return 'cubic-bezier(' + ', '.join(_num(v) for v in value) + ')'
    if kind == 'strokeStyle':
        if isinstance(value, str):
            return value
        raise ValueError('custom strokeStyle has no single CSS value; use a supported stroke style')
    if kind == 'border':
        return ' '.join([css_value('dimension', value['width']), css_value('strokeStyle', value['style']),
                         css_value('color', value['color'])])
    if kind == 'shadow':
        shadows = value if isinstance(value, list) else [value]
        return ', '.join(('inset ' if s.get('inset') else '') + ' '.join(
            [css_value('dimension', s[k]) for k in ('offsetX', 'offsetY', 'blur', 'spread')]
            + [css_value('color', s['color'])]) for s in shadows)
    if kind == 'gradient':
        return 'linear-gradient(' + ', '.join(css_value('color', v['color']) + ' ' +
                     _num(v['position'] * 100) + '%' for v in value) + ')'
    if kind == 'transition':
        return ' '.join([css_value('duration', value['duration']),
                         css_value('cubicBezier', value['timingFunction']), css_value('duration', value['delay'])])
    if kind == 'typography':
        return ' '.join([css_value('fontWeight', value['fontWeight']),
            css_value('dimension', value['fontSize']) + '/' + _num(value['lineHeight']),
            css_value('fontFamily', value['fontFamily'])])
    raise ValueError(f'unsupported DTCG type: {kind}')


def resolve_tokens(documents):
    raw, locations = {}, {}
    for document in documents:
        _validate(document)
        def walk(node, prefix='', inherited=None):
            kind = node.get('$type', inherited)
            if '$extends' in node:
                raise ValueError('group $extends is unsupported; express aliases explicitly')
            if '$value' in node or '$ref' in node:
                if not prefix:
                    raise ValueError('token must have a name')
                if prefix in raw:
                    raise ValueError(f'duplicate token path: {prefix}')
                raw[prefix] = {**node, '$type': kind}
                locations[prefix] = document
                return
            for key, child in node.items():
                if key == '$root':
                    walk(child, prefix, kind)
                elif not key.startswith('$'):
                    if not re.fullmatch(r'[a-z][a-z0-9-]*', key):
                        raise ValueError('token names must use lower-case kebab-case')
                    walk(child, f'{prefix}.{key}' if prefix else key, kind)
        walk(document)
    if not raw:
        raise ValueError('zero token targets')
    resolved, active, css_names = {}, set(), set()
    def resolve(path):
        if path in resolved:
            return resolved[path]
        if path in active or path not in raw:
            raise ValueError(f'cyclic or missing token reference: {path}')
        active.add(path)
        node = raw[path]
        kind = node.get('$type')
        def expand(value):
            if isinstance(value, str) and re.fullmatch(r'\{[^{}]+\}', value):
                return copy.deepcopy(resolve(value[1:-1])['value'])
            if isinstance(value, dict) and set(value) == {'$ref'}:
                reference = value['$ref']
                if not isinstance(reference, str) or not reference.startswith('#/'):
                    raise ValueError('only local DTCG JSON pointers are allowed')
                cursor = locations[path]
                for component in reference[2:].split('/'):
                    key = component.replace('~1', '/').replace('~0', '~')
                    cursor = cursor[int(key)] if isinstance(cursor, list) else cursor[key]
                return expand(cursor)
            if isinstance(value, dict):
                return {k: expand(v) for k, v in value.items()}
            if isinstance(value, list):
                return [expand(v) for v in value]
            return value
        value = node.get('$value', {'$ref': node.get('$ref')})
        inherited_origin = None
        if isinstance(value, str) and re.fullmatch(r'\{[^{}]+\}', value):
            target = resolve(value[1:-1])
            if kind is not None and kind != target['type']:
                raise ValueError(f'alias type mismatch: {path}')
            kind = kind or target['type']
            inherited_origin = target.get('origin')
        if not kind:
            raise ValueError(f'token type is missing: {path}')
        expanded = expand(value)
        _validate({'token': {'$type': kind, '$value': expanded}})
        css = css_value(kind, expanded)
        css_name = '--' + path.replace('.', '-')
        if css_name in css_names:
            raise ValueError(f'CSS custom property collision: {path}')
        css_names.add(css_name)
        extension = node.get('$extensions', {}).get(EXTENSION, {})
        if not isinstance(extension, dict):
            raise ValueError('design token extension must be an object')
        resolved[path] = {'type': kind, 'value': expanded, 'css': css, 'css_name': css_name,
                          'value_hash': value_hash({'type': kind, 'value': expanded}),
                          'origin': extension.get('origin', inherited_origin),
                          'extensions': node.get('$extensions', {})}
        active.remove(path)
        return resolved[path]
    try:
        for path in raw:
            resolve(path)
    except (KeyError, TypeError, IndexError, RecursionError) as exc:
        raise ValueError(f'invalid DTCG value or reference: {type(exc).__name__}') from exc
    return dict(sorted(resolved.items()))
