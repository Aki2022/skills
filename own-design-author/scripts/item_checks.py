"""Route every adopted catalog item to one way of verifying it: code, an approved design decision, or AFM.

The map lives in the skill (references/item-checks.json), not in each repository: every repository uses the
same catalog, so one reviewed map serves all of them. Every catalog item has exactly one route there.

- machine: code checks the item on each rendered page. A checker is either gated by L2 (its failure already
  stops the render stage) or collected by L2 per page and enforced per adopted item after rendering.
- decision: the item is a choice the repository makes once (which variant, when to use a component, what the
  text says). It is verified at adoption time: DESIGN.md states the choice and the owner approved that item.
- judgment: only items that need a visual judgment code cannot make go to the on-device model (AFM).
An item the map does not route, or a machine item whose tokens the design system lacks, is uncovered.
"""
import json
from pathlib import Path

MAP = Path(__file__).resolve().parents[1] / 'references/item-checks.json'
ROUTES = ('machine', 'decision', 'judgment')
# Failures of these already stop L2; reaching item verification means they passed on the page.
L2_GATED = {'palette:color', 'palette:fontFamily', 'palette:radius', 'contrast', 'tap-target',
            'heading-order', 'layout-overflow', 'body-line-height', 'line-length'}
# L2 measures these per page without failing; item verification enforces them for adopted items only.
COLLECTED = {'palette:fontSize', 'palette:fontWeight', 'palette:shadow', 'palette:spacing',
             'contrast-nontext', 'contrast-icon',
             'dom:placeholder', 'dom:readonly-input', 'dom:disabled-input', 'dom:maxlength', 'dom:ol',
             'dom:positive-tabindex', 'dom:date-input', 'dom:table-span', 'dom:vague-link-text', 'dom:text-justify',
             'dom:link-cursor', 'dom:disabled-button', 'dom:text-overflow', 'dom:control-height',
             'dom:disclosure-semantics', 'dom:table-headers', 'dom:table-align', 'dom:label-above',
             'dom:input-pattern', 'dom:paste-blocked', 'dom:form-live-region', 'dom:fake-list', 'dom:h1-title',
             'dom:icon-only', 'dom:icon-hidden', 'dom:new-tab-indicator', 'dom:link-distinct',
             'dom:space-alignment', 'dom:long-italic', 'dom:explicit-bg', 'dom:dialog-border', 'dom:file-button',
             'dom:carousel-slides', 'dom:focus-order',
             'axe:button-name', 'axe:list', 'axe:label', 'axe:link-name'}
# A palette checker is active only when the design system defines tokens to compare against.
TOKEN_NEEDS = {'palette:color': 'color', 'palette:fontFamily': 'fontFamily', 'palette:radius': 'dimension',
               'palette:fontSize': 'dimension', 'palette:spacing': 'dimension',
               'palette:fontWeight': 'fontWeight', 'palette:shadow': 'shadow'}
EVIDENCE_TOKENS = {'color': 'palette:color', 'fontFamily': 'palette:fontFamily'}


def document_key(item):
    return str(item.get('document', '')).removesuffix('/index.md')


def heading_path(item):
    """The exact heading path of a catalog item; a sub-heading never inherits its parent's mapping."""
    return ' > '.join(item.get('headings') or [item.get('title', '')])


def load_map(path=MAP):
    data = json.loads(Path(path).read_text())
    if not isinstance(data, dict) or data.get('schema_version') != 2 or set(data) - {'schema_version', 'comment', 'applicability', 'routes'}:
        raise ValueError('invalid item check map')
    applicability = data.get('applicability')
    if not isinstance(applicability, dict) or any(
            not isinstance(v, list) or not v or any(not isinstance(s, str) or not s.strip() for s in v)
            for v in applicability.values()):
        raise ValueError('invalid applicability selectors')
    routes = data.get('routes')
    if not isinstance(routes, list):
        raise ValueError('invalid item routes')
    index = {}
    for row in routes:
        if not isinstance(row, dict) or row.get('route') not in ROUTES or not isinstance(row.get('paths'), list) or not row['paths']:
            raise ValueError('malformed item route row')
        expected = {'document', 'route', 'paths'} | ({'checkers'} if row['route'] == 'machine' else set())
        if set(row) != expected or (row['route'] == 'machine' and (not row['checkers'] or set(row['checkers']) - L2_GATED - COLLECTED)):
            raise ValueError('unknown checker or malformed item route row')
        for path in row['paths']:
            key = (row['document'], path)
            if key in index:
                raise ValueError('item routed twice: ' + ' | '.join(key))
            index[key] = row
    data['_index'] = index
    return data


def active_checkers(tokens):
    types = {token['type'] for token in tokens.values()}
    return {c for c in L2_GATED | COLLECTED if c not in TOKEN_NEEDS or TOKEN_NEEDS[c] in types}


def route(item, evidence, tokens, rules, mapping, active):
    """Return (route, checkers, reason): route is machine, decision, afm or uncovered."""
    for ref in evidence:
        kind, _, target = ref.partition(':')
        if kind == 'rule' and target in rules:
            return 'machine', ['rule:' + target], ''
        if kind == 'token' and target in tokens and EVIDENCE_TOKENS.get(tokens[target]['type']) in active:
            return 'machine', [EVIDENCE_TOKENS[tokens[target]['type']]], ''
    row = mapping['_index'].get((document_key(item), heading_path(item)))
    if row is None:
        return 'uncovered', [], 'the skill map does not route this catalog item'
    if row['route'] == 'decision':
        return 'decision', [], ''
    if row['route'] == 'judgment':
        return 'afm', [], ''
    missing = sorted(TOKEN_NEEDS[c] for c in row['checkers'] if c not in active)
    if missing:
        return 'uncovered', sorted(row['checkers']), 'design system has no ' + '/'.join(missing) + ' tokens to compare against'
    return 'machine', sorted(row['checkers']), ''
