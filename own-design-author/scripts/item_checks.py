"""Route each adopted catalog item to code checks first and AFM only for what code cannot verify.

The map lives in the skill (references/item-checks.json), not in each repository: every repository uses the
same catalog, so one reviewed map serves all of them. A checker is either gated by L2 (its failure already
stops the render stage) or collected by L2 per page and enforced per adopted item after rendering.
"""
import json
from pathlib import Path

MAP = Path(__file__).resolve().parents[1] / 'references/item-checks.json'
# Failures of these already stop L2; reaching item verification means they passed on the page.
L2_GATED = {'palette:color', 'palette:fontFamily', 'palette:radius', 'contrast', 'tap-target',
            'heading-order', 'layout-overflow'}
# L2 measures these per page without failing; item verification enforces them for adopted items only.
COLLECTED = {'palette:fontSize', 'palette:fontWeight', 'palette:shadow', 'palette:spacing',
             'dom:placeholder', 'dom:readonly-input', 'dom:disabled-input', 'dom:maxlength', 'dom:ol',
             'dom:positive-tabindex', 'dom:date-input', 'dom:table-span', 'dom:vague-link-text',
             'dom:text-justify', 'axe:button-name', 'axe:list'}
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
    if not isinstance(data, dict) or data.get('schema_version') != 1 or set(data) - {'schema_version', 'comment', 'applicability', 'checks'}:
        raise ValueError('invalid item check map')
    applicability = data.get('applicability')
    if not isinstance(applicability, dict) or any(
            not isinstance(v, list) or not v or any(not isinstance(s, str) or not s.strip() for s in v)
            for v in applicability.values()):
        raise ValueError('invalid applicability selectors')
    checks = data.get('checks')
    if not isinstance(checks, list):
        raise ValueError('invalid item checks')
    for row in checks:
        if (not isinstance(row, dict) or set(row) != {'document', 'paths', 'checkers'}
                or not isinstance(row['paths'], list) or not row['paths'] or not row['checkers']
                or set(row['checkers']) - L2_GATED - COLLECTED):
            raise ValueError('unknown checker or malformed item check row')
    return data


def active_checkers(tokens):
    types = {token['type'] for token in tokens.values()}
    return {c for c in L2_GATED | COLLECTED if c not in TOKEN_NEEDS or TOKEN_NEEDS[c] in types}


def route(item, evidence, tokens, rules, mapping, active):
    """Return (route, checkers): 'machine' with the checkers that verify the item, or 'afm'."""
    for ref in evidence:
        kind, _, target = ref.partition(':')
        if kind == 'rule' and target in rules:
            return 'machine', ['rule:' + target]
        if kind == 'token' and target in tokens and EVIDENCE_TOKENS.get(tokens[target]['type']) in active:
            return 'machine', [EVIDENCE_TOKENS[tokens[target]['type']]]
    document, path = document_key(item), heading_path(item)
    for row in mapping['checks']:
        if row['document'] == document and path in row['paths']:
            if set(row['checkers']) <= active:
                return 'machine', sorted(row['checkers'])
    return 'afm', []
