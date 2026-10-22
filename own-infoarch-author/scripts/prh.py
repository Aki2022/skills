"""Reader for the documented subset of the prh rule format (textlint-rule-prh's terms file).

Supported: top-level `version` and `rules`; per rule `expected`, `pattern` / `patterns` (strings are
literal, `/body/flags` is a regular expression with flags g/i/m/u), `regexpMustEmpty` (`$N`),
`options.wordBoundary`, `specs` ({from, to}). Anything else (imports, omitted patterns, named groups,
other flags) is rejected instead of being approximated, so a terms file never passes by being ignored.
"""
from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass

GEN_BEGIN = '  # >>> own-infoarch-author:generated (entities; do not edit, run gen-terms)'
GEN_END = '  # <<< own-infoarch-author:generated'
TOP_KEYS = {'version', 'rules'}
RULE_KEYS = {'expected', 'pattern', 'patterns', 'regexpMustEmpty', 'options', 'specs'}
OPTION_KEYS = {'wordBoundary'}
JS_FLAGS = {'g': 0, 'u': 0, 'i': re.IGNORECASE, 'm': re.MULTILINE}


class PrhError(ValueError):
    pass


@dataclass
class Rule:
    expected: str
    regexes: list
    must_empty: int | None
    specs: list
    index: int

    def _replacement(self, m):
        def sub(g):
            tok = g.group(1)
            if tok == '&':
                return m.group(0)
            n = int(tok)
            return m.group(n) or '' if n <= (m.re.groups or 0) else g.group(0)
        return re.sub(r'\$(&|\d)', sub, self.expected)

    def hits(self, text):
        for rx in self.regexes:
            for m in rx.finditer(text):
                if not m.group(0):
                    continue
                if self.must_empty is not None and (m.group(self.must_empty) or ''):
                    continue
                repl = self._replacement(m)
                if repl != m.group(0):
                    yield m.start(), m.end(), m.group(0), repl

    def apply(self, text):
        out, last = [], 0
        spans = sorted(self.hits(text))
        for start, end, _, repl in spans:
            if start < last:
                continue
            out += [text[last:start], repl]
            last = end
        return ''.join(out) + text[last:]


def _compile(pattern, word_boundary, where):
    if not isinstance(pattern, str) or not pattern:
        raise PrhError(f'{where}: pattern must be a non-empty string')
    m = re.fullmatch(r'/(.*)/([a-z]*)', pattern, re.S)
    if m:
        body, flags = m.groups()
        bad = set(flags) - set(JS_FLAGS)
        if bad:
            raise PrhError(f'{where}: unsupported regexp flags {sorted(bad)}')
        if re.search(r'\(\?<(?![=!])', body) or r'\k<' in body:
            raise PrhError(f'{where}: named groups are not supported')
        f = 0
        for c in flags:
            f |= JS_FLAGS[c]
    else:
        body, f = re.escape(pattern), 0
    if word_boundary:
        body = rf'\b(?:{body})\b'
    try:
        return re.compile(body, f)
    except re.error as exc:
        raise PrhError(f'{where}: invalid regexp: {exc}') from exc


def load_rules(data, *, where='terms.prh.yml'):
    if not isinstance(data, dict):
        raise PrhError(f'{where}: top level must be a mapping')
    extra = set(data) - TOP_KEYS
    if extra:
        raise PrhError(f'{where}: unsupported keys {sorted(extra)} (imports are not supported)')
    if data.get('version') != 1:
        raise PrhError(f'{where}: version must be 1')
    rules_data = data.get('rules')
    if not isinstance(rules_data, list):
        raise PrhError(f'{where}: rules must be a list')
    rules = []
    for i, r in enumerate(rules_data):
        w = f'{where} rules[{i}]'
        if not isinstance(r, dict):
            raise PrhError(f'{w}: must be a mapping')
        extra = set(r) - RULE_KEYS
        if extra:
            raise PrhError(f'{w}: unsupported keys {sorted(extra)}')
        if not isinstance(r.get('expected'), str) or not r['expected']:
            raise PrhError(f'{w}: expected is required')
        if 'pattern' in r and 'patterns' in r:
            raise PrhError(f'{w}: use pattern or patterns, not both')
        pats = r.get('patterns', r.get('pattern'))
        if pats is None:
            raise PrhError(f'{w}: patterns omitted (prh derives them; this reader does not approximate that)')
        pats = pats if isinstance(pats, list) else [pats]
        opts = r.get('options') or {}
        if set(opts) - OPTION_KEYS:
            raise PrhError(f'{w}: unsupported options {sorted(set(opts) - OPTION_KEYS)}')
        must_empty = r.get('regexpMustEmpty')
        if must_empty is not None:
            mm = re.fullmatch(r'\$(\d)', str(must_empty))
            if not mm:
                raise PrhError(f'{w}: regexpMustEmpty must be $N')
            must_empty = int(mm.group(1))
        regexes = [_compile(p, bool(opts.get('wordBoundary')), w) for p in pats]
        if must_empty is not None and any(rx.groups < must_empty for rx in regexes):
            raise PrhError(f'{w}: regexpMustEmpty refers to a missing group')
        specs = r.get('specs') or []
        for s in specs:
            if not isinstance(s, dict) or set(s) != {'from', 'to'}:
                raise PrhError(f'{w}: specs entries need exactly from and to')
        rules.append(Rule(r['expected'], regexes, must_empty, specs, i))
    return rules


def run_specs(rules):
    """prh refuses to load a file whose specs fail; return the failures instead of raising."""
    failures = []
    for rule in rules:
        for s in rule.specs:
            got = rule.apply(s['from'])
            if got != s['to']:
                failures.append({'rule': rule.index, 'expected': rule.expected, 'from': s['from'], 'want': s['to'], 'got': got})
    return failures


def check_text(rules, text):
    for rule in rules:
        for start, end, found, repl in rule.hits(text):
            yield {'rule': rule.index, 'found': found, 'expected': repl, 'start': start, 'end': end}


# ---- generation from entities ----------------------------------------------------------

_DASH = 'ー-−‐–—―ｰ－'


def _halfwidth_kana():
    table = {}
    for cp in range(0xFF66, 0xFF9E):
        h = chr(cp)
        f = unicodedata.normalize('NFKC', h)
        if len(f) == 1:
            table.setdefault(f, h)
        for mark in ('ﾞ', 'ﾟ'):
            f2 = unicodedata.normalize('NFKC', h + mark)
            if len(f2) == 1 and f2 != f:
                table.setdefault(f2, h + mark)
    return table


_HW = _halfwidth_kana()


def _char_class(c):
    if c in _DASH:
        return '[' + re.escape(_DASH) + ']'
    if c.isascii() and c.isalpha():
        u, l = c.upper(), c.lower()
        return f'[{u}{l}{chr(ord(u) + 0xFEE0)}{chr(ord(l) + 0xFEE0)}]'
    if c.isascii() and c.isdigit():
        return f'[{c}{chr(ord(c) + 0xFEE0)}]'
    if c in _HW:
        return f'(?:{re.escape(c)}|{re.escape(_HW[c])})'
    if c.isspace():
        return '[ 　]*'
    return re.escape(c)


_ALNUM_EDGE = '[A-Za-z0-9Ａ-Ｚａ-ｚ０-９]'


def _edges(text):
    """Lookarounds so an ASCII-edged name does not match inside a longer word (Web in Website)."""
    head = f'(?<!{_ALNUM_EDGE})' if text[:1].isascii() and text[:1].isalnum() else ''
    tail = f'(?!{_ALNUM_EDGE})' if text[-1:].isascii() and text[-1:].isalnum() else ''
    return head, tail


def folding_regex(name):
    """Regex for case / width / half-width-kana / dash / inserted-space variants of a name."""
    parts = [_char_class(c) for c in name if not c.isspace()]
    head, tail = _edges(name)
    return '/' + head + '[ 　]?'.join(parts) + tail + '/'


def forbidden_pattern(name, form):
    """A forbidden form that is part of the canonical name must not hit inside the name itself."""
    if form in name:
        i = name.index(form)
        before, after = name[:i], name[i + len(form):]
        body = (f'(?<!{re.escape(before)})' if before else '') + re.escape(form) + \
               (f'(?!{re.escape(after)})' if after else '')
        return f'/{body}/'
    return form


def entity_rules(entities):
    rules = []
    for eid in sorted(entities):
        e = entities[eid]
        if e.get('state') != 'proposed' or not e.get('name'):
            continue
        pats = []
        if e.get('variant_folding', True):
            pats.append(folding_regex(e['name']))
        pats += [forbidden_pattern(e['name'], p) for p in sorted(e.get('forbidden', [])) if p != e['name']]
        if pats:
            rules.append((eid, {'expected': e['name'], 'patterns': pats}))
    return rules


def render_generated(entities):
    lines = [GEN_BEGIN]
    for eid, r in entity_rules(entities):
        lines.append(f'  # {eid}')
        lines.append(f'  - expected: {json.dumps(r["expected"], ensure_ascii=False)}')
        lines.append('    patterns:')
        lines += [f'      - {json.dumps(p, ensure_ascii=False)}' for p in r['patterns']]
    lines.append(GEN_END)
    return '\n'.join(lines)


def generated_block(file_text):
    lines = file_text.split('\n')
    if lines.count(GEN_BEGIN) != 1 or lines.count(GEN_END) != 1:
        return None
    b, e = lines.index(GEN_BEGIN), lines.index(GEN_END)
    if e < b:
        return None
    return '\n'.join(lines[b:e + 1])


def replace_generated(file_text, block):
    current = generated_block(file_text)
    if current is None:
        raise PrhError('generated markers missing or duplicated')
    return file_text.replace(current, block, 1)
