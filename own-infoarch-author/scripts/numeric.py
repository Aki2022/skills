"""Japanese numeric normalizer: text -> claim-candidate tokens with normalized value/unit/qualifiers.

All tables live in references/numeric.json and references/non-claim-patterns.json; this module holds none.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal
from functools import lru_cache

from common import load_ref


@dataclass(frozen=True)
class Token:
    start: int
    end: int
    text: str
    value: tuple  # (low, high) as Decimal; equal for a single value
    unit: str
    qualifiers: frozenset

    def key(self):
        return (self.value, self.unit, self.qualifiers)


@lru_cache(maxsize=None)
def _tables():
    num = load_ref('numeric.json')
    nonclaim = load_ref('non-claim-patterns.json')
    small = ''.join(num['kanji_small_units'])
    digits = ''.join(num['kanji_digits']) + small + ''.join(num['vague_digits'])
    large = ''.join(num['kanji_large_units'])
    arabic = r'(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?'
    part = rf'(?:{arabic}[{small}]?|[{digits}]+)'
    compound = rf'(?:(?:{part}\s*[{large}]\s*)+{part}?|{part})'
    unit_aliases = sorted(((a, u['unit'], u.get('factor', 1)) for u in num['units'] for a in u['aliases']),
                          key=lambda x: -len(x[0]))
    unit_re = '|'.join(re.escape(a) for a, _, _ in unit_aliases)
    prefixes, suffixes = [], []
    for qid, spec in num['qualifiers'].items():
        prefixes += [(p, qid) for p in spec.get('prefix', [])]
        suffixes += [(s, qid) for s in spec.get('suffix', [])]
    prefixes.sort(key=lambda x: -len(x[0]))
    suffixes.sort(key=lambda x: -len(x[0]))
    token_re = re.compile(
        rf'(?P<n1>{compound})\s*(?P<u1>{unit_re})?(?:\s*~\s*(?P<n2>{compound})\s*(?P<u2>{unit_re})?)?')
    return {
        'num': num, 'token_re': token_re,
        'units': {a: (u, Decimal(f)) for a, u, f in unit_aliases},
        'prefixes': prefixes, 'suffixes': suffixes,
        'kanji_only': re.compile(rf'^[{digits}{large}\s]+$'),
        'kanji_scaled': re.compile(rf'[{small}{large}]'),
        'arabic_part': re.compile(arabic), 'arabic_small': re.compile(rf'({arabic})([{small}])'), 'large': large,
        'single_units': set(num['single_kanji_units']), 'vague': set(num['vague_digits']),
        'not_after': num.get('prefix_not_after', {}),
        'nonclaim': [(p['id'], re.compile(p['regex'])) for p in nonclaim['patterns']],
    }


def normalize(text):
    """NFKC, then fold the dash and wave variants NFKC leaves alone. Length may change (NFKC)."""
    t = _tables()['num']
    s = unicodedata.normalize('NFKC', text)
    s = s.translate({ord(c): '-' for c in t['dash_chars']})
    s = s.translate({ord(c): '~' for c in t['wave_chars']})
    s = re.sub(r'(?<=\d)ー(?=\d)', '-', s)  # long-vowel mark typed as a hyphen between digits
    return s


def mask_non_claims(norm):
    """Replace non-claim spans with spaces (same length) and return (masked, [(pattern id, span text)])."""
    masked, hits = norm, []
    for pid, rx in _tables()['nonclaim']:
        def repl(m, pid=pid):
            hits.append((pid, m.group(0)))
            return ' ' * len(m.group(0))
        masked = rx.sub(repl, masked)
    return masked, hits


def _kanji_number(s, t):
    digits, small = t['num']['kanji_digits'], t['num']['kanji_small_units']
    if any(ch in t['vague'] for ch in s):
        return Decimal('NaN')  # 数万・何百: a number nobody can match against a registered value
    total, current = 0, 0
    for ch in s:
        if ch in digits:
            current = current * 10 + digits[ch] if current else digits[ch]
        elif ch in small:
            total += (current or 1) * small[ch]
            current = 0
    return Decimal(total + current)


def parse_number(s):
    t = _tables()
    large = t['num']['kanji_large_units']
    s = re.sub(r'\s+', '', s)
    total, buf = Decimal(0), ''
    for ch in s:
        if ch in large:
            total += _part_value(buf, t) * Decimal(large[ch])
            buf = ''
        else:
            buf += ch
    if buf:
        total += _part_value(buf, t)
    return total


def _part_value(part, t):
    if not part:
        return Decimal(1)
    if t['arabic_part'].fullmatch(part):
        return Decimal(part.replace(',', ''))
    m = t['arabic_small'].fullmatch(part)
    if m:
        return Decimal(m.group(1).replace(',', '')) * Decimal(t['num']['kanji_small_units'][m.group(2)])
    return _kanji_number(part, t)


def _qualifiers(masked, start, end, t):
    quals = set()
    before = masked[:start].rstrip()
    progress = True
    while progress:  # stacked prefixes: 最大約100社
        progress = False
        for p, q in t['prefixes']:
            if before.endswith(p) and before[:-len(p)][-1:] not in t['not_after'].get(p, []):
                quals.add(q)
                before = before[:-len(p)].rstrip()
                progress = True
                break
    after = masked[end:]
    consumed = 0
    progress = True
    while progress:
        progress = False
        rest = after[consumed:].lstrip()
        skipped = len(after[consumed:]) - len(rest)
        for s, q in t['suffixes']:
            if rest.startswith(s):
                quals.add(q)
                consumed += skipped + len(s)
                progress = True
                break
    return frozenset(quals), consumed


def candidate_tokens(text, *, normalized=False):
    """Numeric tokens that are claim candidates (not matched by the non-claim table)."""
    t = _tables()
    norm = text if normalized else normalize(text)
    masked, _ = mask_non_claims(norm)
    out = []
    for m in t['token_re'].finditer(masked):
        n1, u1, n2, u2 = m.group('n1'), m.group('u1'), m.group('n2'), m.group('u2')
        unit_alias = u2 or u1 or ''
        if t['kanji_only'].match(n1) and not unit_alias:
            continue  # kanji words (一緒・千葉) are numbers only when a unit follows
        if (t['kanji_only'].match(n1) and not t['kanji_scaled'].search(n1) and len(n1.strip()) == 1
                and unit_alias not in t['single_units'] and n1.strip() not in t['vague']):
            continue  # a single kanji digit with a unit (一人) is too ambiguous; reported as a limit
        if not unit_alias and (_ascii_letter(masked, m.start() - 1) or _ascii_letter(masked, m.end())):
            continue  # part of an identifier (B2B, 5G, Windows11), not a quantity
        unit, factor = t['units'].get(unit_alias, ('', Decimal(1)))
        lo = parse_number(n1) * factor
        hi = parse_number(n2) * factor if n2 else lo
        quals, consumed = _qualifiers(masked, m.start(), m.end(), t)
        out.append(Token(m.start(), m.end() + consumed, norm[m.start():m.end() + consumed].strip(),
                         (lo, hi),
                         unit, quals))
    return [_canon(tok) for tok in out]


def _ascii_letter(s, i):
    return 0 <= i < len(s) and s[i].isascii() and s[i].isalpha()


def _canon(tok):
    lo, hi = (Decimal(int(v)) if v.is_finite() and v == v.to_integral() else v for v in tok.value)
    return Token(tok.start, tok.end, tok.text, (lo, hi), tok.unit, tok.qualifiers)


def claim_key(value):
    """Normalized key of a claim's registered value (info_arch.yaml claims.*.value)."""
    v = value['value']
    lo, hi = (v, v) if not isinstance(v, list) else v
    lo, hi = (Decimal(str(x)) for x in (lo, hi))
    lo, hi = (Decimal(int(x)) if x == x.to_integral() else x for x in (lo, hi))
    return ((lo, hi), value['unit'], frozenset(value['qualifier']))
