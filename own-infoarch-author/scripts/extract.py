"""Text extraction adapters: artifact file -> blocks.

A block is the text of one block-level unit (paragraph, heading, list item, cell, attribute value) with the
claim anchor of every character, so checks see the words around an anchor (約<span>100社</span>) and
hard-wrapped words are joined (業界\\n初 -> 業界初). Two texts are kept: `norm` (NFKC + dash/wave folding,
for numbers and lexicon) and `raw` (for prh, which must see full-width and half-width forms as written).

HTML: text per block (HTML5 optional end tags close implicitly), user-facing attributes, meta content.
Markdown / Quarto: paragraphs, headings, list items, table rows, Pandoc spans and fenced divs with
`data-ia-claim`, raw-HTML attributes, front matter title / subtitle / description.
Constructs whose text is computed or not extracted are reported as uncovered by the caller.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from html.parser import HTMLParser

import numeric

CLAIM_ATTR = 'data-ia-claim'
HTML_SUFFIXES = {'.html', '.htm'}
MD_SUFFIXES = {'.md', '.qmd', '.markdown'}
VOID = {'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr'}
SKIP = {'script', 'style', 'template'}
BLOCK = {'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th', 'dt', 'dd', 'div', 'section', 'article',
         'header', 'footer', 'main', 'nav', 'aside', 'blockquote', 'figcaption', 'button', 'title', 'tr', 'ul',
         'ol', 'dl', 'table', 'thead', 'tbody', 'tfoot', 'form', 'label', 'option', 'summary', 'details', 'figure',
         'body', 'head', 'html', 'pre', 'hr', 'caption', 'legend', 'fieldset', 'select'}
CLOSES_P = BLOCK - {'button', 'title', 'label', 'option', 'summary', 'body', 'head', 'html', 'td', 'th', 'tr', 'li', 'dt', 'dd'}
TEXT_ATTRS = {'alt', 'title', 'aria-label', 'aria-description', 'placeholder', 'label'}
META_NAMES = {'description', 'og:site_name', 'og:title', 'og:description', 'og:image:alt', 'twitter:title', 'twitter:description',
              'twitter:image:alt'}
# tag -> (tags it implicitly closes, tags that stop the search)
IMPLICIT = {
    'li': ({'li'}, {'ul', 'ol', 'menu'}),
    'dt': ({'dt', 'dd'}, {'dl'}), 'dd': ({'dt', 'dd'}, {'dl'}),
    'tr': ({'tr', 'td', 'th'}, {'table', 'thead', 'tbody', 'tfoot'}),
    'td': ({'td', 'th'}, {'tr', 'table'}), 'th': ({'td', 'th'}, {'tr', 'table'}),
    'option': ({'option'}, {'select', 'datalist', 'optgroup'}),
    'thead': ({'tbody', 'thead', 'tfoot', 'tr', 'td', 'th'}, {'table'}),
    'tbody': ({'tbody', 'thead', 'tfoot', 'tr', 'td', 'th'}, {'table'}),
    'tfoot': ({'tbody', 'thead', 'tfoot', 'tr', 'td', 'th'}, {'table'}),
}


def _is_cjk(ch):
    o = ord(ch)
    return (0x3000 <= o <= 0x30FF or 0x3400 <= o <= 0x4DBF or 0x4E00 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF
            or 0xFF01 <= o <= 0xFF9F)


@dataclass
class Block:
    norm: str
    claims: list  # claim id (or None) per character of norm
    raw: str
    line: int
    origin: str  # 'text' or 'attr:<name>'

    def claim_ids(self):
        return {c for c in self.claims if c}

    def claim_over(self, start, end):
        """The claim if every character in [start, end) carries the same one, else None."""
        ids = set(self.claims[start:end])
        return ids.pop() if len(ids) == 1 else None


def _join(pieces, fold):
    """pieces: [(text, claim)] -> (text, claims). Collapse whitespace; drop it between two CJK characters."""
    chars, claims = [], []
    for text, claim in pieces:
        t = numeric.normalize(text) if fold else text
        for ch in t:
            chars.append(ch)
            claims.append(claim)
    out, out_claims, i = [], [], 0
    while i < len(chars):
        if chars[i].isspace():
            j = i
            while j < len(chars) and chars[j].isspace():
                j += 1
            prev = out[-1] if out else ''
            nxt = chars[j] if j < len(chars) else ''
            if prev and nxt and not (_is_cjk(prev) and _is_cjk(nxt)):
                out.append(' ')
                out_claims.append(claims[i] if claims[i] == claims[j] else None)
            i = j
            continue
        out.append(chars[i])
        out_claims.append(claims[i])
        i += 1
    return ''.join(out), out_claims


@dataclass
class _Buf:
    pieces: list = field(default_factory=list)
    line: int = 0

    def add(self, text, claim, line):
        if not self.pieces:
            self.line = line
        self.pieces.append((text, claim))

    def block(self, origin='text'):
        norm, claims = _join(self.pieces, True)
        raw, _ = _join(self.pieces, False)
        return Block(norm, claims, raw, self.line, origin) if norm else None


def _attr_block(value, claim, line, origin):
    buf = _Buf()
    buf.add(value, claim, line)
    return buf.block(origin)


class _HTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack = []  # (tag, own claim)
        self.buf = _Buf()
        self.blocks, self.uncovered, self.anchor_ids = [], [], []
        self.skip_depth = 0

    def _claim(self):
        return next((c for _, c in reversed(self.stack) if c), None)

    def flush(self):
        b = self.buf.block()
        if b:
            self.blocks.append(b)
        self.buf = _Buf()

    def _pop_to(self, idx):
        self.flush()
        del self.stack[idx:]

    def _implicit_close(self, tag):
        closes, stops = IMPLICIT.get(tag, (set(), set()))
        if tag in CLOSES_P:
            closes, stops = closes | {'p'}, stops | {'div', 'li', 'td', 'th', 'dd', 'dt', 'section', 'article',
                                                      'blockquote', 'button', 'table', 'body'}
        for i in range(len(self.stack) - 1, -1, -1):
            t = self.stack[i][0]
            if t in closes:
                self._pop_to(i)
                return
            if t in stops:
                return

    def handle_starttag(self, tag, attrs):
        a = {k: (v or '') for k, v in attrs}
        if tag == 'script' and a.get('type', '').lower() in {'application/ld+json', 'application/json'}:
            self.uncovered.append(f'<script type="{a["type"]}">（構造化データの文言は抽出しない）')
        if tag in SKIP:
            self.skip_depth += 1
            return
        self._implicit_close(tag)
        own_claim = a.get(CLAIM_ATTR, '').strip() or None
        if own_claim:  # validated even when the element carries no text
            self.anchor_ids.append((own_claim, self.getpos()[0]))
        if tag in BLOCK:
            self.flush()
        claim = own_claim or self._claim()
        line = self.getpos()[0]
        for name in sorted(TEXT_ATTRS & set(a)):
            b = _attr_block(a[name], claim, line, f'attr:{name}')
            if b:
                self.blocks.append(b)
        if tag == 'input' and a.get('type', '').lower() in {'submit', 'button'} and a.get('value'):
            self.blocks.append(_attr_block(a['value'], claim, line, 'attr:value'))
        if tag == 'meta' and (a.get('name') in META_NAMES or a.get('property') in META_NAMES) and a.get('content'):
            self.blocks.append(_attr_block(a['content'], claim, line, 'attr:content'))
        if tag == 'br':
            self.buf.add('\n', claim, line)
        if tag not in VOID:
            self.stack.append((tag, own_claim))

    def handle_endtag(self, tag):
        if tag in SKIP:
            self.skip_depth = max(0, self.skip_depth - 1)
            return
        idx = next((i for i in range(len(self.stack) - 1, -1, -1) if self.stack[i][0] == tag), None)
        if idx is None:
            return
        if tag in BLOCK or any(self.stack[j][0] in BLOCK for j in range(idx, len(self.stack))):
            self._pop_to(idx)
        else:
            del self.stack[idx:]

    def handle_data(self, data):
        if not self.skip_depth:
            self.buf.add(data, self._claim(), self.getpos()[0])


def html_blocks(text):
    p = _HTML()
    p.feed(text)
    p.close()
    p.flush()
    return [b for b in p.blocks if b], p.uncovered, p.anchor_ids


_FENCE = re.compile(r'^\s*(```|~~~)')
_DIV_OPEN = re.compile(r'^\s*:{3,}\s*\{([^}]*)\}\s*$')
_DIV_CLOSE = re.compile(r'^\s*:{3,}\s*$')
_SPAN = re.compile(r'\[([^\[\]]*)\]\{([^}]*)\}')
_ATTR_CLAIM = re.compile(r'data-ia-claim\s*=\s*"?([A-Za-z0-9-]+)"?')
_LINK = re.compile(r'!?\[([^\[\]]*)\]\([^)]*\)')
_CODE = re.compile(r'`[^`]*`')
_COMPUTED = re.compile(r'\{\{<.*?>\}\}|`\{[a-z]+\}[^`]*`')
_TAG = re.compile(r'<[^>]+>')
_TAG_ATTR = re.compile(r'''\b(alt|title|aria-label|aria-description|placeholder)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'<>=`]+))''')
_HEADING = re.compile(r'^\s{0,3}#{1,6}\s+')
_ITEM = re.compile(r'^\s*(?:[-*+]|\d+[.)])\s+')
_QUOTE = re.compile(r'^\s*(?:>\s?)+')
_FRONT_KEYS = ('title', 'subtitle', 'description', 'pagetitle')


def _md_front_matter(lines, blocks):
    if not lines or lines[0].strip() != '---':
        return 0
    end = next((j for j in range(1, len(lines)) if lines[j].strip() in {'---', '...'}), None)
    if end is None:
        return 0
    from common import safe_load_text
    data = safe_load_text('\n'.join(lines[1:end])) or {}
    if isinstance(data, dict):
        for key in _FRONT_KEYS:
            if isinstance(data.get(key), str):
                b = _attr_block(data[key], None, 1, f'attr:front-matter.{key}')
                if b:
                    blocks.append(b)
    return end + 1


def markdown_blocks(text):
    lines = text.split('\n')
    blocks, uncovered, anchor_ids = [], [], []
    start = _md_front_matter(lines, blocks)
    fence, divs, buf = None, [], _Buf()

    def flush():
        nonlocal buf
        b = buf.block()
        if b:
            blocks.append(b)
        buf = _Buf()

    for n in range(start, len(lines)):
        line, lineno = lines[n], n + 1
        if fence:
            if line.strip().startswith(fence):
                fence = None
            continue
        m = _FENCE.match(line)
        if m:
            flush()
            fence = m.group(1)
            continue
        m = _DIV_OPEN.match(line)
        if m:
            flush()
            cm = _ATTR_CLAIM.search(m.group(1))
            if cm:
                anchor_ids.append((cm.group(1), lineno))
            divs.append(cm.group(1) if cm else None)
            continue
        if _DIV_CLOSE.match(line):
            flush()
            if divs:
                divs.pop()
            continue
        if not line.strip():
            flush()
            continue
        outer = next((c for c in reversed(divs) if c), None)
        if re.search(r'<[^>]*\bdata-ia-', line):
            uncovered.append(f'行 {lineno}: Markdown 内の生 HTML の data-ia-* アンカー（Pandoc 属性で書く）')
        if _COMPUTED.search(line):
            uncovered.append(f'行 {lineno}: 計算される値（Quarto のショートコード・インラインコード）')
            line = _COMPUTED.sub(' ', line)
        line = _QUOTE.sub('', line)
        if _HEADING.match(line) or _ITEM.match(line) or line.lstrip().startswith('|'):
            flush()
        line = _HEADING.sub('', line)
        line = re.sub(r'^\s*[-*+]\s+', '', line)
        if line.lstrip().startswith('|'):
            line = line.replace('|', ' ')
            if re.fullmatch(r'[\s:\-]*', line):
                continue
        line = _CODE.sub(' ', line)
        line = _LINK.sub(r'\1', line)
        for am in _TAG_ATTR.finditer(line):
            value = next(g for g in am.groups()[1:] if g is not None)
            b = _attr_block(value, outer, lineno, f'attr:{am.group(1)}')
            if b:
                blocks.append(b)
        pos = 0
        for sm in _SPAN.finditer(line):
            buf.add(_TAG.sub(' ', line[pos:sm.start()]), outer, lineno)
            cm = _ATTR_CLAIM.search(sm.group(2))
            if cm:
                anchor_ids.append((cm.group(1), lineno))
            buf.add(_TAG.sub(' ', sm.group(1)), cm.group(1) if cm else outer, lineno)
            pos = sm.end()
        buf.add(_TAG.sub(' ', line[pos:]) + '\n', outer, lineno)
        if _HEADING.match(lines[n]) or lines[n].lstrip().startswith('|'):
            flush()
    flush()
    return blocks, uncovered, anchor_ids


def blocks_for(path):
    """(blocks, uncovered constructs, claim anchor (id, line) list), or None when no adapter exists. Raises UnicodeDecodeError."""
    suffix = path.suffix.lower()
    if suffix not in HTML_SUFFIXES | MD_SUFFIXES:
        return None
    text = path.read_bytes().decode('utf-8')
    return html_blocks(text) if suffix in HTML_SUFFIXES else markdown_blocks(text)
