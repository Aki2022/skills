"""Text extraction adapters: artifact file -> segments (text, line, innermost claim anchor).

HTML: text nodes per block, plus user-facing attributes (alt, title, aria-*, placeholder, meta content).
Markdown / Quarto: lines, Pandoc spans `[..]{data-ia-claim=..}` and fenced divs `::: {data-ia-claim=..}`.
Other formats are reported as uncovered by the caller. Anchors other than claims belong to ISSUE-02.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from html.parser import HTMLParser

CLAIM_ATTR = 'data-ia-claim'
HTML_SUFFIXES = {'.html', '.htm'}
MD_SUFFIXES = {'.md', '.qmd', '.markdown'}
VOID = {'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr'}
SKIP = {'script', 'style', 'template'}
BLOCK = {'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th', 'dt', 'dd', 'div', 'section', 'article',
         'header', 'footer', 'main', 'nav', 'aside', 'blockquote', 'figcaption', 'button', 'title', 'tr', 'ul',
         'ol', 'table', 'form', 'label', 'option', 'summary', 'details', 'figure', 'body', 'head', 'html'}
TEXT_ATTRS = {'alt', 'title', 'aria-label', 'aria-description', 'placeholder'}
META_NAMES = {'description', 'og:title', 'og:description', 'twitter:title', 'twitter:description'}


@dataclass(frozen=True)
class Segment:
    text: str
    line: int
    claim: str | None
    origin: str  # 'text' or 'attr:<name>'


class _HTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack = []  # (tag, claim)
        self.buf, self.buf_line, self.buf_claim = [], 0, None
        self.segments = []
        self.skip_depth = 0

    def _claim(self):
        for _, c in reversed(self.stack):
            if c:
                return c
        return None

    def flush(self):
        text = ' '.join(''.join(self.buf).split())
        if text:
            self.segments.append(Segment(text, self.buf_line, self.buf_claim, 'text'))
        self.buf = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in SKIP:
            self.skip_depth += 1
            return
        own_claim = (a.get(CLAIM_ATTR) or '').strip() or None
        if tag in BLOCK or own_claim:
            self.flush()
        claim = own_claim or self._claim()
        line = self.getpos()[0]
        for name in sorted(TEXT_ATTRS & set(a)):
            if a[name] and a[name].strip():
                self.segments.append(Segment(' '.join(a[name].split()), line, claim, f'attr:{name}'))
        if tag == 'input' and a.get('type') in {'submit', 'button'} and a.get('value'):
            self.segments.append(Segment(a['value'].strip(), line, claim, 'attr:value'))
        if tag == 'meta' and (a.get('name') in META_NAMES or a.get('property') in META_NAMES) and a.get('content'):
            self.segments.append(Segment(' '.join(a['content'].split()), line, claim, 'attr:content'))
        if tag == 'br':
            self.buf.append(' ')
        if tag not in VOID:
            self.stack.append((tag, own_claim))
            if own_claim:
                self.buf_claim = claim

    def handle_endtag(self, tag):
        if tag in SKIP:
            self.skip_depth = max(0, self.skip_depth - 1)
            return
        idx = next((i for i in range(len(self.stack) - 1, -1, -1) if self.stack[i][0] == tag), None)
        if idx is None:
            return
        closing_claim = any(c for _, c in self.stack[idx:])
        if tag in BLOCK or closing_claim:
            self.flush()
        del self.stack[idx:]

    def handle_data(self, data):
        if self.skip_depth:
            return
        if not self.buf or self.buf_claim != self._claim():
            if self.buf and self.buf_claim != self._claim():
                self.flush()
            self.buf_line = self.getpos()[0]
            self.buf_claim = self._claim()
        self.buf.append(data)


def html_segments(text):
    p = _HTML()
    p.feed(text)
    p.close()
    p.flush()
    return p.segments


_FENCE = re.compile(r'^\s*(```|~~~)')
_DIV_OPEN = re.compile(r'^\s*:{3,}\s*\{([^}]*)\}\s*$')
_DIV_CLOSE = re.compile(r'^\s*:{3,}\s*$')
_SPAN = re.compile(r'\[([^\[\]]*)\]\{([^}]*)\}')
_ATTR_CLAIM = re.compile(r'data-ia-claim\s*=\s*"?([A-Za-z0-9-]+)"?')
_LINK = re.compile(r'!?\[([^\[\]]*)\]\([^)]*\)')
_CODE = re.compile(r'`[^`]*`')
_TAG = re.compile(r'<[^>]+>')
_FRONT_KEYS = {'title', 'subtitle', 'description', 'pagetitle'}


def markdown_segments(text):
    lines = text.split('\n')
    segs, i, fence, divs = [], 0, None, []
    if lines and lines[0].strip() == '---':
        end = next((j for j in range(1, len(lines)) if lines[j].strip() in {'---', '...'}), None)
        if end is not None:
            for j in range(1, end):
                m = re.match(r'^(\w+)\s*:\s*(.+)$', lines[j])
                if m and m.group(1) in _FRONT_KEYS:
                    segs.append(Segment(m.group(2).strip().strip('"\''), j + 1, None, 'attr:front-matter'))
            i = end + 1
    for n in range(i, len(lines)):
        line, lineno = lines[n], n + 1
        if fence:
            if line.strip().startswith(fence):
                fence = None
            continue
        m = _FENCE.match(line)
        if m:
            fence = m.group(1)
            continue
        m = _DIV_OPEN.match(line)
        if m:
            cm = _ATTR_CLAIM.search(m.group(1))
            divs.append(cm.group(1) if cm else None)
            continue
        if _DIV_CLOSE.match(line):
            if divs:
                divs.pop()
            continue
        outer = next((c for c in reversed(divs) if c), None)
        line = _CODE.sub(' ', line)
        line = _LINK.sub(r'\1', line)
        pos, rest = 0, []
        for sm in _SPAN.finditer(line):
            rest.append(line[pos:sm.start()])
            cm = _ATTR_CLAIM.search(sm.group(2))
            inner = _TAG.sub(' ', sm.group(1))
            if cm:
                if inner.strip():
                    segs.append(Segment(' '.join(inner.split()), lineno, cm.group(1), 'text'))
            else:
                rest.append(inner)
            rest.append(' ')
            pos = sm.end()
        rest.append(line[pos:])
        plain = ' '.join(_TAG.sub(' ', ''.join(rest)).split())
        plain = re.sub(r'^#{1,6}\s+', '', plain)
        if plain:
            segs.append(Segment(plain, lineno, outer, 'text'))
    return segs


def segments_for(path):
    suffix = path.suffix.lower()
    text = path.read_text(encoding='utf-8')
    if suffix in HTML_SUFFIXES:
        return html_segments(text)
    if suffix in MD_SUFFIXES:
        return markdown_segments(text)
    return None
