"""script (台本) -> semantic HTML skeleton (one page = one script).

The factory (own-web-build) styles this skeleton with the design system. It must not change the text, the order
of blocks and copy, or the data-ia-* attributes: those carry the information design, and `dom` checks them after
rendering. Each block is `data-ia-section` (the v6 anchor name, kept until the factories rename it) with its content
function in `data-ia-function`, each piece of copy `data-ia-copy`, in their `order`. Facts are
expanded to their dictionary value inside <span data-dict="<dictionary id>">. Declared non-claim text is `data-ia-text`.
"""
from __future__ import annotations

import html

import check
import outline as ol


def _copy(text, facts):
    out, pos = [], 0
    for m in check.PLACEHOLDER.finditer(text):
        out.append(html.escape(text[pos:m.start()]))
        fid = m.group(1)
        out.append(f'<span data-dict="{fid}">{html.escape(facts[fid])}</span>')
        pos = m.end()
    out.append(html.escape(text[pos:]))
    return ''.join(out)


def body(o, ssot):
    facts = {fid: f['value'] for fid, f in ssot['facts'].items() if f.get('value')}
    texts = o.get('texts') or {}
    lines = [f'  <nav data-ia-text="{k}">{_copy(t["text"], facts)}</nav>'
             for k, t in sorted(texts.items()) if t['kind'] == 'navigation']
    lines.append(' <main>')
    for sid, blk, copies in ol.ordered_blocks(o):
        lines.append(f'  <section data-ia-section="{sid}" data-ia-function="{blk["function"]}">')
        lines += [f'    <h2 data-ia-text="{k}">{_copy(t["text"], facts)}</h2>'
                  for k, t in sorted(texts.items()) if t['kind'] == 'label' and t.get('block') == sid]
        lines += [f'    <p data-ia-copy="{cid}">{_copy(c["text"], facts)}</p>' for cid, c in copies]
        lines.append('  </section>')
    lines.append(' </main>')
    return '\n'.join(lines)


def render(o, ssot):
    """{output path: HTML} — every output shows the whole script (one page = one script)."""
    page = ('<!doctype html>\n<html lang="ja">\n<head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(o["id"])}</title></head>\n'
            '<body data-ia-outline="' + html.escape(o['id']) + '">\n' + body(o, ssot) + '\n</body>\n</html>\n')
    return {out: page for out in o['outputs']}
