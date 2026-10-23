"""Structured outline -> semantic HTML skeleton (ISSUE-06).

The factory (own-web-design) styles this skeleton with the design system. It must not change the text, the
order of sections or the data-ia-* attributes: those carry the information design, and `dom` checks them
after rendering. Facts are expanded to their SSOT value inside <span data-ia-fact>.
"""
from __future__ import annotations

import html

import check

ROLE_TAG = {'cta': 'p', 'navigation': 'nav'}


def pages_for_outputs(outline):
    """{output: [pages]} — one page per output when the counts match, else every page in every output."""
    pages, outputs = outline['pages'], outline['outputs']
    if len(pages) == len(outputs):
        return {out: [p] for out, p in zip(outputs, pages)}
    return {out: list(pages) for out in outputs}


def _copy(text, facts):
    out, pos = [], 0
    for m in check.PLACEHOLDER.finditer(text):
        out.append(html.escape(text[pos:m.start()]))
        fid = m.group(1)
        out.append(f'<span data-ia-fact="{fid}">{html.escape(facts[fid])}</span>')
        pos = m.end()
    out.append(html.escape(text[pos:]))
    return ''.join(out)


def _section(s, facts):
    attrs = [f'data-ia-role="{s["role"]}"']
    if 'problem' in s:
        attrs.append(f'data-ia-problem="{s["problem"]}"')
    if 'message' in s:
        attrs.append(f'data-ia-msg="{s["message"]}"')
    tag = ROLE_TAG.get(s['role'], 'p')
    return f'  <section {" ".join(attrs)}>\n    <{tag}>{_copy(s["copy"], facts)}</{tag}>\n  </section>'


def render(outline, ssot):
    facts = {fid: f['value'] for fid, f in ssot['facts'].items() if f.get('value')}
    out = {}
    for path, pages in pages_for_outputs(outline).items():
        body = []
        for p in pages:
            label = f' aria-label="{html.escape(p["label"])}"' if p.get('label') else ''
            body.append(f' <main data-ia-page="{p["id"]}"{label}>')
            body += [_section(s, facts) for s in p['sections']]
            body.append(' </main>')
        out[path] = ('<!doctype html>\n<html lang="ja">\n<head><meta charset="utf-8">'
                     f'<meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(outline["id"])}</title></head>\n'
                     '<body data-ia-outline="' + html.escape(outline['id']) + '">\n' + '\n'.join(body) + '\n</body>\n</html>\n')
    return out
