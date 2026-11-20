"""The HTML skeleton of one blueprint screen: the marks the build keeps and the render check reads.

A region is `data-ia-region` (with its priority and shape), a group `data-ia-group`, an action `data-ia-action`
(`data-ia-primary="true"` on the primary one) inside its region, a declared state `data-ia-state` (hidden until the page is
in that state; it is drawn from the start so it can be checked). Source order follows the priority, but the check reads the
rendered position, never the source order. Text is HTML-escaped; nothing here decides how it looks (design does).
"""
from __future__ import annotations

import html


def ordered_groups(screen):
    """[(group id, group, [(region id, region)])] — groups by their smallest priority, regions by priority."""
    by_group = {}
    for rid, reg in screen['regions'].items():
        by_group.setdefault(reg['group'], []).append((rid, reg))
    out = []
    for gid, regs in by_group.items():
        regs.sort(key=lambda r: (r[1].get('priority', 10 ** 9), r[0]))
        out.append((gid, screen['groups'][gid], regs))
    out.sort(key=lambda g: (g[2][0][1].get('priority', 10 ** 9), g[0]))
    return out


def body(screen):
    actions_of = {}
    for aid, act in sorted(screen['actions'].items()):
        actions_of.setdefault(act['region'], []).append((aid, act))
    lines = [f'<main data-ia-screen="{html.escape(screen["id"])}">', f'  <h1>{html.escape(screen["title"])}</h1>']
    for gid, grp, regs in ordered_groups(screen):
        lines.append(f'  <section data-ia-group="{gid}">')
        lines.append(f'    <h2>{html.escape(grp["text"])}</h2>')
        for rid, reg in regs:
            lines.append(f'    <div data-ia-region="{rid}" data-ia-priority="{reg.get("priority", "")}" data-ia-shape="{reg["shape"]}">')
            lines.append(f'      <p>{html.escape(reg["text"])}</p>')
            for aid, act in actions_of.get(rid, []):
                primary = ' data-ia-primary="true"' if act.get('primary') else ''
                lines.append(f'      <button type="button" data-ia-action="{aid}"{primary}>{html.escape(act["text"])}</button>')
            lines.append('    </div>')
        lines.append('  </section>')
    lines.append('  <aside data-ia-states>')
    for name, decl in sorted(screen['states'].items()):
        where = f' data-ia-state-region="{decl["region"]}"' if 'region' in decl else ''
        lines.append(f'    <div data-ia-state="{name}"{where} hidden>{html.escape(decl["text"])}</div>')
    lines.append('  </aside>')
    lines.append('</main>')
    return '\n'.join(lines)


def render(screen, css=''):
    style = f'<style>\n{css}\n</style>' if css else ''
    return ('<!doctype html>\n<html lang="ja">\n<head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>{html.escape(screen["title"])}</title>{style}</head>\n<body>\n{body(screen)}\n</body>\n</html>\n')
