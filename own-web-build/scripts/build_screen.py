#!/usr/bin/env python3
"""Build a product UI screen from its blueprint (infoarch) and design's allocation rules.

`infoarch-author skeleton --screen` gives the marked HTML (regions, groups, actions, states); this adds the classes and the
stylesheet that design's allocation rules (own-design-author/references/allocation-rules.json) derive from each region's
shape and the screen's breakpoints: width by shape, one column on a phone, the primary action filled and the rest outlined.
It never moves a region: the order on the page stays the priority order, and `audit_screen.py` measures the result.
The stylesheet colours are neutral placeholders; the repository's design/ tokens replace them in a real build (L0/L1).

Usage: build_screen.py --repo <repo root> --screen scr-… --out page.html
Exit: 0 written, 2 refused (the blueprint fails, the screen is unknown, the file exists) or the rules are unreadable.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[2]
INFOARCH = SKILLS / 'own-infoarch-author/scripts/infoarch_author.py'
RULES = SKILLS / 'own-design-author/references/allocation-rules.json'


def load_rules(path=RULES):
    rules = json.loads(Path(path).read_text(encoding='utf-8'))
    for key in ('grid', 'breakpoints', 'width_by_shape'):
        if key not in rules:
            raise ValueError(f'allocation rules without {key}')
    return rules


def css_for(rules):
    """The stylesheet the rules derive: a 12-column grid per group, a span per shape and breakpoint, filled primary."""
    g = rules['grid']
    out = [':root{--bg:#fff;--fg:#1b1b1b;--line:#d0d0d0;--accent:#1f4fbf}',
           '@media (prefers-color-scheme:dark){:root{--bg:#161616;--fg:#eee;--line:#444;--accent:#8fb0ff}}',
           f'body{{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 system-ui,sans-serif}}',
           f'main{{max-width:{g["max_width_px"]}px;margin:0 auto;padding:16px}}',
           f'.grp{{display:grid;grid-template-columns:repeat({g["columns"]},minmax(0,1fr));gap:{g["gap_px"]}px;margin:0 0 {g["gap_px"]}px}}',
           '.grp>h2{grid-column:1/-1;font-size:14px;margin:0}',
           '.region{border:1px solid var(--line);border-radius:8px;padding:12px;min-width:0}',
           '.btn{font:inherit;padding:8px 16px;border-radius:6px;border:1.5px solid var(--accent);background:transparent;color:var(--accent);cursor:pointer;min-height:44px}',
           '.btn.primary{background:var(--accent);color:var(--bg)}',
           ':focus-visible{outline:2px solid var(--accent);outline-offset:2px}',
           '[hidden]{display:none!important}']
    for name in rules['width_by_shape']:
        out.append(f'.span-{name}{{grid-column:span {rules["width_by_shape"][name]["desktop"]}}}')
    for bp in rules['breakpoints']:
        if bp['name'] == 'desktop':
            continue
        cond = f'(max-width:{bp["max_width_px"]}px)' if bp['name'] == 'mobile' else f'(min-width:{bp["min_width_px"]}px) and (max-width:{bp["max_width_px"]}px)'
        inner = ''.join(f'.span-{n}{{grid-column:span {w[bp["name"]]}}}' for n, w in rules['width_by_shape'].items())
        if bp['name'] == 'mobile':
            inner += '.grp{display:block}.grp>.region{margin-bottom:%dpx}' % g['gap_px']
        out.append(f'@media {cond}{{{inner}}}')
    return '\n'.join(out)


def classify(page):
    page = re.sub(r'<section data-ia-group=', '<section class="grp" data-ia-group=', page)
    page = re.sub(r'<div data-ia-region="([^"]+)" data-ia-priority="([^"]*)" data-ia-shape="([^"]+)"',
                  lambda m: f'<div class="region span-{m.group(3)}" data-ia-region="{m.group(1)}" data-ia-priority="{m.group(2)}" data-ia-shape="{m.group(3)}"', page)
    page = re.sub(r'<button type="button" data-ia-action="([^"]+)" data-ia-primary="true"', r'<button type="button" class="btn primary" data-ia-action="\1" data-ia-primary="true"', page)
    page = re.sub(r'<button type="button" data-ia-action="([^"]+)">', r'<button type="button" class="btn" data-ia-action="\1">', page)
    return page


def build(repo, screen, rules_path=RULES):
    rules = load_rules(rules_path)
    with tempfile.TemporaryDirectory() as tmp:
        css = Path(tmp) / 'screen.css'
        css.write_text(css_for(rules), encoding='utf-8')
        proc = subprocess.run([os.environ.get('DESIGN_AUTHOR_PYTHON') or sys.executable, str(INFOARCH), 'skeleton', '--root', str(repo), '--screen', screen, '--css', str(css)],
                              capture_output=True, text=True)
    if proc.returncode != 0:
        raise ValueError(f'skeleton refused (exit {proc.returncode}): {proc.stdout.strip()[:300] or proc.stderr.strip()[:300]}')
    return classify(proc.stdout)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo', required=True)
    ap.add_argument('--screen', required=True)
    ap.add_argument('--out', required=True)
    args = ap.parse_args(argv)
    out = Path(args.out)
    try:
        if out.exists() or out.is_symlink():
            raise ValueError(f'上書きしない: {out}')
        page = build(args.repo, args.screen)
    except (ValueError, OSError) as exc:
        print(json.dumps({'exit_code': 2, 'error': str(exc)}, ensure_ascii=False))
        return 2
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(page, encoding='utf-8')
    print(json.dumps({'exit_code': 0, 'written': str(out)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
