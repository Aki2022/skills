"""Information-architecture marks on a factory page (biz_ops WS-20261114 D2/D16; IA v6 requirement 10).

The factory keeps the outline's marks (data-ia-section / data-ia-copy / data-ia-text / data-dict) and adds nothing but attributes: the marks must never change how the page looks. L1 refuses CSS that
selects a mark; with a real browser, the page with and without its marks renders to identical pixels.
"""
import json
import os
from pathlib import Path
import re
import subprocess
import sys

import pytest

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / 'own-web-build/scripts'))
import audit_html  # noqa: E402

PAGE = '''<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>body{{font:16px sans-serif;margin:0;color:#111;background:#fff}} main{{max-width:40rem;margin:auto}} h1{{font-size:2rem}}
a:focus-visible{{outline:2px solid #00f}} @media (max-width:600px){{h1{{font-size:1.5rem}}}} {extra}</style></head><body>
<main><section data-ia-section="sec-0001" data-ia-role="pain"><h1 data-ia-copy="cpy-0001">運用、重くありませんか</h1>
<p data-ia-copy="cpy-0002">運用の手間が半分になります</p></section>
<section data-ia-section="sec-0002" data-ia-role="act"><a href="#" data-ia-copy="cpy-0003">資料請求はこちら</a></section></main>
</body></html>'''


def rows(html):
    return {label: ok for ok, label in audit_html.check_structure(html)}


def mark_row(html):
    hits = [ok for label, ok in rows(html).items() if 'ia-marks-inert' in label]
    return hits[0] if hits else None


def test_marks_without_styling_pass_the_l1_row():
    assert mark_row(PAGE.format(extra='')) is True


@pytest.mark.parametrize('css', ['[data-ia-copy]{color:red}', 'section[data-ia-role="act"] a{font-weight:700}',
                                 '[data-ia-section="sec-0002"]{display:none}'])
def test_css_that_selects_a_mark_fails_the_l1_row(css):
    assert mark_row(PAGE.format(extra=css)) is False


def test_a_page_without_marks_has_no_mark_row():
    plain = re.sub(r' data-ia-[a-z]+="[^"]*"', '', PAGE.format(extra=''))
    assert mark_row(plain) is None  # D2: only an outlined page carries marks


PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')
SHOT = '''import {{ createRequire }} from "node:module";
const pw = createRequire("{base}/")("{base}/node_modules/playwright");
const b = await pw.chromium.launch(); const out = [];
for (const [w, h] of [[375, 812], [1280, 800]]) {{
  const c = await b.newContext({{ viewport: {{ width: w, height: h }}, deviceScaleFactor: 1 }}); const p = await c.newPage();
  for (const f of process.argv.slice(2)) {{ await p.goto("file://" + f); out.push((await p.screenshot({{ fullPage: true }})).toString("base64")); }}
  await c.close();
}}
await b.close(); console.log(JSON.stringify(out));'''


def shots(tmp_path, *pages):
    files = []
    for i, html in enumerate(pages):
        f = tmp_path / f'p{i}.html'
        f.write_text(html, encoding='utf-8')
        files.append(str(f))
    script = tmp_path / 'shot.mjs'
    script.write_text(SHOT.format(base=PW_DIR), encoding='utf-8')
    run = subprocess.run(['node', str(script), *files], capture_output=True, text=True, check=False)
    assert run.returncode == 0, run.stderr
    data = json.loads(run.stdout)
    n = len(pages)
    return [data[i::n] for i in range(n)]  # per page: one screenshot per viewport


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser snapshot not run (reported as skipped, not passed)')
def test_marks_do_not_change_the_rendered_pixels(tmp_path):
    marked = PAGE.format(extra='')
    plain = re.sub(r' data-ia-[a-z]+="[^"]*"', '', marked)
    styled = PAGE.format(extra='[data-ia-copy="cpy-0002"]{color:#c00}')  # the control: a styled mark must show
    a, b, c = shots(tmp_path, marked, plain, styled)
    assert a == b, 'adding the IA marks changed how the page looks'
    assert a != c, 'the snapshot cannot see a change (control failed)'


def test_css_that_selects_a_dictionary_mark_fails_the_l1_row():
    page = PAGE.replace('<h1 data-ia-copy="cpy-0001">', '<h1 data-ia-copy="cpy-0001"><span data-dict="value-answer-time">10分</span>')
    assert mark_row(page.format(extra='')) is True
    assert mark_row(page.format(extra='[data-dict]{font-weight:700}')) is False
