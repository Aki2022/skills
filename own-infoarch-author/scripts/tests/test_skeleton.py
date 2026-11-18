"""The outline (v3) -> semantic HTML skeleton a factory styles without changing text, order or anchors."""
import copy
import json
import os
import subprocess
import sys
from html.parser import HTMLParser
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import dom  # noqa: E402
import skeleton  # noqa: E402
from test_check import OUTLINE, SSOT, write_repo  # noqa: E402

SCRIPTS = Path(__file__).resolve().parents[1]


class Anchors(HTMLParser):
    def __init__(self):
        super().__init__()
        self.seq, self.facts, self.stack, self.text = [], {}, [], []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        for k in ('data-ia-section', 'data-ia-copy'):
            if k in a:
                self.seq.append(a[k])
        if 'data-ia-fact' in a:
            self.stack.append(a['data-ia-fact'])
            self.text = []

    def handle_data(self, data):
        if self.stack:
            self.text.append(data)

    def handle_endtag(self, tag):
        if tag == 'span' and self.stack:
            self.facts[self.stack.pop()] = ''.join(self.text)


def parse(html):
    p = Anchors()
    p.feed(html)
    return p


ORDER = ['sec-0001', 'cpy-0001', 'cpy-0002', 'sec-0002', 'cpy-0003', 'sec-0003', 'cpy-0004', 'cpy-0005']


def test_skeleton_keeps_outline_order_anchors_and_fact_values():
    pages = skeleton.render(OUTLINE, SSOT)
    assert list(pages) == ['site/index.html']
    p = parse(pages['site/index.html'])
    assert p.seq == ORDER
    assert p.facts == {'fct-0001': '100社以上', 'fct-0002': '03-1234-5678'}
    assert '{{' not in pages['site/index.html']
    assert [k[1] for k in dom.expected(OUTLINE)[0]] == ORDER  # dom expects the order the skeleton writes


def test_order_comes_from_the_order_field_not_the_file():
    o = copy.deepcopy(OUTLINE)
    o['sections'] = dict(reversed(list(o['sections'].items())))
    o['copies'] = dict(reversed(list(o['copies'].items())))
    assert parse(skeleton.render(o, SSOT)['site/index.html']).seq == ORDER
    o['sections']['sec-0001']['order'], o['sections']['sec-0002']['order'] = 2, 1
    assert parse(skeleton.render(o, SSOT)['site/index.html']).seq[:3] == ['sec-0002', 'cpy-0003', 'sec-0001']


def test_copy_is_html_escaped():
    o = copy.deepcopy(OUTLINE)
    o['copies']['cpy-0002']['text'] = '<script>alert(1)</script> {{fct-0001}} & <b>運用</b>'
    html = skeleton.render(o, SSOT)['site/index.html']
    assert '<script>alert' not in html and '&lt;script&gt;' in html and '&amp;' in html and '<b>' not in html


def test_every_output_shows_the_whole_outline():
    o = copy.deepcopy(OUTLINE)
    o['outputs'] = ['site/index.html', 'site/print.html']
    pages = skeleton.render(o, SSOT)
    assert all(parse(h).seq == ORDER for h in pages.values()) and len(pages) == 2


def test_cli_writes_skeleton_only_with_write_and_never_overwrites(tmp_path):
    write_repo(tmp_path)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, check=False)
    assert proc.returncode == 0 and 'data-ia-copy="cpy-0001"' in proc.stdout
    assert not (tmp_path / 'site/index.html').exists()
    out = tmp_path / 'skeleton'
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(tmp_path),
                           '--outline', 'lp-municipal', '--out', str(out)], capture_output=True, text=True, check=False)
    assert proc.returncode == 0 and (out / 'site/index.html').is_file()
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(tmp_path),
                           '--outline', 'lp-municipal', '--out', str(out)], capture_output=True, text=True, check=False)
    assert proc.returncode == 2


def test_skeleton_refuses_an_outline_that_fails_check(tmp_path):
    o = copy.deepcopy(OUTLINE)
    o['sections']['sec-0002']['places'] = ['msg-0099']
    write_repo(tmp_path, outlines=(o,))
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, check=False)
    assert proc.returncode == 2
    assert 'IA_SKELETON_REFUSED' in [f['code'] for f in json.loads(proc.stdout)['findings']]


PW_DIR = os.environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR')


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_skeleton_passes_the_rendered_dom_check(tmp_path):
    write_repo(tmp_path)
    for rel, html in skeleton.render(OUTLINE, SSOT).items():
        (tmp_path / rel).parent.mkdir(parents=True, exist_ok=True)
        (tmp_path / rel).write_text(html, encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    assert proc.returncode == 0, json.loads(proc.stdout)['findings']


TEXTS = {'nav': {'kind': 'navigation', 'text': 'サービス 使い方'},
         'label-problem': {'kind': 'label', 'text': 'PROBLEM', 'section': 'sec-0001'}}


def test_declared_texts_carry_their_key():
    o = copy.deepcopy(OUTLINE)
    o['texts'] = TEXTS
    html = skeleton.render(o, SSOT)['site/index.html']
    assert 'data-ia-text="nav"' in html and 'data-ia-text="label-problem"' in html
    assert html.index('data-ia-text="label-problem"') < html.index('data-ia-copy="cpy-0001"')
    assert parse(html).seq == ORDER  # section and copy anchors unchanged
    assert 'data-ia-text' not in skeleton.render(OUTLINE, SSOT)['site/index.html']


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
def test_skeleton_of_a_declared_outline_leaves_no_unanchored_text(tmp_path):
    from test_check import DECK
    o = copy.deepcopy(OUTLINE)
    o['texts'] = TEXTS
    write_repo(tmp_path, outlines=(o, DECK))
    for rel, html in skeleton.render(o, SSOT).items():
        (tmp_path / rel).parent.mkdir(parents=True, exist_ok=True)
        (tmp_path / rel).write_text(html, encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'dom', '--root', str(tmp_path),
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    assert proc.returncode == 0 and not report['unanchored'], report['findings']
