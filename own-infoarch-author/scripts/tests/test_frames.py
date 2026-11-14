"""Swappable frames of an LP (SPEC-information-architecture-standardization v5, requirements 16 and 17; D1, D2).

Frame types live in the media knowhow; an outline declares frame instances (unique ids) and anchors element
frames on sections. The skeleton carries data-ia-frame; dom checks the instances; the mapping table lets an
external experiment tool point at "which version, which frame, which atoms".
"""
import copy
import json
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import dom  # noqa: E402
import skeleton  # noqa: E402
from test_check import DECK, OUTLINE, SSOT, check, cli, codes, dump, write_repo  # noqa: E402
from test_dom import green_snapshot, anchor  # noqa: E402

FRAMES = {'media': 'lp', 'instances': [
    {'id': 'hero', 'type': 'first-view-message'}, {'id': 'cta-main', 'type': 'cta'},
    {'id': 'proof-1', 'type': 'proof-position'}, {'id': 'order', 'type': 'section-order'}]}


def framed():
    o = copy.deepcopy(OUTLINE)
    o['frames'] = copy.deepcopy(FRAMES)
    sec = o['pages'][0]['sections']
    sec[1]['frame'] = 'hero'
    sec[2]['frame'] = 'proof-1'
    sec[3]['frame'] = 'cta-main'
    return o


def framed_snapshot():
    snap = green_snapshot()
    snap['anchors'] += [anchor('frame', 'hero', 120), anchor('frame', 'proof-1', 900), anchor('frame', 'cta-main', 1500)]
    return snap


def dom_codes(snap, outline):
    return [c for c, level, _ in dom.compare(outline, SSOT, snap) if level == 'fail']


def test_lp_frame_types_are_the_four_of_d1():
    knowhow = json.loads((Path(__file__).resolve().parents[2] / 'references/media/frames.json').read_text(encoding='utf-8'))
    assert knowhow['media']['lp'] == {
        'first-view-message': {'target': 'element', 'min': 1, 'max': 1},
        'cta': {'target': 'element', 'min': 1, 'max': None},
        'proof-position': {'target': 'element', 'min': 0, 'max': None},
        'section-order': {'target': 'sequence', 'min': 1, 'max': 1}}


def test_framed_outline_is_green(tmp_path):
    write_repo(tmp_path, outlines=(framed(), DECK))
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']


@pytest.mark.parametrize('mutate,code', [
    (lambda o: o['frames']['instances'].append({'id': 'hero', 'type': 'cta'}), 'IA_FRAME_DUPLICATE_ID'),
    (lambda o: o['frames']['instances'].append({'id': 'hero-2', 'type': 'first-view-message'}), 'IA_FRAME_MULTIPLICITY'),
    (lambda o: o['frames']['instances'].append({'id': 'order-2', 'type': 'section-order'}), 'IA_FRAME_MULTIPLICITY'),
    (lambda o: o['frames']['instances'].remove({'id': 'cta-main', 'type': 'cta'}) or o['pages'][0]['sections'][3].pop('frame'), 'IA_FRAME_MULTIPLICITY'),
    (lambda o: o['frames']['instances'].append({'id': 'banner', 'type': 'hero-image'}), 'IA_FRAME_UNKNOWN_TYPE'),
    (lambda o: o['frames'].update(media='poster'), 'IA_FRAME_UNKNOWN_MEDIA'),
    (lambda o: o['pages'][0]['sections'][0].update(frame='nowhere'), 'IA_FRAME_UNDECLARED'),
    (lambda o: o['pages'][0]['sections'][0].update(frame='order'), 'IA_FRAME_SEQUENCE_ANCHORED'),
    (lambda o: o['pages'][0]['sections'][0].update(frame='hero'), 'IA_FRAME_DUPLICATE_USE'),
    (lambda o: o['pages'][0]['sections'][2].pop('frame'), 'IA_FRAME_UNUSED'),
])
def test_bad_frame_declarations_fail(tmp_path, mutate, code):  # acceptance (g), outline side
    o = framed()
    mutate(o)
    write_repo(tmp_path, outlines=(o, DECK))
    rc, report = check(tmp_path)
    assert rc == 2 and code in codes(report), report['findings']


def test_skeleton_adds_only_the_frame_attribute():  # D2: nothing but the attribute changes
    plain = skeleton.render(OUTLINE, SSOT)['site/index.html']
    html = skeleton.render(framed(), SSOT)['site/index.html']
    assert 'data-ia-frame="hero"' in html and 'data-ia-frame="order"' not in html
    assert re.sub(r' data-ia-frame="[^"]*"', '', html) == plain


def test_unframed_outline_renders_no_frame_attribute():
    assert 'data-ia-frame' not in skeleton.render(OUTLINE, SSOT)['site/index.html']


def test_framed_snapshot_is_green():
    assert dom_codes(framed_snapshot(), framed()) == []


@pytest.mark.parametrize('mutate,code', [
    (lambda s: s['anchors'].remove(next(a for a in s['anchors'] if a['id'] == 'hero')), 'IA_DOM_FRAME_MISSING'),
    (lambda s: s['anchors'].append(anchor('frame', 'hero', 130)), 'IA_DOM_FRAME_DUPLICATE'),
    (lambda s: s['anchors'].append(anchor('frame', 'banner', 50)), 'IA_DOM_FRAME_UNDECLARED'),
])
def test_bad_frame_anchors_fail(mutate, code):  # acceptance (g), rendered side
    snap = framed_snapshot()
    mutate(snap)
    assert code in dom_codes(snap, framed())


def test_mapping_table_binds_version_frames_and_atom_content(tmp_path):  # requirement 17
    write_repo(tmp_path, outlines=(framed(), DECK))
    rc, table = cli('mapping', '--root', str(tmp_path), '--outline', 'lp-municipal')
    assert rc == 0, table
    assert table['deliverable'] == 'lp-municipal' and len(table['version_hash']) == 64
    rows = {r['id']: r for r in table['frames']}
    assert rows['hero']['type'] == 'first-view-message' and 'msg-0001' in rows['hero']['atoms']
    assert rows['order']['type'] == 'section-order' and rows['order']['atoms'] == ['pan-0001', 'msg-0001', 'msg-0002', 'msg-0003']
    before = rows['hero']['atoms']['msg-0001']
    changed = copy.deepcopy(SSOT)
    changed['messages']['msg-0001']['text'] = '運用の手間が三分の一になる'
    dump(tmp_path / 'design/info_arch/info_arch.yaml', changed)
    rc, again = cli('mapping', '--root', str(tmp_path), '--outline', 'lp-municipal')
    hero = {r['id']: r for r in again['frames']}['hero']
    assert hero['atoms']['msg-0001'] != before  # same frame instance, same message id, new content -> new hash
    assert again['version_hash'] != table['version_hash']


@pytest.mark.skipif(not __import__('os').environ.get('INFOARCH_TEST_PLAYWRIGHT_DIR'),
                    reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('drop,expect', [(False, 0), (True, 2)])
def test_real_browser_reads_frame_anchors_of_the_skeleton(tmp_path, drop, expect):
    import os
    import subprocess
    write_repo(tmp_path, outlines=(framed(), DECK))
    html = skeleton.render(framed(), SSOT)['site/index.html']
    if drop:
        html = html.replace(' data-ia-frame="hero"', '')
    page = tmp_path / 'site/index.html'
    page.parent.mkdir(parents=True, exist_ok=True)
    page.write_text(html, encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=os.environ['INFOARCH_TEST_PLAYWRIGHT_DIR'])
    proc = subprocess.run([sys.executable, str(Path(__file__).resolve().parents[1] / 'infoarch_author.py'), 'dom',
                           '--root', str(tmp_path), '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    report = json.loads(proc.stdout)
    assert proc.returncode == expect, report
    if drop:
        assert 'IA_DOM_FRAME_MISSING' in [f['code'] for f in report['findings']]
