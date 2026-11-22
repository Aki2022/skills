"""The approval request page (SPEC-deliverable-production-layers requirement 12): built by machine from the data the
approval hash binds, so what the owner reads is what the hash covers."""
import copy
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

import approval_request as ar
import positioning as pos
import positioning_model as pm
from ia_fixtures import SSOT, confirm, write_positioning

SCRIPTS = Path(__file__).resolve().parents[1]


def model(ssot=SSOT):
    return pm.Model(root=Path('.'), ssot=ssot)


def embedded(page):
    return json.loads(re.search(r'<script id="data" type="application/json">(.*?)</script>', page, re.S).group(1).replace('<\\/', '</'))


def test_every_value_path_has_one_card_whose_hash_is_the_approval_hash():
    m = model()
    items = ar.build(m)
    targets = pos.approval_targets(m)
    assert {i['key']: i['value_hash'] for i in items} == targets and len(items) == 3
    page = ar.page(items)
    cards = embedded(page)['cards']
    assert {c['key']: c['value_hash'] for c in cards.values()} == targets
    for i in items:
        assert page.count(f'data-key="{i["key"]}"') == 1 and page.count(f'>{i["hash12"]}<') == 1


def test_the_hashes_on_the_page_are_exactly_what_the_ledger_verifies():
    """Rows built from the page's own data approve every path: the page and the ledger bind the same bytes."""
    m = model()
    rows = {c['key']: c['value_hash'] for c in embedded(ar.page(ar.build(m)))['cards'].values()}
    assert pos.approved_keys(m, rows) == {'val-0001', 'val-0002', 'val-0003'}


def test_each_card_shows_the_whole_path_in_reading_order():
    m = model()
    item = next(i for i in ar.build(m) if i['id'] == 'val-0001')
    assert [r['role'] for r in item['roles']] == ['claim', 'basis', 'premise', 'evidence']
    shown = {x['key'] for r in item['roles'] for x in r['rows']}
    assert shown == set(pos.paths(m)['val-0001'].roles)  # nothing the hash covers is left out, nothing else is added
    page = ar.page([item])
    for key in shown:
        assert key in page


def test_changing_one_target_changes_the_page_and_only_that_targets_hash():
    m = model()
    before = ar.build(m)
    page_before = ar.page(before)
    changed = copy.deepcopy(SSOT)
    changed['alternatives']['alt-0002']['text'] = '保守を別の会社に委託する'  # on seg-0001's paths, not seg-0002's
    m2 = model(changed)
    after = ar.build(m2)
    page_after = ar.page(after)
    assert page_after != page_before and '別の会社に' in page_after and '別の会社に' not in page_before
    old, new = ({i['id']: i['value_hash'] for i in items} for items in (before, after))
    assert old['val-0001'] != new['val-0001']  # the path the changed alternative stands on
    assert old['val-0002'] == new['val-0002'] and old['val-0003'] == new['val-0003']  # a path it is not on
    # the new page matches the new data, and the old page's hashes no longer verify
    assert {c['key']: c['value_hash'] for c in embedded(page_after)['cards'].values()} == pos.approval_targets(m2)
    old_rows = {c['key']: c['value_hash'] for c in embedded(page_before)['cards'].values()}
    assert pos.approved_keys(m2, old_rows) == {'val-0002', 'val-0003'}


def test_approved_and_required_paths_are_marked():
    m = model()
    targets = pos.approval_targets(m)
    items = ar.build(m, approved={'val-0001'}, must=['ia:path:val-0002'])
    by = {i['id']: i for i in items}
    assert by['val-0001']['approved'] and not by['val-0002']['approved'] and by['val-0002']['required']
    page = ar.page(items)
    assert page.count('承認済み') >= 1 and page.count('出荷に必要') >= 1 and len(targets) == 3


def test_a_disqualified_path_lists_what_the_owner_must_decide():
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if not (e['kind'] == 'fulfills' and e['from'] == 'alt-0002' and e['to'] == 'pan-0001')]
    m = model(ssot)
    items = ar.build(m)
    assert any(i['problems'] for i in items)
    page = ar.page(items)
    assert '決めてほしいこと' in page and any(p['code'] in page for i in items for p in i['problems'])


def test_the_page_offers_circle_and_cross_and_saves_to_the_db_collection():
    page = ar.page(ar.build(model()))
    assert '○ 承認する' in page and '× 差し戻す' in page
    assert "claude.use('db')" in page and embedded(page)['collection'] == 'decisions'
    assert 'id="nodb"' in page and 'id="copy"' in page  # a page that cannot save says so and offers a copy


def test_the_page_is_a_body_for_the_artifact_tool_and_keeps_the_page_contract():
    page = ar.page(ar.build(model()))
    head = page[:600]
    assert head.lstrip().startswith('<title>') and 'claude:viewport-fit-sides' in head
    assert not re.search(r'<(!doctype|html|head|body)[\s>]', page.lower())  # <header> inside a card is fine
    assert 'padding-left:max(16px,env(safe-area-inset-left,0px))' in page and 'padding-right:max(16px,env(safe-area-inset-right,0px))' in page
    assert page.count('color-scheme:dark') >= 2 and ':root:not([data-theme="light"])' in page and ':root[data-theme="dark"]' in page
    assert 'background:var(--bg)' in page


def test_text_is_escaped_in_the_page_and_in_the_embedded_data():
    ssot = copy.deepcopy(SSOT)
    ssot['values']['val-0001']['text'] = '</script><script>alert(1)</script>'
    page = ar.page(ar.build(model(ssot)))
    assert '<script>alert(1)' not in page
    assert page.count('</script>') == 2  # the data block and the code block, nothing the text could add


@pytest.mark.skipif(shutil.which('node') is None, reason='node is not installed')
def test_the_pages_script_is_valid_javascript(tmp_path):
    page = ar.page(ar.build(model()))
    code = re.findall(r'<script>(.*?)</script>', page, re.S)[-1]
    f = tmp_path / 'page.js'
    f.write_text(code, encoding='utf-8')
    proc = subprocess.run(['node', '--check', str(f)], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr


def test_saved_decisions_are_read_back_against_the_current_hashes():
    items = ar.build(model())
    by = {i['id']: i for i in items}
    row = lambda i, decision, hash_=None: {'id': i, 'data': {'key': by[i]['key'], 'value_hash': hash_ or by[i]['value_hash'], 'decision': decision}}
    out = ar.decisions_from_rows(items, [row('val-0001', 'ok'), row('val-0002', 'ng'), row('val-0003', 'ok', 'f' * 64),
                                         {'id': 'val-9999', 'data': {'key': 'ia:path:val-9999'}}])
    assert out == {'ok': ['ia:path:val-0001'], 'ng': ['ia:path:val-0002'], 'stale': ['ia:path:val-0003'],
                   'undecided': [], 'unknown': ['val-9999']}
    assert ar.decisions_from_rows(items, [])['undecided'] == sorted(i['key'] for i in items)


def test_request_rows_have_the_shape_approve_prepare_binds():
    items = ar.build(model())
    assert ar.request_rows(items) == [{'key': i['key'], 'value_hash': i['value_hash']} for i in items]


# ---- the command --------------------------------------------------------------------------------------------------

def cli(*args):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'positioning_author.py'), *args], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


def repo(tmp_path, ssot=SSOT):
    write_positioning(tmp_path, ssot)
    confirm(tmp_path, ssot)
    return tmp_path


def test_the_command_writes_the_page_and_the_rows_and_refuses_to_overwrite(tmp_path):
    root = repo(tmp_path)
    out = tmp_path / 'out'
    rc, res = cli('request', '--root', str(root), '--out', str(out))
    assert rc == 0 and res['cards'] == 3 and res['approved'] == 3
    page = (out / 'approval-request.html').read_text(encoding='utf-8')
    rows = json.loads((out / 'approval-request.json').read_text(encoding='utf-8'))['rows']
    assert {r['key']: r['value_hash'] for r in rows} == {c['key']: c['value_hash'] for c in embedded(page)['cards'].values()}
    rc, res = cli('request', '--root', str(root), '--out', str(out))
    assert rc == 2 and '上書きしない' in res['error']


def test_the_command_reads_the_owners_answers_back(tmp_path):
    root = repo(tmp_path)
    m = model()
    targets = pos.approval_targets(m)
    saved = [{'id': 'val-0001', 'data': {'key': 'ia:path:val-0001', 'value_hash': targets['ia:path:val-0001'], 'decision': 'ok'}},
             {'id': 'val-0002', 'data': {'key': 'ia:path:val-0002', 'value_hash': 'a' * 64, 'decision': 'ok'}}]
    f = tmp_path / 'decisions.json'
    f.write_text(json.dumps(saved), encoding='utf-8')
    rc, res = cli('request', '--root', str(root), '--read-decisions', str(f))
    assert rc == 1  # a stale answer is not silently counted
    assert res['ok'] == ['ia:path:val-0001'] and res['stale'] == ['ia:path:val-0002'] and res['undecided'] == ['ia:path:val-0003']


def test_the_command_marks_the_paths_shipping_needs(tmp_path):
    root = repo(tmp_path)
    must = tmp_path / 'targets.json'
    must.write_text(json.dumps({'approval_hashes': {}, 'must_approval_keys': ['ia:path:val-0001']}), encoding='utf-8')
    rc, _ = cli('request', '--root', str(root), '--out', str(tmp_path / 'o'), '--must-file', str(must))
    assert rc == 0
    assert (tmp_path / 'o/approval-request.html').read_text(encoding='utf-8').count('出荷に必要') >= 1


def test_no_paths_is_not_a_page(tmp_path):
    ssot = copy.deepcopy(SSOT)
    for block in ('segments', 'jobs', 'pains', 'gains', 'alternatives', 'categories', 'strengths', 'values', 'proofs'):
        ssot[block] = {}
    ssot['edges'] = []
    write_positioning(tmp_path, ssot)
    rc, res = cli('request', '--root', str(tmp_path), '--out', str(tmp_path / 'o'))
    assert rc == 2 and not (tmp_path / 'o').exists()


def test_the_answers_are_read_back_in_the_shape_an_artifact_read_writes_them(tmp_path):
    """`ArtifactData` list with out_dir writes <dir>/decisions/<doc id>.json holding the document body alone (observed
    on a real artifact, 2026-11-20); the page's own copy is a list of {id, data}; both must read back the same."""
    m = model()
    t = pos.approval_targets(m)
    folder = tmp_path / 'decisions'
    folder.mkdir()
    body = {'key': 'ia:path:val-0001', 'value_hash': t['ia:path:val-0001'], 'decision': 'ok', 'note': '', 'decided_at': '2026-11-20T00:00:00Z'}
    (folder / 'val-0001.json').write_text(json.dumps(body), encoding='utf-8')
    (folder / 'val-0002.json').write_text(json.dumps({**body, 'key': 'ia:path:val-0002', 'value_hash': 'b' * 64, 'decision': 'ng'}), encoding='utf-8')
    rows = ar.load_rows(folder)
    assert [r['id'] for r in rows] == ['val-0001', 'val-0002']
    out = ar.decisions_from_rows(ar.build(m), rows)
    assert out['ok'] == ['ia:path:val-0001'] and out['stale'] == ['ia:path:val-0002'] and out['unknown'] == []
    page_copy = tmp_path / 'copy.json'
    page_copy.write_text(json.dumps(rows), encoding='utf-8')
    assert ar.decisions_from_rows(ar.build(m), ar.load_rows(page_copy)) == out
    single = tmp_path / 'val-0001.json'
    single.write_text(json.dumps(body), encoding='utf-8')
    assert ar.load_rows(single) == [{'id': 'val-0001', 'data': body}]
    assert [r['id'] for r in ar.load_rows(folder / 'val-0001.json')] == ['val-0001']


# ---- the diagram: the path drawn as boxes and lines --------------------------------------------------------------

def test_the_diagram_draws_every_atom_and_arrow_the_hash_covers_and_nothing_else():
    m = model()
    for item in ar.build(m):
        path = pos.paths(m)[item['id']]
        atoms = {k for k in path.members if k in pos.atoms(m, proposed_only=False)}
        arrows = set(path.members) - atoms
        g = item['graph']
        assert {n['id'] for n in g['nodes']} == atoms and {e['key'] for e in g['edges']} == arrows
        svg = ar._diagram(g)
        assert svg.count('<g class="node') == len(atoms) and svg.count('marker-end') == len(arrows)
        for k in atoms:
            assert svg.count(f'class="nid">{k}<') == 1


def test_the_diagram_colours_the_degree_and_marks_the_claim():
    svg = ar._diagram(ar.build(model())[0]['graph'])
    assert 'deg-full' in svg or 'deg-partial' in svg or 'deg-none' in svg
    assert 'role-claim' in svg and 'role-premise' in svg and 'role-evidence' in svg
    assert '満たす' in svg and 'viewBox' in svg


def test_the_page_shows_the_diagram_before_the_text_list_and_keeps_the_list_available():
    page = ar.page(ar.build(model()))
    assert page.count('<svg class="graph"') == 3 and page.count('文章でも読む') == 3
    assert page.index('<svg class="graph"') < page.index('文章でも読む')
    assert 'overflow-x:auto' in page  # a wide diagram scrolls inside its own box, never the page
