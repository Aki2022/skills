"""blueprint (the product UI's infoarch): the green fixture, every disqualification, one approval per item."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

from blueprint_fixtures import BP_DIR, OVERVIEW, PERSON, PRODUCT, SCREENS, apply_ops, confirm_items, item_hashes, write_blueprint
from ia_fixtures import SSOT, dump

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
DATE = '2026-11-20'
TABLE = json.loads((SKILL / 'fixtures/blueprint-red-cases.json').read_text(encoding='utf-8'))
CONDITIONS = {c['id']: c for c in TABLE['conditions']}
CASES = TABLE['cases']


def cli(*args):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), *args], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


def check(root):
    rc, report = cli('check', '--root', str(root), '--date', DATE)
    assert report['exit_code'] == rc
    return rc, report


def codes(report, level='fail'):
    return [f['code'] for f in report['findings'] if f['level'] == level]


# ---- the table: every condition has a red case, every case a condition -------------------------------------------

def test_every_condition_has_a_red_case_and_every_case_names_a_condition():
    covered = {c['condition'] for c in CASES}
    assert sorted(set(CONDITIONS) - covered) == []  # 対応の無い失格条件 0 件
    assert sorted(covered - set(CONDITIONS)) == []
    assert len({c['id'] for c in CASES}) == len(CASES)
    assert len(CONDITIONS) >= 20 and len(CASES) >= 30


def test_every_blueprint_code_is_raised_by_the_checker():
    """A condition whose code the checker never raises would sit in the table forever green."""
    source = (SKILL / 'scripts/blueprint.py').read_text(encoding='utf-8')
    for c in CONDITIONS.values():
        if c['code'] != 'IA_SCHEMA':
            assert f"'{c['code']}'" in source, c


def write_case(root, case):
    product, screens, unapproved = apply_ops(case['ops'])
    write_blueprint(root, product, screens, confirmed=False)
    if product is not None:
        confirm_items(root, product, skip=unapproved)


@pytest.mark.parametrize('case', CASES, ids=[c['id'] for c in CASES])
def test_red_case_fails_with_its_condition_code(case, tmp_path):
    write_case(tmp_path, case)
    rc, report = check(tmp_path)
    assert rc == 2, report['findings']
    assert CONDITIONS[case['condition']]['code'] in codes(report), codes(report)


# ---- green ---------------------------------------------------------------------------------------------------------

def test_green_fixture_passes_and_counts_what_it_checked(tmp_path):
    write_blueprint(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
    [info] = [f for f in report['findings'] if f['code'] == 'IA_BLUEPRINT_CHECKED']
    assert info['detail'] == {'screens': 2, 'items': 10}
    assert report['targets'] == {'channels': 0, 'outlines': 0, 'sections': 0}  # the bridge's three counts stay as they were


def test_a_repository_with_no_blueprint_and_no_outline_still_fails(tmp_path):
    dump(tmp_path / 'design/info_arch/info_arch.yaml', SSOT)
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_NO_OUTLINES' in codes(report)


def test_a_draft_screen_that_cites_an_unapproved_item_only_warns(tmp_path):
    draft = copy.deepcopy(OVERVIEW)
    draft['ship'] = False
    write_blueprint(tmp_path, screens=(draft, PERSON), confirmed=False)
    confirm_items(tmp_path, PRODUCT, skip=('pri-0001',))
    rc, report = check(tmp_path)
    warned = [f for f in report['findings'] if f['level'] == 'warn' and f['code'] == 'IA_BLUEPRINT_ITEM_UNAPPROVED']
    # PERSON is shipped and cites pri-0001 too, so this stays a failure; the draft screen alone only warns
    assert rc == 2 and warned and all('scr-overview' in f['message'] for f in warned)
    person = copy.deepcopy(PERSON)
    person['ship'] = False
    write_blueprint(tmp_path / 'b', screens=(draft, person), confirmed=False)
    confirm_items(tmp_path / 'b', PRODUCT, skip=('pri-0001',))
    rc, report = check(tmp_path / 'b')
    assert rc == 1 and 'IA_BLUEPRINT_ITEM_UNAPPROVED' in codes(report, 'warn') and not codes(report), report['findings']


def test_a_stale_approval_is_not_an_approval(tmp_path):
    write_blueprint(tmp_path)
    changed = copy.deepcopy(PRODUCT)
    changed['principles']['pri-0001']['text'] = '最も大事な状態を最上段に置く'
    dump(tmp_path / BP_DIR / 'product.yaml', changed)  # the owner's row is for the old text
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_BLUEPRINT_ITEM_UNAPPROVED' in codes(report)


def test_unapproved_values_under_a_principle_warn_without_stopping(tmp_path):
    write_blueprint(tmp_path)
    from ia_fixtures import confirm
    rows = tmp_path / 'design/approvals.yaml'
    import yaml
    kept = [r for r in yaml.safe_load(rows.read_text(encoding='utf-8')) if r['key'] != 'ia:path:val-0001']
    dump(rows, kept)
    rc, report = check(tmp_path)
    assert rc == 1 and 'IA_BLUEPRINT_VALUE_UNAPPROVED' in codes(report, 'warn') and not codes(report), report['findings']


# ---- approval is per item ------------------------------------------------------------------------------------------

def test_adding_an_item_moves_no_other_items_approval():
    before = item_hashes(PRODUCT)
    more = copy.deepcopy(PRODUCT)
    more['ledger']['next']['con'] = 4
    more['constraints']['con-0003'] = {'state': 'proposed', 'provenance': 'owner', 'text': '色だけで状態を伝えない'}
    after = item_hashes(more)
    assert {k: v for k, v in after.items() if k in before} == before  # every earlier approval still matches
    assert set(after) - set(before) == {'ia:item:con-0003'}


def test_changing_one_item_moves_only_that_items_hash():
    before = item_hashes(PRODUCT)
    changed = copy.deepcopy(PRODUCT)
    changed['stories']['sto-0001']['text'] = '運用の状態を、経営にひと目で伝えたい'
    after = item_hashes(changed)
    assert {k for k in before if before[k] != after[k]} == {'ia:item:sto-0001'}


def test_provenance_is_not_part_of_what_the_owner_approved():
    before = item_hashes(PRODUCT)
    changed = copy.deepcopy(PRODUCT)
    changed['stories']['sto-0002']['provenance'] = 'extracted:docs/x.md#y'
    assert item_hashes(changed) == before


def test_the_check_sees_an_added_item_as_still_approved_for_the_others(tmp_path):
    write_blueprint(tmp_path)
    more = copy.deepcopy(PRODUCT)
    more['ledger']['next']['con'] = 4
    more['constraints']['con-0003'] = {'state': 'proposed', 'provenance': 'proposal', 'text': '色だけで状態を伝えない'}
    dump(tmp_path / BP_DIR / 'product.yaml', more)  # con-0003 has no row; nothing cites it yet
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']


def test_targets_list_every_proposed_item_and_the_ones_shipping_needs(tmp_path):
    write_blueprint(tmp_path)
    rc, out = cli('targets', '--root', str(tmp_path))
    assert rc == 0
    items = sorted(k for k in out['approval_hashes'] if k.startswith('ia:item:'))
    assert items == sorted('ia:item:' + i for b in ('stories', 'principles', 'constraints', 'objects', 'entries') for i in PRODUCT[b])
    assert out['approval_hashes'] == {**{k: v for k, v in out['approval_hashes'].items() if k.startswith('ia:path:')}, **item_hashes(PRODUCT)}
    must = set(out['must_approval_keys'])
    assert {'ia:item:sto-0001', 'ia:item:sto-0002', 'ia:item:pri-0001', 'ia:item:con-0001', 'ia:item:ojc-0001', 'ia:item:ojc-0002', 'ia:item:ojc-0003'} <= must
    assert 'ia:item:con-0002' not in must and 'ia:item:pri-0002' not in must  # no shipped screen cites them


def test_the_bridge_accepts_item_keys():
    import ia_module
    assert all(ia_module.KEY.fullmatch(k) for k in item_hashes(PRODUCT))
    assert ia_module.KEY.fullmatch('ia:path:val-0001') and not ia_module.KEY.fullmatch('ia:item:xyz-0001')
    assert not ia_module.KEY.fullmatch('ia:positioning:val-0001')


def test_impact_lists_the_screens_that_cite_an_item(tmp_path):
    write_blueprint(tmp_path)
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'ojc-0003')
    assert rc == 0 and out['screens'] == ['scr-overview', 'scr-person']
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'con-0001')
    assert out['screens'] == ['scr-overview']
    rc, out = cli('impact', '--root', str(tmp_path), '--id', 'sto-9999')
    assert rc == 2


def test_an_outline_repository_is_unchanged_by_the_blueprint(tmp_path):
    from test_check import write_repo
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0 and not [f for f in report['findings'] if f['code'].startswith('IA_BLUEPRINT')]


def test_outlines_and_a_blueprint_live_in_one_repository(tmp_path):
    from test_check import write_repo
    write_repo(tmp_path)
    write_blueprint(tmp_path)  # same positioning, plus the product UI
    rc, report = check(tmp_path)
    assert rc == 0, report['findings']
    assert report['targets']['outlines'] == 2


def test_the_blueprint_command_reports_its_own_counts(tmp_path):
    write_blueprint(tmp_path)
    rc, report = cli('blueprint', '--root', str(tmp_path), '--date', DATE)
    assert rc == 0 and report['targets'] == {'screens': 2, 'items': 10}, report
    write_case(tmp_path / 'red', next(c for c in CASES if c['id'] == 'two-primary-actions'))
    rc, report = cli('blueprint', '--root', str(tmp_path / 'red'), '--date', DATE)
    assert rc == 2 and 'IA_BLUEPRINT_PRIMARY_MANY' in codes(report)


def test_the_blueprint_command_is_not_a_pass_without_a_blueprint(tmp_path):
    dump(tmp_path / 'design/info_arch/info_arch.yaml', SSOT)
    rc, report = cli('blueprint', '--root', str(tmp_path), '--date', DATE)
    assert rc == 2 and 'IA_BLUEPRINT_NO_SCREENS' in codes(report)


def test_dotfiles_in_screens_are_ignored_and_other_files_fail(tmp_path):
    write_blueprint(tmp_path)
    (tmp_path / BP_DIR / 'screens/.DS_Store').write_bytes(b'x')
    assert check(tmp_path)[0] == 0
    (tmp_path / BP_DIR / 'screens/notes.txt').write_text('x', encoding='utf-8')
    rc, report = check(tmp_path)
    assert rc == 2 and 'IA_BLUEPRINT_SCREEN_FILE' in codes(report)


# ---- the approval request for the product items --------------------------------------------------------------------

def embedded(page):
    import re
    return json.loads(re.search(r'<script id="data" type="application/json">(.*?)</script>', page, re.S).group(1).replace('<\\/', '</'))


def test_one_card_per_item_and_its_hash_is_the_hash_the_ledger_binds():
    import blueprint
    import blueprint_request
    bp = blueprint.Blueprint(Path('.'), present=True, product=PRODUCT)
    items = blueprint_request.build(bp)
    assert {i['key']: i['value_hash'] for i in items} == item_hashes(PRODUCT) and len(items) == 10
    page = blueprint_request.page(items)
    cards = embedded(page)['cards']
    assert {c['key']: c['value_hash'] for c in cards.values()} == item_hashes(PRODUCT)
    for i in items:
        assert page.count(f'data-key="{i["key"]}"') == 1 and page.count(f'>{i["hash12"]}<') == 1
    assert '○ 承認する' in page and '× 差し戻す' in page and 'decisions' in page


def test_changing_one_item_changes_the_page_and_only_that_cards_hash():
    import blueprint
    import blueprint_request
    bp = blueprint.Blueprint(Path('.'), present=True, product=PRODUCT)
    changed = copy.deepcopy(PRODUCT)
    changed['principles']['pri-0002']['text'] = '今の状態を、いつも画面に示す'
    bp2 = blueprint.Blueprint(Path('.'), present=True, product=changed)
    before, after = blueprint_request.build(bp), blueprint_request.build(bp2)
    old, new = ({i['id']: i['value_hash'] for i in items} for items in (before, after))
    assert {k for k in old if old[k] != new[k]} == {'pri-0002'}
    assert blueprint_request.page(before) != blueprint_request.page(after) and 'いつも画面に示す' in blueprint_request.page(after)
    rows = {c['key']: c['value_hash'] for c in embedded(blueprint_request.page(before))['cards'].values()}
    assert blueprint.approved_items(bp2, rows) == set(new) - {'pri-0002'}  # the stale page approves everything but the changed item


def test_every_element_the_pages_script_treats_as_a_card_has_the_controls_it_needs(tmp_path):
    """The page's script selects cards with one selector and then reads a textarea, a state line and two buttons from each.
    The overview section is also a `.card` box: a selector that matches it throws on the first card, so no button gets a
    click handler and nothing is saved (the owner could not click). Every element the selector matches must be an
    approval card with its controls, and the number of them must equal the embedded cards."""
    import re
    from html.parser import HTMLParser

    write_blueprint(tmp_path)
    out = tmp_path / 'o'
    rc, res = cli('request', '--root', str(tmp_path), '--out', str(out))
    assert rc == 0
    page = (out / 'approval-request.html').read_text(encoding='utf-8')
    selectors = set(re.findall(r"document\.querySelectorAll\('(\.card[^']*)'\)", page))
    assert selectors, 'the script no longer selects cards by .card: update this test with the new selector'

    class Boxes(HTMLParser):
        """Every element with class `card`, with the controls found inside it (nesting counted by depth)."""

        def __init__(self):
            super().__init__()
            self.cards, self.open = [], []

        def handle_starttag(self, tag, attrs):
            a = dict(attrs)
            classes = (a.get('class') or '').split()
            for box in self.open:
                box['buttons'] += tag == 'button'
                box['textarea'] += tag == 'textarea'
                box['state'] += 'state' in classes
                box['depth'] += 1
            if 'card' in classes:
                box = {'tag': tag, 'attrs': a, 'buttons': 0, 'textarea': 0, 'state': 0, 'depth': 1}
                self.cards.append(box)
                self.open.append(box)

        def handle_endtag(self, tag):
            for box in list(self.open):
                box['depth'] -= 1
                if box['depth'] == 0:
                    self.open.remove(box)

    parser = Boxes()
    parser.feed(page)
    embedded_ids = set(embedded(page)['cards'])
    for selector in selectors:
        needs_id = '[data-id]' in selector
        matched = [c for c in parser.cards if (not needs_id) or 'data-id' in c['attrs']]
        for c in matched:
            assert c['attrs'].get('data-id') in embedded_ids and (c['buttons'], c['textarea'], c['state']) == (2, 1, 1), \
                f"{selector} also matches a {c['tag']}.card that is not an approval card: the page's script throws on it"
        assert len(matched) == len(embedded_ids)


def test_the_command_writes_the_page_and_reads_the_answers_back(tmp_path):
    write_blueprint(tmp_path)
    out = tmp_path / 'out'
    rc, res = cli('request', '--root', str(tmp_path), '--out', str(out))
    assert rc == 0 and res['cards'] == 10 and res['approved'] == 10
    rows = json.loads((out / 'approval-request.json').read_text(encoding='utf-8'))['rows']
    assert {r['key']: r['value_hash'] for r in rows} == item_hashes(PRODUCT)
    rc, res = cli('request', '--root', str(tmp_path), '--out', str(out))
    assert rc == 2 and '上書きしない' in res['error']
    saved = [{'id': 'sto-0001', 'data': {'key': 'ia:item:sto-0001', 'value_hash': item_hashes(PRODUCT)['ia:item:sto-0001'], 'decision': 'ok'}},
             {'id': 'sto-0002', 'data': {'key': 'ia:item:sto-0002', 'value_hash': 'a' * 64, 'decision': 'ng'}}]
    (tmp_path / 'saved.json').write_text(json.dumps(saved), encoding='utf-8')
    rc, res = cli('request', '--root', str(tmp_path), '--read-decisions', str(tmp_path / 'saved.json'))
    assert rc == 1 and res['ok'] == ['ia:item:sto-0001'] and res['stale'] == ['ia:item:sto-0002'] and len(res['undecided']) == 8


def test_a_failing_blueprint_makes_no_request(tmp_path):
    write_case(tmp_path, next(c for c in CASES if c['id'] == 'two-primary-actions'))
    rc, res = cli('request', '--root', str(tmp_path), '--out', str(tmp_path / 'o'))
    assert rc == 2 and not (tmp_path / 'o').exists()


def test_the_overview_draws_every_item_the_screens_and_the_values_they_stand_on(tmp_path):
    import blueprint
    import blueprint_request
    import positioning
    from ia_fixtures import SSOT
    import positioning_model as pm
    bp = blueprint.Blueprint(Path('.'), present=True, product=PRODUCT, screens={s['id']: (Path(f'{s["id"]}.yaml'), s) for s in SCREENS})
    values = {k: a['text'] for k, a in positioning.atoms(pm.Model(Path('.'), SSOT), proposed_only=False).items() if k.startswith('val-')}
    g = blueprint_request.overview_graph(bp, values)
    ids = {n['id'] for n in g['nodes']}
    assert set(blueprint.items(bp)) <= ids and {'scr-overview', 'scr-person', 'val-0001', 'val-0003'} <= ids
    assert len(ids) == len(g['nodes'])  # no box twice
    pairs = {(e['from'], e['to']) for e in g['edges']}
    assert ('pri-0001', 'val-0001') in pairs and ('pri-0001', 'sto-0001') in pairs and ('ntr-0001', 'scr-overview') in pairs
    assert ('scr-person', 'sto-0002') in pairs and ('ojc-0001', 'val-0001') in pairs
    html = blueprint_request.overview(bp, values)
    assert html.count('<g class="node') == len(ids) and 'プロダクト全体の関係図' in html


def test_the_page_opens_with_the_overview_then_one_card_per_item(tmp_path):
    write_blueprint(tmp_path)
    out = tmp_path / 'o'
    rc, res = cli('request', '--root', str(tmp_path), '--out', str(out))
    assert rc == 0
    page = (out / 'approval-request.html').read_text(encoding='utf-8')
    assert page.index('プロダクト全体の関係図') < page.index('data-key="ia:item:sto-0001"')
    assert page.count('data-key="ia:item:') == 10


# ---- the screen skeleton (what the build keeps and the render check reads) ----------------------------------------

def test_the_screen_skeleton_marks_regions_groups_actions_and_states(tmp_path):
    write_blueprint(tmp_path)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(tmp_path), '--screen', 'scr-overview'],
                          capture_output=True, text=True)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    html = proc.stdout
    import re
    assert re.findall(r'data-ia-region="(rgn-\d+)"', html) == ['rgn-0001', 'rgn-0002', 'rgn-0003']  # priority order
    assert html.count('data-ia-group=') == 2 and html.count('data-ia-primary="true"') == 1
    assert all(f'data-ia-state="{s}"' in html for s in OVERVIEW['states']) and 'hidden' in html
    assert 'data-ia-action="act-0001"' in html and 'rgn-0003' in html


def test_the_screen_skeleton_is_refused_for_a_failing_blueprint_and_never_overwrites(tmp_path):
    write_case(tmp_path, next(c for c in CASES if c['id'] == 'two-primary-actions'))
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(tmp_path), '--screen', 'scr-overview'], capture_output=True, text=True)
    assert proc.returncode == 2
    ok = tmp_path / 'ok'
    write_blueprint(ok)
    out = ok / 'x.html'
    args = [sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'skeleton', '--root', str(ok), '--screen', 'scr-overview', '--out', str(out)]
    assert subprocess.run(args, capture_output=True, text=True).returncode == 0
    assert subprocess.run(args, capture_output=True, text=True).returncode == 2
