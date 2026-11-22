"""Layout overrides are approved one row at a time (SPEC-deliverable-production-layers requirement 4, ADR-20261121).

A row allows the screen, the pair of regions and the widths it names. The width cannot be left out, and a reason alone passes nothing. The
owner's approval sits in the ledger as `ds:web:override:<id>`, bound by a hash of the row and of the two regions it points at: a change to
the row or to those regions un-approves that row alone; a copy inherits nothing; a new row moves no other row's approval.

Every red case below is a change to a conforming repository that the check must stop; the approved repository must pass. The CLI cases
drive the real commands in a tmp repository (request page, the owner's words and date into the ledger, check).
"""
import copy
import json
import subprocess
import sys

import pytest
import yaml

from test_core import SKILL, approve, check, project, write  # noqa: F401  (project is a fixture)
import layer_paths as lp  # noqa: E402  (test_core puts scripts/ on the path)
import layout_overrides as lo  # noqa: E402

sys.path.insert(0, str(SKILL.parent / 'own-infoarch-author/scripts/tests'))
from test_ia_bridge import attach_blueprint  # noqa: E402  (a product UI: positioning, a blueprint of two screens, the IA module on)
from blueprint_fixtures import OVERVIEW as SCREEN, PERSON as OTHER  # noqa: E402  (scr-overview: rgn-0001..0003; scr-person: rgn-0001..0002)

WORDS = '配置の例外 ovr-0001 を承認する'
DAY = '2026-11-22'

ROW_A = {'id': 'ovr-0001', 'screen': 'scr-overview', 'above': 'rgn-0002', 'below': 'rgn-0001', 'viewports': ['desktop'], 'reason': '部署の比較を先に見たい'}
ROW_B = {'id': 'ovr-0002', 'screen': 'scr-overview', 'above': 'rgn-0003', 'below': 'rgn-0002', 'viewports': ['desktop', 'mobile'], 'reason': '担当者の一覧を先に見たい'}


def dump(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding='utf-8')


def screen_file(root, screen):
    return lp.path(root, 'infoarch', 'blueprint', 'screens', f'{screen["id"]}.yaml')


def setup(project, *rows):
    """A product UI (the real blueprint fixture, the IA module on) with the given override rows written and nothing of them approved."""
    root, catalogs = project
    attach_blueprint(project)
    dump(lp.path(root, 'layout_overrides'), [copy.deepcopy(r) for r in rows])
    return root, catalogs


def cli(root, catalogs, *args):
    return subprocess.run([sys.executable, str(SKILL / 'scripts/design_author.py'), *args[:1], '--repo', str(root),
                           *(['--catalog-dir', str(catalogs)] if args[0] in ('check', 'approve') else []), *args[1:]], capture_output=True, text=True)


def record(project, tmp_path, *ids, words=WORDS, day=DAY):
    """The owner approves rows: prepare the request for exactly those rows, then record it with their words and date (the real CLI)."""
    root, catalogs = project
    prepared = cli(root, catalogs, 'approve', '--prepare', '--date', day, *[a for i in ids for a in ('--target', lo.key_for(i))])
    assert prepared.returncode == 0, prepared.stdout + prepared.stderr
    request = tmp_path / f'request-{"-".join(ids)}.json'
    request.write_text(prepared.stdout, encoding='utf-8')
    done = cli(root, catalogs, 'approve', '--request', str(request), '--approver', 'owner', '--owner-words', words, '--approved-on', day)
    assert done.returncode in (0, 1), done.stdout + done.stderr
    return json.loads(done.stdout)


def state(root):
    return {r.id: r.status for r in lo.evaluate_repo(root).rows}


def codes(root):
    return sorted({c for c, _ in lo.evaluate_repo(root).findings})


def edit_rows(root, fn):
    rows = yaml.safe_load(lp.path(root, 'layout_overrides').read_text(encoding='utf-8'))
    fn(rows)  # edits the list in place (what it returns is not the file)
    dump(lp.path(root, 'layout_overrides'), rows)


def edit_screen(root, screen, fn):
    path = screen_file(root, screen)
    data = yaml.safe_load(path.read_text(encoding='utf-8'))
    fn(data)
    dump(path, data)


@pytest.fixture
def approved(project, tmp_path):
    """Both rows written and the owner's approval of both recorded; tokens approved too, so the check is otherwise green."""
    approve(project)
    root, _ = setup(project, ROW_A, ROW_B)
    record(project, tmp_path, 'ovr-0001', 'ovr-0002')
    assert state(root) == {'ovr-0001': 'approved', 'ovr-0002': 'approved'}
    return project


# ---- the red cases: a row that must be disqualified ----------------------------------------------------------------

def test_no_overrides_file_changes_nothing(project):
    approve(project)
    result = check(project)
    assert result['exit_code'] == 0 and 'overrides' not in result['counts']
    assert not lp.path(project[0], 'layout_overrides').exists()


def test_a_row_with_no_approval_is_unapproved_and_the_ship_gate_stops_it(project):
    approve(project)
    root, _ = setup(project, ROW_A)
    assert state(root) == {'ovr-0001': 'unapproved'}
    plain, ship = check(project), check(project, ship=True)
    assert plain['exit_code'] == 1 and any('layout overrides awaiting owner approval: 1' in w for w in plain['warnings'])
    assert ship['exit_code'] == 2 and any('ship gate: unapproved layout overrides 1' in e for e in ship['errors'])


def test_a_reason_alone_passes_nothing(project):
    """A row that is only a reason (the old way to pass anything): disqualified for everything it lacks."""
    approve(project)
    root, _ = setup(project, {'reason': '部署の比較を先に見たい'})
    assert {'OVERRIDE_ID', 'OVERRIDE_NO_TARGET', 'OVERRIDE_NO_WIDTH'} <= set(codes(root))
    assert check(project)['exit_code'] == 2


@pytest.mark.parametrize('viewports', [None, [], ['phone'], ['desktop', 'desktop'], 'desktop'])
def test_a_row_with_no_usable_width_is_disqualified(project, viewports):
    """The width cannot be left out: leaving it out meant every width. An empty list, an unknown name and a repeated one are no width either."""
    approve(project)
    row = {k: v for k, v in ROW_A.items() if k != 'viewports'}
    if viewports is not None:
        row['viewports'] = viewports
    root, _ = setup(project, row)
    assert {'OVERRIDE_NO_WIDTH', 'OVERRIDE_WIDTH_UNKNOWN'} & set(codes(root))
    assert state(root) == {'ovr-0001': 'invalid'}
    assert check(project)['exit_code'] == 2


@pytest.mark.parametrize('missing', ['screen', 'above', 'below'])
def test_a_row_with_no_screen_or_region_is_disqualified(project, missing):
    approve(project)
    root, _ = setup(project, {k: v for k, v in ROW_A.items() if k != missing})
    assert 'OVERRIDE_NO_TARGET' in codes(root)
    assert check(project)['exit_code'] == 2


@pytest.mark.parametrize('reason', [None, '', '   '])
def test_a_row_with_no_reason_is_disqualified(project, reason):
    approve(project)
    row = {k: v for k, v in ROW_A.items() if k != 'reason'}
    if reason is not None:
        row['reason'] = reason
    root, _ = setup(project, row)
    assert 'OVERRIDE_NO_REASON' in codes(root)
    assert check(project)['exit_code'] == 2


def test_a_row_with_an_unknown_field_is_disqualified_not_ignored(project):
    approve(project)
    root, _ = setup(project, {**ROW_A, 'approved': True})
    assert 'OVERRIDE_UNKNOWN_FIELD' in codes(root) and state(root) == {'ovr-0001': 'invalid'}


def test_a_row_naming_a_screen_or_region_that_is_not_there_is_disqualified(project):
    approve(project)
    for row, code in (({**ROW_A, 'screen': 'scr-nowhere'}, 'OVERRIDE_SCREEN_MISSING'), ({**ROW_A, 'above': 'rgn-0099'}, 'OVERRIDE_REGION_MISSING'),
                      ({**ROW_A, 'below': 'rgn-0099'}, 'OVERRIDE_REGION_MISSING')):
        root, _ = setup(project, row)
        assert codes(root) == [code], row
        assert check(project)['exit_code'] == 2


def test_a_row_pointing_at_the_same_region_twice_is_disqualified(project):
    approve(project)
    root, _ = setup(project, {**ROW_A, 'below': 'rgn-0002'})
    assert 'OVERRIDE_SAME_REGION' in codes(root)


def test_a_copied_row_with_the_same_id_is_disqualified(project):
    approve(project)
    root, _ = setup(project, ROW_A, {**ROW_A, 'screen': 'scr-person'})
    assert state(root) == {'ovr-0001': 'invalid'} and codes(root) == ['OVERRIDE_DUPLICATE']


def test_a_copied_row_with_a_new_id_is_disqualified_and_has_no_approval(approved, tmp_path):
    """Copying an approved row changes only its id: the copy repeats the scope (disqualified) and its new id holds no approval."""
    root, _ = approved
    edit_rows(root, lambda rows: rows.append({**ROW_A, 'id': 'ovr-0009'}))
    assert 'OVERRIDE_DUPLICATE_SCOPE' in codes(root)
    assert state(root)['ovr-0009'] == 'invalid'
    assert lo.key_for('ovr-0009') not in lo.read_approvals(root)
    assert check(approved)['exit_code'] == 2


def test_a_copy_in_a_row_that_differs_only_by_its_id_inherits_no_approval(approved):
    """Without the duplicate rule the copy would still be unapproved: the hash binds the id."""
    root, _ = approved
    original = next(r for r in lo.evaluate_repo(root).rows if r.id == 'ovr-0001')
    copy_row = {**ROW_A, 'id': 'ovr-0009', 'viewports': ['mobile']}  # a new scope, so only the id (and width) differ
    edit_rows(root, lambda rows: rows.append(copy_row))
    new = next(r for r in lo.evaluate_repo(root).rows if r.id == 'ovr-0009')
    assert new.status == 'unapproved' and new.hash != original.hash


def test_an_id_that_is_not_an_id_is_disqualified(project):
    approve(project)
    root, _ = setup(project, {**ROW_A, 'id': 'first'})
    assert 'OVERRIDE_ID' in codes(root)


def test_a_file_that_is_not_a_list_of_rows_is_an_error_not_an_empty_pass(project):
    approve(project)
    root, _ = setup(project)
    lp.path(root, 'layout_overrides').write_text('reason: 全部許す\n', encoding='utf-8')
    result = check(project)
    assert result['exit_code'] == 2 and any('layout-overrides' in e for e in result['errors'])


# ---- the green case: an approved row passes -------------------------------------------------------------------------

def test_approved_rows_pass(approved):
    root, _ = approved
    result = check(approved)
    assert result['exit_code'] == 0 and result['counts']['overrides'] == {'rows': 2, 'approved': 2}
    assert not [w for w in result['warnings'] if 'override' in w]
    assert codes(root) == []


def test_the_ledger_holds_the_owner_row_for_each_approved_row(approved):
    root, _ = approved
    rows = {r['key']: r for r in json.loads(lp.path(root, 'ledger').read_text(encoding='utf-8'))}
    for rid in ('ovr-0001', 'ovr-0002'):
        row = rows[lo.key_for(rid)]
        assert row['approver'] == 'owner' and row['date'] == DAY and len(row['value_hash']) == 64


# ---- a change un-approves exactly that row ---------------------------------------------------------------------------

CHANGES = {
    'reason': lambda r: r.update(reason='別の理由にした'),
    'width added': lambda r: r.update(viewports=['desktop', 'mobile']),
    'width changed': lambda r: r.update(viewports=['mobile']),
    'above': lambda r: r.update(above='rgn-0003'),
    'below': lambda r: r.update(below='rgn-0003', above='rgn-0002'),
    'screen': lambda r: r.update(screen='scr-person', above='rgn-0002', below='rgn-0001'),
}


@pytest.mark.parametrize('what', sorted(CHANGES))
def test_changing_one_thing_in_an_approved_row_un_approves_that_row_only(approved, what):
    root, _ = approved
    edit_rows(root, lambda rows: CHANGES[what](rows[0]))
    after = state(root)
    assert after['ovr-0001'] == 'unapproved', what
    assert after['ovr-0002'] == 'approved', 'the other row lost its approval'
    result = check(approved, ship=True)
    assert any('ship gate: unapproved layout overrides 1' in e for e in result['errors'])


def test_the_width_order_is_not_content(approved):
    root, _ = approved
    edit_rows(root, lambda rows: rows[1].update(viewports=['mobile', 'desktop']))
    assert state(root)['ovr-0002'] == 'approved'


def test_adding_a_row_leaves_every_other_approval_in_place(approved):
    root, _ = approved
    edit_rows(root, lambda rows: rows.append({**ROW_B, 'id': 'ovr-0003', 'above': 'rgn-0003', 'below': 'rgn-0001'}))
    assert state(root) == {'ovr-0001': 'approved', 'ovr-0002': 'approved', 'ovr-0003': 'unapproved'}


def test_removing_a_row_leaves_the_other_approval_and_the_check_does_not_crash(approved):
    root, _ = approved
    edit_rows(root, lambda rows: rows.pop(0))
    assert state(root) == {'ovr-0002': 'approved'}
    result = check(approved)
    assert result['exit_code'] == 1 and any('no longer in the file: 1' in w for w in result['warnings'])


def test_a_row_whose_approval_is_removed_and_written_back_is_unapproved(approved):
    root, _ = approved
    ledger = json.loads(lp.path(root, 'ledger').read_text(encoding='utf-8'))
    write(lp.path(root, 'ledger'), [r for r in ledger if r['key'] != lo.key_for('ovr-0001')])
    assert state(root) == {'ovr-0001': 'unapproved', 'ovr-0002': 'approved'}


def test_a_ledger_row_with_a_wrong_hash_is_not_an_approval(approved):
    root, _ = approved
    ledger = json.loads(lp.path(root, 'ledger').read_text(encoding='utf-8'))
    for r in ledger:
        if r['key'] == lo.key_for('ovr-0001'):
            r['value_hash'] = '0' * 64
    write(lp.path(root, 'ledger'), ledger)
    assert state(root) == {'ovr-0001': 'unapproved', 'ovr-0002': 'approved'}


# ---- the screen and the regions a row points at ------------------------------------------------------------------------

@pytest.mark.parametrize('what,change', [
    ('priority', lambda s: s['regions']['rgn-0002'].update(priority=5)),
    ('text', lambda s: s['regions']['rgn-0002'].update(text='部署ごとの比較（言い換え）')),
    ('shape', lambda s: s['regions']['rgn-0001'].update(shape='status')),
    ('group', lambda s: s['regions']['rgn-0001'].update(group='grp-0009')),
])
def test_changing_a_region_a_row_points_at_drops_its_approval(approved, what, change):
    root, _ = approved
    edit_screen(root, SCREEN, change)
    after = state(root)
    assert after['ovr-0001'] == 'unapproved', what  # points at rgn-0001 and rgn-0002
    if what in ('priority', 'text'):  # ovr-0002 points at rgn-0002 and rgn-0003
        assert after['ovr-0002'] == 'unapproved', what
    else:
        assert after['ovr-0002'] == 'approved', what


def test_changing_a_region_no_row_points_at_keeps_every_approval(approved):
    root, _ = approved
    edit_screen(root, SCREEN, lambda s: s['regions']['rgn-0003'].update(text='担当者の一覧（言い換え）'))
    assert state(root)['ovr-0001'] == 'approved'  # ovr-0001 does not point at rgn-0003; ovr-0002 does
    assert state(root)['ovr-0002'] == 'unapproved'
    edit_screen(root, SCREEN, lambda s: s['regions'].update({'rgn-0004': {'group': 'grp-0002', 'priority': 4, 'shape': 'text', 'text': '注記'}}))
    assert state(root)['ovr-0001'] == 'approved'


def test_a_region_that_disappears_disqualifies_the_row_and_nothing_carries_over(approved):
    root, _ = approved
    edit_screen(root, SCREEN, lambda s: s['regions'].pop('rgn-0002'))
    assert state(root) == {'ovr-0001': 'invalid', 'ovr-0002': 'invalid'}
    assert codes(root) == ['OVERRIDE_REGION_MISSING']
    assert any('OVERRIDE_REGION_MISSING' in e for e in check(approved)['errors'])
    edit_screen(root, SCREEN, lambda s: s['regions'].update({'rgn-0002': {'group': 'grp-0002', 'priority': 2, 'shape': 'list', 'text': '別の領域'}}))
    assert state(root)['ovr-0001'] == 'unapproved'  # a different region under the same id is not the one that was approved


def test_a_screen_that_disappears_disqualifies_its_rows(approved):
    root, _ = approved
    screen_file(root, SCREEN).unlink()
    assert state(root) == {'ovr-0001': 'invalid', 'ovr-0002': 'invalid'} and codes(root) == ['OVERRIDE_SCREEN_MISSING']
    assert any('OVERRIDE_SCREEN_MISSING' in e for e in check(approved)['errors'])


def test_the_same_row_on_another_screen_is_not_approved(approved):
    root, _ = approved
    edit_rows(root, lambda rows: rows.append({**ROW_A, 'id': 'ovr-0007', 'screen': 'scr-person'}))
    assert state(root)['ovr-0007'] == 'unapproved' and state(root)['ovr-0001'] == 'approved'


# ---- the owner's approval path --------------------------------------------------------------------------------------------

def test_the_must_items_never_carry_an_override_with_them(project):
    """`approve --prepare --must` is a batch; an exception is approved row by row, named, never swept in with the rest."""
    approve(project)
    root, catalogs = setup(project, ROW_A)
    result = cli(root, catalogs, 'approve', '--prepare', '--must')
    text = result.stdout + result.stderr
    assert lo.key_for('ovr-0001') not in text
    named = cli(root, catalogs, 'approve', '--prepare', '--target', lo.key_for('ovr-0001'))
    assert named.returncode == 0 and json.loads(named.stdout)['rows'][0]['key'] == lo.key_for('ovr-0001')


def test_an_invalid_row_cannot_be_prepared_for_approval(project):
    approve(project)
    root, catalogs = setup(project, {k: v for k, v in ROW_A.items() if k != 'viewports'})
    result = cli(root, catalogs, 'approve', '--prepare', '--target', lo.key_for('ovr-0001'))
    assert result.returncode != 0


def test_recording_needs_the_owners_words_and_date(project, tmp_path):
    approve(project)
    root, catalogs = setup(project, ROW_A)
    prepared = cli(root, catalogs, 'approve', '--prepare', '--date', DAY, '--target', lo.key_for('ovr-0001'))
    request = tmp_path / 'request.json'
    request.write_text(prepared.stdout, encoding='utf-8')
    refused = cli(root, catalogs, 'approve', '--request', str(request), '--approver', 'owner')
    assert refused.returncode != 0 and state(root) == {'ovr-0001': 'unapproved'}


def test_the_prepared_request_goes_stale_if_the_row_changes_before_recording(project, tmp_path):
    approve(project)
    root, catalogs = setup(project, ROW_A)
    prepared = cli(root, catalogs, 'approve', '--prepare', '--date', DAY, '--target', lo.key_for('ovr-0001'))
    request = tmp_path / 'request.json'
    request.write_text(prepared.stdout, encoding='utf-8')
    edit_rows(root, lambda rows: rows[0].update(viewports=['desktop', 'mobile']))
    stale = cli(root, catalogs, 'approve', '--request', str(request), '--approver', 'owner', '--owner-words', WORDS, '--approved-on', DAY)
    assert stale.returncode != 0 and state(root) == {'ovr-0001': 'unapproved'}


# ---- the request page, through the real CLI -------------------------------------------------------------------------------

def test_the_request_page_the_owners_words_and_the_check_through_the_cli(project, tmp_path):
    """request page -> the owner's words and date into the ledger -> check green (the whole path in a tmp repository)."""
    approve(project)
    root, catalogs = setup(project, ROW_A, ROW_B)
    out = tmp_path / 'req'
    made = cli(root, catalogs, 'override-request', '--out', str(out))
    assert made.returncode == 0, made.stdout + made.stderr
    page = (out / 'approval-request.html').read_text(encoding='utf-8')
    rows = json.loads((out / 'approval-request.json').read_text(encoding='utf-8'))['rows']
    assert [r['key'] for r in rows] == [lo.key_for('ovr-0001'), lo.key_for('ovr-0002')]
    for r in rows:
        assert r['key'] in page and r['value_hash'][:12] in page  # a card per row, with the hash the page prints
    for text in ('部署の比較を先に見たい', 'desktop', 'rgn-0002', '部署ごとの比較', 'data-d="ok"', 'data-d="ng"'):
        assert text in page
    assert check(project)['exit_code'] == 1 and check(project, ship=True)['exit_code'] == 2  # nothing recorded yet
    assert made.returncode == 0 and json.loads(made.stdout)['approved'] == 0

    done = record(project, tmp_path, 'ovr-0001', 'ovr-0002')
    assert done['status'] == 'recorded' and done['recorded_rows'] == 2 and WORDS in done['commit_trailer'] and DAY in done['commit_trailer']
    ledger = {r['key']: r for r in json.loads(lp.path(root, 'ledger').read_text(encoding='utf-8'))}
    assert {ledger[k]['date'] for k in (lo.key_for('ovr-0001'), lo.key_for('ovr-0002'))} == {DAY}
    after = check(project)
    assert after['exit_code'] == 0 and after['counts']['overrides'] == {'rows': 2, 'approved': 2}
    assert not any('override' in e for e in check(project, ship=True)['errors'])  # the ship gate has nothing left about the rows

    again = cli(root, catalogs, 'override-request', '--out', str(tmp_path / 'req2'))
    assert json.loads(again.stdout)['approved'] == 2  # a card knows it is already approved


def test_the_request_page_is_never_overwritten_and_refuses_a_failing_file(project, tmp_path):
    approve(project)
    root, catalogs = setup(project, ROW_A)
    out = tmp_path / 'req'
    assert cli(root, catalogs, 'override-request', '--out', str(out)).returncode == 0
    twice = cli(root, catalogs, 'override-request', '--out', str(out))
    assert twice.returncode == 2 and '上書きしない' in twice.stdout
    edit_rows(root, lambda rows: rows[0].pop('viewports'))
    bad = cli(root, catalogs, 'override-request', '--out', str(tmp_path / 'req3'))
    assert bad.returncode == 2 and 'OVERRIDE_NO_WIDTH' in bad.stdout and not (tmp_path / 'req3').exists()


def test_the_request_has_nothing_to_ask_without_rows(project, tmp_path):
    approve(project)
    root, catalogs = project
    result = cli(root, catalogs, 'override-request', '--out', str(tmp_path / 'req'))
    assert result.returncode == 2 and '0 件' in result.stdout


def test_the_owners_saved_answers_are_read_back_against_the_current_cards(project, tmp_path):
    approve(project)
    root, catalogs = setup(project, ROW_A, ROW_B)
    out = tmp_path / 'req'
    cli(root, catalogs, 'override-request', '--out', str(out))
    cards = json.loads((out / 'approval-request.json').read_text(encoding='utf-8'))['rows']
    answers = [{'id': 'ovr-0001', 'data': {**cards[0], 'decision': 'ok'}},
               {'id': 'ovr-0002', 'data': {**cards[1], 'value_hash': '1' * 64, 'decision': 'ok'}}]  # answered an older version of the row
    saved = tmp_path / 'answers.json'
    saved.write_text(json.dumps(answers), encoding='utf-8')
    read = cli(root, catalogs, 'override-request', '--read-decisions', str(saved))
    out_json = json.loads(read.stdout)
    assert out_json['ok'] == [lo.key_for('ovr-0001')] and out_json['stale'] == [lo.key_for('ovr-0002')] and read.returncode == 1


# ---- the key and the ledger ---------------------------------------------------------------------------------------------

def test_the_ledger_accepts_the_override_key_and_nothing_looser():
    from approvals import verify_rows
    good = {'key': 'ds:web:override:ovr-0001', 'value_hash': 'a' * 64, 'approver': 'owner', 'date': DAY}
    assert verify_rows([good], None) == {'ds:web:override:ovr-0001': 'a' * 64}
    with pytest.raises(ValueError):
        verify_rows([{**good, 'key': 'ds:web:overrides:ovr-0001'}], None)
    with pytest.raises(ValueError):
        verify_rows([{**good, 'key': 'ds:web:override:'}], None)


def test_a_stale_override_approval_does_not_pass_for_a_row_that_came_back_changed(approved):
    root, _ = approved
    edit_rows(root, lambda rows: rows.pop(0))
    edit_rows(root, lambda rows: rows.insert(0, {**ROW_A, 'viewports': ['desktop', 'mobile']}))  # the same id, a wider row
    assert state(root) == {'ovr-0001': 'unapproved', 'ovr-0002': 'approved'}
