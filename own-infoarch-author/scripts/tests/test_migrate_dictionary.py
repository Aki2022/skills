"""migrate-dictionary: schema 4 (facts, entities, voice in the SSOT) -> schema 5 and the docs dictionary."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

from ia_fixtures import SSOT, confirm
from test_check import CHANNEL, DECK, DRAFT, OUTLINE, SCRIPTS, check, codes, dump, write_repo

DATE = '2026-10-23'
V4_BLOCKS = {
    'facts': {'fct-0001': {'state': 'proposed', 'provenance': 'owner', 'label': '導入社数', 'value': '100社以上', 'as_of': '2026-10-01'},
              'fct-0002': {'state': 'proposed', 'provenance': 'owner', 'label': '問い合わせ電話', 'value': '03-1234-5678', 'evergreen': True}},
    'entities': {'ent-0001': {'state': 'proposed', 'provenance': 'owner', 'type': 'product', 'name': 'Webサーバー'}},
    'voices': {'vce-0001': {'state': 'proposed', 'provenance': 'owner', 'traits': [
        {'trait': t, 'is': 'できることを具体的に言う', 'is_not': '大げさな約束はしない'} for t in ('誠実', '前向き', 'わかりやすい')]}}}
MAP = {'fct-0001': {'id': 'value-customer-count', 'expires': '2027-10-01'},
       'fct-0002': {'id': 'value-phone', 'evergreen': True},
       'ent-0001': {'id': 'name-web-server', 'variants': ['ウェブサーバー']}}
SSOT_TEXT = """# 伝え方の SSOT（コメントは残る）
schema_version: 4
ledger:
  next: {prf: 4, fct: 3, ent: 2, vce: 2, seg: 3, job: 3, pan: 3, gan: 3, alt: 4, cat: 3, str: 3, val: 4}
  tombstones: {}
"""


def v4_ssot():
    ssot = {**copy.deepcopy(SSOT), 'schema_version': 4, **copy.deepcopy(V4_BLOCKS)}
    ssot['ledger']['next'] = {**SSOT['ledger']['next'], 'fct': 3, 'ent': 2, 'vce': 2}
    return ssot


def outline_v4(o):
    o = copy.deepcopy(o)
    for c in o['copies'].values():
        c['text'] = c['text'].replace('{{value-customer-count}}', '{{fct-0001}}').replace('{{value-phone}}', '{{fct-0002}}')
    return o


HTML = ('<body data-ia-outline="lp-municipal"><section data-ia-section="sec-0002"><p data-ia-copy="cpy-0003">導入は'
        '<span data-ia-fact="fct-0001">100社以上</span></p></section><footer><span data-ia-entity="ent-0001">Webサーバー</span></footer></body>')


@pytest.fixture
def repo(tmp_path):
    ssot = v4_ssot()
    write_repo(tmp_path, ssot, outlines=(outline_v4(OUTLINE), outline_v4(DECK)), dictionary=None)
    (tmp_path / 'site').mkdir()
    (tmp_path / 'site/index.html').write_text(HTML, encoding='utf-8')
    return tmp_path


def mapfile(tmp_path, mapping=MAP):
    f = tmp_path.parent / f'{tmp_path.name}-map.json'
    f.write_text(json.dumps(mapping, ensure_ascii=False), encoding='utf-8')
    return f


def migrate(root, *extra, mapping=MAP):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'migrate-dictionary', '--root', str(root),
                           '--map', str(mapfile(root, mapping)), '--date', DATE, *extra], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


def snapshot(root):
    return {str(p.relative_to(root)): p.read_bytes() for p in sorted(root.rglob('*')) if p.is_file()}


def test_the_schema_4_tree_checks_with_one_warning_before(repo):
    rc, report = check(repo)
    assert rc == 1 and not codes(report) and codes(report, 'warn') == ['IA_SSOT_LEGACY_V4']


def test_dry_run_reports_the_plan_and_writes_nothing(repo):
    before = snapshot(repo)
    rc, out = migrate(repo)
    assert rc == 0 and out['written'] is False and not out['unmigrated']
    assert out['id_map'] == {'fct-0001': 'value-customer-count', 'fct-0002': 'value-phone', 'ent-0001': 'name-web-server'}
    assert out['check'] == {'before': {'exit_code': 1, 'fail_codes': []}, 'after': {'exit_code': 0, 'fail_codes': []}}
    assert snapshot(repo) == before


def test_write_migrates_the_tree_and_the_check_result_is_the_same_without_the_warning(repo):
    rc, out = migrate(repo, '--write')
    assert rc == 0 and out['written'] is True
    ssot = yaml.safe_load((repo / 'design/info_arch/info_arch.yaml').read_text(encoding='utf-8'))
    assert ssot['schema_version'] == 5 and not {'facts', 'entities', 'voices'} & set(ssot)
    assert set(ssot['ledger']['next']) == set(SSOT['ledger']['next'])
    dictionary = json.loads((repo / 'docs/dictionary/product.json').read_text(encoding='utf-8'))
    assert {e['id']: e['kind'] for e in dictionary['entries']} == {'value-customer-count': 'value', 'value-phone': 'value', 'name-web-server': 'name'}
    by = {e['id']: e for e in dictionary['entries']}
    assert by['value-customer-count']['expires'] == '2027-10-01' and by['value-customer-count']['as_of'] == '2026-10-01'
    assert by['value-phone']['evergreen'] is True and 'expires' not in by['value-phone']
    assert by['name-web-server']['variants'] == ['ウェブサーバー'] and by['name-web-server']['text'] == 'Webサーバー'
    lp = yaml.safe_load((repo / 'design/info_arch/outlines/lp-municipal.yaml').read_text(encoding='utf-8'))
    assert '{{value-customer-count}}' in lp['copies']['cpy-0003']['text'] and 'fct-' not in json.dumps(lp)
    html = (repo / 'site/index.html').read_text(encoding='utf-8')
    assert 'data-dict="value-customer-count"' in html and 'data-dict="name-web-server"' in html and 'data-ia-fact' not in html
    rc, report = check(repo)
    assert rc == 0 and not codes(report) and not codes(report, 'warn')  # the same failures (none), and the warning is gone


def test_the_ssot_comments_survive_the_rewrite(tmp_path):
    ssot = v4_ssot()
    write_repo(tmp_path, ssot, outlines=(outline_v4(OUTLINE), outline_v4(DECK)), dictionary=None)
    p = tmp_path / 'design/info_arch/info_arch.yaml'
    p.write_text('# 手で書いたコメント\n' + p.read_text(encoding='utf-8'), encoding='utf-8')
    rc, _ = migrate(tmp_path, '--write')
    assert rc == 0 and p.read_text(encoding='utf-8').startswith('# 手で書いたコメント\n')


def test_a_second_run_changes_nothing(repo):
    migrate(repo, '--write')
    before = snapshot(repo)
    rc, out = migrate(repo, '--write')
    assert rc == 0 and out['status'] == 'already-migrated' and snapshot(repo) == before


# ---- what cannot be moved mechanically: exit 2 and nothing written ---------------------------------------------

BAD = {
    'not-mapped': lambda m: m.pop('fct-0002'),
    'lifetime': lambda m: m['fct-0001'].pop('expires'),
    'dictionary-id': lambda m: m['fct-0001'].update(id='fct-0001'),
    'map-key': lambda m: m['fct-0001'].update(color='red'),
}


@pytest.mark.parametrize('reason', sorted(BAD))
def test_an_unmappable_item_is_a_failure_and_writes_nothing(repo, reason):
    mapping = copy.deepcopy(MAP)
    BAD[reason](mapping)
    before = snapshot(repo)
    rc, out = migrate(repo, '--write', mapping=mapping)
    assert rc == 2 and reason in [u['reason'] for u in out['unmigrated']]
    assert snapshot(repo) == before


def test_a_value_with_both_a_date_and_evergreen_is_refused(repo):
    mapping = copy.deepcopy(MAP)
    mapping['fct-0001']['evergreen'] = True
    rc, out = migrate(repo, mapping=mapping)
    assert rc == 2 and 'lifetime' in [u['reason'] for u in out['unmigrated']]


def test_an_existing_dictionary_file_is_never_overwritten(repo):
    (repo / 'docs/dictionary').mkdir(parents=True)
    (repo / 'docs/dictionary/product.json').write_text('{}', encoding='utf-8')
    before = snapshot(repo)
    rc, out = migrate(repo, '--write')
    assert rc == 2 and 'dictionary-exists' in [u['reason'] for u in out['unmigrated']] and snapshot(repo) == before


def test_a_placeholder_naming_no_fact_is_a_failure(repo):
    o = outline_v4(OUTLINE)
    o['copies']['cpy-0003']['text'] = '導入は{{fct-0099}}'
    dump(repo / 'design/info_arch/outlines/lp-municipal.yaml', o)
    rc, out = migrate(repo, '--write')
    assert rc == 2 and 'placeholder' in [u['reason'] for u in out['unmigrated']]
    assert yaml.safe_load((repo / 'design/info_arch/info_arch.yaml').read_text(encoding='utf-8'))['schema_version'] == 4


def test_a_rendered_mark_naming_no_entity_is_a_failure(repo):
    (repo / 'site/index.html').write_text(HTML.replace('ent-0001', 'ent-0042'), encoding='utf-8')
    rc, out = migrate(repo, '--write')
    assert rc == 2 and 'html-mark' in [u['reason'] for u in out['unmigrated']]


def test_a_migration_that_would_make_the_check_fail_is_refused(repo):
    """A value that is already expired for a shipped outline: the dictionary check fails it, the old tree did not."""
    mapping = copy.deepcopy(MAP)
    mapping['fct-0001']['expires'] = '2026-10-01'
    rc, out = migrate(repo, '--write', mapping=mapping)
    assert rc == 2 and 'check-regressed' in [u['reason'] for u in out['unmigrated']]
    assert (repo / 'design/info_arch/info_arch.yaml').read_text(encoding='utf-8').count('schema_version: 4') == 1


def test_an_unreadable_ssot_is_a_failure(tmp_path):
    (tmp_path / 'design/info_arch').mkdir(parents=True)
    rc, out = migrate(tmp_path)
    assert rc == 2 and out['unmigrated'][0]['reason'] == 'unreadable'
