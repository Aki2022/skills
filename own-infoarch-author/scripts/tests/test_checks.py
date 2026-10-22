"""Green fixture repo and every required red case through the real CLI entry point."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
import prh  # noqa: E402

DATE = '2026-10-22'
RED = json.loads((SKILL / 'fixtures/red-cases.json').read_text(encoding='utf-8'))['cases']

INFO = {
    'schema_version': 1,
    'ledger': {'next': {'aud': 3, 'prb': 2, 'msg': 5, 'clm': 3, 'trm': 1, 'ent': 2}, 'tombstones': {'msg-0004': 'msg-0003'}},
    'audiences': {
        'aud-0001': {'state': 'proposed', 'provenance': 'owner', 'situation': '自治体の情報部門', 'knows': '既存システムの運用負荷',
                     'decides': '移行するか', 'next_action': '資料請求'},
        'aud-0002': {'state': 'proposed', 'provenance': 'owner', 'situation': '中小企業の経営者', 'knows': '人手不足',
                     'decides': '導入するか', 'next_action': '問い合わせ'},
    },
    'problems': {'prb-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '運用が重い', 'audience': 'aud-0001',
                              'evidence_claims': ['clm-0001'], 'rank': 1}},
    'messages': {
        'msg-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '導入が進んでいる', 'role': 'benefit', 'audience': 'aud-0001',
                     'addresses': ['prb-0001'], 'supports': ['clm-0001'], 'rank': 1,
                     'renditions': {'web': '{{clm-0001}}の企業が導入'}},
        'msg-0002': {'state': 'proposed', 'provenance': 'owner', 'text': '自治体で使われている', 'role': 'proof', 'audience': 'aud-0001',
                     'supports': ['clm-0002'], 'rank': 2},
        'msg-0003': {'state': 'proposed', 'provenance': 'owner', 'text': '資料を請求する', 'role': 'cta', 'audience': 'aud-0001', 'rank': 3},
    },
    'claims': {
        'clm-0001': {'state': 'proposed', 'provenance': 'owner', 'kind': 'number', 'statement': '導入社数は 100 社以上',
                     'value': {'value': 100, 'unit': '社', 'qualifier': ['gte']}, 'evidence': [{'source': '導入台帳'}],
                     'scope': {'audiences': ['aud-0001'], 'as_of': '2026-10-01'}, 'expires': '2027-12-31'},
        'clm-0002': {'state': 'proposed', 'provenance': 'owner', 'kind': 'social_proof', 'statement': '多くの自治体で活用されています',
                     'evidence': [{'source': '導入台帳'}], 'scope': {'audiences': ['aud-0001'], 'as_of': '2026-10-01'},
                     'expires': '2027-12-31'},
    },
    'entities': {'ent-0001': {'state': 'proposed', 'provenance': 'owner', 'type': 'product', 'name': 'Webサーバー',
                              'forbidden': ['ウェブサーバー']}},
}
CHANNEL = {'schema_version': 1, 'media': 'web', 'preset': 'lp', 'outputs': ['site/**/*.html', 'site/*.md'], 'ship': True,
           'ship_command': 'mise run deploy', 'messages': ['msg-0001', 'msg-0002', 'msg-0003'],
           'order': {'mode': 'pattern', 'pattern': ['benefit', 'proof', 'cta']}}
HTML = '''<!doctype html><html><head><title>Webサーバー</title>
<meta name="description" content="Webサーバーの紹介"></head><body>
<h1>Webサーバーで運用を軽くする</h1>
<p><span data-ia-claim="clm-0001">100社以上</span>の企業が導入</p>
<p data-ia-claim="clm-0002">多くの自治体で活用されています</p>
<p>2026年10月22日更新。お問い合わせ 03-1234-5678。STEP 1 から始める3つのポイント。</p>
<img src="a.png" alt="Webサーバーの画面">
<script>var n = 999;</script>
</body></html>
'''
MD = '# 会社概要\n\n[100社以上]{data-ia-claim=clm-0001}の導入。\n\n```\n12345\n```\n'
MANUAL_RULE = '  - expected: お問い合わせ\n    pattern: お問合せ\n'


def write_repo(root, info=INFO, channel=CHANNEL, html=HTML, md=MD, terms=None):
    base = root / 'design/info_arch'
    (base / 'channels').mkdir(parents=True, exist_ok=True)
    (base / 'info_arch.yaml').write_text(yaml.safe_dump(info, allow_unicode=True, sort_keys=False), encoding='utf-8')
    if channel is not None:
        (base / 'channels/web.yaml').write_text(yaml.safe_dump(channel, allow_unicode=True, sort_keys=False), encoding='utf-8')
    if terms is None:
        terms = f'version: 1\nrules:\n{prh.render_generated(info.get("entities", {}))}\n{MANUAL_RULE}'
    (base / 'terms.prh.yml').write_text(terms, encoding='utf-8')
    (root / 'site').mkdir(exist_ok=True)
    if html is not None:
        (root / 'site/index.html').write_text(html, encoding='utf-8')
    if md is not None:
        (root / 'site/about.md').write_text(md, encoding='utf-8')


def run(root, level='l1'):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'check', '--root', str(root),
                           '--level', level, '--date', DATE], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    report = json.loads(proc.stdout)
    assert report['exit_code'] == proc.returncode
    return proc.returncode, report


def codes(report, level='fail'):
    return [f['code'] for f in report['findings'] if f['level'] == level]


def test_green_fixture_passes_with_real_targets(tmp_path):
    write_repo(tmp_path)
    rc, report = run(tmp_path)
    assert rc == 0, report['findings']
    assert report['targets']['files'] == 2 and report['targets']['blocks'] >= 6


def add_html(html, snippet):
    return html.replace('</body>', f'{snippet}\n</body>')


def mutate(case, root):
    info, channel = copy.deepcopy(INFO), copy.deepcopy(CHANNEL)
    kw = {}
    if case == 'unanchored-9wari':
        kw['html'] = add_html(HTML, '<p>お客様の9割が満足</p>')
    elif case == 'unanchored-markdown':
        kw['md'] = MD + '\n導入社数は120社です。\n'
    elif case == 'lexicon-hit':
        kw['html'] = add_html(HTML, '<p>業界初のサービスです</p>')
    elif case == 'six-variant-forms':
        forms = ['webサーバー', 'ＷＥＢサーバー', 'Webｻｰﾊﾞｰ', 'Web サーバー', 'Webサ−バ−', 'ウェブサーバー']
        kw['html'] = add_html(HTML, '\n'.join(f'<p>{f}を使う</p>' for f in forms))
    elif case == 'qualifier-mismatch':
        kw['html'] = HTML.replace('>100社以上<', '>約100社<')
    elif case == 'value-mismatch':
        kw['html'] = HTML.replace('>100社以上<', '>120社以上<')
    elif case == 'unknown-claim-anchor':
        kw['html'] = add_html(HTML, '<p data-ia-claim="clm-0099">50%削減</p>')
    elif case == 'rendition-raw-number':
        info['messages']['msg-0001']['renditions']['web'] = '100社以上の企業が導入'
    elif case == 'prh-generated-hand-edit':
        block = prh.render_generated(info['entities'])
        kw['terms'] = f'version: 1\nrules:\n{block.replace("ウェブサーバー", "ウェッブサーバー")}\n{MANUAL_RULE}'
    elif case == 'prh-spec-fails':
        bad = '  - expected: サーバー\n    pattern: サーバ\n    specs:\n      - from: サーバー\n        to: サーバー\n'
        kw['terms'] = f'version: 1\nrules:\n{prh.render_generated(info["entities"])}\n{MANUAL_RULE}{bad}'
    elif case == 'zero-target-files':
        channel['outputs'] = ['nothing/*.html']
    elif case == 'zero-target-text':
        kw['html'] = '<html><body></body></html>'
        kw['md'] = ''
    elif case == 'no-channels':
        channel = None
    elif case == 'state-approved-written':
        info['claims']['clm-0001']['state'] = 'approved'
    elif case == 'ship-without-ship-command':
        del channel['ship_command']
    elif case == 'id-not-issued':
        info['ledger']['next']['clm'] = 2
    elif case == 'ref-to-tombstone':
        channel['messages'] = ['msg-0001', 'msg-0002', 'msg-0003', 'msg-0004']
    elif case == 'claim-expired':
        info['claims']['clm-0001']['expires'] = '2026-10-21'
    elif case == 'claim-out-of-scope':
        info['messages']['msg-0001']['audience'] = 'aud-0002'
    elif case == 'preset-role-missing':
        channel['messages'] = ['msg-0001', 'msg-0002']
    elif case == 'empty-message-referenced':
        info['messages']['msg-0003'] = {'state': 'empty', 'provenance': 'proposal'}
    else:
        raise KeyError(case)
    write_repo(root, info=info, channel=channel, **kw)


@pytest.mark.parametrize('case', RED, ids=[c['id'] for c in RED])
def test_red_case_fails_with_exit_2(case, tmp_path):
    mutate(case['id'], tmp_path)
    rc, report = run(tmp_path)
    assert rc == 2, report['findings']
    found = codes(report)
    assert case['code'] in found, found
    if 'count' in case:
        assert found.count(case['code']) == case['count'], report['findings']


def test_red_case_table_and_mutations_agree(tmp_path):
    for c in RED:
        mutate(c['id'], tmp_path / c['id'])
    with pytest.raises(KeyError):
        mutate('not-a-case', tmp_path / 'x')


def test_l0_only_does_not_read_outputs(tmp_path):
    write_repo(tmp_path, html=add_html(HTML, '<p>お客様の9割が満足</p>'))
    rc, report = run(tmp_path, 'l0')
    assert rc == 0 and report['targets'] == {}


def test_unsupported_output_format_is_uncovered_not_passed(tmp_path):
    channel = copy.deepcopy(CHANNEL)
    channel['outputs'] = channel['outputs'] + ['site/*.tsx']
    write_repo(tmp_path, channel=channel)
    (tmp_path / 'site/App.tsx').write_text('export const A = () => <p>9割</p>;', encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 1 and report['uncovered'][0]['path'] == 'site/App.tsx'


def test_unreferenced_proposed_claim_is_info_only(tmp_path):
    info = copy.deepcopy(INFO)
    info['ledger']['next']['clm'] = 4
    info['claims']['clm-0003'] = copy.deepcopy(info['claims']['clm-0001'])
    write_repo(tmp_path, info=info)
    rc, report = run(tmp_path)
    assert rc == 0 and 'IA_CLAIMS_UNREFERENCED' in codes(report, 'info')


def test_gen_terms_write_repairs_drift(tmp_path):
    mutate('prh-generated-hand-edit', tmp_path)
    subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'gen-terms', '--root', str(tmp_path), '--write'],
                   check=True, capture_output=True)
    rc, report = run(tmp_path)
    assert rc == 0, report['findings']


def test_template_is_schema_valid_and_not_green_until_filled(tmp_path):
    import shutil
    shutil.copytree(SKILL / 'templates/design', tmp_path / 'design')
    rc, report = run(tmp_path, 'l0')
    assert 'IA_SCHEMA' not in codes(report) and 'IA_TERMS_GENERATED_DRIFT' not in codes(report), report['findings']
    assert rc == 2 and 'IA_PRESET_BLOCK' in codes(report)


def test_wrapper_uses_design_author_python(tmp_path):
    import os
    write_repo(tmp_path)
    env = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable)
    proc = subprocess.run([str(SCRIPTS / 'infoarch-author'), 'check', '--root', str(tmp_path), '--date', DATE],
                          capture_output=True, text=True, env=env, check=False)
    assert proc.returncode == 0, proc.stdout + proc.stderr
