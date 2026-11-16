"""One atomic v2 -> v4 migration (SPEC-information-architecture-standardization v5, requirement 18; Q9, Q11).

The tool moves what is mechanical (problem -> pain, the attribute rename, audience -> segment, the module pin),
reports what needs drafting (jobs, values, serves, outline strategy), the old -> new id map, the missed count
(must be 0), the IA judgments the rename invalidates, and how to undo it. It is rerunnable: a v4 tree is a no-op.
"""
import json
import subprocess
import sys
from pathlib import Path

import yaml

SCRIPTS = Path(__file__).resolve().parents[1]

V2 = {'schema_version': 2,
      'ledger': {'next': {'aud': 2, 'prb': 3, 'msg': 2, 'prf': 1, 'fct': 1, 'ent': 1, 'vce': 1},
                 'tombstones': {'prb-0002': None}},
      'audiences': {'aud-0001': {'state': 'proposed', 'provenance': 'owner', 'label': '人事責任者', 'situation': '見極めが勘頼み',
                                 'desired_action': '申し込む', 'priorities': ['msg-0001']}},
      'problems': {'prb-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '判断指標がない', 'audience': 'aud-0001'}},
      'messages': {'msg-0001': {'state': 'proposed', 'provenance': 'owner', 'text': '見極めを数字で', 'role': 'benefit',
                                'addresses': ['prb-0001']}},
      'proofs': {}, 'facts': {}, 'entities': {}, 'voices': {}}
OUTLINE = {'schema_version': 1, 'id': 'lp-top', 'channel': 'web',
           'brief': {'purpose': '申し込み', 'action': '申し込む', 'medium': 'LP'}, 'audience': 'aud-0001',
           'outputs': ['site/index.html'],
           'pages': [{'id': 'top', 'sections': [{'role': 'problem', 'problem': 'prb-0001', 'copy': '指標がない'},
                                                {'role': 'benefit', 'message': 'msg-0001', 'copy': '数字で見極める'}]}]}
HTML = ('<section data-ia-role="problem" data-ia-problem="prb-0001" class="x">指標がない</section>\n'
        '<section data-ia-role="benefit" data-ia-msg="msg-0001">数字で見極める</section>\n')


def repo(tmp_path):
    base = tmp_path / 'design/info_arch'
    (base / 'outlines').mkdir(parents=True)
    (base / 'channels').mkdir()
    (base / 'info_arch.yaml').write_text(yaml.safe_dump(V2, allow_unicode=True, sort_keys=False), encoding='utf-8')
    (base / 'outlines/lp-top.yaml').write_text(yaml.safe_dump(OUTLINE, allow_unicode=True, sort_keys=False), encoding='utf-8')
    (base / 'channels/web.yaml').write_text(yaml.safe_dump({'schema_version': 2, 'media': 'web', 'outputs': ['site/*.html'],
                                                            'ship': False}), encoding='utf-8')
    (tmp_path / 'site').mkdir()
    (tmp_path / 'site/index.html').write_text(HTML, encoding='utf-8')
    (tmp_path / 'design/design.yaml').write_text(json.dumps({'schema_version': 1, 'modules': {'ia': '0.2.0'}}), encoding='utf-8')
    (tmp_path / 'design/judgments.json').write_text(json.dumps({'schema_version': 3, 'rows': [
        {'candidate': {'input': {'kind': 'ia-unregistered-claim'}}}, {'candidate': {'input': {'kind': 'doc-relevance'}}}]}),
        encoding='utf-8')
    return tmp_path


def migrate(root, *extra):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'migrate', '--root', str(root), *extra],
                          capture_output=True, text=True, check=False)
    return proc.returncode, json.loads(proc.stdout)


def test_preview_writes_nothing(tmp_path):
    root = repo(tmp_path)
    before = (root / 'design/info_arch/info_arch.yaml').read_bytes()
    rc, out = migrate(root)
    assert rc == 0 and out['status'] == 'preview', out
    assert (root / 'design/info_arch/info_arch.yaml').read_bytes() == before


def test_write_moves_everything_mechanical_and_reports_the_rest(tmp_path):
    root = repo(tmp_path)
    rc, out = migrate(root, '--write')
    assert rc == 0 and out['status'] == 'migrated', out
    assert out['id_map'] == {'prb-0001': 'pan-0001', 'prb-0002': 'pan-0002', 'aud-0001': 'seg-0001'}
    assert out['missed'] == []
    ssot = yaml.safe_load((root / 'design/info_arch/info_arch.yaml').read_text(encoding='utf-8'))
    assert ssot['schema_version'] == 4 and 'problems' not in ssot and ssot['edges'] == []
    assert ssot['pains']['pan-0001']['text'] == '判断指標がない' and ssot['pains']['pan-0001']['segment'] == 'seg-0001'
    assert ssot['audiences']['aud-0001']['segment'] == 'seg-0001'
    assert ssot['segments']['seg-0001']['qualifier'] == '人事責任者'
    assert ssot['messages']['msg-0001']['addresses'] == ['pan-0001']
    assert ssot['ledger']['tombstones'] == {'pan-0002': None} and ssot['ledger']['next']['pan'] == 3
    outline = yaml.safe_load((root / 'design/info_arch/outlines/lp-top.yaml').read_text(encoding='utf-8'))
    assert outline['schema_version'] == 2 and outline['pages'][0]['sections'][0] == {'role': 'pain', 'pain': 'pan-0001', 'copy': '指標がない'}
    html = (root / 'site/index.html').read_text(encoding='utf-8')
    assert 'data-ia-pain="pan-0001"' in html and 'data-ia-role="pain"' in html and 'problem' not in html
    assert json.loads((root / 'design/design.yaml').read_text())['modules']['ia'] == '0.4.0'
    assert out['invalidated_judgments'] == {'ia_rows': 1}
    drafting = set(out['drafting_required'])
    assert {'jobs:job-0001', 'messages:msg-0001.serves', 'outlines:lp-top.strategy'} <= drafting
    assert 'git' in out['recovery']


def test_rerun_is_a_no_op(tmp_path):
    root = repo(tmp_path)
    migrate(root, '--write')
    snapshot = {p: p.read_bytes() for p in root.rglob('*') if p.is_file()}
    rc, out = migrate(root, '--write')
    assert rc == 0 and out['status'] == 'already-v4' and out['changes'] == 0, out
    assert {p: p.read_bytes() for p in root.rglob('*') if p.is_file()} == snapshot


def test_a_leftover_old_reference_is_reported_as_missed(tmp_path):
    root = repo(tmp_path)
    (root / 'design/info_arch/INFO_ARCH.md').write_text('課題は prb-0001 を見る\n', encoding='utf-8')
    rc, out = migrate(root, '--write')
    assert rc == 2 and out['missed'] == ['design/info_arch/INFO_ARCH.md'], out


def test_an_old_id_in_an_html_comment_is_reported_not_missed(tmp_path):
    root = repo(tmp_path)
    (root / 'site/index.html').write_text('<!-- 課題（prb-0001） -->\n' + HTML, encoding='utf-8')
    rc, out = migrate(root, '--write')
    assert rc == 0 and out['missed'] == [] and out['comment_mentions'] == ['site/index.html'], out
    assert '<!-- 課題（prb-0001） -->' in (root / 'site/index.html').read_text(encoding='utf-8')  # D5: attributes only


def test_a_v5_tree_is_refused_not_rewritten(tmp_path):
    root = repo(tmp_path)
    path = root / 'design/info_arch/info_arch.yaml'
    data = yaml.safe_load(path.read_text(encoding='utf-8'))
    data['schema_version'] = 3
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding='utf-8')
    before = path.read_bytes()
    rc, out = migrate(root, '--write')
    assert rc == 2 and 'ISSUE-17' in out['error'] and path.read_bytes() == before, out
