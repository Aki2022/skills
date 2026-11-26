"""A small repository with a design layer whose L0 has run (a real manifest), a script (台本) and a layout plan.

The design layer uses a synthetic item catalog (never upstream classifications), the way own-design-author's own tests do;
the layout pattern vocabulary is the skill's real catalog. Content is neutral sample text."""
import copy
import json
import os
import runpy
import shutil
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
SCRIPTS = SKILLS / 'own-web-build/scripts'
DESIGN_SCRIPTS = SKILLS / 'own-design-author/scripts'
for _path in (DESIGN_SCRIPTS, SCRIPTS):
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))
import layer_paths as lp  # noqa: E402
from classify import classify_response, load_policy  # noqa: E402
from design_author import check_design, init_design  # noqa: E402

# The infoarch layer is own-infoarch-author's own test fixture (positioning, channels, dictionary, two scripts), so L0 resolves
# a real IA module; the script of this kit is added beside them.
IA_FIXTURE = runpy.run_path(str(SKILLS / 'own-infoarch-author/scripts/tests/test_check.py'))

SCRIPT = {
    'schema_version': 1, 'id': 'lp-sample', 'channel': 'web', 'outputs': ['site/sample.html'],
    'target': {'segment': 'seg-0001', 'condition': '初めて検討する人'}, 'touchpoint': 'lp', 'tone': '落ち着いた調子',
    'objective': {'id': 'obj-0001', 'action': '資料を請求する'},
    'messages': {'msg-0001': {'value': 'val-0001', 'rank': 'lead', 'type': 'normal'},
                 'msg-0002': {'value': 'val-0001', 'rank': 'support', 'type': 'cta', 'achieves': 'obj-0001'}},
    # written out of order on purpose: the order is the `order` field
    'blocks': {'blk-0003': {'order': 3, 'function': 'fact'},
               'blk-0001': {'order': 1, 'function': 'identify'},
               'blk-0002': {'order': 2, 'function': 'claim', 'places': ['msg-0001']},
               'blk-0004': {'order': 4, 'function': 'action', 'places': ['msg-0002']},
               'blk-0005': {'order': 5, 'function': 'source'}},
    'copies': {'cpy-0001': {'block': 'blk-0001', 'order': 1, 'text': 'サンプル事業'},
               'cpy-0003': {'block': 'blk-0002', 'order': 2, 'text': '準備の時間を短くします', 'message': 'msg-0001'},
               'cpy-0002': {'block': 'blk-0002', 'order': 1, 'text': '毎日の手間を減らす', 'message': 'msg-0001'},
               'cpy-0004': {'block': 'blk-0003', 'order': 1, 'text': '利用者の 8 割が続けている', 'proposal': True},
               'cpy-0005': {'block': 'blk-0004', 'order': 1, 'text': '資料を請求する', 'message': 'msg-0002'},
               'cpy-0006': {'block': 'blk-0005', 'order': 1, 'text': '運営：サンプル運営室'}},
    'texts': {'nav-top': {'kind': 'navigation', 'text': 'はじめに'}},
}

PLAN = {
    'schema_version': 1, 'deliverable': 'lp-sample', 'script': 'lp-sample',
    'blocks': {'blk-0001': {'pattern': 'header-bar'},
               'blk-0002': {'pattern': 'split-2', 'details': '絵は右の列から上の帯へ少しはみ出す。見出しの下に細い線'},
               'blk-0003': {'pattern': 'stats-row', 'narrow': 'stack-1'},
               'blk-0004': {'pattern': 'band'},
               'blk-0005': {'pattern': 'source-footer'}},
    'comp': {'reviewed': False},
}

# The build brief (ux/build/<deliverable id>/brief.yaml): which of design's tokens the deliverable uses, its style
# references and, per block, a visual requirement. The kit's design has one theme, so it takes every token.
BRIEF = {'schema_version': 1, 'deliverable': 'lp-sample', 'script': 'lp-sample', 'tokens': {}, 'style_references': []}


def write(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n')


def _item(ident='dads-000001'):
    row = {'id': ident, 'match_key': ident, 'document': 'fixture', 'title': ident, 'headings': [ident],
           'text': 'A synthetic fixture policy.', 'source_url': 'https://example.org/',
           'evidence_types': ['token', 'doc', 'rule']}
    policy = load_policy()
    choice = 'global1-lp0-product0-presentation0'
    answer = {'type': 'choice', 'choice': choice, 'confidence': 1.0,
              'probabilities': {k: float(k == choice) for k in policy['outcomes']}}
    row['classification'] = classify_response(row, answer, policy['model'], policy)
    return row


class Repo:
    def __init__(self, tmp_path, git=True):
        self.root = tmp_path / 'repo'
        self.root.mkdir()
        # L0 resolves the IA module from git history, so the repository starts in git; a kit made with git=False leaves
        # git after its first L0 (the manifest it wrote stays fresh: it hashes files, not history).
        self.git = git
        subprocess.run(['git', 'init', '-q', str(self.root)], check=True)
        self.catalogs = tmp_path / 'catalogs'
        write(self.catalogs / 'test-v1.json', {'schema_version': 1, 'version': 'test-v1', 'items': [_item()]})
        write(self.catalogs / 'id-ledger.json', {'schema_version': 1, 'next_id': 2,
              'entries': [{'id': 'dads-000001', 'match_key': 'dads-000001', 'retired': False}]})
        init_design(self.root, 'lp', 'test-v1', catalog_dir=self.catalogs)
        self.design_config(items={'dads-000001': {'state': 'defined', 'evidence': ['token:color.primary']}},
                           modules={'ia': '0.4.0'})
        ssot = copy.deepcopy(IA_FIXTURE['SSOT'])
        ssot['proofs']['prf-0001']['as_of'] = '2099-01-01'
        dictionary = copy.deepcopy(IA_FIXTURE['DICTIONARY'])
        dictionary['entries'][0]['expires'] = '2099-01-01'
        IA_FIXTURE['write_repo'](self.root, ssot, outlines=(IA_FIXTURE['OUTLINE'], IA_FIXTURE['DECK'], SCRIPT),
                                 dictionary=dictionary)
        write(lp.path(self.root, 'design_dir', 'tokens/brand.tokens.json'), {
            'color': {'$type': 'color',
                      'primary': {'$value': {'colorSpace': 'srgb', 'components': [0, 0.2, 0.4], 'hex': '#003366'},
                                  '$extensions': {'org.own.design': {'origin': 'proposal'}}},
                      'surface': {'$value': {'colorSpace': 'srgb', 'components': [1, 1, 1], 'hex': '#ffffff'},
                                  '$extensions': {'org.own.design': {'origin': 'proposal'}}}},
            'font': {'$type': 'fontFamily', 'body': {'$value': ['Sample Sans', 'sans-serif'],
                                                     '$extensions': {'org.own.design': {'origin': 'proposal'}}}}})
        self.plan(PLAN)
        self.brief(BRIEF)

    def design_config(self, **changes):
        path = lp.path(self.root, 'design_config')
        config = json.loads(path.read_text())
        config.update(changes)
        write(path, config)

    def l0(self):
        result = check_design(self.root, catalog_dir=self.catalogs)
        assert result['exit_code'] < 2, result['errors']
        if not self.git and (self.root / '.git').is_dir():
            shutil.move(str(self.root / '.git'), str(self.root.parent / 'git-moved-away'))
        return result

    def script(self, data, ident=None):
        IA_FIXTURE['dump'](lp.path(self.root, 'infoarch', 'scripts', f'{ident or data["id"]}.yaml'), data)

    def plan(self, data, deliverable=None):
        path = lp.path(self.root, 'build', deliverable or data['deliverable'], 'plan.yaml')
        write(path, data)
        return path

    def brief(self, data, deliverable=None):
        path = lp.path(self.root, 'build', deliverable or data['deliverable'], 'brief.yaml')
        write(path, data)
        return path

    def brief_path(self, deliverable='lp-sample'):
        return lp.path(self.root, 'build', deliverable, 'brief.yaml')

    def plan_text(self, text, deliverable='lp-sample'):
        path = lp.path(self.root, 'build', deliverable, 'plan.yaml')
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text)

    def comp_dir(self, deliverable='lp-sample'):
        return lp.path(self.root, 'build', deliverable, 'comp')

    def ignore_comps(self):
        shutil.copyfile(SKILLS / 'own-web-build/templates/build/.gitignore', lp.path(self.root, 'build', '.gitignore'))


def plan(**changes):
    data = copy.deepcopy(PLAN)
    data.update(changes)
    return data


def brief(**changes):
    data = copy.deepcopy(BRIEF)
    data.update(changes)
    return data


def run(script, *args, env=None):
    merged = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable)
    merged.update(env or {})
    result = subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)], capture_output=True, text=True,
                            env=merged, timeout=120)
    try:
        out = json.loads(result.stdout)
    except json.JSONDecodeError:
        out = None
    return result.returncode, out, result


def codes(report, level=None):
    return [f['code'] for f in report['findings'] if level is None or f['level'] == level]
