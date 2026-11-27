#!/usr/bin/env python3
"""Check a deliverable's layout plan (組み方の計画; SPEC-deliverable-production-layers requirements 14, 15, 18, 21, 22).

The plan sits at ux/build/<deliverable id>/plan.yaml (the layer rule's `build` entry) and names, for every block of its
script (ux/infoarch/scripts/<script id>.yaml), a layout pattern of design's vocabulary, the narrow-screen shape when it
differs from the pattern's default, and the details the comp added. Checked:
- the plan reads (unique keys) and follows references/schema/plan.v1.schema.json; `deliverable` is its directory name;
- the script exists and the plan covers each of its blocks exactly once (a missing or an extra block fails; zero
  blocks is nothing checked, never a pass);
- every pattern and narrow name is in the resolved vocabulary (the L0 manifest; a missing or stale manifest fails);
- a pattern not suited to the block's content function warns; a block without `narrow` takes the pattern's default,
  and the report says which (info);
- the comp images stay out of git: an existing comp dir that git would commit warns; outside git it is reported as
  not checked (a warning, never a silent pass);
- a variant (`variant_of`, `overrides`) is compared with its base plan: an undeclared difference warns (交絡あり), a
  declared override that does not differ warns. `compare` is the one comparison the infoarch diff reads too;
- a block's emphasis (計画の強調: which characters of one copy or one dictionary value are drawn larger than the rest of it)
  names a copy of the block or a value one of its copies shows, and its text is in that target with something left over;
- with --html, the drawn page is measured at each width (measure_emphasis.mjs): an emphasized text drawn no larger than the
  rest of its target, or a target the block does not draw, fails; a page that cannot be drawn is not measured and fails
  (never a pass). Free-text details are not compared with the page; only the emphasis is.

Usage: plan_check.py --repo <repo> --deliverable <id> [--media web] [--html <drawn page>]
Exit: 0 pass, 1 warnings, 2 failing or nothing checked. The report has the five categories (実行・スキップ・欠落・
未カバー・screened) and the findings (fail / warn / info).
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
SKILLS = SKILL.parent
sys.path.insert(0, str(SKILLS / 'own-design-author/scripts'))
import layer_paths  # noqa: E402  (the one rule for where a repository keeps its layers)
import unicodedata  # noqa: E402
from reporting import empty_report, skip, stage  # noqa: E402  (the five report categories)

SCHEMA = SKILL / 'references/schema/plan.v1.schema.json'
FUNCTION_NAMES = {'identify': '名乗り', 'claim': '主張', 'fact': '事実', 'evidence': '証拠と声', 'offer': '提供条件',
                  'action': '行動', 'source': '出典と運営者'}  # docs dictionary term-function-*
CONFOUNDED_LABEL = '交絡あり'
REEXEC_GUARD = 'PLAN_CHECK_REEXEC'
MEASURE_EMPHASIS = Path(__file__).resolve().with_name('measure_emphasis.mjs')
VIEWPORTS = [{'name': 'mobile', 'width': 375, 'height': 812}, {'name': 'desktop', 'width': 1280, 'height': 900}]


def fold(text):
    """Characters as compared: width-folded (NFKC), whitespace removed (the way copy and dictionary values are read)."""
    return ''.join(unicodedata.normalize('NFKC', text or '').split())


def measure_emphasis(html, emphasis, repo, viewports=VIEWPORTS):
    """Draw the page at each width and measure each emphasis. -> rows; raises RuntimeError when it cannot be drawn."""
    env = {**os.environ, 'SCREEN_PLAYWRIGHT_DIR': str(repo)}
    proc = subprocess.run(['node', str(MEASURE_EMPHASIS), str(html), json.dumps(emphasis, ensure_ascii=False),
                           json.dumps(viewports)], capture_output=True, text=True, timeout=180, env=env)
    if proc.returncode != 0:
        raise RuntimeError(f'exit {proc.returncode}: {proc.stderr.strip()[:300]}')
    return json.loads(proc.stdout)


def start_again_under_the_dependency_interpreter():
    """PyYAML and jsonschema are L0's dependencies; when the python3 on PATH lacks them, start once more under
    DESIGN_AUTHOR_PYTHON (as audit_screen does). Without either, the reader reports the missing dependency."""
    try:
        import jsonschema  # noqa: F401
        import yaml  # noqa: F401
        return
    except ImportError:
        pass
    target = os.environ.get('DESIGN_AUTHOR_PYTHON')
    if not target or os.environ.get(REEXEC_GUARD):
        return
    os.environ[REEXEC_GUARD] = '1'
    try:
        os.execv(target, [target, str(Path(__file__).resolve()), *sys.argv[1:]])
    except OSError:
        pass


def read_data(path):
    from design_author import read_data as read  # JSON-compatible YAML with unique keys (L0's reader)
    return read(Path(path))


def plan_path(repo, deliverable):
    return layer_paths.path(repo, 'build', deliverable, 'plan.yaml')


def comp_dir(repo, deliverable):
    return layer_paths.path(repo, 'build', deliverable, 'comp')


def script_path(repo, script):
    return layer_paths.path(repo, 'infoarch', 'scripts', f'{script}.yaml')


def manifest_path(repo, media):
    return layer_paths.path(repo, 'resolved', f'{media}.json')


def find_plans(repo, script):
    """{deliverable id: plan} of every readable plan under the build entry whose `script` is `script`."""
    base = layer_paths.path(repo, 'build')
    out = {}
    if base.is_dir():
        for path in sorted(base.glob('*/plan.yaml')):
            try:
                data = read_data(path)
            except (OSError, ValueError):
                continue
            if isinstance(data, dict) and data.get('script') == script:
                out[path.parent.name] = data
    return out


def elements(plan):
    """{element key: content} of a plan, for comparisons (the comp review state is not a layout element)."""
    return {f'blocks/{bid}': block for bid, block in sorted((plan.get('blocks') or {}).items())}


_ABSENT = object()


def compare(base, variant, declared=()):
    """The plan side of a variant comparison (requirement 18): what changed between two plans, which changes the variant
    did not declare (a confounded comparison) and which declared overrides did not change, with a warning each."""
    ea, eb = elements(base), elements(variant)
    changes = []
    for key in sorted(set(ea) | set(eb)):
        if key not in eb:
            changes.append({'layer': 'plan', 'element': key, 'change': 'removed'})
        elif key not in ea:
            changes.append({'layer': 'plan', 'element': key, 'change': 'added'})
        elif ea[key] != eb[key]:
            attrs = sorted(k for k in set(ea[key]) | set(eb[key]) if ea[key].get(k) != eb[key].get(k))
            changes.append({'layer': 'plan', 'element': key, 'change': 'changed', 'attributes': attrs})
    changed = [c['element'] for c in changes]
    declared = set(declared or ())
    undeclared = sorted(set(changed) - declared)
    unchanged = sorted(k for k in declared if k not in changed and (k in ea or k in eb))
    findings = []
    if undeclared:
        findings.append({'level': 'warn', 'code': 'PLAN_UNDECLARED_DIFFERENCE',
                         'message': f'{CONFOUNDED_LABEL}：基本の計画と違うのに上書きとして宣言していないまとまり {undeclared}'})
    if unchanged:
        findings.append({'level': 'warn', 'code': 'PLAN_OVERRIDE_UNCHANGED',
                         'message': f'上書きとして宣言したのに基本の計画と違わないまとまり {unchanged}'})
    return {'changes': changes, 'changed': changed, 'undeclared': undeclared, 'unchanged': unchanged,
            'confounded': bool(undeclared), 'findings': findings}


def comp_ignored(repo, deliverable):
    """True when git ignores the comp dir, False when git would commit it, None when this is not a git work tree."""
    inside = subprocess.run(['git', '-C', str(repo), 'rev-parse', '--is-inside-work-tree'], capture_output=True, text=True)
    if inside.returncode != 0 or inside.stdout.strip() != 'true':
        return None
    probe = subprocess.run(['git', '-C', str(repo), 'check-ignore', '-q', layer_paths.rel('build', deliverable, 'comp') + '/'],
                           capture_output=True, text=True)
    if probe.returncode in (0, 1):
        return probe.returncode == 0
    return None


class Check:
    def __init__(self, repo, deliverable, media, html=None):
        self.repo, self.deliverable, self.media, self.html = Path(repo), deliverable, media, html
        self.findings, self.report = [], empty_report()
        self.out = {'deliverable': deliverable, 'plan': layer_paths.rel('build', deliverable, 'plan.yaml'), 'media': media,
                    'script': None, 'blocks': 0, 'narrow': {}, 'variant': None, 'emphasis': 0, 'emphasis_drawn': None}

    def add(self, level, code, message, path=None):
        self.findings.append({'level': level, 'code': code, 'message': message,
                              **({'path': path} if path else {})})

    def done(self):
        levels = {f['level'] for f in self.findings}
        code = 2 if 'fail' in levels else 1 if 'warn' in levels else 0
        return {**self.out, 'exit_code': code, 'findings': self.findings, 'report': self.report}

    def _schema(self, plan):
        try:
            import jsonschema
        except ImportError:
            self.add('fail', 'PLAN_UNREADABLE', 'jsonschema が無い（L0 の依存環境で動かす。DESIGN_AUTHOR_PYTHON）')
            return False
        schema = json.loads(SCHEMA.read_text(encoding='utf-8'))
        errors = sorted(jsonschema.Draft202012Validator(schema).iter_errors(plan), key=lambda e: list(e.absolute_path))
        for error in errors:
            where = '/'.join(map(str, error.absolute_path)) or '(root)'
            self.add('fail', 'PLAN_SCHEMA', f'{where}: {error.message}', self.out['plan'])
        return not errors

    def _vocabulary(self):
        try:
            from resolved_contract import load_manifest
            manifest = load_manifest(manifest_path(self.repo, self.media))
            vocabulary = manifest.get('layout_patterns')
            if not isinstance(vocabulary, dict) or not isinstance(vocabulary.get('patterns'), dict):
                raise ValueError('manifest has no layout pattern vocabulary; rerun L0 with this version')
            return vocabulary['patterns']
        except layer_paths.LayerLayoutError:
            raise
        except (OSError, ValueError, KeyError, TypeError) as exc:
            text = str(exc).replace(str(self.repo.resolve()), '<repo>').replace(str(self.repo), '<repo>')
            self.add('fail', 'PLAN_MANIFEST_UNAVAILABLE',
                     f'組み方の型の語彙を読む解決済みマニフェスト（{layer_paths.rel("resolved", self.media + ".json")}）が無いか古い：{text}。'
                     'design-author check（L0）を先に流す')
            self.report['missing'].append({'id': 'layout-vocabulary', 'reason': 'resolved manifest unavailable or stale'})
            return None

    def run(self):
        layer_paths.guard(self.repo)
        path = plan_path(self.repo, self.deliverable)
        skip(self.report, 'script-check', '台本そのものの検査は infoarch-author check（L0 の IA）が持つ。ここでは読むだけ')
        if not path.is_file():
            self.add('fail', 'PLAN_MISSING', f'組み方の計画 {self.out["plan"]} が無い（確かめる対象が無い）', self.out['plan'])
            self.report['missing'].append({'id': 'plan', 'reason': 'no plan file'})
            return self.done()
        try:
            plan = read_data(path)
        except (OSError, ValueError) as exc:
            self.add('fail', 'PLAN_UNREADABLE', f'組み方の計画が読めない：{exc}', self.out['plan'])
            return self.done()
        if not self._schema(plan):
            stage(self.report, 'plan', 2)
            return self.done()
        self.out['script'] = plan['script']
        if plan['deliverable'] != self.deliverable:
            self.add('fail', 'PLAN_DELIVERABLE_MISMATCH',
                     f'deliverable {plan["deliverable"]} が置き場所のディレクトリ名 {self.deliverable} と違う', self.out['plan'])
        script_rel = layer_paths.rel('infoarch', 'scripts', f'{plan["script"]}.yaml')
        try:
            script = read_data(script_path(self.repo, plan['script']))
            blocks = script.get('blocks') if isinstance(script, dict) else None
            if not isinstance(blocks, dict):
                raise ValueError('blocks が無い')
        except (OSError, ValueError) as exc:
            self.add('fail', 'PLAN_SCRIPT_MISSING', f'台本 {script_rel} が無いか読めない（{type(exc).__name__}）', script_rel)
            stage(self.report, 'plan', 2)
            return self.done()
        vocabulary = self._vocabulary()
        self._coverage(plan, blocks, vocabulary, script_rel)
        emphasis = self._emphasis(plan, script)
        if self.html is not None:
            self._drawn(emphasis)
        self._comp(plan)
        self._variant(plan)
        result = self.done()
        stage(self.report, 'plan', result['exit_code'], len(blocks))
        return {**result, 'report': self.report}

    def _coverage(self, plan, blocks, vocabulary, script_rel):
        self.out['blocks'] = len(blocks)
        if not blocks:
            self.add('fail', 'PLAN_NOTHING_CHECKED', f'台本 {script_rel} にまとまりが 0 個（対象 0 件を合格にしない）', script_rel)
            return
        planned = plan['blocks']
        missing = sorted(set(blocks) - set(planned))
        extra = sorted(set(planned) - set(blocks))
        if missing:
            self.add('fail', 'PLAN_BLOCK_MISSING', f'台本のまとまり {missing} の組み方が計画に無い', self.out['plan'])
        if extra:
            self.add('fail', 'PLAN_BLOCK_UNKNOWN', f'計画のまとまり {extra} が台本に無い', self.out['plan'])
        if vocabulary is None:
            return
        defaults = []
        for bid in sorted(set(blocks) & set(planned)):
            row, function = planned[bid], (blocks[bid] or {}).get('function')
            pattern = vocabulary.get(row['pattern'])
            if pattern is None:
                self.add('fail', 'PLAN_PATTERN_UNKNOWN', f'{bid}: 型 {row["pattern"]} が design の語彙に無い', self.out['plan'])
            elif function not in pattern.get('functions', []):
                self.add('warn', 'PLAN_PATTERN_UNSUITED',
                         f'{bid}: 型 {row["pattern"]} は働き「{FUNCTION_NAMES.get(function, function)}」に向く型として書かれていない'
                         f'（向く働き：{"・".join(FUNCTION_NAMES.get(f, f) for f in pattern.get("functions", []))}）', self.out['plan'])
            if 'narrow' in row:
                if row['narrow'] not in vocabulary:
                    self.add('fail', 'PLAN_NARROW_UNKNOWN', f'{bid}: 狭い画面の型 {row["narrow"]} が design の語彙に無い', self.out['plan'])
                else:
                    self.out['narrow'][bid] = {'pattern': row['narrow'], 'source': 'plan'}
            elif pattern is not None:
                self.out['narrow'][bid] = {'pattern': pattern['narrow'], 'source': 'default'}
                defaults.append(f'{bid}→{pattern["narrow"]}')
        if defaults:
            self.add('info', 'PLAN_NARROW_DEFAULT', f'narrow を書いていないまとまりは型の既定で崩す：{", ".join(defaults)}')
        if not plan['comp']['reviewed']:
            self.add('info', 'PLAN_COMP_NOT_REVIEWED', 'カンプの門をまだ通っていない（comp.reviewed: false）。HTML はカンプを見て計画を直してから組む')

    def _comp(self, plan):
        folder = comp_dir(self.repo, self.deliverable)
        rel = layer_paths.rel('build', self.deliverable, 'comp') + '/'
        if not folder.exists():
            skip(self.report, 'comp-ignore', f'{rel} がまだ無い')
            return
        ignored = comp_ignored(self.repo, self.deliverable)
        if ignored is None:
            self.add('warn', 'PLAN_COMP_IGNORE_NOT_CHECKED', f'git の作業ツリーでないので {rel} が git に入らないかを確かめていない', rel)
            self.report['uncovered'].append({'id': 'comp-ignore', 'reason': 'not a git work tree'})
            return
        stage(self.report, 'comp-ignore', 0 if ignored else 1, 1)
        if not ignored:
            self.add('warn', 'PLAN_COMP_NOT_IGNORED',
                     f'{rel} が git に入る（カンプの画像は git に入れない）。own-web-build/templates/build/.gitignore を '
                     f'{layer_paths.rel("build", ".gitignore")} に写す', rel)

    def _emphasis(self, plan, script):
        """The plan's emphasis, read against the script: -> [{block, target, text}] of the ones that can be drawn."""
        rows = [(bid, e) for bid, row in sorted(plan['blocks'].items()) for e in row.get('emphasis', [])]
        self.out['emphasis'] = len(rows)
        if not rows:
            self.add('info', 'PLAN_EMPHASIS_NONE', '計画に強調（emphasis）が無い：大きく描く文字を描いた画面で確かめる対象が無い')
            skip(self.report, 'plan-emphasis', 'no emphasis in the plan')
            return []
        from comp_request import _ia_check, dictionary_values
        values, _ = dictionary_values(self.repo)
        placeholder = _ia_check().PLACEHOLDER
        copies = script.get('copies') or {}
        good = []
        for bid, e in rows:
            mine = {cid: c for cid, c in copies.items() if isinstance(c, dict) and c.get('block') == bid}
            shown = {m.group(1) for c in mine.values() for m in placeholder.finditer(str(c.get('text', '')))}
            if e['target'] in mine:
                text = placeholder.sub(lambda m: values.get(m.group(1), m.group(0)), str(mine[e['target']].get('text', '')))
            elif e['target'] in shown and e['target'] in values:
                text = values[e['target']]
            else:
                self.add('fail', 'PLAN_EMPHASIS_TARGET_UNKNOWN',
                         f'{bid}: 強調の対象 {e["target"]} は、このまとまりの文案でも、その文案が示す辞書の値でもない', self.out['plan'])
                continue
            want, have = fold(e['text']), fold(text)
            if want not in have or want == have:
                why = '対象の文字に無い' if want not in have else '対象の文字の全部で、残りが無い（何より大きくするかが無い）'
                self.add('fail', 'PLAN_EMPHASIS_TEXT_ABSENT', f'{bid}: 強調の文字「{e["text"]}」が {e["target"]} の{why}', self.out['plan'])
                continue
            good.append({'block': bid, 'target': e['target'], 'text': e['text']})
        stage(self.report, 'plan-emphasis', 2 if len(good) < len(rows) else 0, len(rows))
        return good

    def _drawn(self, emphasis):
        """Measure the emphasis on the drawn page (--html)."""
        if not emphasis:
            skip(self.report, 'plan-emphasis-drawn', 'nothing to measure')
            return
        try:
            rows = measure_emphasis(Path(self.html).resolve(), emphasis, self.repo.resolve())
        except (OSError, RuntimeError, ValueError, subprocess.TimeoutExpired) as exc:
            self.add('fail', 'PLAN_EMPHASIS_NOT_MEASURED', f'描いた画面で強調を測れなかった（合格にしない）：{exc}')
            self.report['missing'].append({'id': 'plan-emphasis-drawn', 'reason': 'the page could not be drawn'})
            return
        failed = 0
        for r in rows:
            where = f'{r["block"]} {r["target"]}「{r["text"]}」（{r["viewport"]}）'
            if not r['drawn']:
                failed += 1
                self.add('fail', 'PLAN_EMPHASIS_TARGET_NOT_DRAWN', f'{where}: 対象がこのまとまりの中に描かれていない')
                continue
            bad = [x for x in r['elements'] if not x.get('larger')]
            if bad:
                failed += 1
                sizes = '、'.join(f'強調 {x.get("emphasized_px")}px・残り {x.get("rest_px")}px' if x.get('found') else '強調の文字が無い'
                                 for x in bad)
                self.add('fail', 'PLAN_EMPHASIS_NOT_RENDERED', f'{where}: 強調の文字が残りより大きく描かれていない（{sizes}）')
        self.out['emphasis_drawn'] = {'checked': len(rows), 'failed': failed}
        stage(self.report, 'plan-emphasis-drawn', 2 if failed else 0, len(rows))

    def _variant(self, plan):
        base_id, declared = plan.get('variant_of'), plan.get('overrides')
        if base_id is None:
            if declared is not None:
                self.add('fail', 'PLAN_OVERRIDES_WITHOUT_BASE', 'overrides を書いたのに variant_of（基本の計画）が無い', self.out['plan'])
            else:
                skip(self.report, 'plan-variant', '変種ではない')
            return
        base_rel = layer_paths.rel('build', base_id, 'plan.yaml')
        try:
            base = read_data(plan_path(self.repo, base_id)) if base_id != self.deliverable else None
            if not isinstance(base, dict) or not isinstance(base.get('blocks'), dict):
                raise ValueError('not a plan')
        except (OSError, ValueError):
            self.add('fail', 'PLAN_VARIANT_BASE_MISSING', f'基本の計画 {base_rel} が無いか読めない（自分自身も基本にできない）', base_rel)
            return
        declared = declared or []
        known = set(elements(base)) | set(elements(plan))
        unknown = sorted(k for k in declared if k not in known)
        if unknown:
            self.add('fail', 'PLAN_OVERRIDE_UNKNOWN', f'上書きの宣言 {unknown} がどちらの計画のまとまりでもない', self.out['plan'])
        compared = compare(base, plan, declared)
        for f in compared['findings']:
            self.add(f['level'], f['code'], f['message'], self.out['plan'])
        self.out['variant'] = {'variant_of': base_id, **{k: compared[k] for k in ('changed', 'undeclared', 'unchanged', 'confounded')},
                               'confounded_label': CONFOUNDED_LABEL if compared['confounded'] else None}
        stage(self.report, 'plan-variant', 1 if compared['findings'] else 0, len(known))


def check_plan(repo, deliverable, media='web', html=None):
    return Check(repo, deliverable, media, html).run()


def main(argv=None):
    parser = argparse.ArgumentParser(description='check a deliverable layout plan (組み方の計画)')
    parser.add_argument('--repo', default='.')
    parser.add_argument('--deliverable', required=True)
    parser.add_argument('--media', default='web')
    parser.add_argument('--html', help='the drawn page: measure the plan emphasis on it at each width')
    args = parser.parse_args(argv)
    try:
        result = check_plan(Path(args.repo), args.deliverable, args.media, args.html)
    except layer_paths.LayerLayoutError as exc:
        result = {'deliverable': args.deliverable, 'exit_code': 2, 'report': empty_report(),
                  'findings': [{'level': 'fail', 'code': layer_paths.FOREIGN_CODE, 'message': str(exc)}]}
    except Exception as exc:  # noqa: BLE001  an uncaught error would exit 1, which reads as a warning
        result = {'deliverable': args.deliverable, 'exit_code': 2, 'report': empty_report(),
                  'findings': [{'level': 'fail', 'code': 'PLAN_CHECK_FAILED', 'message': f'{type(exc).__name__}: {exc}'}]}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result['exit_code']


if __name__ == '__main__':
    start_again_under_the_dependency_interpreter()
    raise SystemExit(main())
