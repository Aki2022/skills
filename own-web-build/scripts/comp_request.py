#!/usr/bin/env python3
"""Write the comp request of a deliverable (カンプ; SPEC-deliverable-production-layers requirements 14, 22, 23).

The comp is the whole page drawn wide, block by block in the script's order, in one image or split into `parts`
contiguous images. Its inputs:
- the script (台本): blocks, their content function and copy, and the style it is written for (tone, target condition,
  objective action). Dictionary values ({{value-…}}) in the copy are filled in from the repository's dictionary; a copy
  longer than the comp settings' text_max_chars is listed but drawn as gray placeholder lines. The copy is placeholder
  text the comp may show; its wording is never used.
- the build brief (ux/build/<deliverable id>/brief.yaml, required; references/schema/brief.v1.schema.json): which of
  design's tokens the deliverable uses (a design with several brand themes passes only one), the style reference
  images (own-image-gen's anchors) and, per block, a visual requirement (emphasis).
- the composition mode: `plan` (default) draws the layout plan (each block's pattern and details; the plan is checked
  and required); `free` is an experiment pending the owner's decision on the flow order (requirement 14): the plan is
  not read, and the image model proposes each block's composition while keeping the block order and functions.
- the resolved design (the manifest): colors, fonts and the comp settings (`comp`: width, text_max_chars, parts).

This writes own-image-gen's field files (the ten labels of its prompt shell) and a request with the output paths under
ux/build/<deliverable id>/comp/ (git-ignored), then stops: it never generates an image. Run own-image-gen with the
commands the request holds, and show the comp to the owner (the comp gate).

Usage: comp_request.py --repo <repo> --deliverable <id> [--media web] [--parts N] [--composition plan|free]
Exit: 0 written, 1 written with warnings (the plan check warned, a dictionary value is unresolved, a token glob matches
nothing, or the ignore of the comp dir could not be checked), 2 refused: the brief is missing or broken, the plan check
fails (plan mode), the comp dir would be committed, the comp settings are unknown or the parts are out of range.
Nothing is written on 2.
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import plan_check  # noqa: E402  (also puts own-design-author/scripts on sys.path)
from plan_check import FUNCTION_NAMES, SKILL, SKILLS, layer_paths  # noqa: E402

IMAGE_GEN = '<skills>/own-image-gen/scripts/image_gen.py'  # <skills> is the skills root: no local path is written
BRIEF = 'brief.yaml'
BRIEF_SCHEMA = SKILL / 'references/schema/brief.v1.schema.json'
REQUEST = 'request.json'
MODES = ('plan', 'free')
GRAY = '灰色の線'


def brief_path(repo, deliverable):
    return layer_paths.path(repo, 'build', deliverable, BRIEF)


def _names(parts):
    """[(fields file, output image)] per part: one part keeps the plain names, a split numbers them from 1."""
    if parts == 1:
        return [('fields.json', 'comp.png')]
    return [(f'fields-{k}.json', f'comp-{k}.png') for k in range(1, parts + 1)]


# --- dictionary values ----------------------------------------------------------------------------------------------

def _ia_check():
    """own-infoarch-author's check module: its placeholder rule and dictionary view are the one reading of values."""
    path = str(SKILLS / 'own-infoarch-author/scripts')
    if path not in sys.path:
        sys.path.insert(0, path)
    import check as ia_check  # noqa: E402  (its common module puts own-doc-update's dictionary reader on sys.path)
    return ia_check


def dictionary_values(root):
    """({dictionary id: text} of the repository's `value` entries, [errors]) through the infoarch reader."""
    ia_check = _ia_check()
    import dictionary
    entries, _, errors = dictionary.load(Path(root))
    facts, _ = ia_check.terms_of(entries)
    return {k: str(v['value']) for k, v in facts.items()}, errors


def fill_values(text, values):
    """(text with every {{value-…}} the dictionary has filled in, [placeholder ids left as they were])."""
    missing = []

    def one(m):
        if m.group(1) in values:
            return values[m.group(1)]
        missing.append(m.group(1))
        return m.group(0)
    return _ia_check().PLACEHOLDER.sub(one, text), missing


# --- the script and the design --------------------------------------------------------------------------------------

def _ordered(script):
    """[(block id, block, [copy text, ...])] in `order`, copies in their own `order` (ties by id), as the script reads."""
    copies = script.get('copies') or {}
    blocks = sorted((script.get('blocks') or {}).items(), key=lambda kv: (kv[1].get('order', 0), kv[0]))
    return [(bid, b, [c['text'] for _, c in sorted(((cid, c) for cid, c in copies.items() if c.get('block') == bid),
                                                    key=lambda kv: (kv[1].get('order', 0), kv[0]))])
            for bid, b in blocks]


def filter_tokens(tokens, rule):
    """(tokens kept by the brief's include/exclude globs, [globs that matched no token])."""
    include, exclude = rule.get('include'), rule.get('exclude') or []
    unused = [g for g in (include or []) + exclude if not any(fnmatch.fnmatchcase(p, g) for p in tokens)]
    kept = {p: t for p, t in tokens.items()
            if (include is None or any(fnmatch.fnmatchcase(p, g) for g in include))
            and not any(fnmatch.fnmatchcase(p, g) for g in exclude)}
    return kept, unused


def _rows(tokens, kind):
    return [f'{path} {token["css"]}' for path, token in sorted(tokens.items())
            if token.get('type') == kind and isinstance(token.get('css'), str)]


def split(blocks, parts):
    """`blocks` in order cut into `parts` contiguous groups whose sizes differ by at most one (the larger ones first)."""
    size, extra = divmod(len(blocks), parts)
    out, start = [], 0
    for k in range(parts):
        end = start + size + (1 if k < extra else 0)
        out.append(blocks[start:end])
        start = end
    return out


# --- the fields -----------------------------------------------------------------------------------------------------

def _line(n, bid, block, mode, plan, vocabulary, emphasis):
    function = FUNCTION_NAMES.get(block.get('function'), block.get('function'))
    if mode == 'plan':
        row = plan['blocks'][bid]
        pattern = vocabulary[row['pattern']]
        line = (f'{n}. {bid}（働き：{function}）型 {row["pattern"]}：{pattern["description"]}。'
                f'列 {pattern["columns"]}、帯 {"あり" if pattern["band"] else "なし"}、'
                f'絵の置き方 {"/".join(pattern["images"])}、区切り {"/".join(pattern["separators"])}')
        if row.get('details'):
            line += f'。細部：{row["details"]}'
    else:
        line = f'{n}. {bid}（働き：{function}）組み方は提案する'
    if emphasis.get(bid):
        line += f'。見た目の要望：{emphasis[bid]}'
    return line


def _content(n, bid, texts, limit):
    if not texts:
        return f'{n}. {bid}：（文案なし）'
    shown = [f'［本文：{GRAY}で描く・約 {len(t)} 字］「{t}」' if len(t) > limit else f'「{t}」' for t in texts]
    return f'{n}. {bid}：' + ' / '.join(shown)


def build_fields(script, plan, tokens, settings, *, mode='plan', vocabulary=None, emphasis=None, references=(),
                 part=(1, 1), numbered=None):
    """The ten fields of own-image-gen's prompt shell for one comp image. Every value is text.

    `numbered` is [(page position, block id, block, [copy])] of the blocks this image draws (all blocks by default);
    `part` is (k, of). The copy is drawn as given: fill dictionary values in first (fill_values)."""
    emphasis = emphasis or {}
    width, limit = settings['width'], settings['text_max_chars']
    if numbered is None:
        numbered = [(n, bid, b, texts) for n, (bid, b, texts) in enumerate(_ordered(script), 1)]
    k, of = part
    composition = [_line(n, bid, b, mode, plan, vocabulary, emphasis) for n, bid, b, _ in numbered]
    content = [_content(n, bid, texts, limit) for n, bid, _, texts in numbered]
    palette, fonts = _rows(tokens, 'color'), _rows(tokens, 'fontFamily')
    ids = [bid for _, bid, _, _ in numbered]
    span = ids[0] if len(ids) == 1 else f'{ids[0]}〜{ids[-1]}'
    if of == 1:
        asset = f'full mockup：ページ全体を上から下まで描いた 1 枚のカンプ（幅 {width}px の広い画面、縦長）'
    else:
        asset = (f'full mockup：ページを上から {of} 枚に分けたカンプの {of} 枚のうち {k} 枚目（まとまり {span}、'
                 f'幅 {width}px の広い画面）。前後の枚と続けて 1 ページに見える描き方にする')
    style = (f'台本のトーン「{script["tone"]}」で描いた完成見本。読む人：{script["target"]["condition"]}。'
             f'読んだ人にしてほしい行動：{script["objective"]["action"]}。色と書体は下の値だけを使う'
             + (f'。書体：{"、".join(fonts)}' if fonts else ''))
    if mode == 'plan':
        house = f'台本のまとまり {len(numbered)} 個を、番号の順に上から 1 列に積む。まとまりごとに、組み方の計画が指す型で組む'
        head = ''
        keep = 'まとまりの数と順を変えない。各まとまりの型（列の数・帯・絵の置き方・区切り）を守る。'
    else:
        house = (f'台本のまとまり {len(numbered)} 個を、番号の順に上から積む。まとまりごとの組み方は、'
                 'その働きを最も強く伝える形を自由に提案する')
        head = ('各まとまりの組み方を自由に提案する：格子と列の数、帯、絵の大きさと置き場所、重なり、はみ出し（裁ち落とし）、'
                '何を目立たせるか。まとまりの順と、各まとまりの働きは変えない。見た目の要望があるまとまりはそれに応える。\n')
        keep = 'まとまりの数と順、各まとまりの働きを変えない。'
    text_rule = (f'文字として描くのは {limit} 字以下の短い文言（見出し・数字・ボタン）だけ。'
                 f'「{GRAY}で描く」と印を付けた本文は、読める文字にせず灰色の線の束で描く')
    return {
        'Use case': 'ui-mockup',
        'Asset type': asset,
        'Style references': (f'style_ref images（{len(references)} 枚、--anchor で渡す）：{", ".join(references)}'
                             if references else ''),
        'Style': style,
        'House structure': house,
        'Composition': head + '\n'.join(composition),
        'Content': ('文言は仮置き。カンプは見せてよいが、カンプに描かれた文言は使わない（使うのは台本の文言）。'
                    f'{text_rule}。\n' + '\n'.join(content)),
        'Color palette': '、'.join(palette) if palette else '（design に色の値が無い）',
        'Constraints': (f'幅 {width}px の 1 枚。{keep}'
                        '言葉で決めきれない細部（絵と文字の重なり・はみ出し・接地・区切り方）はここで描いて決める。'
                        f'{text_rule}。狭い画面は描かない'),
        'Avoid': 'まとまりの追加・削除・並べ替え、台本に無い主張や数字、実在のロゴ・商標・人物写真、読めない細かい文字の詰め込み',
    }


# --- the request ----------------------------------------------------------------------------------------------------

def _refuse(out, code, message, **extra):
    out['findings'].append({'level': 'fail', 'code': code, 'message': message})
    out.update(extra)
    out['exit_code'] = 2
    return out


def _read_brief(repo, deliverable, out):
    """The brief, or None after a refusal was added to `out`."""
    rel = layer_paths.rel('build', deliverable, BRIEF)
    path = brief_path(repo, deliverable)
    if not path.is_file():
        _refuse(out, 'COMP_BRIEF_MISSING',
                f'制作の指示 {rel} が無い。使うトークン（tokens）・見本の画像（style_references）を書いてから頼む'
                '（形は own-web-build/references/plan-and-comp.md の「制作の指示」）')
        return None
    from design_author import read_data
    try:
        brief = read_data(path)
    except (OSError, ValueError) as exc:
        _refuse(out, 'COMP_BRIEF_INVALID', f'{rel} が読めない：{exc}')
        return None
    import jsonschema
    schema = json.loads(BRIEF_SCHEMA.read_text(encoding='utf-8'))
    errors = sorted(jsonschema.Draft202012Validator(schema).iter_errors(brief), key=lambda e: list(e.absolute_path))
    if errors:
        _refuse(out, 'COMP_BRIEF_INVALID', f'{rel} の形が schema に合わない：' + '；'.join(
            f'{"/".join(map(str, e.absolute_path)) or "(root)"}: {e.message}' for e in errors))
        return None
    if brief['deliverable'] != deliverable:
        _refuse(out, 'COMP_BRIEF_INVALID', f'{rel} の deliverable {brief["deliverable"]} が置き場所のディレクトリ名 {deliverable} と違う')
        return None
    missing = [r for r in brief['style_references'] if not (Path(repo) / r).is_file()]
    if missing:
        _refuse(out, 'COMP_BRIEF_REFERENCE_MISSING', f'{rel} の見本の画像 {missing} が無い')
        return None
    return brief


def _comp_settings(manifest):
    """{width, text_max_chars, parts} from the manifest's `comp`, or None when one is missing (an older L0)."""
    comp = manifest.get('comp') or {}
    keys = ('width', 'text_max_chars', 'parts')
    if any(isinstance(comp.get(k), bool) or not isinstance(comp.get(k), int) or comp.get(k) <= 0 for k in keys):
        return None
    return {k: comp[k] for k in keys}


def comp_request(repo, deliverable, media='web', parts=None, composition='plan'):
    repo = Path(repo)
    out = {'deliverable': deliverable, 'composition': composition, 'exit_code': 0, 'findings': []}
    if composition not in MODES:
        return _refuse(out, 'COMP_REQUEST_FAILED', f'composition は {MODES} のどれか')
    layer_paths.guard(repo)
    brief = _read_brief(repo, deliverable, out)
    if brief is None:
        return out
    from design_author import read_data
    from resolved_contract import load_manifest
    plan, vocabulary = None, None
    if composition == 'plan':
        checked = plan_check.check_plan(repo, deliverable, media)
        out['plan_check'] = {'exit_code': checked['exit_code'],
                             'findings': [f for f in checked['findings'] if f['level'] != 'info']}
        if checked['exit_code'] == 2:
            return _refuse(out, 'COMP_PLAN_FAILED', '組み方の計画の検査が失格。計画を直してから頼む（カンプは計画を絵にする）')
        plan = read_data(plan_check.plan_path(repo, deliverable))
        if plan['script'] != brief['script']:
            return _refuse(out, 'COMP_BRIEF_SCRIPT_MISMATCH',
                           f'制作の指示の script {brief["script"]} が組み方の計画の script {plan["script"]} と違う')
    else:
        out['plan_check'] = {'skipped': '組み方を自由に提案させる試し（composition free）。計画は読まない'}
    script_rel = layer_paths.rel('infoarch', 'scripts', f'{brief["script"]}.yaml')
    try:
        script = read_data(plan_check.script_path(repo, brief['script']))
        if not isinstance(script, dict) or not isinstance(script.get('blocks'), dict) or not script['blocks']:
            raise ValueError('まとまりが無い')
    except (OSError, ValueError) as exc:
        return _refuse(out, 'COMP_SCRIPT_MISSING', f'台本 {script_rel} が無いか読めないか、まとまりが 0 個（{type(exc).__name__}）')
    unknown = sorted(set(brief.get('emphasis') or {}) - set(script['blocks']))
    if unknown:
        return _refuse(out, 'COMP_BRIEF_BLOCK_UNKNOWN', f'制作の指示の emphasis のまとまり {unknown} が台本 {script_rel} に無い')
    folder = plan_check.comp_dir(repo, deliverable)
    rel = layer_paths.rel('build', deliverable, 'comp')
    ignored = plan_check.comp_ignored(repo, deliverable)
    if ignored is False:
        return _refuse(out, 'COMP_DIR_NOT_IGNORED',
                       f'{rel}/ が git に入る（カンプの画像は git に入れない）。own-web-build/templates/build/.gitignore を '
                       f'{layer_paths.rel("build", ".gitignore")} に写してから頼む')
    try:
        manifest = load_manifest(plan_check.manifest_path(repo, media))
    except (OSError, ValueError) as exc:
        return _refuse(out, 'COMP_MANIFEST_UNAVAILABLE', f'解決済みマニフェストが無いか古い（{type(exc).__name__}）。design-author check（L0）を先に流す')
    if composition == 'plan':
        vocabulary = manifest['layout_patterns']['patterns']
    settings = _comp_settings(manifest)
    if settings is None:
        return _refuse(out, 'COMP_SIZE_MISSING',
                       '解決済みマニフェストにカンプの設定（comp.width・comp.text_max_chars・comp.parts）が無い。design-author check（L0）をこの版で流し直す')
    blocks = _ordered(script)
    parts = settings['parts'] if parts is None else parts
    if not 1 <= parts <= len(blocks):
        return _refuse(out, 'COMP_PARTS_INVALID', f'parts {parts} は 1 から まとまりの数 {len(blocks)} までにする')
    tokens, unused = filter_tokens(manifest.get('tokens', {}), brief['tokens'])
    if not _rows(tokens, 'color'):
        return _refuse(out, 'COMP_BRIEF_NO_COLOR', f'制作の指示の tokens で絞ると色が 0 個になる（{brief["tokens"]}）')
    warnings = []
    if unused:
        warnings.append(('COMP_BRIEF_GLOB_UNUSED', f'制作の指示の tokens の glob {unused} がどのトークンにも一致しない（書き間違いなら絞り込みが効いていない）'))
    values, dict_errors = dictionary_values(repo)
    unresolved = []
    numbered = []
    for n, (bid, block, texts) in enumerate(blocks, 1):
        filled = []
        for text in texts:
            text, missing = fill_values(text, values)
            unresolved += [m for m in missing if m not in unresolved]
            filled.append(text)
        numbered.append((n, bid, block, filled))
    if unresolved:
        why = f'（辞書が読めない：{dict_errors}）' if dict_errors else ''
        warnings.append(('COMP_VALUE_UNRESOLVED', f'辞書の値 {unresolved} が見つからず、{{{{…}}}} のまま渡す{why}'))
    references = list(brief['style_references'])
    groups = split(numbered, parts)
    for stale in list(folder.glob('fields*.json')) if folder.is_dir() else []:
        stale.unlink()  # an earlier split's field files would read as part of this request
    folder.mkdir(parents=True, exist_ok=True)
    workdir = '.' if references else rel  # image_gen checks an anchor from the cwd and codex reads it from the workdir
    anchors = [a for r in references for a in ('--anchor', r)] + (['--take-latest'] if references else [])
    part_list, fields_rels, out_rels, commands = [], [], [], []
    for k, (group, (fname, oname)) in enumerate(zip(groups, _names(parts)), 1):
        fields = build_fields(script, plan, tokens, settings, mode=composition, vocabulary=vocabulary,
                              emphasis=brief.get('emphasis'), references=references, part=(k, parts), numbered=group)
        fields_rel, out_rel = f'{rel}/{fname}', f'{rel}/{oname}'
        (folder / fname).write_text(json.dumps(fields, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        command = ['python3', IMAGE_GEN, '--fields', fields_rel, '--out', out_rel, *anchors, '--workdir', workdir]
        part_list.append({'part': k, 'blocks': [bid for _, bid, _, _ in group], 'fields': fields_rel, 'out': out_rel,
                          'command': command})
        fields_rels.append(fields_rel)
        out_rels.append(out_rel)
        commands.append(command)
    request = {'deliverable': deliverable, 'script': brief['script'], 'media': media, 'composition': composition,
               'brief': layer_paths.rel('build', deliverable, BRIEF), 'width': settings['width'],
               'text_max_chars': settings['text_max_chars'], 'parts': parts, 'style_references': references,
               'fields': fields_rels, 'out': out_rels, 'commands': commands, 'part_list': part_list,
               'note': 'repo の root で上の commands を順に流す（<skills> は skills の root。--fields・--out・--anchor・--workdir は repo の root からの相対）。'
                       '描いた絵をオーナーに見せ（カンプの門）、足された細部を計画の details に写して comp.reviewed を true にする'
                       + ('。composition free は試し（流れの順はオーナーの決定待ち）：描いた組み方を計画に書き起こしてから HTML を組む'
                          if composition == 'free' else '')}
    (folder / REQUEST).write_text(json.dumps(request, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    out.update({'fields': fields_rels, 'out': out_rels, 'request': f'{rel}/{REQUEST}', 'width': settings['width'],
                'parts': parts})
    for code, message in warnings:
        out['findings'].append({'level': 'warn', 'code': code, 'message': message})
    if not references:
        out['findings'].append({'level': 'info', 'code': 'COMP_NO_STYLE_REFERENCE',
                                'message': '見本の画像（style_references）が無い。文と値だけで描く'})
    if ignored is None:
        out['findings'].append({'level': 'warn', 'code': 'COMP_IGNORE_NOT_CHECKED',
                                'message': f'git の作業ツリーでないので {rel}/ が git に入らないかを確かめていない'})
    levels = {f['level'] for f in out['findings']} | {f['level'] for f in out['plan_check'].get('findings', [])}
    out['exit_code'] = 1 if 'warn' in levels else 0
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description='write the comp request (own-image-gen fields) of a deliverable')
    parser.add_argument('--repo', default='.')
    parser.add_argument('--deliverable', required=True)
    parser.add_argument('--media', default='web')
    parser.add_argument('--parts', type=int, default=None, help='split the page into N images (default: the comp settings)')
    parser.add_argument('--composition', choices=MODES, default='plan',
                        help='plan: draw the layout plan; free: an experiment, the model proposes the composition')
    args = parser.parse_args(argv)
    try:
        result = comp_request(args.repo, args.deliverable, args.media, args.parts, args.composition)
    except layer_paths.LayerLayoutError as exc:
        result = {'deliverable': args.deliverable, 'exit_code': 2,
                  'findings': [{'level': 'fail', 'code': layer_paths.FOREIGN_CODE, 'message': str(exc)}]}
    except Exception as exc:  # noqa: BLE001  an uncaught error would exit 1, which reads as 'written with warnings'
        result = {'deliverable': args.deliverable, 'exit_code': 2,
                  'findings': [{'level': 'fail', 'code': 'COMP_REQUEST_FAILED', 'message': f'{type(exc).__name__}: {exc}'}]}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result['exit_code']


if __name__ == '__main__':
    plan_check.start_again_under_the_dependency_interpreter()
    raise SystemExit(main())
