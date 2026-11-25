#!/usr/bin/env python3
"""Write the comp request of a deliverable (カンプ; SPEC-deliverable-production-layers requirements 14, 22, 23).

The comp is one wide image of the whole page, block by block in the script's order: for each block its content function,
the layout pattern of the plan and the details the plan already holds, with the script's copy as placeholder text the comp
may show (its wording is never used). The input is media-neutral (script, plan, design's resolved values); only the size
differs per media, and it comes from the resolved profile (the manifest's `comp`).

This writes own-image-gen's field file (the ten labels of its prompt shell) and a request with the output path under
ux/build/<deliverable id>/comp/ (git-ignored), then stops: it never generates an image. Run own-image-gen with the
command the request holds, and show the comp to the owner (the comp gate).

Usage: comp_request.py --repo <repo> --deliverable <id> [--media web]
Exit: 0 written, 1 written with warnings (the plan check warned, or the ignore of the comp dir could not be checked),
2 refused: the plan check fails, the comp dir would be committed, or the size is unknown. Nothing is written on 2.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import plan_check  # noqa: E402  (also puts own-design-author/scripts on sys.path)
from plan_check import FUNCTION_NAMES, layer_paths  # noqa: E402

IMAGE_GEN = '<skills>/own-image-gen/scripts/image_gen.py'  # <skills> is the skills root: no local path is written
OUTPUT = 'comp.png'
FIELDS = 'fields.json'
REQUEST = 'request.json'


def _ordered(script):
    """[(block id, block, [copy text, ...])] in `order`, copies in their own `order` (ties by id), as the script reads."""
    copies = script.get('copies') or {}
    blocks = sorted((script.get('blocks') or {}).items(), key=lambda kv: (kv[1].get('order', 0), kv[0]))
    return [(bid, b, [c['text'] for _, c in sorted(((cid, c) for cid, c in copies.items() if c.get('block') == bid),
                                                    key=lambda kv: (kv[1].get('order', 0), kv[0]))])
            for bid, b in blocks]


def _palette(manifest):
    rows = [f'{path} {token["css"]}' for path, token in sorted(manifest.get('tokens', {}).items())
            if token.get('type') == 'color' and isinstance(token.get('css'), str)]
    return rows


def _fonts(manifest):
    return [f'{path} {token["css"]}' for path, token in sorted(manifest.get('tokens', {}).items())
            if token.get('type') == 'fontFamily' and isinstance(token.get('css'), str)]


def build_fields(script, plan, manifest, width):
    """The ten fields of own-image-gen's prompt shell for one comp. Every value is text."""
    vocabulary = manifest['layout_patterns']['patterns']
    blocks = _ordered(script)
    composition, content = [], []
    for n, (bid, block, texts) in enumerate(blocks, 1):
        row = plan['blocks'][bid]
        pattern = vocabulary[row['pattern']]
        function = FUNCTION_NAMES.get(block.get('function'), block.get('function'))
        line = (f'{n}. {bid}（働き：{function}）型 {row["pattern"]}：{pattern["description"]}。'
                f'列 {pattern["columns"]}、帯 {"あり" if pattern["band"] else "なし"}、'
                f'絵の置き方 {"/".join(pattern["images"])}、区切り {"/".join(pattern["separators"])}')
        if row.get('details'):
            line += f'。細部：{row["details"]}'
        composition.append(line)
        content.append(f'{n}. {bid}：' + ' / '.join(f'「{t}」' for t in texts) if texts else f'{n}. {bid}：（文案なし）')
    palette = _palette(manifest)
    fonts = _fonts(manifest)
    return {
        'Use case': 'ui-mockup',
        'Asset type': f'full mockup：ページ全体を上から下まで描いた 1 枚のカンプ（幅 {width}px の広い画面、縦長）',
        'Style references': '',
        'Style': ('design の値に沿った落ち着いた完成見本。色と書体は下の値だけを使う'
                  + (f'。書体：{"、".join(fonts)}' if fonts else '')),
        'House structure': f'台本のまとまり {len(blocks)} 個を、番号の順に上から 1 列に積む。まとまりごとに、組み方の計画が指す型で組む',
        'Composition': '\n'.join(composition),
        'Content': ('文言は仮置き。カンプは見せてよいが、カンプに描かれた文言は使わない（使うのは台本の文言）。\n'
                    + '\n'.join(content)),
        'Color palette': '、'.join(palette) if palette else '（design に色の値が無い）',
        'Constraints': (f'幅 {width}px の 1 枚。まとまりの数と順を変えない。各まとまりの型（列の数・帯・絵の置き方・区切り）を守る。'
                        '言葉で決めきれない細部（絵と文字の重なり・はみ出し・接地・区切り方）はここで描いて決める。狭い画面は描かない'),
        'Avoid': 'まとまりの追加・削除・並べ替え、台本に無い主張や数字、実在のロゴ・商標・人物写真、読めない細かい文字の詰め込み',
    }


def _refuse(out, code, message, **extra):
    out['findings'].append({'level': 'fail', 'code': code, 'message': message})
    out.update(extra)
    out['exit_code'] = 2
    return out


def comp_request(repo, deliverable, media='web'):
    repo = Path(repo)
    out = {'deliverable': deliverable, 'exit_code': 0, 'findings': []}
    checked = plan_check.check_plan(repo, deliverable, media)
    out['plan_check'] = {'exit_code': checked['exit_code'],
                         'findings': [f for f in checked['findings'] if f['level'] != 'info']}
    if checked['exit_code'] == 2:
        return _refuse(out, 'COMP_PLAN_FAILED', '組み方の計画の検査が失格。計画を直してから頼む（カンプは計画を絵にする）')
    folder = plan_check.comp_dir(repo, deliverable)
    rel = layer_paths.rel('build', deliverable, 'comp')
    ignored = plan_check.comp_ignored(repo, deliverable)
    if ignored is False:
        return _refuse(out, 'COMP_DIR_NOT_IGNORED',
                       f'{rel}/ が git に入る（カンプの画像は git に入れない）。own-web-build/templates/build/.gitignore を '
                       f'{layer_paths.rel("build", ".gitignore")} に写してから頼む')
    from design_author import read_data
    from resolved_contract import load_manifest
    plan = read_data(plan_check.plan_path(repo, deliverable))
    script = read_data(plan_check.script_path(repo, plan['script']))
    manifest = load_manifest(plan_check.manifest_path(repo, media))
    width = (manifest.get('comp') or {}).get('width')
    if isinstance(width, bool) or not isinstance(width, int) or width <= 0:
        return _refuse(out, 'COMP_SIZE_MISSING', f'解決済みマニフェストにカンプの大きさ（comp.width）が無い。design-author check（L0）をこの版で流し直す')
    fields = build_fields(script, plan, manifest, width)
    folder.mkdir(parents=True, exist_ok=True)
    fields_rel, out_rel = f'{rel}/{FIELDS}', f'{rel}/{OUTPUT}'
    (folder / FIELDS).write_text(json.dumps(fields, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    request = {'deliverable': deliverable, 'script': plan['script'], 'media': media, 'width': width,
               'fields': fields_rel, 'out': [out_rel],
               'command': ['python3', IMAGE_GEN, '--fields', fields_rel, '--out', out_rel, '--workdir', rel],
               'note': 'repo の root で上の command を流す（<skills> は skills の root。--fields・--out・--workdir は repo の root からの相対）。'
                       '生成した 1 枚をオーナーに見せ（カンプの門）、足された細部を計画の details に写して comp.reviewed を true にする'}
    (folder / REQUEST).write_text(json.dumps(request, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    out.update({'fields': fields_rel, 'out': [out_rel], 'request': f'{rel}/{REQUEST}', 'width': width})
    if ignored is None:
        out['findings'].append({'level': 'warn', 'code': 'COMP_IGNORE_NOT_CHECKED',
                                'message': f'git の作業ツリーでないので {rel}/ が git に入らないかを確かめていない'})
    levels = {f['level'] for f in out['findings']} | {f['level'] for f in out['plan_check']['findings']}
    out['exit_code'] = 1 if 'warn' in levels else 0
    return out


def main(argv=None):
    parser = argparse.ArgumentParser(description='write the comp request (own-image-gen fields) of a deliverable')
    parser.add_argument('--repo', default='.')
    parser.add_argument('--deliverable', required=True)
    parser.add_argument('--media', default='web')
    args = parser.parse_args(argv)
    try:
        result = comp_request(args.repo, args.deliverable, args.media)
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
