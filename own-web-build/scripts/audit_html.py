#!/usr/bin/env python3
"""own-web-build — HTML [自動] Validator.

media/html の type: constraint を機械検証する。LLMの自己申告に頼らず、
生成したHTMLファイルを実際に解析してPASS/FAILを出す。手順10で必ず実行する。

使い方:
    python3 scripts/audit_html.py <file.html> --manifest ux/design/.resolved/web.json [--tailwind tailwind.config.cjs]

終了コード: 全PASSなら0、FAILが1つでもあれば1。
純標準ライブラリのみ（外部依存なし・system python3で動く）。

検証範囲:
  - 完全に機械判定できる構造制約（lang / viewport / landmark / svg-label / focus / img-alt）
外部 <link rel="stylesheet"> のローカルCSSも読み、@media/focus を評価する
（実サイトのように CSS を別ファイルに分けても誤検出しない）。
コントラストはここでは判定しない。描画後の実測（own-web-build の measure_render.mjs）だけが持ち、
計算は own-design-author の contrast_kit.mjs の 1 か所にある。以前はここで CSS の文字列から静的に
も測っていて、同じ式が 2 か所にあった（静的な判定は、リテラルの色の組しか読めず、描画後の結果と食い違えた）。
出力には「コントラストは未カバー」と必ず書く。
"""
import argparse
import json
from pathlib import Path
from html.parser import HTMLParser
import os
import re
import sys


sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'own-design-author/scripts'))
from resolved_contract import load_manifest, check_tailwind, check_asset_exports
import layer_paths


class Targets(HTMLParser):
    def __init__(self):
        super().__init__()
        self.count = 0

    def handle_starttag(self, tag, attrs):
        if tag not in ('html', 'head', 'body', 'meta', 'link', 'title', 'script', 'style'):
            self.count += 1


def load_linked_css(html, base_dir):
    """<link rel="stylesheet" href="..."> のローカルCSSを読み、結合テキストを返す。
    http(s):// や // で始まる外部URLはスキップ（読めないため誤検出しない）。"""
    css_parts = []
    for tag in re.findall(r"<link\b[^>]*>", html, re.I):
        if not re.search(r'rel\s*=\s*["\']?stylesheet', tag, re.I):
            continue
        m = re.search(r'href\s*=\s*["\']([^"\']+)["\']', tag, re.I)
        if not m:
            continue
        href = m.group(1).strip()
        if re.match(r"(https?:)?//", href):
            continue  # 外部CSSは対象外
        href = href.split("?", 1)[0].split("#", 1)[0]
        path = os.path.normpath(os.path.join(base_dir, href))
        try:
            with open(path, encoding="utf-8") as f:
                css_parts.append(f.read())
        except (OSError, UnicodeDecodeError):
            continue  # 読めないCSSは黙ってスキップ（[自己点検]に委ねる）
    return "\n".join(css_parts)


# ---------- structural constraints ----------
def check_structure(html, linked_css=""):
    # html+linkedCSS を結合して評価するもの（focus / @media は外部CSSにあってもOK）
    styles = html + "\n" + linked_css
    res = []
    res.append((bool(re.search(r"<html[^>]*\blang\s*=", html, re.I)),
                "html lang属性 (html-lang)"))
    res.append((bool(re.search(r'<meta[^>]*name\s*=\s*["\']viewport["\']', html, re.I)),
                "viewport meta (html-viewport)"))
    res.append(("<main" in html.lower(), "main ランドマーク (html-semantic)"))
    res.append((bool(re.search(r"<h1[\s>]", html, re.I)), "h1 見出し (html-semantic)"))
    res.append((":focus-visible" in styles or ":focus" in styles,
                "focus 可視化 (html-keyboard)"))

    # img alt
    imgs = re.findall(r"<img\b[^>]*>", html, re.I)
    img_ok = all(re.search(r"\balt\s*=", t, re.I) for t in imgs)
    res.append((img_ok, f"img に alt ({len(imgs)}件) (html-alt)"))

    # svg accessible name: role=img+aria-label OR aria-hidden
    svgs = re.findall(r"<svg\b[^>]*>", html, re.I)
    def svg_ok(t):
        has_label = re.search(r'role\s*=\s*["\']img["\']', t, re.I) and re.search(r"aria-label\s*=", t, re.I)
        hidden = re.search(r'aria-hidden\s*=\s*["\']true["\']', t, re.I)
        return bool(has_label or hidden)
    svg_all = all(svg_ok(t) for t in svgs)
    res.append((svg_all, f"svg にアクセシブル名/aria-hidden ({len(svgs)}件) (html-svg-label)"))

    # responsive hint: at least one media query（外部CSSにあってもOK）
    res.append((bool(re.search(r"@media", styles)), "レスポンシブ @media (html-responsive)"))

    # 情報設計の印（outline がある成果物だけが持つ。D2）は属性だけで、見た目を変えない。CSS が印を選ぶと、
    # 印を付けたり外したりで見た目が変わり、印の付け替えが見た目の差分になる。
    if re.search(r"\bdata-(ia-(section|copy)|dict)\s*=", html):
        css = "\n".join(re.findall(r"<style\b[^>]*>(.*?)</style>", html, re.I | re.S)) + "\n" + linked_css
        styled = sorted(set(re.findall(r"\[\s*(data-(?:ia-[a-z-]+|dict))", css)))
        res.append((not styled, f"情報設計の印にスタイルを当てない{(' ' + str(styled)) if styled else ''} (ia-marks-inert)"))
    return res


def audit(path, manifest):
    with open(path, encoding="utf-8") as f:
        html = f.read()
    base_dir = os.path.dirname(os.path.abspath(path))
    linked_css = load_linked_css(html, base_dir)
    results = check_structure(html, linked_css)
    targets = Targets()
    targets.feed(html)
    results.append((targets.count > 0, f"HTML targets: {targets.count}"))

    print(f"\n=== audit: {Path(path).name} ===")
    failed = 0
    for ok, label in results:
        print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
        if not ok:
            failed += 1
    print("  [未カバー] コントラスト: 描画後の実測（measure_render.mjs）だけが判定する。L1 の合格は文字の読みやすさを保証しない")
    print("  [未カバー] computed palette / axe / dynamic themes: L2 required")
    for ident in manifest.get("coverage", {}).get("uncovered", []):
        print(f"  [未カバー] {ident}: no executed rule runner")
    print(f"  -> {len(results) - failed}/{len(results)} passed, {failed} failed")
    return failed == 0


def main(argv):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('files', nargs='+')
    parser.add_argument('--manifest', required=True, type=Path)
    parser.add_argument('--tailwind', action='append', type=Path, default=[])
    args = parser.parse_args(argv[1:])
    try:
        # Where the repository is comes from the manifest path text alone, before any file is opened: a repository
        # that moved to a layout this version does not read stops here, even when the manifest is missing.
        repo = layer_paths.repo_from_resolved(args.manifest)
        manifest = load_manifest(args.manifest)
        check_asset_exports(repo, manifest)
        configs = [repo / row['target'] for row in (manifest.get('tailwind_templates') or {}).get('configs', [])]
        configs.extend(args.tailwind)
        if not configs:
            skipped = ('.git', 'node_modules', Path(layer_paths.rel('resolved')).name, '.venv')
            for directory, dirs, files in os.walk(repo):
                dirs[:] = [d for d in dirs if d not in skipped]
                configs.extend(Path(directory) / name for name in files if name.startswith('tailwind.config.'))
        for config in configs:
            check_tailwind(config, manifest, repo=repo)
        all_ok = all([audit(p, manifest) for p in args.files])
    except layer_paths.LayerLayoutError as exc:
        print(json.dumps({'exit_code': 2, 'errors': [str(exc)]}, ensure_ascii=False))
        return 2
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print(json.dumps({'exit_code': 2, 'errors': [type(exc).__name__ + ': invalid or unavailable input']}, ensure_ascii=False))
        return 2
    print("RESULT:", "L1 PASS (L2 required)" if all_ok else "FAIL")
    return 0 if all_ok else 1


if __name__ == '__main__':
    sys.exit(main(sys.argv))
