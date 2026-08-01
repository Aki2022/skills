#!/usr/bin/env python3
"""origin-design-runtime — HTML [自動] Validator.

media/html の type: constraint を機械検証する。LLMの自己申告に頼らず、
生成したHTMLファイルを実際に解析してPASS/FAILを出す。手順10で必ず実行する。

使い方:
    python3 scripts/audit_html.py <file.html> [file2.html ...]

終了コード: 全PASSなら0、FAILが1つでもあれば1。
純標準ライブラリのみ（外部依存なし・system python3で動く）。

検証範囲:
  - 完全に機械判定できる構造制約（lang / viewport / landmark / svg-label / focus / img-alt）
  - 判定可能な範囲のコントラスト（同一ルール内 color+background、CSS変数解決を含む）
外部 <link rel="stylesheet"> のローカルCSSも読み、@media/focus/リテラルcolor+background を評価する
（実サイトのように CSS を別ファイルに分けても誤検出しない）。
コントラストは画像上テキスト・グラデーション・CSS変数のテーマ切替(data-theme)・JS動的スタイルは判定不能。
その分は [自己点検]（validators/common/contrast.md）に委ねる旨を出力する。
"""
import os
import re
import sys


# ---------- contrast ----------
def _srgb_to_lin(c):
    c = c / 255.0
    return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4


def _luminance(rgb):
    r, g, b = rgb
    return 0.2126 * _srgb_to_lin(r) + 0.7152 * _srgb_to_lin(g) + 0.0722 * _srgb_to_lin(b)


def contrast_ratio(fg, bg):
    l1, l2 = _luminance(fg), _luminance(bg)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


_NAMED = {
    "white": (255, 255, 255), "black": (0, 0, 0), "red": (255, 0, 0),
    "gray": (128, 128, 128), "grey": (128, 128, 128), "silver": (192, 192, 192),
    "transparent": None,
}


def parse_color(v):
    """CSS色文字列を (r,g,b) に。解決できなければ None。"""
    if v is None:
        return None
    v = v.strip().lower()
    if v in _NAMED:
        return _NAMED[v]
    m = re.fullmatch(r"#([0-9a-f]{3})", v)
    if m:
        h = m.group(1)
        return tuple(int(c * 2, 16) for c in h)
    m = re.fullmatch(r"#([0-9a-f]{6})", v)
    if m:
        h = m.group(1)
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    # rgb()/rgba(): アルファが 1 未満なら背景合成が必要で判定不能 → None
    m = re.match(r"rgba?\(\s*([0-9]+)\s*,\s*([0-9]+)\s*,\s*([0-9]+)\s*(?:,\s*([0-9.]+)\s*)?\)", v)
    if m:
        if m.group(4) is not None and float(m.group(4)) < 1.0:
            return None  # 半透明は合成なしに判定できない（[自己点検]へ）
        return tuple(int(m.group(i)) for i in (1, 2, 3))
    return None


def resolve_vars(text):
    """基底 :root {} の --var: value のみを集め、var(--x) を解決した文字列を返す（1段）。

    テーマ上書き（:root[data-theme=...]{}）の変数は集めない。集めると最後に定義された
    テーマ値が勝ち、別テーマの背景と組み合わさった「実在しないペア」を誤検出するため。
    テーマ切替時のコントラストは判定不能として [自己点検]（validators/common/contrast.md）に委ねる。
    基底 :root が無い場合は、後方互換で全 --var を集める（単一テーマのHTML断片向け）。
    """
    base_blocks = re.findall(r":root\s*\{([^{}]*)\}", text)
    if base_blocks:
        varmap = {}
        for block in base_blocks:
            for k, val in re.findall(r"(--[\w-]+)\s*:\s*([^;{}]+)", block):
                varmap[k] = val.strip()
    else:
        varmap = dict(re.findall(r"(--[\w-]+)\s*:\s*([^;{}]+)", text))

    def sub(m):
        name = m.group(1).strip()
        fallback = m.group(2)
        if name in varmap:
            return varmap[name].strip()
        return (fallback or "").strip()

    return re.sub(r"var\(\s*(--[\w-]+)\s*(?:,\s*([^)]*))?\)", sub, text)


def literal_color_pairs(css_text, source="css"):
    """CSSテキストのルール単位で color+background の両指定を探し、解決可能なら判定を返す。
    <style>ブロックでも外部CSSでも共通で使う。"""
    results = []
    resolved = resolve_vars(css_text)
    for block in re.findall(r"\{([^{}]*)\}", resolved):
        fg = re.search(r"(?<!-)\bcolor\s*:\s*([^;]+)", block)
        bg = re.search(r"background(?:-color)?\s*:\s*([^;{]+)", block)
        if fg and bg:
            fgc = parse_color(fg.group(1))
            bgc = parse_color(bg.group(1))
            if fgc and bgc:
                r = round(contrast_ratio(fgc, bgc), 2)
                results.append((r >= 4.5,
                                f"[{source}] contrast {fg.group(1).strip()} on {bg.group(1).strip()} = {r}:1 (min 4.5)"))
    return results


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


def check_contrast(html, linked_css=""):
    """同一CSSルール内で color と background が両方指定された箇所を判定。
    <style>ブロック・インライン style・外部リンクCSS を対象にする。"""
    results = []
    # <style> ブロック
    for style in re.findall(r"<style[^>]*>(.*?)</style>", html, re.S | re.I):
        results += literal_color_pairs(style, source="style")
    # インライン style="color:..;background:.."
    resolved = resolve_vars(html)
    for st in re.findall(r'style\s*=\s*"([^"]*)"', resolved):
        fg = re.search(r"(?<!-)\bcolor\s*:\s*([^;]+)", st)
        bg = re.search(r"background(?:-color)?\s*:\s*([^;]+)", st)
        if fg and bg:
            fgc, bgc = parse_color(fg.group(1)), parse_color(bg.group(1))
            if fgc and bgc:
                r = round(contrast_ratio(fgc, bgc), 2)
                results.append((r >= 4.5, f"[inline] contrast = {r}:1 (min 4.5)"))
    # 外部リンクCSS（:root のCSS変数も resolve_vars で解決）
    if linked_css:
        results += literal_color_pairs(linked_css, source="linked-css")
    return results


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
    return res


def audit(path):
    with open(path, encoding="utf-8") as f:
        html = f.read()
    base_dir = os.path.dirname(os.path.abspath(path))
    linked_css = load_linked_css(html, base_dir)
    results = check_structure(html, linked_css)
    contrast = check_contrast(html, linked_css)
    results += contrast

    print(f"\n=== audit: {path} ===")
    failed = 0
    for ok, label in results:
        print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
        if not ok:
            failed += 1
    if not contrast:
        print("  [note] 判定可能なcolor+backgroundペアなし。コントラストは[自己点検]で確認すること")
    print(f"  -> {len(results) - failed}/{len(results)} passed, {failed} failed")
    return failed == 0


def main(argv):
    if len(argv) < 2:
        print("usage: python3 audit_html.py <file.html> [more.html ...]", file=sys.stderr)
        return 2
    all_ok = True
    for p in argv[1:]:
        try:
            if not audit(p):
                all_ok = False
        except FileNotFoundError:
            print(f"  [ERROR] file not found: {p}", file=sys.stderr)
            all_ok = False
    print("\nRESULT:", "ALL PASS" if all_ok else "FAIL (手順11で修正)")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
