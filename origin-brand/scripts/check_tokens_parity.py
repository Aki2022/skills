#!/usr/bin/env python3
"""check_tokens_parity.py — 倉庫 my_company の pptx 媒体プロファイル ↔
origin-pptx/style-guide/tokens.json の現行値が同値かを検査する。

決定6 により、倉庫 my_company の pptx 媒体プロファイルは origin-pptx の tokens.json
現行値と同値でなければならない（将来 pptx が倉庫を参照しても出力が変わらないことの
保証）。比較するのは下表18項目だけ。

scope: この検査は上表18項目の値が一致するかだけを見る。web プロファイルと dataViz は
比較しない（媒体が違うので同値である必要がない）。

exit: 0 18項目すべて一致 / 1 いずれかの値が不一致 / 2 どちらかのパスが辿れない
      （キーが無い）、または比較できた件数が18件に満たない（欠落を「一致」と混同しない）。

使い方:
  check_tokens_parity.py [--warehouse-color <path>] [--warehouse-typography <path>]
                          [--pptx-tokens <path>]
  （省略時は正典3ファイルの既定パスを使う）
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

SCOPE_LINE = (
    "scope: この検査は上表18項目の値が一致するかだけを見る。"
    "web プロファイルと dataViz は比較しない（媒体が違うので同値である必要がない）。"
)

# origin-brand/scripts/check_tokens_parity.py から見た正典パス
_SCRIPTS_DIR = Path(__file__).resolve().parent
_ORIGIN_BRAND_DIR = _SCRIPTS_DIR.parent
_REPO_ROOT = _ORIGIN_BRAND_DIR.parent

DEFAULT_WAREHOUSE_COLOR = _ORIGIN_BRAND_DIR / "plugins" / "my_company" / "tokens" / "color.json"
DEFAULT_WAREHOUSE_TYPOGRAPHY = (
    _ORIGIN_BRAND_DIR / "plugins" / "my_company" / "tokens" / "typography.json"
)
DEFAULT_PPTX_TOKENS = _REPO_ROOT / "origin-pptx" / "style-guide" / "tokens.json"

# (表示名, 倉庫側ファイル種別, 倉庫側パス, origin-pptx側パス)
# これだけを比べる。増やさない。
MAPPINGS: list[tuple[str, str, str, str]] = [
    ("color: text.body", "color", "mediaProfiles.pptx.text.body", "color.gray.body.value"),
    ("color: text.heading", "color", "mediaProfiles.pptx.text.heading", "color.gray.heading.value"),
    ("color: shared.gray.panel", "color", "shared.gray.panel", "color.gray.panel.value"),
    ("color: shared.gray.footer", "color", "shared.gray.footer", "color.gray.footer.value"),
    ("color: shared.gray.copyright", "color", "shared.gray.copyright", "color.gray.copyright.value"),
    ("color: shared.tone.neutral", "color", "shared.tone.neutral", "color.message.neutral.value"),
    ("color: shared.tone.positive", "color", "shared.tone.positive", "color.message.positive.value"),
    ("color: shared.tone.negative", "color", "shared.tone.negative", "color.message.negative.value"),
    ("color: shared.background.default", "color", "shared.background.default", "color.background.default.value"),
    ("color: shared.background.border", "color", "shared.background.border", "color.background.border.value"),
    ("color: contrastPolicy.bodyMin", "color", "mediaProfiles.pptx.contrastPolicy.bodyMin", "color.contrastPolicy.bodyMin"),
    ("typography: fontFamily.sans", "typography", "mediaProfiles.pptx.fontFamily.sans", "typography.fontFamily.value"),
    ("typography: scale.title.sizePt", "typography", "mediaProfiles.pptx.scale.title.sizePt", "typography.scale.title.sizePt"),
    ("typography: scale.keyMessage.sizePt", "typography", "mediaProfiles.pptx.scale.keyMessage.sizePt", "typography.scale.keyMessage.sizePt"),
    ("typography: scale.heading.sizePt", "typography", "mediaProfiles.pptx.scale.heading.sizePt", "typography.scale.heading.sizePt"),
    ("typography: scale.body.sizePt", "typography", "mediaProfiles.pptx.scale.body.sizePt", "typography.scale.body.sizePt"),
    ("typography: scale.label.sizePt", "typography", "mediaProfiles.pptx.scale.label.sizePt", "typography.scale.label.sizePt"),
    ("typography: policy.bodyMinPt", "typography", "mediaProfiles.pptx.policy.bodyMinPt", "typography.policy.bodyMinPt"),
]
EXPECTED_COUNT = 18
assert len(MAPPINGS) == EXPECTED_COUNT  # 表がずれたらこの検査自体が壊れる


class _NotFound:
    """辞書に無い場合の番兵。None も正当な値でありうるため区別する。"""


_NOT_FOUND = _NotFound()


def resolve(obj: object, dotted_path: str):
    """`a.b.c` のドット区切りパスをたどる。辿れなければ _NOT_FOUND を返す。"""
    cur = obj
    for part in dotted_path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return _NOT_FOUND
        cur = cur[part]
    return cur


def _load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--warehouse-color", default=str(DEFAULT_WAREHOUSE_COLOR),
        help="倉庫 my_company の tokens/color.json のパス",
    )
    ap.add_argument(
        "--warehouse-typography", default=str(DEFAULT_WAREHOUSE_TYPOGRAPHY),
        help="倉庫 my_company の tokens/typography.json のパス",
    )
    ap.add_argument(
        "--pptx-tokens", default=str(DEFAULT_PPTX_TOKENS),
        help="origin-pptx/style-guide/tokens.json のパス",
    )
    args = ap.parse_args(argv)

    print(SCOPE_LINE)

    color_path = Path(args.warehouse_color)
    typography_path = Path(args.warehouse_typography)
    pptx_path = Path(args.pptx_tokens)

    for label, path in (
        ("--warehouse-color", color_path),
        ("--warehouse-typography", typography_path),
        ("--pptx-tokens", pptx_path),
    ):
        if not path.is_file():
            print(f"NG: {label} が存在しない: {path}")
            return 2

    try:
        warehouse = {"color": _load_json(color_path), "typography": _load_json(typography_path)}
        pptx_tokens = _load_json(pptx_path)
    except json.JSONDecodeError as exc:
        print(f"NG: JSON として読めない: {exc}")
        return 2

    compared = 0
    missing: list[str] = []
    mismatches: list[tuple[str, object, object]] = []

    for label, file_key, warehouse_path, pptx_dotted_path in MAPPINGS:
        w_val = resolve(warehouse[file_key], warehouse_path)
        p_val = resolve(pptx_tokens, pptx_dotted_path)
        w_found = w_val is not _NOT_FOUND
        p_found = p_val is not _NOT_FOUND

        if not (w_found and p_found):
            side = []
            if not w_found:
                side.append(f"倉庫側 {warehouse_path} が無い")
            if not p_found:
                side.append(f"pptx側 {pptx_dotted_path} が無い")
            missing.append(f"{label}: {'; '.join(side)}")
            print(f"MISSING {label}: {'; '.join(side)}")
            continue

        compared += 1
        if w_val != p_val:
            mismatches.append((label, w_val, p_val))
            print(f"NG   {label}: 倉庫={w_val!r} / pptx={p_val!r}")
        else:
            print(f"OK   {label}: {w_val!r}")

    print(f"compared: {compared}/{EXPECTED_COUNT}")

    if compared != EXPECTED_COUNT:
        print(
            f"NG（比較失敗）: 期待件数={EXPECTED_COUNT} に対し比較できたのは {compared} 件。"
            "欠落（キーが辿れない）を一致とは判定しない。"
        )
        return 2

    if mismatches:
        print(f"NG（不一致）: {len(mismatches)}件")
        return 1

    print(f"OK: {EXPECTED_COUNT}項目すべて一致")
    return 0


if __name__ == "__main__":
    sys.exit(main())
