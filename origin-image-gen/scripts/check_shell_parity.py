#!/usr/bin/env python3
"""check_shell_parity.py — プロンプトの器（10フィールド）の「骨」を2箇所で比較する。

段階移行のあいだ、プロンプトの器の定義は次の2箇所に**二重に**存在する:

  側1: origin-image-gen/references/prompt-shell.md
       「## 順序と各フィールドの役割」表。各行 `| <番号> | \\`<ラベル>\\` | ... |`
  側2: origin-pptx/style-guide/imagegen-prompt-convention.md
       「## 1. プロンプト構造 — ラベル付きフィールドの固定順序」直後のコードブロック。
       各行 `<ラベル>: <説明>`

どちらもラベルは出現順（＝表の番号順）に並んでいる想定。本スクリプトは**ラベル文字列の
出現順の列**だけを2箇所から抜き出し、完全一致するかを見る。

scope: 本スクリプトが見るのは器の骨（ラベルと順序）だけ。各フィールドの中身（説明文・
値の書き方など）が2箇所でズレていても、それは検出しない。

exit: 0 一致 / 1 不一致 / 2 どちらかの抽出が10件にならなかった（抽出失敗を「一致」と
      混同しない — 両側とも0件でも「一致」とは判定しない）

使い方:
  check_shell_parity.py [--shell <path>] [--convention <path>]
  （省略時は正典2ファイルの既定パスを使う）
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

SCOPE_LINE = (
    "scope: この検査は器の骨（ラベルと順序）のみを比較する。各フィールドの中身のズレは検出しない。"
)

EXPECTED_COUNT = 10

# origin-image-gen/scripts/check_shell_parity.py から見た正典パス
_SCRIPTS_DIR = Path(__file__).resolve().parent
_ORIGIN_IMAGE_GEN_DIR = _SCRIPTS_DIR.parent
_REPO_ROOT = _ORIGIN_IMAGE_GEN_DIR.parent

DEFAULT_SHELL = _ORIGIN_IMAGE_GEN_DIR / "references" / "prompt-shell.md"
DEFAULT_CONVENTION = _REPO_ROOT / "origin-pptx" / "style-guide" / "imagegen-prompt-convention.md"

SHELL_HEADING = "## 順序と各フィールドの役割"
CONVENTION_HEADING = "## 1. プロンプト構造 — ラベル付きフィールドの固定順序"

_ROW_RE = re.compile(r"^\|\s*\d+\s*\|\s*`([^`]+)`\s*\|")
_FIELD_LINE_RE = re.compile(r"^([A-Za-z][A-Za-z ]*):\s")


def _slice_section(text: str, heading: str) -> str:
    """`heading` から次の `## ` 見出しの手前までを切り出す。見出しが無ければ空文字。"""
    idx = text.find(heading)
    if idx == -1:
        return ""
    rest = text[idx + len(heading):]
    m = re.search(r"^## ", rest, re.M)
    return rest[: m.start()] if m else rest


def extract_shell_labels(path: Path) -> list[str]:
    """prompt-shell.md の表からラベルを出現順（＝物理的な行順）に抜き出す。

    表を丸ごと入れ替えるような編集ミスが起きたとき、番号セルを頼りに並べ直して
    しまうと変化を見逃す。そのため番号は使わず、行の出現順をそのまま採用する。
    """
    text = path.read_text(encoding="utf-8")
    section = _slice_section(text, SHELL_HEADING)
    labels = []
    for line in section.splitlines():
        m = _ROW_RE.match(line)
        if m:
            labels.append(m.group(1).strip())
    return labels


def extract_convention_labels(path: Path) -> list[str]:
    """imagegen-prompt-convention.md の見出し直後の最初のコードブロックからラベルを抜き出す。"""
    text = path.read_text(encoding="utf-8")
    idx = text.find(CONVENTION_HEADING)
    if idx == -1:
        return []
    rest = text[idx + len(CONVENTION_HEADING):]
    m = re.search(r"```[^\n]*\n(.*?)```", rest, re.S)
    if not m:
        return []
    block = m.group(1)
    labels = []
    for line in block.splitlines():
        m2 = _FIELD_LINE_RE.match(line)
        if m2:
            labels.append(m2.group(1).strip())
    return labels


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--shell", default=str(DEFAULT_SHELL), help="prompt-shell.md のパス")
    ap.add_argument("--convention", default=str(DEFAULT_CONVENTION),
                     help="imagegen-prompt-convention.md のパス")
    args = ap.parse_args(argv)

    print(SCOPE_LINE)

    shell_path = Path(args.shell)
    convention_path = Path(args.convention)

    if not shell_path.is_file():
        print(f"NG: --shell が存在しない: {shell_path}")
        return 2
    if not convention_path.is_file():
        print(f"NG: --convention が存在しない: {convention_path}")
        return 2

    shell_labels = extract_shell_labels(shell_path)
    convention_labels = extract_convention_labels(convention_path)

    # 抽出の失敗（10件に満たない）を「一致」と混同しない。両方0件でも一致とはしない。
    if len(shell_labels) != EXPECTED_COUNT or len(convention_labels) != EXPECTED_COUNT:
        print(
            "NG（抽出失敗）: 期待件数="
            f"{EXPECTED_COUNT} / shell={len(shell_labels)}件 {shell_labels} "
            f"/ convention={len(convention_labels)}件 {convention_labels}"
        )
        return 2

    if shell_labels == convention_labels:
        print(f"OK: 10フィールドのラベルと順序が一致 {shell_labels}")
        return 0

    print("NG（不一致）: index (shell側 / convention側)")
    max_len = max(len(shell_labels), len(convention_labels))
    for i in range(max_len):
        left = shell_labels[i] if i < len(shell_labels) else "(欠落)"
        right = convention_labels[i] if i < len(convention_labels) else "(欠落)"
        mark = "  " if left == right else "!="
        print(f"  [{i}] {mark} {left!r} / {right!r}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
