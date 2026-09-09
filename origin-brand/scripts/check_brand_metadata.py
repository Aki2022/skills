#!/usr/bin/env python3
"""check_brand_metadata.py — origin-brand/plugins/*/plugin.yaml の等級（tier）と
取得メタ（source）の形を検査する。

決定8 により、各ブランド倉庫の箱（plugin.yaml）は次を満たさなければならない:

  1. tier が存在し、`owned` か `observed` のどちらか（`licensed` は廃止済み）。
  2. source が存在し、upstream / upstream_version / fetched_at の3キーすべてを持つ。
  3. tier: owned なら upstream: none かつ fetched_at: n/a（自社資産に取得日は無い）。
  4. tier: observed なら upstream が http:// か https:// で始まる。
  5. upstream_version は「具体的な版」または unversioned / unrecorded / n/a。空文字は違反。
  6. fetched_at は YYYY-MM-DD 形式か n/a。

scope: この検査は等級（tier）と取得メタ（source）の形だけを見る。上流の内容が実際に
新しいかは見ない。

exit: 0 全箱が合格 / 1 いずれかの箱が違反 / 2 --plugins-dir が存在しない、または
      plugin.yaml を持つ箱が0件（対象0件を「全部合格」にしない）。

決定的な制約: 既定の python3 に pyyaml は入っていない。plugin.yaml は
`id: ... / tier: ... / source: <nested>` という最大2段ネストの形しか使わないため、
本スクリプトは YAML 全体ではなくこの形だけを行単位で自前解析する。

使い方:
  check_brand_metadata.py [--plugins-dir <path>]
  （省略時は origin-brand/plugins を既定値として使う）
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Optional

SCOPE_LINE = (
    "scope: この検査は等級（tier）と取得メタ（source）の形だけを見る。"
    "上流の内容が実際に新しいかは見ない。"
)

VALID_TIERS = ("owned", "observed")
REQUIRED_SOURCE_KEYS = ("upstream", "upstream_version", "fetched_at")
_GENERIC_VERSION_TOKENS = {"unversioned", "unrecorded", "n/a"}
_FETCHED_AT_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# origin-brand/scripts/check_brand_metadata.py から見た正典パス
_SCRIPTS_DIR = Path(__file__).resolve().parent
_ORIGIN_BRAND_DIR = _SCRIPTS_DIR.parent
DEFAULT_PLUGINS_DIR = _ORIGIN_BRAND_DIR / "plugins"

_TOP_LEVEL_KEY_RE = re.compile(r"^source:\s*(.*)$")
_NESTED_KEY_RE = re.compile(r"^[ \t]+([A-Za-z0-9_]+):\s*(.*)$")


def _strip_comment(value: str) -> str:
    """`value  # comment` の末尾コメントを落として前後の引用符・空白も剥がす。"""
    idx = value.find(" #")
    if idx != -1:
        value = value[:idx]
    elif value.startswith("#"):
        value = ""
    return value.strip().strip('"').strip("'")


def extract_top_level_scalar(lines: list[str], key: str) -> Optional[str]:
    """トップレベル（行頭に空白の無い）`key: value` のスカラー値を返す。"""
    pattern = re.compile(r"^" + re.escape(key) + r":\s*(.*)$")
    for line in lines:
        m = pattern.match(line)
        if m:
            return _strip_comment(m.group(1))
    return None


def extract_source_block(lines: list[str]) -> Optional[dict[str, str]]:
    """トップレベル `source:` 直下（1段インデント）のマッピングを抜き出す。

    ネストは2段までという契約（id/tier/source.upstream 等）以上を汎用に解釈しようと
    せず、`source:` の次の非空行から、インデントが外れる（トップレベルに戻る）行の
    手前までだけを見る。
    """
    source_start = None
    for i, line in enumerate(lines):
        m = _TOP_LEVEL_KEY_RE.match(line)
        if m:
            val = _strip_comment(m.group(1))
            if val:
                # `source: something` という形（マッピングではなくスカラー）は
                # この契約が想定する形ではない。無いのと同様に扱う。
                return None
            source_start = i + 1
            break
    if source_start is None:
        return None

    result: dict[str, str] = {}
    for line in lines[source_start:]:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if not line[:1].isspace():
            break  # インデントが外れた = source ブロックの終わり
        m2 = _NESTED_KEY_RE.match(line)
        if m2:
            result[m2.group(1)] = _strip_comment(m2.group(2))
    return result


def validate_box(box_name: str, plugin_yaml: Path) -> list[str]:
    text = plugin_yaml.read_text(encoding="utf-8")
    lines = text.splitlines()
    violations: list[str] = []

    tier = extract_top_level_scalar(lines, "tier")
    if tier is None:
        violations.append(f"{box_name}: tier キーが無い")
    elif tier not in VALID_TIERS:
        violations.append(
            f"{box_name}: tier が owned/observed のどちらでもない (tier: {tier!r})"
        )

    source = extract_source_block(lines)
    if source is None:
        violations.append(f"{box_name}: source キーが無い（またはマッピング形式でない）")
        source = {}

    missing_keys = [k for k in REQUIRED_SOURCE_KEYS if k not in source]
    if missing_keys:
        violations.append(f"{box_name}: source に必須キーが無い: {missing_keys}")

    if tier == "owned":
        upstream = source.get("upstream")
        if upstream is not None and upstream != "none":
            violations.append(
                f"{box_name}: tier: owned なのに source.upstream が 'none' でない "
                f"(upstream: {upstream!r})"
            )
        fetched_at = source.get("fetched_at")
        if fetched_at is not None and fetched_at != "n/a":
            violations.append(
                f"{box_name}: tier: owned なのに source.fetched_at が 'n/a' でない "
                f"(fetched_at: {fetched_at!r})"
            )
    elif tier == "observed":
        upstream = source.get("upstream", "")
        if not (upstream.startswith("http://") or upstream.startswith("https://")):
            violations.append(
                f"{box_name}: tier: observed なのに source.upstream が http(s):// で"
                f"始まらない (upstream: {upstream!r})"
            )

    if "upstream_version" in source:
        uv = source["upstream_version"]
        if uv == "":
            violations.append(f"{box_name}: source.upstream_version が空文字")

    if "fetched_at" in source:
        fa = source["fetched_at"]
        if fa != "n/a" and not _FETCHED_AT_RE.match(fa):
            violations.append(
                f"{box_name}: source.fetched_at が YYYY-MM-DD でも n/a でもない "
                f"(fetched_at: {fa!r})"
            )

    return violations


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument(
        "--plugins-dir",
        default=str(DEFAULT_PLUGINS_DIR),
        help="箱（<box>/plugin.yaml）を並べたディレクトリ（既定: origin-brand/plugins）",
    )
    args = ap.parse_args(argv)

    print(SCOPE_LINE)

    plugins_dir = Path(args.plugins_dir)
    if not plugins_dir.is_dir():
        print(f"NG: --plugins-dir が存在しない: {plugins_dir}")
        return 2

    box_dirs = sorted(
        d for d in plugins_dir.iterdir() if d.is_dir() and (d / "plugin.yaml").is_file()
    )
    if not box_dirs:
        print(f"NG（対象0件）: {plugins_dir} 配下に plugin.yaml を持つ箱が無い")
        return 2

    all_violations: list[str] = []
    for box_dir in box_dirs:
        all_violations.extend(validate_box(box_dir.name, box_dir / "plugin.yaml"))

    if all_violations:
        print(f"NG: {len(all_violations)}件の違反 ({len(box_dirs)}箱中)")
        for v in all_violations:
            print(f"  - {v}")
        return 1

    print(f"OK: {len(box_dirs)}箱すべて合格 ({[d.name for d in box_dirs]})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
