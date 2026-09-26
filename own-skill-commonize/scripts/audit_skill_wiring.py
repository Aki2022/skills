#!/usr/bin/env python3
"""配線先を**正典の側から辿って**監査する。台帳を持たない。

なぜ台帳を持たないか
--------------------
2026-09-22 に配布先の一覧をファイル（alias-roots.txt）に置いた。人の記憶よりは
強いが、**人が書くものは書き落とす**。実測で一覧は 6 件だったが、正典を指す配線先は
**13 件**あった。落ちていたのは ~/.claude-private/skills・~/.claude-seat2/skills・
~/.gemini/antigravity-cli/skills の3経路（いずれも生きていた）と、古いバックアップ4件。
一覧方式は最初から半分しか見ていなかった。

代わりに「**誰が正典を指しているか**」を辿る。配線は必ず正典への symlink なので、
正典を指す symlink を探せば配線先は自分で名乗る。新しい道具を入れて配線した瞬間から
対象になる。一覧に足す作業が要らない。

見るもの（人間の指定した観点）
------------------------------
1. 正典そのものが正しいか        → skill_lint の S1-S9 が持つ。ここでは見ない
2. 配線先と正典の skill 個数が同じ
3. symlink が網羅的に配線されている（不足が無い）
4. ファイル名が同じ（余分が無い）
5. 個別対処した skill は配線先でも解決するか

中身は照合しない。配線は正典の同じ実体を指すので原理的にズレない
（72 skill × 5 root = 360 通りを realpath で照合し別実体 0 件・2026-09-22）。
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from skill_catalog import load_skill_catalog

# 正典直下の、skill ではない管理ディレクトリ
NON_SKILL = {".git", "docs", "synced", "node_modules"}
# 配線先直下の、その製品が自分で書くもの（比較から外す）
PRODUCT_OWNED = {".system"}
# 退役した配線先。名前で分かるものは監査対象にしない
RETIRED_MARKERS = ("_backup_", "_old_", ".bak_", ".orphaned_", ".disabled")


def canon_skills(canon: Path) -> set[str]:
    return load_skill_catalog(canon).active


def is_retired(path: Path) -> bool:
    return any(m in path.name for m in RETIRED_MARKERS)


def discover(canon: Path, search_roots: list[Path], max_depth: int) -> dict[Path, str]:
    """正典を指す symlink から配線先を見つける。

    戻り値は {配線先ディレクトリ: "whole" | "per-skill"}。
    whole は「棚ごと1本」(root 自体が正典への symlink)、
    per-skill は「1冊ずつ」(root の中に正典内を指す symlink が並ぶ)。
    """
    found: dict[Path, str] = {}
    for base in search_roots:
        if not base.is_dir():
            continue
        for path in _walk(base, max_depth):
            try:
                if not path.is_symlink():
                    continue
                target = path.resolve()
            except (OSError, RuntimeError):
                continue
            if target == canon:
                found.setdefault(path, "whole")
            elif canon in target.parents:
                found.setdefault(path.parent, "per-skill")
    return {p: k for p, k in found.items() if not is_retired(p)}


def _walk(base: Path, max_depth: int):
    """深さを区切って走査する。symlink は辿らない（ループを避ける）。"""
    base_depth = len(base.parts)
    for dirpath, dirnames, filenames in os.walk(base, followlinks=False):
        d = Path(dirpath)
        if len(d.parts) - base_depth >= max_depth:
            dirnames[:] = []
        for name in list(dirnames) + filenames:
            yield d / name


def audit(
    canon: Path,
    wiring: dict[Path, str],
    base: set[str] | None = None,
) -> tuple[list[str], list[str], dict]:
    base = canon_skills(canon) if base is None else base
    failures: list[str] = []
    notes: list[str] = []
    coverage = {"roots": len(wiring), "whole": 0, "per_skill": 0}

    for root in sorted(wiring):
        kind = wiring[root]
        shown = str(root).replace(str(Path.home()), "~")
        if kind == "whole":
            coverage["whole"] += 1
            notes.append(f"OK  {shown}: 棚ごと1本（正典そのもの・{len(base)} 件）")
            continue

        coverage["per_skill"] += 1
        names = {
            n for n in os.listdir(root)
            if n not in PRODUCT_OWNED and not n.startswith(".")
        }
        missing = sorted(base - names)
        extra = sorted(names - base)
        broken = sorted(n for n in names if not (root / n).exists())
        # 名前だけを見ると、正典と同じ名前で実体を置かれたときに素通りする。
        # Codex の $skill-installer は $CODEX_HOME/skills/<名前> へ実体を書き込むので、
        # 正典にある名前を上書きされるとその skill だけ Codex 専用になり、
        # Claude / Gemini とズレたまま「すべて一致」と報告される（2026-09-23 に実測）。
        detached = sorted(n for n in names if not (root / n).is_symlink())
        # 正典の外を指す symlink も同じ結果になる。
        outside = sorted(
            n for n in names
            if (root / n).is_symlink()
            and canon.resolve() not in Path(os.path.realpath(root / n)).parents
        )

        if len(names) != len(base):
            failures.append(
                f"{shown}: 個数が違う（正典 {len(base)} / 配線先 {len(names)}）")
        for n in missing:
            failures.append(f"{shown}: 配線が無い: {n}")
        for n in extra:
            failures.append(f"{shown}: 正典に無い名前: {n}")
        for n in broken:
            failures.append(f"{shown}: 宙を指している: {n}")
        for n in detached:
            failures.append(
                f"{shown}: symlink でない実体が置かれている"
                f"（正典から切り離され、このツール専用になっている）: {n}")
        for n in outside:
            failures.append(
                f"{shown}: 正典の外を指している: {n} → {os.readlink(root / n)}")
        if not (missing or extra or broken or detached or outside) \
                and len(names) == len(base):
            notes.append(f"OK  {shown}: 1冊ずつ {len(names)} 件すべて一致")

    return failures, notes, coverage


def audit_individual(canon: Path) -> list[str]:
    """個別対処した skill（正典の外を指す symlink）が解決するか。"""
    out = []
    for p in sorted(canon.iterdir()):
        if not p.is_symlink():
            continue
        shown = str(p).replace(str(Path.home()), "~")
        if p.exists():
            out.append(f"OK  {shown}: 個別対処・解決する → {os.readlink(p)}")
        else:
            out.append(f"FAIL {shown}: 個別対処・宙を指している → {os.readlink(p)}")
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--canonical", type=Path, default=Path.home() / ".agents/skills")
    ap.add_argument("--search", type=Path, action="append", default=[],
                    help="配線先を探す起点。既定はホーム直下の隠しディレクトリ全部")
    ap.add_argument("--max-depth", type=int, default=3)
    args = ap.parse_args(argv)

    canon = args.canonical.expanduser().resolve()
    if not canon.is_dir():
        print(f"ERROR 正典が無い: {canon}", file=sys.stderr)
        return 2
    catalog = load_skill_catalog(canon)
    if catalog.errors:
        for error in catalog.errors:
            print(f"FAIL canonical catalog: {error}")
        print(f"RESULT: FAIL ({len(catalog.errors)} canonical catalog errors)")
        return 2
    if not catalog.active:
        # 対象0件を黙って通さない
        print(f"ERROR 正典に skill が1件も無い: {canon}", file=sys.stderr)
        return 2
    if catalog.retired_physical:
        print(
            "NOTE retired physical skills excluded from active set: "
            + ", ".join(sorted(catalog.retired_physical))
        )

    search = [p.expanduser() for p in args.search] or [
        d for d in Path.home().glob(".*") if d.is_dir() and not d.is_symlink()
    ]

    wiring = discover(canon, search, args.max_depth)
    failures, notes, cov = audit(canon, wiring, catalog.active)
    individual = audit_individual(canon)

    for line in notes + [l for l in individual if l.startswith("OK")]:
        print(line)
    for line in failures + [l for l in individual if l.startswith("FAIL")]:
        print(line)

    # 何を見て何を見なかったかを**常に**言う。
    # 緑と「何も見ていない」を出力で区別できるようにするため。
    print(f"coverage: 配線先 {cov['roots']} 件を検査"
          f"（棚ごと {cov['whole']} / 1冊ずつ {cov['per_skill']}）"
          f" 正典 {len(catalog.active)} 件 / 個別対処 {len(individual)} 件")
    if not wiring:
        # 配線0件には2つの意味がある。区別しないと空振りを通す。
        #   A まだ配線していない正典（新しい環境・差し替えた $HOME の fixture）→ 正常
        #   B 探索範囲の指定を誤った                                            → 誤り
        # 探索起点が1つも実在しなければ B。起点は実在するが配線が無ければ A。
        missing = [p for p in search if not p.is_dir()]
        if len(missing) == len(search):
            print("ERROR 探索起点が1つも実在しない（--search の指定を確認する): "
                  + ", ".join(str(p) for p in missing), file=sys.stderr)
            return 2
        print("NOTE 配線先が1件も見つからない（この正典はまだどこからも配線されていない）")
        print("RESULT: SKIP")
        return 0

    bad = len(failures) + sum(1 for l in individual if l.startswith("FAIL"))
    print("RESULT: OK" if bad == 0 else f"RESULT: FAIL ({bad})")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
