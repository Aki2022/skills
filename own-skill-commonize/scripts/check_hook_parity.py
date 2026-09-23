#!/usr/bin/env python3
"""hook の参照が全ツール・全席で揃っているかを検査する。

なぜ要るか: hook の **script 実体**は `~/.agents/hooks/` などに1つ置いて共有できるが、
**「どの設定ファイルがそれを呼ぶか」は共通化できない**。Claude の `settings.json` は
permission や environment という account 固有の可変 state を含むため、丸ごと symlink に
できない（commonize の不変条件6）。結果、参照だけがツールごとに手書きで残る。

この非対称が 2026-09-21 の `origin-` → `own-` 改名で実害を出した。
`~/.claude/settings.json` は直したが `~/.claude-seat2/settings.json` を取りこぼし、
seat2 の guard hook と session_start hook が2日間死んでいた（`rc=127 No such file`）。
さらに `stop_nudge.sh` は main と Codex にあるのに seat2 だけ落ちていた。
どちらも、どこにも赤が出ないまま動き続けた。

検査は2つ。

1. **実在検査** — 設定に書かれた script が実在するか。死んだ hook は黙って無効になる。
2. **一致検査** — 同じ製品の席どうしで、呼んでいる script の集合が同じか。

**引数は比較しない。** `aiphetamine_hook.py --account-name main` と `--account-name alias`
のように、account 識別子を引数で渡す hook は席ごとに違って正しい。script のパスだけを
見れば、正当な差を誤検知せずに「片方にしか無い hook」を捕まえられる。

製品をまたぐ比較はしない。Codex と Gemini はイベント名もスキーマも違い、
揃っていないことが正常だから。
"""
from __future__ import annotations

import json
import os
import shlex
import sys
from pathlib import Path

# 同じ製品の席どうしは一致していなければならない。
PEER_GROUPS: dict[str, list[str]] = {
    "Claude": ["~/.claude/settings.json",
               "~/.claude-seat2/settings.json",
               "~/.claude-private/settings.json"],
}
# 席をまたいで共有される（＝一致検査の対象外の）単独ファイル。
STANDALONE: list[str] = [
    "~/.agents/codex/hooks.json",       # Codex 3席が symlink で共有
    "~/.gemini/config/hooks.json",      # Antigravity。スキーマが別
]


def _shown(p: str) -> str:
    return p.replace(str(Path.home()), "~")


def _expand(cmd: str) -> str:
    return cmd.replace("${HOME}", str(Path.home())).replace("$HOME", str(Path.home()))


def _scripts(cmd: str) -> list[str]:
    """コマンド文字列から、実在を確かめるべき script のパスを取り出す。"""
    out = []
    try:
        toks = shlex.split(_expand(cmd))
    except ValueError:
        return out
    for t in toks:
        if "/" in t and t.endswith((".sh", ".py")):
            out.append(t)
    return out


def _collect(path: str) -> dict[tuple[str, str], str] | None:
    """{(イベント, script のファイル名): 絶対パス} を返す。見つからなければ None。"""
    f = os.path.expanduser(path)
    if not os.path.exists(f):
        return None
    d = json.load(open(f))
    out: dict[tuple[str, str], str] = {}

    def add(ev: str, groups):
        for g in groups:
            for h in g.get("hooks", []):
                for s in _scripts(h.get("command", "")):
                    out[(ev, os.path.basename(s))] = s

    if "hooks" in d and isinstance(d["hooks"], dict):     # Claude / Codex
        for ev, groups in d["hooks"].items():
            add(ev, groups)
    else:                                                  # Antigravity: 名前付きグループ
        for _name, evs in d.items():
            if isinstance(evs, dict):
                for ev, groups in evs.items():
                    add(ev, groups)
    return out


def main(argv: list[str]) -> int:
    failures: list[str] = []
    checked_files = 0
    checked_hooks = 0

    def existence(path: str, hooks: dict[tuple[str, str], str]) -> None:
        nonlocal checked_hooks
        checked_hooks += len(hooks)
        for (ev, name), abspath in sorted(hooks.items()):
            if not os.path.exists(abspath):
                failures.append(
                    f"{_shown(path)} [{ev}]: 参照先が実在しない → {_shown(abspath)}")

    for product, paths in PEER_GROUPS.items():
        present: dict[str, dict[tuple[str, str], str]] = {}
        for p in paths:
            h = _collect(p)
            if h is None:
                continue
            present[p] = h
            checked_files += 1
            existence(p, h)
        if len(present) < 2:
            continue
        union: set[tuple[str, str]] = set()
        for h in present.values():
            union |= set(h)
        for key in sorted(union):
            missing = [p for p, h in present.items() if key not in h]
            if missing:
                have = [p for p, h in present.items() if key in h]
                failures.append(
                    f"{product}: [{key[0]}] {key[1]} が "
                    f"{', '.join(_shown(m) for m in missing)} に無い"
                    f"（{', '.join(_shown(x) for x in have)} にはある）")

    for p in STANDALONE:
        h = _collect(p)
        if h is None:
            continue
        checked_files += 1
        existence(p, h)

    if checked_files == 0:
        print("ERROR 検査対象の設定ファイルが1つも見つからない", file=sys.stderr)
        return 2

    for f in failures:
        print(f"FAIL {f}")
    print(f"coverage: 設定 {checked_files} 件 / hook 参照 {checked_hooks} 件を検査")
    print(f"RESULT: {'FAIL (%d)' % len(failures) if failures else 'OK'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
