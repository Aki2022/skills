#!/usr/bin/env python3
"""check_mirrors.py — 第三者 skill ミラーの同値・鮮度検査（決定論・LLM 不使用）。

使い方:
  python3 check_mirrors.py [skills_root] [--max-age-days N]

検査の範囲（ここに無いものは見ない）:
  parity   : mirrors.yaml の各 dir が local_copy と**バイト同一**か（全ファイルの相対パスと内容の
             sha256）。local_copy が無い entry は SKIP（built-in 等は比較不可）。不一致は FAIL（exit 1）。
  freshness: fetched_at から max_age_days を超えていれば WARN（exit 0）。**上流の最新版は見ない**
             （ネットワークに出ない）ので、上流が更新されたかは分からない。空振りを含む警告として扱う。
  drift    : entry に plugin（"<plugin>@<marketplace>"）と plugin_path があれば、Claude Code の
             installed_plugins.json が指す導入版の <installPath>/<plugin_path> と正典を比べ、
             違えば WARN drift（exit 0）。同期は mirror_skill.sh --update-existing で所有者が行う。
  RESULT 行に parity の比較数・SKIP 数と drift 数を出す（緑でも比べていない件数を隠さない）。
exit: 0 = OK/WARN のみ, 1 = FAIL あり, 2 = 台帳が無い・読めない
PyYAML に依存しない（環境に無いため）。台帳は「- key: value」のフラットな list のみを想定する。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
from datetime import date
from pathlib import Path


def parse_manifest(text: str) -> list[dict]:
    entries: list[dict] = []
    cur: dict | None = None
    in_list = False
    for raw in text.splitlines():
        line = raw.split(" #", 1)[0].rstrip() if not raw.lstrip().startswith("#") else ""
        if not line.strip():
            continue
        if line.strip() == "mirrors:":
            in_list = True
            continue
        # トップレベルの別キー（retired: 等）で mirrors: の並びは終わる。
        # 終端を見ていなかったため、2026-09-23 に retired: を足した時点で
        # 退役済み 23 件を現役ミラーとして読み、「mirror dir missing」で赤くなった。
        # 空振りではなく誤検知だが、原因が台帳の書式にあることは出力から読めなかった。
        if in_list and line[:1] not in (" ", "\t") and line.rstrip().endswith(":"):
            in_list = False
            continue
        if not in_list:
            continue
        s = line.strip()
        if s.startswith("- "):
            cur = {}
            entries.append(cur)
            s = s[2:]
        if cur is None or ":" not in s:
            continue
        k, v = s.split(":", 1)
        value = v.strip()
        if value.startswith('"') and value.endswith('"'):
            try:
                value = json.loads(value)
            except json.JSONDecodeError:
                pass
        elif value.startswith("'") and value.endswith("'"):
            value = value[1:-1].replace("''", "'")
        cur[k.strip()] = value
    return entries


def tree_digest(root: Path) -> str:
    h = hashlib.sha256()
    for p in sorted(x for x in root.rglob("*") if x.is_file()):
        rel = p.relative_to(root).as_posix()
        h.update(rel.encode())
        h.update(b"\0")
        h.update(p.read_bytes())
        h.update(b"\0")
    return h.hexdigest()


def load_installed_plugins(paths: list[Path]) -> dict[str, list[dict]]:
    """plugin@marketplace -> install records, merged over every readable file."""
    found: dict[str, list[dict]] = {}
    for path in paths:
        try:
            data = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        plugins = data.get("plugins", data) if isinstance(data, dict) else {}
        for key, records in plugins.items():
            records = records if isinstance(records, list) else [records]
            found.setdefault(key, []).extend(r for r in records if isinstance(r, dict))
    return found


def load_marketplaces(paths: list[Path]) -> dict[str, Path]:
    """marketplace name -> local clone (installLocation) that exists on disk."""
    found: dict[str, Path] = {}
    for path in paths:
        try:
            data = json.loads(path.read_text())
        except (OSError, json.JSONDecodeError):
            continue
        for name, record in (data.items() if isinstance(data, dict) else []):
            loc = record.get("installLocation") if isinstance(record, dict) else None
            if loc and Path(loc).is_dir() and name not in found:
                found[name] = Path(loc)
    return found


def clone_head(clone: Path) -> str:
    """Commit SHA of a marketplace clone, read from .git without running git.

    The official marketplace is delivered as a snapshot without .git; it carries
    the commit in a `.gcs-sha` file (measured 2026-11-08, not documented).
    """
    marker = clone / ".gcs-sha"
    if not (clone / ".git").exists() and marker.is_file():
        value = marker.read_text().strip()
        if len(value) == 40 and all(ch in "0123456789abcdef" for ch in value):
            return value
    git = clone / ".git"
    try:
        head = (git / "HEAD").read_text().strip()
        if not head.startswith("ref: "):
            return head
        ref = head[5:]
        ref_file = git / ref
        if ref_file.is_file():
            return ref_file.read_text().strip()
        for line in (git / "packed-refs").read_text().splitlines():
            if line.endswith(" " + ref):
                return line.split(" ", 1)[0]
    except OSError:
        pass
    return "unknown"


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root", nargs="?", default=str(Path.home() / ".agents" / "skills"))
    ap.add_argument("--max-age-days", type=int, default=90)
    ap.add_argument("--plugins-file", action="append", type=Path, default=[],
                    help="Claude Code installed_plugins.json (repeatable; default: ~/.claude/plugins/)")
    ap.add_argument("--marketplaces-file", action="append", type=Path, default=[],
                    help="Claude Code known_marketplaces.json (repeatable; default: ~/.claude/plugins/)")
    ap.add_argument("--only-drift", action="store_true",
                    help="print only WARN drift lines (for the session-start hook); silent otherwise")
    a = ap.parse_args(argv)
    if a.only_drift:
        import io
        from contextlib import redirect_stdout
        rest = [x for x in (sys.argv[1:] if argv is None else argv) if x != "--only-drift"]
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = main(rest)
        for line in buf.getvalue().splitlines():
            if line.startswith("WARN drift"):
                print(line)
        return rc
    plugin_files = a.plugins_file or [Path.home() / ".claude" / "plugins" / "installed_plugins.json"]
    installed = load_installed_plugins(plugin_files)
    market_files = a.marketplaces_file or [Path.home() / ".claude" / "plugins" / "known_marketplaces.json"]
    clones = load_marketplaces(market_files)
    compared = skipped = drift = 0
    root = Path(a.root).expanduser()
    manifest = root / "mirrors.yaml"
    print(f"scope: parity = local_copy とのバイト同一（local_copy 無しは SKIP）/ "
          f"freshness = fetched_at から {a.max_age_days} 日で WARN。上流の最新版は見ない。")
    if not manifest.is_file():
        print(f"ERROR manifest not found: {manifest}")
        return 2
    entries = parse_manifest(manifest.read_text())
    if not entries:
        print(f"ERROR manifest has no mirrors: {manifest}")
        return 2
    fail = False
    for e in entries:
        d = e.get("dir", "?")
        target = root / d
        if not target.is_dir():
            print(f"FAIL parity {d}: mirror dir missing: {target}")
            fail = True
            continue
        lc = e.get("local_copy")
        if lc:
            lcp = Path(os.path.expanduser(lc))
            if not lcp.is_dir():
                print(f"SKIP parity {d}: local_copy not present: {lcp}")
                skipped += 1
            elif tree_digest(target) != tree_digest(lcp):
                print(f"FAIL parity {d}: differs from local_copy {lcp}")
                fail = True
                compared += 1
            else:
                print(f"OK parity {d}")
                compared += 1
        else:
            print(f"SKIP parity {d}: no local_copy (built-in 等・比較不可)")
            skipped += 1
        plugin = e.get("plugin")
        if plugin:
            records = installed.get(plugin, [])
            sources = [Path(r["installPath"]) / e.get("plugin_path", "") for r in records
                       if r.get("installPath")]
            sources = [s for s in sources if s.is_dir()]
            if not sources:
                print(f"SKIP plugin {d}: {plugin} has no install path on disk (not installed, or its cache was cleaned)")
            else:
                record = next(r for r in records
                              if (Path(r.get("installPath", "")) / e.get("plugin_path", "")).is_dir())
                pin = record.get("gitCommitSha") or record.get("version") or "unknown"
                if tree_digest(target) != tree_digest(sources[0]):
                    print(f"WARN drift {d}: plugin {plugin} {pin} differs from the canon; "
                          f"sync with mirror_skill.sh --update-existing --source-dir {sources[0]} --version {pin}")
                    drift += 1
                else:
                    print(f"OK plugin {d}: matches {plugin} {pin}")
        market = e.get("marketplace")
        if market:
            clone = clones.get(market)
            source = clone / e.get("marketplace_path", "") if clone else None
            if source is None or not source.is_dir():
                print(f"SKIP marketplace {d}: {market} clone or {e.get('marketplace_path', '')} not on disk")
            else:
                pin = clone_head(clone)
                if tree_digest(target) != tree_digest(source):
                    print(f"WARN drift {d}: marketplace {market} {pin} differs from the canon; "
                          f"sync with mirror_skill.sh --update-existing --source-dir {source} --version {pin}")
                    drift += 1
                else:
                    print(f"OK marketplace {d}: matches {market} {pin}")
        fa = e.get("fetched_at")
        try:
            age = (date.today() - date.fromisoformat(fa)).days if fa else None
        except ValueError:
            age = None
        if age is None:
            print(f"WARN stale {d}: fetched_at missing or unparsable")
        elif age > a.max_age_days:
            print(f"WARN stale {d}: fetched {age} days ago (> {a.max_age_days})")
        else:
            print(f"OK fresh {d}: {age} days")
    print("RESULT:", "FAIL" if fail else "OK",
          f"(parity compared {compared}, skipped {skipped}; plugin drift {drift})")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
