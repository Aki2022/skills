#!/usr/bin/env python3
"""prepare_publish_materials.py — ⑤の vault 公開素材を outline.md から作る。

スライド見出しやキーメッセージは解析しない（デッキごとに書式が違い、解析すると黙って別物になる）。
outline.md はそのまま vault の document note の本文にし、note の summary は outline の題 1 行にする。
追加の記録ファイルは作らない。opt-out の記録は outline.md の先頭 1 行
`<!-- vault_publish: publish|opted_out -->` だけ。project の確定値は vault の note の frontmatter が持つ。

出力:
  --out-dir/outline.md  先頭の vault_publish 行だけを除いた公開用 outline（opt-out のときは何も作らない）
  標準出力の JSON       {"vault_publish": "publish" | "opted_out" | null, "summary": <題>, "source_repo": <名>}
    vault_publish null = 記録が無い（既存デッキ）。⑤で opt-out を聞いて outline.md の先頭に行を足す。
    summary     = outline の最初の `# ` 見出し（`outline.md —` の接頭辞を除く）。無ければ --fallback-title、それも無ければ非 0。
    source_repo = git remote（origin、無ければ最初）の URL の最後の `/` か `:` より後ろから `.git` を除いた名前。
                  remote が無ければ main チェックアウトのディレクトリ名。git worktree のディレクトリ名は使わない。

使い方:
  python3 prepare_publish_materials.py <デッキdir>/process/outline.md --out-dir "$(mktemp -d)" [--fallback-title <デッキ名>]
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

_MARKER_RE = re.compile(r"^<!--\s*vault_publish:\s*(\S+?)\s*-->\s*$")
_VALID = {"publish", "opted_out"}
_H1_RE = re.compile(r"^# (?!スライド\d)(.+)$", re.M)
_REPO_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")


def split_marker(text: str) -> tuple[str | None, str]:
    """Return (recorded choice or None, text without the marker line). Two markers are an error."""
    lines = text.splitlines(keepends=True)
    markers = [i for i, line in enumerate(lines) if _MARKER_RE.match(line.strip())]
    if not markers:
        return None, text
    if len(markers) > 1:
        raise SystemExit("more than one vault_publish line in outline.md; keep exactly one")
    index = markers[0]
    if any(line.strip() for line in lines[:index]):
        return None, text  # 先頭の行ではない（本文中の例示など）。記録とはみなさない
    value = _MARKER_RE.match(lines[index].strip()).group(1)
    if value not in _VALID:
        raise SystemExit(f"vault_publish must be publish or opted_out, got {value!r}")
    return value, "".join(lines[:index] + lines[index + 1 :])


def find_title(text: str, fallback: str | None) -> str:
    match = _H1_RE.search(text)
    if match:
        title = re.sub(r"^outline\.md\s*[—–-]+\s*", "", match.group(1)).strip()
        if title:
            return title
    if fallback:
        return fallback
    raise SystemExit("outline has no title line (# ...); pass --fallback-title")


def _git(directory: Path, *args: str) -> str:
    try:
        result = subprocess.run(
            ["git", "-C", str(directory), *args], capture_output=True, text=True, timeout=15, check=False
        )
    except (OSError, subprocess.SubprocessError) as exc:
        raise SystemExit(f"git is not available: {exc}")
    if result.returncode != 0:
        raise SystemExit(f"git {args[0]} failed in the outline's directory (is it inside a git repository?)")
    return result.stdout


def source_repo(directory: Path) -> str:
    remotes = _git(directory, "remote").split()
    if remotes:
        remote = "origin" if "origin" in remotes else remotes[0]
        url = _git(directory, "remote", "get-url", remote).strip().rstrip("/")
        name = re.split(r"[/:]", url)[-1]
        name = name[: -len(".git")] if name.endswith(".git") else name
    else:
        first = next(
            (line[len("worktree ") :] for line in _git(directory, "worktree", "list", "--porcelain").splitlines()
             if line.startswith("worktree ")),
            "",
        )
        name = Path(first).name
    if not _REPO_NAME_RE.match(name):
        raise SystemExit(f"cannot derive a repository name (got {name!r})")
    return name


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("outline", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--fallback-title", help="outline に題の行が無いときの summary（デッキ名など）")
    args = parser.parse_args()
    if not args.outline.is_file():
        raise SystemExit(f"not a file: {args.outline.name}")
    # BOM は先頭だけでなく、⑤が先頭へ行を足したあとの 2 行目に来ることもある。全部落とす
    choice, published = split_marker(args.outline.read_text(encoding="utf-8").replace("\ufeff", ""))
    summary = find_title(published, args.fallback_title)
    repo = source_repo(args.outline.resolve().parent)
    if choice != "opted_out":  # 出さないときは何も作らない
        args.out_dir.mkdir(parents=True, exist_ok=True)
        (args.out_dir / "outline.md").write_text(published, encoding="utf-8")
    json.dump({"vault_publish": choice, "summary": summary, "source_repo": repo}, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
