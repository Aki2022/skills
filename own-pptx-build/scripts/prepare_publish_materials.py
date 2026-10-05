#!/usr/bin/env python3
"""prepare_publish_materials.py — ⑤の vault 公開素材を outline.md から作る。

スライド見出しやキーメッセージは解析しない（デッキごとに書式が違い、解析すると黙って別物になる）。
outline.md はそのまま vault の document note の本文にし、note の summary は outline の題 1 行にする。
追加の記録ファイルは作らない。opt-out の記録は outline.md の先頭 1 行
`<!-- vault_publish: publish|opted_out -->` だけ。project の確定値は vault の note の frontmatter が持つ。

出力:
  --out-dir/outline.md  先頭の vault_publish 行だけを除いた公開用 outline（opt-out のときは何も作らない）
  --out-dir/summary.txt summary を 1 行で書いたもの。`attach-document --summary-file` へ渡す（題に `"` `$(` `` ` `` が
                        あってもシェルを通らない）
  標準出力の JSON       {"vault_publish": "publish" | "opted_out" | null, "summary": <題>, "source_repo": <名>}
    vault_publish null = 記録が無い（既存デッキ）。⑤で opt-out を聞いて outline.md の先頭に行を足す。
    summary     = outline の題。`# outline.md — 題` を優先し、無ければ**最初のスライド見出しより前**の `# ` 見出し
                  （コードフェンス内は除く。`# スライド<n>`・`# S<n>`・`# Slide <n>` 形式は題にしない）。
                  無ければ --fallback-title、それも無ければ非 0。題として妥当かは呼び出し側が確認する。
    source_repo = git remote（origin、無ければ最初）の URL の最後の `/` か `:` より後ろから `.git` を除いた名前。
                  remote が無ければ main チェックアウトのディレクトリ名。git worktree のディレクトリ名は使わない。
                  導出できない名前（日本語・空白・先頭が `_` や `.`）は `--source-repo <ASCII 名>` で固定する。

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
_SLIDE_HEADING_RE = re.compile(r"^#\s*(?:スライド|slide|s)\s*\d", re.I)
_OUTLINE_TITLE_RE = re.compile(r"^#\s+outline\.md\s*[—–-]+\s*(.+)$")
_REPO_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")


def split_marker(text: str) -> tuple[str | None, str]:
    """Return (recorded choice or None, text without the marker line).

    The marker must be the first non-empty line. A marker anywhere else (or a second one) is an error, not a
    silent "unset": otherwise ⑤ would prepend a new line and leave a stale one behind in the published text.
    """
    lines = text.splitlines(keepends=True)
    markers = [i for i, line in enumerate(lines) if _MARKER_RE.match(line.strip())]
    if not markers:
        return None, text
    if len(markers) > 1:
        raise SystemExit("more than one vault_publish line in outline.md; keep exactly one")
    index = markers[0]
    if any(line.strip() for line in lines[:index]):
        raise SystemExit(
            f"the vault_publish line must be the first line of outline.md (found at line {index + 1}); "
            "move it to the top or remove it"
        )
    value = _MARKER_RE.match(lines[index].strip()).group(1)
    if value not in _VALID:
        raise SystemExit(f"vault_publish must be publish or opted_out, got {value!r}")
    return value, "".join(lines[:index] + lines[index + 1 :])


def _headings_outside_fences(text: str) -> list[str]:
    found: list[str] = []
    in_fence = False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            in_fence = not in_fence
            continue
        if not in_fence and line.startswith("# "):
            found.append(line)
    return found


def find_title(text: str, fallback: str | None) -> str:
    """The outline's title: `# outline.md — 題`, else a `# ` heading before the first slide heading."""
    headings = _headings_outside_fences(text)
    for heading in headings:
        match = _OUTLINE_TITLE_RE.match(heading)
        if match and match.group(1).strip():
            return match.group(1).strip()
    for heading in headings:
        if _SLIDE_HEADING_RE.match(heading):
            break  # 最初のスライド見出しより後ろの見出しは題ではない
        title = heading[2:].strip()
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


def source_repo(directory: Path, override: str | None = None) -> str:
    if override is not None:
        if not _REPO_NAME_RE.match(override):
            raise SystemExit(f"--source-repo must be an ASCII repository name, got {override!r}")
        return override
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
    parser.add_argument("--fallback-title", help="outline に題の行が無いときの summary（デッキの dir 名 yyyymmdd_内容 など）")
    parser.add_argument("--source-repo", help="git remote から導出できない名前（日本語・空白等）のときに固定する ASCII 名")
    args = parser.parse_args()
    if not args.outline.is_file():
        raise SystemExit(f"not a file: {args.outline.name}")
    # BOM は先頭だけでなく、⑤が先頭へ行を足したあとの 2 行目に来ることもある。全部落とす
    choice, published = split_marker(args.outline.read_text(encoding="utf-8").replace("\ufeff", ""))
    summary = find_title(published, args.fallback_title)
    repo = source_repo(args.outline.resolve().parent, args.source_repo)
    if choice != "opted_out":  # 出さないときは何も作らない
        args.out_dir.mkdir(parents=True, exist_ok=True)
        (args.out_dir / "outline.md").write_text(published, encoding="utf-8")
        (args.out_dir / "summary.txt").write_text(summary + "\n", encoding="utf-8")
    json.dump({"vault_publish": choice, "summary": summary, "source_repo": repo}, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
