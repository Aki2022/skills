#!/usr/bin/env python3
"""prepare_publish_materials.py — ⑤の vault 公開素材を outline.md から機械的に作る。

`digest.md` や `deck_meta.json` を別ファイルとして残さず、公開のたびに outline.md から作る（LLM を使わない）。
opt-out の記録は outline.md の先頭 1 行 `<!-- vault_publish: publish|opted_out -->` だけ。project と source_repo は
vault の note（frontmatter）が持つので、ここでは扱わない（SPEC-document-publish）。

出力（`--out-dir` 配下。opt-out のときは何も作らない）:
  digest.md   1 行目 = outline の H1 の題（`outline.md —` の接頭辞を除く）、以降 = `- S<n> <title>: <キーメッセージ>`
  outline.md  outline.md から先頭の vault_publish 行だけを除いたもの

標準出力は JSON: {"vault_publish": "publish" | "opted_out" | null, "slides": <枚数>}（null = 記録が無い）。
キーメッセージが 1 つも見つからない outline は、空のダイジェストを黙って作らず非 0 で止める。

使い方:
  python3 prepare_publish_materials.py <デッキdir>/process/outline.md --out-dir "$(mktemp -d)"
"""

import argparse
import json
import re
import sys
from pathlib import Path

_MARKER_RE = re.compile(r"^<!--\s*vault_publish:\s*(\S+?)\s*-->\s*$")
_VALID = {"publish", "opted_out"}
_SLIDE_SPLIT_RE = re.compile(r"^# スライド(\d+): ?", re.M)
_KEY_MESSAGE_RE = re.compile(r"^## キーメッセージ[^\n]*\n+(.+?)\n", re.M)
_H1_RE = re.compile(r"^# (?!スライド\d+:)(.+)$", re.M)


def split_marker(text: str) -> tuple[str | None, str]:
    """Return (recorded choice or None, text without the marker line)."""
    lines = text.splitlines(keepends=True)
    first = next((i for i, line in enumerate(lines) if line.strip()), None)
    if first is None:
        return None, text
    match = _MARKER_RE.match(lines[first].strip())
    if match is None:
        return None, text
    value = match.group(1)
    if value not in _VALID:
        raise SystemExit(f"vault_publish must be publish or opted_out, got {value!r}")
    return value, "".join(lines[:first] + lines[first + 1 :])


def build_digest(text: str) -> tuple[str, int]:
    title_match = _H1_RE.search(text)
    if title_match is None:
        raise SystemExit("outline has no title line (# ...)")
    title = re.sub(r"^outline\.md\s*[—–-]+\s*", "", title_match.group(1)).strip()
    parts = _SLIDE_SPLIT_RE.split(text)
    lines = [title]
    slides = 0
    for index in range(1, len(parts), 2):
        slides += 1
        number, body = parts[index], parts[index + 1]
        slide_title = body.split("\n", 1)[0].strip()
        message = _KEY_MESSAGE_RE.search(body)
        if message:
            lines.append(f"- S{number} {slide_title}: {message.group(1).strip()}")
    if len(lines) == 1:
        raise SystemExit("no key message found in any slide; refusing to make an empty digest")
    return "\n".join(lines) + "\n", slides


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("outline", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    args = parser.parse_args()
    if not args.outline.is_file():
        raise SystemExit(f"not a file: {args.outline.name}")
    choice, published = split_marker(args.outline.read_text(encoding="utf-8"))
    if choice == "opted_out":  # 出さないので素材は作らない（キーメッセージが無くても止めない）
        slides = len(_SLIDE_SPLIT_RE.findall(published))
    else:
        digest, slides = build_digest(published)
        args.out_dir.mkdir(parents=True, exist_ok=True)
        (args.out_dir / "digest.md").write_text(digest, encoding="utf-8")
        (args.out_dir / "outline.md").write_text(published, encoding="utf-8")
    json.dump({"vault_publish": choice, "slides": slides}, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
