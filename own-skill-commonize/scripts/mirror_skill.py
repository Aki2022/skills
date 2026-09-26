#!/usr/bin/env python3
"""Import one third-party skill into the canonical root and wire all per-skill roots."""
from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def manifest_names(path: Path) -> dict[str, set[str]]:
    result = {"mirrors": set(), "retired": set()}
    section = None
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line[:1].isspace():
            section = line.split(":", 1)[0] if line.endswith(":") else None
        match = re.match(r"\s+- dir:\s*(.*?)\s*$", line)
        if section in result and match:
            name = match.group(1).strip().strip("\"'")
            result[section].add(name)
    return result


def frontmatter_name(path: Path) -> str | None:
    for line in path.read_text(encoding="utf-8").splitlines():
        match = re.match(r"^name:\s*(.*?)\s*$", line)
        if match:
            return match.group(1).strip().strip("\"'")
    return None


def replace_retired_entry(lines: list[str], name: str) -> list[str]:
    out: list[str] = []
    section = None
    skipping = False
    for line in lines:
        if not line[:1].isspace():
            if skipping:
                skipping = False
            section = line.split(":", 1)[0] if line.endswith(":") else None
        if section == "retired" and re.match(r"^  - dir:\s*" + re.escape(name) + r"\s*$", line):
            skipping = True
            continue
        if skipping:
            if re.match(r"^  - dir:", line):
                skipping = False
            else:
                continue
        out.append(line)
    return out


def write_ledger(
    path: Path, *, name: str, upstream: str, upstream_path: str,
    version: str, license_name: str, reinstall: str, note: str,
    local_copy: str, reactivate: bool,
) -> None:
    lines = path.read_text(encoding="utf-8").splitlines()
    if reactivate:
        lines = replace_retired_entry(lines, name)
    block = [
        f"  - dir: {name}",
        f"    upstream: {json.dumps(upstream, ensure_ascii=False)}",
        f"    upstream_path: {json.dumps(upstream_path, ensure_ascii=False)}",
        f"    upstream_version: {json.dumps(version, ensure_ascii=False)}",
        f"    fetched_at: {datetime.date.today().isoformat()}",
        f"    license: {json.dumps(license_name, ensure_ascii=False)}",
        f"    reinstall: {json.dumps(reinstall, ensure_ascii=False)}",
    ]
    if local_copy:
        block.append(f"    local_copy: {json.dumps(local_copy, ensure_ascii=False)}")
    if note:
        block.append(f"    note: {json.dumps(note, ensure_ascii=False)}")
    try:
        marker = next(i for i, line in enumerate(lines) if line == "retired:")
    except StopIteration:
        lines.extend(["", "retired:"])
        marker = len(lines) - 1
    while marker > 0 and not lines[marker - 1].strip():
        marker -= 1
    lines[marker:marker] = ["", *block]
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--upstream", required=True)
    parser.add_argument("--path", dest="upstream_path", required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--license", dest="license_name", required=True)
    parser.add_argument("--version", default="unknown")
    parser.add_argument("--note", default="")
    parser.add_argument("--source-dir", type=Path, help="copy a reviewed local source instead of fetching")
    parser.add_argument("--local-copy", default="", help="local comparison source recorded in mirrors.yaml")
    parser.add_argument("--reactivate-retired", action="store_true")
    parser.add_argument("--alias-root", action="append", type=Path, default=[])
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    name = args.name
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", name) or name.startswith("own-"):
        print("ERROR invalid third-party skill name", file=sys.stderr)
        return 2
    canonical = Path(os.environ.get("AGENTS_SKILLS_ROOT", "~/.agents/skills")).expanduser()
    ledger = canonical / "mirrors.yaml"
    if not ledger.is_file():
        print("ERROR mirrors.yaml is missing", file=sys.stderr)
        return 2
    entries = manifest_names(ledger)
    if name in entries["mirrors"]:
        print(f"ERROR {name} is already active in mirrors.yaml", file=sys.stderr)
        return 2
    if name in entries["retired"] and not args.reactivate_retired:
        print(f"ERROR {name} is retired; pass --reactivate-retired", file=sys.stderr)
        return 2
    if args.reactivate_retired and name not in entries["retired"]:
        print(f"ERROR {name} is not retired", file=sys.stderr)
        return 2
    destination = canonical / name
    if not args.reactivate_retired and destination.exists():
        print(f"ERROR canonical directory already exists: {name}", file=sys.stderr)
        return 2

    home = Path.home()
    alias_roots = args.alias_root or [
        home / ".codex/skills",
        home / ".codex-private/skills",
        home / ".codex-seat2/skills",
        home / ".gemini/config/skills",
    ]
    for root in alias_roots:
        if not root.is_dir():
            continue
        alias = root / name
        if alias.is_symlink() and alias.resolve(strict=False) == destination.resolve(strict=False):
            continue
        if alias.exists() or alias.is_symlink():
            print(f"ERROR alias root has an existing entry for {name}: {root.name}", file=sys.stderr)
            return 2

    staged: Path | None = None
    if args.reactivate_retired and not args.source_dir:
        if not (destination / "SKILL.md").is_file():
            print(f"ERROR retired skill source is missing: {name}", file=sys.stderr)
            return 2
    else:
        with tempfile.TemporaryDirectory(prefix=".mirror-skill-", dir=canonical) as tmp:
            staged = Path(tmp) / name
            if args.source_dir:
                source = args.source_dir.expanduser()
                skill_file = source / "SKILL.md"
                if not skill_file.is_file():
                    print("ERROR source directory has no SKILL.md", file=sys.stderr)
                    return 2
                source_name = frontmatter_name(skill_file)
                if source_name and source_name != name:
                    print(f"ERROR source frontmatter name {source_name!r} does not match {name!r}", file=sys.stderr)
                    return 2
                shutil.copytree(source, staged, symlinks=True)
            else:
                try:
                    subprocess.run(
                        ["npx", "--yes", "degit", f"{args.upstream}/{args.upstream_path}", str(staged)],
                        check=True,
                    )
                except (OSError, subprocess.CalledProcessError):
                    print("ERROR upstream fetch failed", file=sys.stderr)
                    return 1
            skill_file = staged / "SKILL.md"
            if not skill_file.is_file():
                print("ERROR fetched source has no SKILL.md", file=sys.stderr)
                return 1
            source_name = frontmatter_name(skill_file)
            if source_name and source_name != name:
                print(f"ERROR source frontmatter name {source_name!r} does not match {name!r}", file=sys.stderr)
                return 1
            if args.reactivate_retired:
                backup = Path(tmp) / "retired-source-backup"
                shutil.move(str(destination), str(backup))
                try:
                    shutil.move(str(staged), str(destination))
                except OSError:
                    shutil.move(str(backup), str(destination))
                    raise
            else:
                shutil.move(str(staged), str(destination))

    reinstall = f"npx degit {args.upstream}/{args.upstream_path} ~/.agents/skills/{name}"
    write_ledger(
        ledger, name=name, upstream=args.upstream, upstream_path=args.upstream_path,
        version=args.version, license_name=args.license_name, reinstall=reinstall,
        note=args.note, local_copy=args.local_copy, reactivate=args.reactivate_retired,
    )
    gitignore = canonical / ".gitignore"
    ignore_lines = gitignore.read_text(encoding="utf-8").splitlines() if gitignore.exists() else []
    ignore_entry = f"/{name}/"
    if ignore_entry not in ignore_lines:
        ignore_lines.append(ignore_entry)
        gitignore.write_text("\n".join(ignore_lines) + "\n", encoding="utf-8")

    for root in alias_roots:
        if not root.is_dir():
            print(f"SKIP missing alias root: {root.name}")
            continue
        alias = root / name
        if alias.is_symlink() and alias.resolve(strict=False) == destination.resolve(strict=False):
            continue
        alias.symlink_to(destination)
        print(f"WIRED {root.name}/{name}")

    script_dir = Path(__file__).resolve().parent
    audit_args = ["python3", str(script_dir / "audit_skill_wiring.py"), "--canonical", str(canonical)]
    for root in alias_roots:
        audit_args.extend(["--search", str(root)])
    result = subprocess.run(audit_args, check=False)
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
