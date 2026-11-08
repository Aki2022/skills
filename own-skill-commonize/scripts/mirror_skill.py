#!/usr/bin/env python3
"""Import or update one third-party skill in the canonical root and wire all per-skill roots.

--update-existing replaces an active mirror from a new source (normally the
Claude Code plugin copy of the skill). Without --apply it only prints the diff
between the canonical mirror and the new source and changes nothing; the owner
reviews that diff and re-runs with --apply. Ledger fields that are not given on
the command line (note, local_copy, plugin, ...) are kept.
"""
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

# A pin must name one upstream state: a full 40-hex commit SHA or a version tag.
# Short SHAs are refused: degit cannot resolve them (measured 2026-11-08:
# "could not find commit hash for 2b1ffd2").
PIN_RE = re.compile(r"^(?:[0-9a-f]{40}|v?\d+(?:\.\d+)*)$")
FULL_SHA_RE = re.compile(r"^[0-9a-f]{40}$")
UPDATED_KEYS = ("upstream", "upstream_path", "upstream_version", "fetched_at", "license", "reinstall")


def manifest_names(path: Path) -> dict[str, set[str]]:
    result = {"mirrors": set(), "retired": set()}
    section = None
    for line in path.read_text(encoding="utf-8").splitlines():
        if line and not line[:1].isspace():
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


def entry_span(lines: list[str], section_name: str, name: str) -> tuple[int, int] | None:
    """Line range [start, end) of one `  - dir: <name>` entry, trailing blanks excluded."""
    section = None
    start = None
    for i, line in enumerate(lines):
        top = bool(line) and not line[:1].isspace()
        if start is not None and (top or re.match(r"^  - dir:", line)):
            end = i
            while end > start and not lines[end - 1].strip():
                end -= 1
            return start, end
        if top:
            section = line.split(":", 1)[0]
        if section == section_name and re.match(r"^  - dir:\s*" + re.escape(name) + r"\s*$", line):
            start = i
    if start is None:
        return None
    end = len(lines)
    while end > start and not lines[end - 1].strip():
        end -= 1
    return start, end


def replace_retired_entry(lines: list[str], name: str) -> list[str]:
    span = entry_span(lines, "retired", name)
    if span is None:
        return lines
    start, end = span
    while end < len(lines) and not lines[end].strip():
        end += 1
    return lines[:start] + lines[end:]


def ledger_values(
    *, upstream: str, upstream_path: str, version: str, license_name: str,
    reinstall: str, note: str, local_copy: str, plugin: str = "", plugin_path: str = "",
) -> dict[str, str]:
    values = {
        "upstream": json.dumps(upstream, ensure_ascii=False),
        "upstream_path": json.dumps(upstream_path, ensure_ascii=False),
        "upstream_version": json.dumps(version, ensure_ascii=False),
        "fetched_at": datetime.date.today().isoformat(),
        "license": json.dumps(license_name, ensure_ascii=False),
        "reinstall": json.dumps(reinstall, ensure_ascii=False),
    }
    if local_copy:
        values["local_copy"] = json.dumps(local_copy, ensure_ascii=False)
    if note:
        values["note"] = json.dumps(note, ensure_ascii=False)
    if plugin:
        values["plugin"] = json.dumps(plugin, ensure_ascii=False)
    if plugin_path:
        values["plugin_path"] = json.dumps(plugin_path, ensure_ascii=False)
    return values


def write_ledger(path: Path, *, name: str, values: dict[str, str], reactivate: bool) -> None:
    lines = path.read_text(encoding="utf-8").splitlines()
    if reactivate:
        lines = replace_retired_entry(lines, name)
    block = [f"  - dir: {name}"] + [f"    {key}: {value}" for key, value in values.items()]
    try:
        marker = next(i for i, line in enumerate(lines) if line == "retired:")
    except StopIteration:
        lines.extend(["", "retired:"])
        marker = len(lines) - 1
    while marker > 0 and not lines[marker - 1].strip():
        marker -= 1
    lines[marker:marker] = ["", *block]
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def update_ledger(path: Path, *, name: str, values: dict[str, str]) -> None:
    """Rewrite one active entry in place, keeping every field not in `values`."""
    lines = path.read_text(encoding="utf-8").splitlines()
    span = entry_span(lines, "mirrors", name)
    if span is None:
        raise ValueError(f"{name} is not active in mirrors.yaml")
    start, end = span
    block = [lines[start]]
    seen: set[str] = set()
    for line in lines[start + 1:end]:
        match = re.match(r"^    ([A-Za-z_]+):", line)
        key = match.group(1) if match else None
        if key in values:
            block.append(f"    {key}: {values[key]}")
            seen.add(key)
        else:
            block.append(line)
    block += [f"    {key}: {value}" for key, value in values.items() if key not in seen]
    lines[start:end] = block
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def tree_files(root: Path) -> dict[str, Path]:
    files: dict[str, Path] = {}
    for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
        for entry in dirnames + filenames:
            full = Path(dirpath) / entry
            rel = full.relative_to(root).as_posix()
            if full.is_symlink():
                files[rel] = full
            elif entry in filenames:
                files[rel] = full
    return files


def same_entry(a: Path, b: Path) -> bool:
    if a.is_symlink() or b.is_symlink():
        return a.is_symlink() and b.is_symlink() and os.readlink(a) == os.readlink(b)
    return a.read_bytes() == b.read_bytes()


def print_diff(current: Path, new: Path) -> int:
    old_files, new_files = tree_files(current), tree_files(new)
    changes = 0
    for rel in sorted(set(old_files) | set(new_files)):
        if rel not in new_files:
            print(f"REMOVED {rel}")
        elif rel not in old_files:
            print(f"ADDED {rel}")
        elif not same_entry(old_files[rel], new_files[rel]):
            print(f"CHANGED {rel}")
        else:
            continue
        changes += 1
    return changes


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
    parser.add_argument("--plugin", default="",
                        help='Claude Code plugin the source comes from, "<plugin>@<marketplace>" (drift check)')
    parser.add_argument("--plugin-path", default="", help="skill path inside that plugin, e.g. skills/<name>")
    parser.add_argument("--reactivate-retired", action="store_true")
    parser.add_argument("--update-existing", action="store_true",
                        help="replace an active mirror; without --apply only print the diff")
    parser.add_argument("--apply", action="store_true", help="with --update-existing: write the change")
    parser.add_argument("--alias-root", action="append", type=Path, default=[])
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    name = args.name
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", name) or name.startswith("own-"):
        print("ERROR invalid third-party skill name", file=sys.stderr)
        return 2
    if args.update_existing and args.reactivate_retired:
        print("ERROR --update-existing and --reactivate-retired cannot be combined", file=sys.stderr)
        return 2
    if args.apply and not args.update_existing:
        print("ERROR --apply is only for --update-existing", file=sys.stderr)
        return 2
    if args.update_existing and not PIN_RE.fullmatch(args.version):
        print("ERROR --update-existing needs --version as a full 40-hex commit SHA or a version tag",
              file=sys.stderr)
        return 2
    canonical = Path(os.environ.get("AGENTS_SKILLS_ROOT", "~/.agents/skills")).expanduser()
    ledger = canonical / "mirrors.yaml"
    if not ledger.is_file():
        print("ERROR mirrors.yaml is missing", file=sys.stderr)
        return 2
    entries = manifest_names(ledger)
    destination = canonical / name
    if args.update_existing:
        if name not in entries["mirrors"]:
            print(f"ERROR {name} is not active in mirrors.yaml", file=sys.stderr)
            return 2
        if destination.is_symlink() or not (destination / "SKILL.md").is_file():
            print(f"ERROR active mirror source is missing or not a directory: {name}", file=sys.stderr)
            return 2
    else:
        if name in entries["mirrors"]:
            print(f"ERROR {name} is already active in mirrors.yaml; use --update-existing", file=sys.stderr)
            return 2
        if name in entries["retired"] and not args.reactivate_retired:
            print(f"ERROR {name} is retired; pass --reactivate-retired", file=sys.stderr)
            return 2
        if args.reactivate_retired and name not in entries["retired"]:
            print(f"ERROR {name} is not retired", file=sys.stderr)
            return 2
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

    ref = f"#{args.version}" if args.update_existing and FULL_SHA_RE.fullmatch(args.version) else ""
    reinstall = f"npx degit {args.upstream}/{args.upstream_path}{ref} ~/.agents/skills/{name}"
    values = ledger_values(
        upstream=args.upstream, upstream_path=args.upstream_path, version=args.version,
        license_name=args.license_name, reinstall=reinstall, note=args.note, local_copy=args.local_copy,
        plugin=args.plugin, plugin_path=args.plugin_path,
    )

    if args.reactivate_retired and not args.source_dir:
        if not (destination / "SKILL.md").is_file():
            print(f"ERROR retired skill source is missing: {name}", file=sys.stderr)
            return 2
    else:
        with tempfile.TemporaryDirectory(prefix=".mirror-skill-", dir=canonical) as tmp:
            staged = Path(tmp) / name
            if args.source_dir:
                source = args.source_dir.expanduser()
                if not (source / "SKILL.md").is_file():
                    print("ERROR source directory has no SKILL.md", file=sys.stderr)
                    return 2
                shutil.copytree(source, staged, symlinks=True)
            else:
                try:
                    subprocess.run(
                        ["npx", "--yes", "degit", f"{args.upstream}/{args.upstream_path}{ref}", str(staged)],
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
            if args.update_existing and not args.apply:
                changes = print_diff(destination, staged)
                print(f"DRY-RUN {name}: {changes} change(s) to {args.version}; nothing written. "
                      "Review the diff, then re-run with --apply.")
                return 0
            if destination.exists():
                backup = Path(tmp) / "previous-source"
                shutil.move(str(destination), str(backup))
                try:
                    shutil.move(str(staged), str(destination))
                    if args.update_existing:
                        update_ledger(ledger, name=name, values=values)
                except Exception:
                    if destination.exists():
                        shutil.rmtree(destination)
                    shutil.move(str(backup), str(destination))
                    raise
            else:
                shutil.move(str(staged), str(destination))

    if not args.update_existing:
        write_ledger(ledger, name=name, values=values, reactivate=args.reactivate_retired)
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
