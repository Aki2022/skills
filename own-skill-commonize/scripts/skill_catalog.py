#!/usr/bin/env python3
"""Resolve the active canonical skill set from mirrors.yaml.

Repository-local skill roots without mirrors.yaml keep the legacy behavior of
enumerating their direct SKILL.md directories. In the global canonical root,
only own-* directories and entries under mirrors: are active. retired: is
retained as recovery metadata and must never create aliases.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


SKILL_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
DIR_ENTRY = re.compile(r"^\s*-\s+dir:\s*(.*?)\s*$")
TOP_LEVEL = re.compile(r"^(mirrors|retired):(?:\s*(?:\[\])?)?\s*(?:#.*)?$")


@dataclass
class SkillCatalog:
    active: set[str]
    retired: set[str]
    retired_physical: set[str]
    errors: list[str]
    legacy: bool = False


def parse_manifest(path: Path) -> tuple[list[str], list[str], list[str]]:
    """Read only the dir fields from the two flat skill lists."""

    sections: dict[str, list[str]] = {"mirrors": [], "retired": []}
    errors: list[str] = []
    section: str | None = None
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as exc:
        return [], [], [f"cannot read mirrors.yaml: {exc.__class__.__name__}"]

    for number, raw in enumerate(lines, start=1):
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        if not raw[0].isspace():
            match = TOP_LEVEL.match(raw)
            section = match.group(1) if match else None
            continue
        if section not in sections:
            continue
        clean = raw.split(" #", 1)[0].rstrip()
        match = DIR_ENTRY.match(clean)
        if not match:
            continue
        name = match.group(1).strip().strip("\"'")
        if not SKILL_NAME.fullmatch(name) or name in {".", ".."}:
            errors.append(f"invalid {section} skill name at mirrors.yaml:{number}")
            continue
        sections[section].append(name)

    for key, values in sections.items():
        seen: set[str] = set()
        duplicates = sorted(value for value in values if value in seen or seen.add(value))
        if duplicates:
            errors.append(f"duplicate {key} entries: {', '.join(duplicates)}")
    overlap = sorted(set(sections["mirrors"]) & set(sections["retired"]))
    if overlap:
        errors.append(f"entries appear in mirrors and retired: {', '.join(overlap)}")
    return sections["mirrors"], sections["retired"], errors


def load_skill_catalog(root: Path) -> SkillCatalog:
    root = root.expanduser()
    physical = {
        child.name
        for child in root.iterdir()
        if child.name and not child.name.startswith(".")
        and child.is_dir() and (child / "SKILL.md").is_file()
    }
    manifest = root / "mirrors.yaml"
    if not manifest.is_file():
        return SkillCatalog(
            active=physical,
            retired=set(),
            retired_physical=set(),
            errors=[],
            legacy=True,
        )

    mirrors, retired, errors = parse_manifest(manifest)
    active = {name for name in physical if name.startswith("own-")} | set(mirrors)
    retired_set = set(retired)

    for name in sorted(set(mirrors) - physical):
        errors.append(f"active mirror directory missing: {name}")
    unregistered = physical - active - retired_set
    for name in sorted(unregistered):
        errors.append(f"unregistered physical skill: {name}")

    return SkillCatalog(
        active=active,
        retired=retired_set,
        retired_physical=physical & retired_set,
        errors=errors,
    )
