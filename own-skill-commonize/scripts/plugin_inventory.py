#!/usr/bin/env python3
"""Inventory enabled user plugins and their on-disk functional assets.

Claude and Codex are queried through their local read-only plugin-list
commands. Project-scoped Claude plugins, disabled plugins, uninstalled
catalog entries, and cache directories without an installed/enabled record
are excluded. Output never prints installation paths or config values.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10 and earlier
    import tomli as tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SKIP_PARTS = {".git", "node_modules", ".venv", "__pycache__"}
MANIFESTS = (
    ".claude-plugin/plugin.json",
    ".codex-plugin/plugin.json",
    "plugin.json",
)


@dataclass(frozen=True)
class PluginAssets:
    skills: tuple[str, ...]
    commands: int
    agents: int
    hooks: bool
    mcp_servers: int
    lsp_servers: int
    apps: int
    license: str
    skill_license_counts: tuple[tuple[str, int], ...]
    manifest_found: bool

    @property
    def route_kind(self) -> str:
        other = any(
            (self.commands, self.agents, self.hooks, self.mcp_servers,
             self.lsp_servers, self.apps)
        )
        if self.skills and other:
            return "hybrid"
        if self.skills:
            return "skill-only"
        if other:
            return "integration"
        return "unclassified"


def _count(value: Any) -> int:
    if isinstance(value, (dict, list, tuple, set)):
        return len(value)
    return int(bool(value))


def _manifest(root: Path) -> dict[str, Any]:
    for relative in MANIFESTS:
        path = root / relative
        if not path.is_file():
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        return data if isinstance(data, dict) else {}
    return {}


def _component_files(root: Path) -> list[Path]:
    result: list[Path] = []
    for path in root.rglob("*"):
        if any(part in SKIP_PARTS for part in path.parts):
            continue
        try:
            if path.is_file():
                result.append(path)
        except OSError:
            continue
    return result


def _mcp_file_count(path: Path) -> int:
    if path.suffix.lower() == ".json":
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return 0
        if not isinstance(data, dict):
            return 0
        servers = data.get("mcpServers", data.get("servers"))
        return _count(servers)
    if path.suffix.lower() == ".toml":
        try:
            data = tomllib.loads(path.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError):
            return 0
        return _count(data.get("mcpServers", data.get("servers")))
    return 0


def _license_name(text: str) -> str:
    normalized = text.lower().replace(" ", "-").replace("_", "-")
    if normalized.strip() == "mit":
        return "MIT"
    if "proprietary" in normalized or "all-rights-reserved" in normalized:
        return "Proprietary"
    if "apache-license" in normalized or "apache-2.0" in normalized:
        return "Apache-2.0"
    if "mit-license" in normalized or "spdx-license-identifier:-mit" in normalized:
        return "MIT"
    if "bsd-3-clause" in normalized:
        return "BSD-3-Clause"
    return "unspecified"


def _skill_license_counts(files: list[Path]) -> tuple[tuple[str, int], ...]:
    skill_dirs = {path.parent for path in files if path.name == "SKILL.md"}
    counts: dict[str, int] = {}
    for skill_dir in skill_dirs:
        license_files = sorted(
            path for path in skill_dir.iterdir()
            if path.is_file() and path.name.lower().startswith("license")
        )
        label = "unspecified"
        if license_files:
            try:
                label = _license_name(license_files[0].read_text(encoding="utf-8", errors="replace"))
            except OSError:
                pass
        else:
            try:
                for line in (skill_dir / "SKILL.md").read_text(encoding="utf-8").splitlines():
                    match = re.match(r"^license:\s*(.*?)\s*$", line, re.IGNORECASE)
                    if match:
                        label = _license_name(match.group(1))
                        break
            except OSError:
                pass
        counts[label] = counts.get(label, 0) + 1
    return tuple(sorted(counts.items()))


def scan_plugin_assets(root: Path) -> PluginAssets:
    root = root.expanduser()
    manifest_path = next((root / p for p in MANIFESTS if (root / p).is_file()), None)
    manifest = _manifest(root)
    files = _component_files(root)

    skills = sorted({
        path.parent.name for path in files if path.name == "SKILL.md"
    })
    command_files = [
        path for path in files
        if "commands" in path.relative_to(root).parts and path.suffix.lower() in {".md", ".toml", ".json"}
    ]
    agent_files = [
        path for path in files
        if "agents" in path.relative_to(root).parts and path.suffix.lower() in {".md", ".toml", ".json"}
    ]

    mcp_files = [
        path for path in files
        if path.name in {".mcp.json", "mcp.json", "mcp.toml"}
    ]
    mcp_count = max(_count(manifest.get("mcpServers")), sum(_mcp_file_count(p) for p in mcp_files))
    hooks = bool(manifest.get("hooks")) or any(
        path.name in {"hooks.json", "hooks.toml"}
        or ("hooks" in path.relative_to(root).parts and path.suffix in {".sh", ".js", ".py"})
        for path in files
    )
    lsp_count = max(
        _count(manifest.get("lspServers")),
        _count(manifest.get("languageServers")),
        sum(
            1 for path in files
            if path.name in {"lsp.json", "lsp.toml"} or "lsp" in path.relative_to(root).parts
        ),
    )
    apps = _count(manifest.get("apps"))
    return PluginAssets(
        skills=tuple(skills),
        commands=max(len(command_files), _count(manifest.get("commands"))),
        agents=max(len(agent_files), _count(manifest.get("agents"))),
        hooks=hooks,
        mcp_servers=mcp_count,
        lsp_servers=lsp_count,
        apps=apps,
        license=str(manifest.get("license", "unspecified")),
        skill_license_counts=_skill_license_counts(files),
        manifest_found=manifest_path is not None,
    )


def active_codex_plugins(data: dict[str, Any]) -> list[str]:
    return sorted({
        str(item.get("pluginId"))
        for item in data.get("installed", [])
        if isinstance(item, dict)
        and item.get("installed") is True
        and item.get("enabled") is True
        and item.get("pluginId")
    })


def active_claude_plugins(
    rows: list[dict[str, Any]],
    user_enabled: dict[str, bool] | None = None,
) -> tuple[list[dict[str, Any]], int, int]:
    user_enabled = user_enabled or {}
    def is_enabled(row: dict[str, Any]) -> bool:
        enabled = row.get("enabled") is True
        plugin_id = str(row.get("id", ""))
        # Claude can report a shared plugin ID as enabled because a project
        # scope enables it, even while the account's user setting is false.
        # Respect the explicit account setting for user-scope inventory.
        if row.get("scope") == "user" and plugin_id in user_enabled:
            enabled = enabled and user_enabled[plugin_id]
        return enabled

    active = [
        row for row in rows
        if row.get("scope") == "user" and is_enabled(row)
    ]
    project = sum(1 for row in rows if row.get("scope") != "user" and is_enabled(row))
    disabled = sum(1 for row in rows if row.get("scope") == "user" and not is_enabled(row))
    active.sort(key=lambda row: str(row.get("id", "")))
    return active, project, disabled


def format_plugin(
    provider: str,
    account: str,
    record: dict[str, Any],
    assets: PluginAssets,
) -> str:
    plugin_id = record.get("id") or record.get("pluginId") or "unknown"
    version = record.get("version", "unknown")
    route_kind = assets.route_kind
    lsp_servers = assets.lsp_servers
    if not assets.manifest_found and "lsp" in str(plugin_id).lower():
        route_kind = "integration"
        lsp_servers = max(1, lsp_servers)
    elif route_kind == "unclassified":
        # The product's installed+enabled CLI record proves an active plugin
        # integration even when its cache omits a component manifest.
        route_kind = "integration"
    skills = ",".join(assets.skills) if assets.skills else "-"
    license_summary = assets.license
    if any(label != "unspecified" for label, _ in assets.skill_license_counts):
        labels = ",".join(f"{label}:{count}" for label, count in assets.skill_license_counts)
        license_summary = f"per-skill[{labels}]"
    components = (
        f"skills={len(assets.skills)}[{skills}] "
        f"commands={assets.commands} agents={assets.agents} "
        f"hooks={int(assets.hooks)} mcp={assets.mcp_servers} "
        f"lsp={lsp_servers} apps={assets.apps}"
    )
    manifest = "" if assets.manifest_found else " manifest=missing"
    return (
        f"  {provider}/{account} {plugin_id} v{version} "
        f"route={route_kind} {components} license={license_summary}{manifest}"
    )


def _account_spec(raw: str) -> tuple[str, Path]:
    label, sep, path = raw.partition("=")
    if not sep or not re.fullmatch(r"[A-Za-z0-9_-]+", label) or not path:
        raise argparse.ArgumentTypeError("account must be LABEL=PATH")
    return label, Path(path).expanduser()


def _run_json(command: list[str], env: dict[str, str]) -> Any:
    try:
        result = subprocess.run(command, capture_output=True, text=True, env=env, check=False)
    except OSError as exc:
        raise RuntimeError("required plugin-list command is unavailable") from exc
    if result.returncode != 0:
        raise RuntimeError("plugin-list command failed")
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("plugin-list command returned invalid JSON") from exc


def _path_exists(raw: Any) -> Path | None:
    if not isinstance(raw, str) or not raw:
        return None
    path = Path(raw).expanduser()
    return path if path.is_dir() else None


def _codex_source_path(root: Path, record: dict[str, Any]) -> Path | None:
    source = record.get("source", {})
    direct_path = _path_exists(source.get("path") if isinstance(source, dict) else None)
    if direct_path is not None:
        return direct_path

    # Marketplace records from current Codex builds carry marketplaceName,
    # pluginId, and version instead of a source.path. The installed+enabled CLI
    # record is the authority that allows inspecting this exact-version cache.
    marketplace = record.get("marketplaceName")
    plugin_id = str(record.get("pluginId", "")).partition("@")[0]
    version = record.get("version")
    parts = (marketplace, plugin_id, version)
    safe_parts = all(
        isinstance(part, str)
        and part
        and Path(part).name == part
        and part not in {".", ".."}
        for part in parts
    )
    if not safe_parts:
        return None
    candidate = root / "plugins" / "cache" / str(marketplace) / plugin_id / str(version)
    return candidate if candidate.is_dir() else None


def _run_claude(label: str, root: Path) -> tuple[int, int]:
    if not (root / "settings.json").is_file():
        print(f"SKIP Claude/{label}: no user settings file")
        return 0, 0
    env = os.environ.copy()
    env["CLAUDE_CONFIG_DIR"] = str(root)
    rows = _run_json(["claude", "plugin", "list", "--json"], env)
    if not isinstance(rows, list):
        raise RuntimeError("Claude plugin list has unexpected shape")
    configured: dict[str, bool] = {}
    try:
        settings = json.loads((root / "settings.json").read_text(encoding="utf-8"))
        plugins = settings.get("enabledPlugins", {})
        if isinstance(plugins, dict):
            configured = {
                str(key): value for key, value in plugins.items()
                if isinstance(value, bool)
            }
    except (OSError, json.JSONDecodeError):
        pass
    active, project_count, disabled_count = active_claude_plugins(rows, configured)
    listed_user = {
        str(row.get("id")) for row in rows
        if row.get("scope") == "user" and row.get("id")
    }
    unresolved = sorted(
        plugin_id for plugin_id, enabled in configured.items()
        if enabled and plugin_id not in listed_user
    )
    active_user_ids = {str(row.get("id")) for row in active if row.get("id")}
    config_cli_mismatch = sorted(
        plugin_id for plugin_id, enabled in configured.items()
        if enabled and plugin_id in listed_user and plugin_id not in active_user_ids
    )
    print(
        f"Claude/{label}: enabled user plugins={len(active)} "
        f"ignored project plugins={project_count} disabled user plugins={disabled_count}"
    )
    missing_sources = 0
    for record in active:
        plugin_id = str(record.get("id", "unknown"))
        path = _path_exists(record.get("installPath"))
        if path is None:
            print(f"  UNRESOLVED Claude/{label} {plugin_id}: installed source is missing")
            missing_sources += 1
            continue
        print(format_plugin("Claude", label, record, scan_plugin_assets(path)))
    for plugin_id in unresolved:
        print(f"  CONFIG-ONLY Claude/{label} {plugin_id}: not installed and enabled")
    for plugin_id in config_cli_mismatch:
        print(f"  CONFIG-MISMATCH Claude/{label} {plugin_id}: user setting enabled but CLI record is not enabled")
    return len(active), len(unresolved) + len(config_cli_mismatch) + missing_sources


def _run_codex(label: str, root: Path) -> tuple[int, int]:
    config_path = root / "config.toml"
    if not config_path.is_file():
        print(f"SKIP Codex/{label}: no config.toml")
        return 0, 0
    env = os.environ.copy()
    env["CODEX_HOME"] = str(root)
    data = _run_json(["codex", "plugin", "list", "--json"], env)
    if not isinstance(data, dict):
        raise RuntimeError("Codex plugin list has unexpected shape")
    records = data.get("installed", [])
    if not isinstance(records, list):
        raise RuntimeError("Codex installed plugin list has unexpected shape")
    enabled_records = [
        row for row in records
        if isinstance(row, dict) and row.get("installed") is True and row.get("enabled") is True
    ]
    try:
        config = tomllib.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError):
        config = {}
    plugin_config = config.get("plugins", {})
    configured = {
        str(plugin_id) for plugin_id, settings in plugin_config.items()
        if isinstance(settings, dict) and settings.get("enabled") is True
    } if isinstance(plugin_config, dict) else set()
    installed_ids = {
        str(row.get("pluginId")) for row in records
        if isinstance(row, dict) and row.get("installed") is True and row.get("pluginId")
    }
    unresolved = sorted(configured - installed_ids)
    disabled_installed = {
        str(row.get("pluginId")) for row in records
        if isinstance(row, dict) and row.get("installed") is True
        and row.get("enabled") is not True and row.get("pluginId")
    }
    config_cli_mismatch = sorted(configured & disabled_installed)
    print(
        f"Codex/{label}: installed+enabled plugins={len(enabled_records)} "
        f"disabled installed plugins={sum(1 for row in records if isinstance(row, dict) and row.get('installed') is True and row.get('enabled') is not True)} "
        f"enabled config without installed record={len(unresolved)} "
        f"enabled config with disabled installed record={len(config_cli_mismatch)}"
    )
    missing_sources = 0
    for record in sorted(enabled_records, key=lambda row: str(row.get("pluginId", ""))):
        plugin_id = str(record.get("pluginId", "unknown"))
        path = _codex_source_path(root, record)
        if path is None:
            print(f"  UNRESOLVED Codex/{label} {plugin_id}: installed source is missing")
            missing_sources += 1
            continue
        print(format_plugin("Codex", label, record, scan_plugin_assets(path)))
    for plugin_id in unresolved:
        print(f"  CONFIG-ONLY Codex/{label} {plugin_id}: not installed and enabled")
    for plugin_id in config_cli_mismatch:
        print(f"  CONFIG-MISMATCH Codex/{label} {plugin_id}: config enabled but installed record disabled")
    return len(enabled_records), len(unresolved) + len(config_cli_mismatch) + missing_sources


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--claude", action="append", default=[], type=_account_spec, metavar="LABEL=PATH")
    parser.add_argument("--codex", action="append", default=[], type=_account_spec, metavar="LABEL=PATH")
    args = parser.parse_args(argv)
    if not args.claude and not args.codex:
        parser.error("at least one --claude or --codex account is required")
    for provider, accounts in (("Claude", args.claude), ("Codex", args.codex)):
        labels = [label for label, _ in accounts]
        if len(labels) != len(set(labels)):
            parser.error(f"{provider} account labels must be unique")

    print("scope: installed and enabled user plugins; project scope, disabled entries, and cache-only directories excluded")
    failures = 0
    for label, root in args.claude:
        try:
            _, mismatch = _run_claude(label, root)
            failures += mismatch
        except RuntimeError as exc:
            print(f"ERROR Claude/{label}: {exc}")
            failures += 1
    for label, root in args.codex:
        try:
            _, mismatch = _run_codex(label, root)
            failures += mismatch
        except RuntimeError as exc:
            print(f"ERROR Codex/{label}: {exc}")
            failures += 1
    print("RESULT: OK" if failures == 0 else f"RESULT: REVIEW ({failures} unresolved inventory entries)")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
