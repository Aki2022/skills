#!/usr/bin/env python3
"""Temporary, fail-closed launcher for GPT-5.6 Luna leaf workers."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from datetime import datetime, date
from zoneinfo import ZoneInfo


MODEL = "gpt-5.6-luna"
EXPIRES_AFTER = date(2026, 9, 30)
TIMEZONE = ZoneInfo("Asia/Tokyo")
LEAF_MARKER = "OWN_LUNA_RUN_LEAF"
# 旧マーカーも読む。これは再帰呼び出しを止める安全機構で、旧名を立てた親から
# 新スクリプトを呼ぶと、読む側が新名しか見なければ再帰が黙って通る。
# 立てる側は新名だけでよい（旧名の子はもう作られない）。
LEGACY_LEAF_MARKER = "ORIGIN_AUGUST_LUNA_LOOP_LEAF"

ROLE_CONFIG = {
    "worker": {
        "effort": "high",
        "sandbox": "workspace-write",
        "instruction": (
            "Implement the bounded task directly. Keep changes minimal, use TDD "
            "when practical, run relevant verification, and report changed files, "
            "tests, contract impacts, risks, and assumptions."
        ),
    },
    "explorer": {
        "effort": "high",
        "sandbox": "read-only",
        "instruction": (
            "Investigate the codebase without editing files. Return concise evidence "
            "with relevant file paths and unresolved uncertainty."
        ),
    },
    "researcher": {
        "effort": "high",
        "sandbox": "read-only",
        "instruction": (
            "Perform source-oriented technical research without editing files. "
            "Prefer primary sources and distinguish sourced facts from inference."
        ),
    },
    "reviewer": {
        "effort": "max",
        "sandbox": "read-only",
        "instruction": (
            "Review independently with fresh context and without editing files. "
            "Prioritize correctness, security, regressions, and missing tests; return "
            "actionable findings before any summary."
        ),
    },
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("prompt", nargs="?", help="Task prompt; stdin is used when omitted")
    parser.add_argument("--role", choices=ROLE_CONFIG)
    parser.add_argument("--prompt-file", type=Path)
    parser.add_argument("--cwd", type=Path, default=Path.cwd())
    parser.add_argument("--check", action="store_true", help="Check bridge status only")
    parser.add_argument("--dry-run", action="store_true", help="Print launch metadata only")
    args = parser.parse_args()
    if not args.check and not args.role:
        parser.error("--role is required unless --check is used")
    if args.prompt is not None and args.prompt_file is not None:
        parser.error("use only one of prompt or --prompt-file")
    return args


def resolve_codex() -> str:
    configured = os.environ.get("CODEX_BIN")
    candidates = [
        configured,
        "/opt/homebrew/bin/codex",
        "/usr/local/bin/codex",
        "/Applications/ChatGPT.app/Contents/Resources/codex",
        shutil.which("codex"),
    ]
    for candidate in candidates:
        if not candidate or not Path(candidate).is_file() or not os.access(candidate, os.X_OK):
            continue
        launcher = str(Path(candidate).expanduser())
        try:
            probe = subprocess.run(
                [launcher, "--version"],
                check=False,
                capture_output=True,
                text=True,
                timeout=10,
            )
        except (OSError, subprocess.SubprocessError):
            continue
        if probe.returncode == 0:
            return launcher
    raise RuntimeError("Codex CLI was not found; set CODEX_BIN to its executable path")


def inspect_luna(codex: str) -> tuple[str | None, str | None]:
    completed = subprocess.run(
        [codex, "debug", "models"],
        check=False,
        capture_output=True,
        text=True,
        timeout=30,
    )
    if completed.returncode != 0:
        return None, f"model catalog inspection failed with exit {completed.returncode}"
    try:
        catalog = json.loads(completed.stdout)
        models = catalog if isinstance(catalog, list) else catalog.get("models", [])
        luna = next(model for model in models if model.get("slug") == MODEL)
        return luna.get("multi_agent_version"), None
    except (json.JSONDecodeError, StopIteration, AttributeError) as exc:
        return None, f"could not parse Luna model metadata: {exc}"


def bridge_status(codex: str) -> dict[str, object]:
    today = datetime.now(TIMEZONE).date()
    multi_agent_version, catalog_warning = inspect_luna(codex)
    return {
        "available": today <= EXPIRES_AFTER and multi_agent_version != "v2",
        "codex_cli": Path(codex).name,
        "model": MODEL,
        "luna_multi_agent_version": multi_agent_version,
        "catalog_warning": catalog_warning,
        "expires_after_jst": EXPIRES_AFTER.isoformat(),
        "today_jst": today.isoformat(),
    }


def enforce_status(status: dict[str, object]) -> None:
    if status["today_jst"] > status["expires_after_jst"]:
        raise RuntimeError(
            "temporary Luna bridge expired; re-check native Luna support before changing routing"
        )
    if status["luna_multi_agent_version"] == "v2":
        raise RuntimeError(
            "Luna is Multi-Agent V2 compatible; remove this bridge and restore native routing"
        )


def read_prompt(args: argparse.Namespace) -> str:
    if args.prompt_file is not None:
        prompt = args.prompt_file.read_text(encoding="utf-8")
    elif args.prompt is not None:
        prompt = args.prompt
    elif not sys.stdin.isatty():
        prompt = sys.stdin.read()
    else:
        raise RuntimeError("provide a prompt, --prompt-file, or piped stdin")
    if not prompt.strip():
        raise RuntimeError("task prompt must not be empty")
    return prompt


def main() -> int:
    args = parse_args()
    if os.environ.get(LEAF_MARKER) == "1" or os.environ.get(LEGACY_LEAF_MARKER) == "1":
        raise RuntimeError("recursive Luna bridge invocation is forbidden")

    codex = resolve_codex()
    status = bridge_status(codex)
    if args.check:
        print(json.dumps(status, indent=2, sort_keys=True))
        return 0 if status["available"] else 2
    enforce_status(status)

    role = ROLE_CONFIG[args.role]
    prompt = read_prompt(args)
    cwd = args.cwd.resolve()
    if not cwd.is_dir():
        raise RuntimeError(f"working directory does not exist: {cwd}")

    leaf_prompt = (
        f"[{LEAF_MARKER}]\n"
        "You are a temporary Luna leaf process. Perform this task yourself. "
        "Do not delegate, spawn agents, or invoke the Luna bridge.\n\n"
        f"Role instructions: {role['instruction']}\n\n"
        f"Task:\n{prompt.strip()}\n"
    )
    command = [
        codex,
        "exec",
        "--model",
        MODEL,
        "--config",
        f'model_reasoning_effort="{role["effort"]}"',
        "--sandbox",
        role["sandbox"],
        "--cd",
        str(cwd),
        "--ephemeral",
        "--json",
        "-",
    ]
    metadata = {
        **status,
        "role": args.role,
        "effort": role["effort"],
        "sandbox": role["sandbox"],
        "prompt_chars": len(leaf_prompt),
    }
    print(json.dumps(metadata, sort_keys=True), file=sys.stderr)
    if args.dry_run:
        safe_command = ["codex", *command[1:]]
        safe_command[safe_command.index(str(cwd))] = "<working-directory>"
        print(json.dumps({**metadata, "command": safe_command}, indent=2, sort_keys=True))
        return 0

    environment = os.environ.copy()
    environment[LEAF_MARKER] = "1"
    completed = subprocess.run(command, input=leaf_prompt, text=True, env=environment)
    return completed.returncode


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"own-luna-run: {exc}", file=sys.stderr)
        raise SystemExit(2) from exc
