#!/usr/bin/env python3
"""Install and inspect the per-user skill usage LaunchAgent.

The checked-in plist is a home-independent template.  Installation is explicit
and refuses to overwrite a different existing plist.  Loading the job is also
explicit; ``--check`` performs no writes and only reports whether the template,
installed file, and launchd registration agree.
"""

from __future__ import annotations

import argparse
import os
import plistlib
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Sequence


LABEL = "com.origin.skill-usage-metrics"
SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRIPT_DIR.parent / "references/com.origin.skill-usage-metrics.plist"


def _home_path(home: str | Path | None = None) -> Path:
    return Path(home).expanduser() if home is not None else Path.home()


def launchd_domain(uid: int | None = None) -> str:
    user_id = os.getuid() if uid is None else uid
    return f"gui/{user_id}"


def installed_path(home: str | Path | None = None) -> Path:
    return _home_path(home) / "Library/LaunchAgents" / f"{LABEL}.plist"


def state_path(home: str | Path | None = None) -> Path:
    return _home_path(home) / ".local/state/origin-skill-usage"


def template_bytes() -> bytes:
    try:
        data = TEMPLATE_PATH.read_bytes()
        parsed = plistlib.loads(data)
    except (OSError, plistlib.InvalidFileException, ValueError) as error:
        raise RuntimeError("launchd template is invalid") from error
    if parsed.get("Label") != LABEL:
        raise RuntimeError("launchd template label is incorrect")
    interval = parsed.get("StartCalendarInterval")
    if not isinstance(interval, dict) or interval.get("Hour") != 3 or interval.get("Minute") != 30:
        raise RuntimeError("launchd template schedule is incorrect")
    if parsed.get("RunAtLoad") is not False:
        raise RuntimeError("launchd template must not run at load")
    return data


def _write_private_state(path: Path) -> None:
    if path.is_symlink():
        raise RuntimeError("state directory must not be a symlink")
    path.mkdir(parents=True, exist_ok=True, mode=0o700)
    mode = path.stat().st_mode & 0o777
    if mode & 0o077:
        path.chmod(0o700)


def install_plist(home: str | Path | None = None, *, replace: bool = False) -> Path:
    """Install the exact template atomically, refusing an unexpected file."""

    data = template_bytes()
    destination = installed_path(home)
    if destination.is_symlink():
        raise RuntimeError("installed plist must not be a symlink")
    if destination.parent.is_symlink():
        raise RuntimeError("LaunchAgents directory must not be a symlink")
    destination.parent.mkdir(parents=True, exist_ok=True)
    _write_private_state(state_path(home))
    if destination.exists():
        current = destination.read_bytes()
        if current != data and not replace:
            raise RuntimeError("an existing plist differs; pass --replace after review")
        if current == data:
            destination.chmod(0o600)
            return destination
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "wb", dir=destination.parent, prefix=f".{destination.name}.", delete=False
        ) as handle:
            temporary_name = handle.name
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.chmod(temporary_name, 0o600)
        os.replace(temporary_name, destination)
        temporary_name = None
    finally:
        if temporary_name:
            try:
                os.unlink(temporary_name)
            except OSError:
                pass
    return destination


def _launchctl(*arguments: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["/bin/launchctl", *arguments],
        capture_output=True,
        text=True,
        check=False,
    )


def is_loaded(*, uid: int | None = None) -> bool:
    result = _launchctl("print", f"{launchd_domain(uid)}/{LABEL}")
    return result.returncode == 0


def bootstrap(home: str | Path | None = None, *, uid: int | None = None) -> None:
    destination = installed_path(home)
    if not destination.is_file() or destination.is_symlink():
        raise RuntimeError("install the plist before bootstrapping")
    domain = launchd_domain(uid)
    if is_loaded(uid=uid):
        return
    result = _launchctl("bootstrap", domain, str(destination))
    if result.returncode != 0:
        raise RuntimeError("launchd bootstrap failed")


def kickstart(*, uid: int | None = None) -> None:
    if not is_loaded(uid=uid):
        raise RuntimeError("launchd job is not loaded")
    result = _launchctl("kickstart", "-k", f"{launchd_domain(uid)}/{LABEL}")
    if result.returncode != 0:
        raise RuntimeError("launchd kickstart failed")


def check(home: str | Path | None = None, *, uid: int | None = None) -> dict[str, bool]:
    data = template_bytes()
    destination = installed_path(home)
    installed = destination.is_file() and not destination.is_symlink()
    exact = installed and destination.read_bytes() == data
    return {
        "template_valid": True,
        "plist_installed": installed,
        "plist_exact": exact,
        "loaded": is_loaded(uid=uid),
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Manage the per-user skill usage LaunchAgent.")
    actions = parser.add_mutually_exclusive_group(required=True)
    actions.add_argument(
        "--check", action="store_true", help="read-only template/install/load check"
    )
    actions.add_argument("--install", action="store_true", help="install the exact template")
    parser.add_argument("--replace", action="store_true", help="replace a reviewed differing plist")
    parser.add_argument(
        "--bootstrap", action="store_true", help="load the installed job into launchd"
    )
    parser.add_argument(
        "--kickstart", action="store_true", help="run one immediate scheduled collection"
    )
    args = parser.parse_args(argv)
    try:
        if args.check:
            result = check()
            print(" ".join(f"{key}={str(value).lower()}" for key, value in result.items()))
            return 0 if result["plist_exact"] and result["loaded"] else 1
        if args.replace and is_loaded():
            raise RuntimeError(
                "cannot replace a loaded plist; bootout the job and review before retrying"
            )
        install_plist(replace=args.replace)
        if args.bootstrap:
            bootstrap()
        if args.kickstart:
            kickstart()
        return 0
    except (OSError, RuntimeError, ValueError, TypeError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
