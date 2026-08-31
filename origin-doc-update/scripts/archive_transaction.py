"""Small transactional helpers for archiving one document and its index row."""

from __future__ import annotations

import os
import stat
import tempfile
from pathlib import Path
from typing import Optional


def stage_text(path: Path, content: str, mode_source: Optional[Path] = None) -> Path:
    """Write content beside its destination without touching the destination."""
    path = Path(path)
    mode_source = Path(mode_source) if mode_source is not None else path
    fd, staged_name = tempfile.mkstemp(
        prefix=f".{path.name}.",
        dir=str(path.parent),
        text=True,
    )
    staged = Path(staged_name)
    try:
        with os.fdopen(fd, "w") as handle:
            handle.write(content)
        if mode_source.exists():
            os.chmod(staged, stat.S_IMODE(mode_source.stat().st_mode))
        return staged
    except BaseException:
        try:
            staged.unlink()
        except FileNotFoundError:
            pass
        raise


def cleanup_staged(paths: list[Optional[Path]]) -> None:
    """Remove only temporary paths created by :func:`stage_text`."""
    for path in paths:
        if path is None:
            continue
        try:
            Path(path).unlink()
        except FileNotFoundError:
            pass


def apply_archive(
    source: Path,
    destination: Path,
    staged_destination: Path,
    staged_source_restore: Path,
    index_path: Optional[Path] = None,
    staged_index: Optional[Path] = None,
    staged_index_restore: Optional[Path] = None,
) -> None:
    """Install staged archive/index contents, restoring all files on failure.

    All staged files must already exist. The source and destination are expected to
    be on the same filesystem, as they are for ``docs/*/archive``. A failure in a
    later replacement therefore does not leave an archived document with a stale
    active index row.
    """
    source = Path(source)
    destination = Path(destination)
    staged_destination = Path(staged_destination)
    staged_source_restore = Path(staged_source_restore)
    index_path = Path(index_path) if index_path is not None else None
    staged_index = Path(staged_index) if staged_index is not None else None
    staged_index_restore = (
        Path(staged_index_restore) if staged_index_restore is not None else None
    )
    if index_path is not None and (
        staged_index is None or staged_index_restore is None
    ):
        raise ValueError("index transaction requires both staged index files")

    moved = False
    installed_new_index = False
    try:
        os.replace(source, destination)
        moved = True
        os.replace(staged_destination, destination)
        if index_path is not None and staged_index is not None:
            os.replace(staged_index, index_path)
            installed_new_index = True
    except BaseException as error:
        rollback_errors: list[str] = []

        if installed_new_index and index_path is not None and staged_index_restore is not None:
            try:
                os.replace(staged_index_restore, index_path)
            except OSError as rollback_error:
                rollback_errors.append(f"index rollback failed: {rollback_error}")

        if moved:
            try:
                if source.exists():
                    raise OSError(f"source reappeared during rollback: {source}")
                if destination.exists():
                    os.replace(destination, source)
                os.replace(staged_source_restore, source)
            except OSError as rollback_error:
                rollback_errors.append(f"document rollback failed: {rollback_error}")

        detail = f"archive transaction failed: {error}"
        if rollback_errors:
            detail += "; " + "; ".join(rollback_errors)
        raise RuntimeError(detail) from error
