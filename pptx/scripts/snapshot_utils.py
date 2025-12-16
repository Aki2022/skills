#!/usr/bin/env python3
"""Snapshot utilities for audit and reproducibility.

This module creates snapshots of templates and styles used at generation time,
stored in powerpoint/processing/snapshot/ for audit purposes.
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

from scripts.logging_utils import get_logger


def create_generation_snapshot(project_dir: Optional[str] = None) -> None:
    """Create snapshot of templates used at generation time.

    Copies current templates from ~/.claude/skills/pptx/templates/ to
    {project_dir}/powerpoint/processing/snapshot/ for audit purposes.

    Args:
        project_dir: Project directory path. If None, uses current working directory.

    Creates:
        powerpoint/processing/snapshot/
        ├── template.pptx    # Template used at generation time
        ├── template.crtx    # Chart template used
        ├── style.yaml       # Style config used
        └── timestamp.txt    # Generation timestamp and metadata
    """
    logger = get_logger()

    if project_dir is None:
        project_dir = os.getcwd()

    # Paths
    snapshot_dir = os.path.join(project_dir, 'powerpoint', 'processing', 'snapshot')
    skill_templates = os.path.expanduser('~/.claude/skills/pptx/templates')

    # Create snapshot directory
    os.makedirs(snapshot_dir, exist_ok=True)
    logger.info(f"Creating generation snapshot in {snapshot_dir}")

    # Template files to snapshot
    files_to_copy = [
        ('template.pptx', 'PowerPoint template'),
        ('template.crtx', 'Chart template'),
        ('style.yaml', 'Style configuration'),
        ('TEMPLATE.md', 'Layout documentation'),
    ]

    # Copy each file
    for filename, description in files_to_copy:
        src = os.path.join(skill_templates, filename)
        dst = os.path.join(snapshot_dir, filename)

        if not os.path.exists(src):
            logger.warning(f"{description} not found: {src}")
            continue

        try:
            shutil.copy2(src, dst)
            logger.debug(f"Copied {description}: {filename}")
        except Exception as e:
            logger.error(f"Failed to copy {filename}: {e}")

    # Create timestamp file
    timestamp_path = os.path.join(snapshot_dir, 'timestamp.txt')
    try:
        with open(timestamp_path, 'w', encoding='utf-8') as f:
            f.write(f"Generation Timestamp: {datetime.now().isoformat()}\n")
            f.write(f"Skill Templates Path: {skill_templates}\n")
            f.write(f"Project Directory: {project_dir}\n")
            f.write(f"\nPurpose: Audit trail of templates/styles used at generation time\n")
            f.write(f"Note: Regeneration always uses latest templates from skill directory\n")
        logger.debug(f"Created timestamp file: {timestamp_path}")
    except Exception as e:
        logger.error(f"Failed to create timestamp file: {e}")

    logger.info(f"✅ Generation snapshot created successfully")


def get_snapshot_info(project_dir: Optional[str] = None) -> dict:
    """Get information about the generation snapshot.

    Args:
        project_dir: Project directory path. If None, uses current working directory.

    Returns:
        Dict with snapshot information:
            - exists: bool - whether snapshot exists
            - timestamp: str - generation timestamp (if exists)
            - files: list - list of snapshot files (if exists)
    """
    if project_dir is None:
        project_dir = os.getcwd()

    snapshot_dir = os.path.join(project_dir, 'powerpoint', 'processing', 'snapshot')

    if not os.path.exists(snapshot_dir):
        return {'exists': False}

    # Read timestamp
    timestamp_path = os.path.join(snapshot_dir, 'timestamp.txt')
    timestamp = None
    if os.path.exists(timestamp_path):
        try:
            with open(timestamp_path, 'r', encoding='utf-8') as f:
                first_line = f.readline().strip()
                if first_line.startswith('Generation Timestamp: '):
                    timestamp = first_line.split(': ', 1)[1]
        except Exception:
            pass

    # List files
    files = []
    if os.path.exists(snapshot_dir):
        files = [f for f in os.listdir(snapshot_dir) if os.path.isfile(os.path.join(snapshot_dir, f))]

    return {
        'exists': True,
        'timestamp': timestamp,
        'files': files,
        'path': snapshot_dir
    }


if __name__ == '__main__':
    # Demo usage
    print("Snapshot Utils Demo")
    print("=" * 50)

    # Create snapshot
    create_generation_snapshot()

    # Get snapshot info
    info = get_snapshot_info()
    print(f"\nSnapshot exists: {info.get('exists')}")
    if info.get('exists'):
        print(f"Timestamp: {info.get('timestamp')}")
        print(f"Files: {', '.join(info.get('files', []))}")
        print(f"Path: {info.get('path')}")
