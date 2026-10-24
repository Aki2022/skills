"""Installed command links must reach the canonical skill implementation."""
import os
from pathlib import Path
import subprocess
import sys
from test_core import SKILL


def test_relative_command_symlinks_reach_all_three_cli_entrypoints(tmp_path):
    for name in ['design-author','design-check','design-ship']:
        target = tmp_path / (name + '-target')
        target.symlink_to(SKILL / 'scripts' / name)
        alias = tmp_path / name
        alias.symlink_to(target.name)
        run = subprocess.run([str(alias),'--help'],env=dict(os.environ,DESIGN_AUTHOR_PYTHON=sys.executable),capture_output=True,text=True)
        assert run.returncode == 0, run.stderr
        assert 'usage:' in run.stdout
