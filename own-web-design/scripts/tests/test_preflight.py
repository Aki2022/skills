"""Exercise the factory entrypoint with the real L0, never a mocked green response."""
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[3]
ENTRY = ROOT / 'own-web-design/scripts/preflight.py'

def test_factory_stops_without_design_from_nested_directory(tmp_path):
    subprocess.run(['git', 'init', '-q', str(tmp_path)], check=True)
    nested = tmp_path / 'app'
    nested.mkdir()
    env = dict(os.environ, DESIGN_AUTHOR_PYTHON=sys.executable)
    run = subprocess.run([sys.executable, str(ENTRY)], cwd=nested, env=env,
                         capture_output=True, text=True)
    assert run.returncode == 2
    result = json.loads(run.stdout)
    assert result['exit_code'] == 2
    assert result['errors']
    assert not (tmp_path / 'design').exists()
    assert not list(tmp_path.rglob('*.html'))


def test_factory_never_falls_back_when_l0_command_is_missing(tmp_path):
    env = dict(os.environ, DESIGN_AUTHOR_COMMAND=str(tmp_path / 'missing'))
    run = subprocess.run([sys.executable, str(ENTRY)], cwd=tmp_path, env=env,
                         capture_output=True, text=True)
    assert run.returncode == 2
    assert json.loads(run.stdout)['errors']
