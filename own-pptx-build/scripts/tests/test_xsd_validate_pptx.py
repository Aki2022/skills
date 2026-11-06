"""xsd_validate_pptx.py のスキーマの場所の扱い。

スキーマは --schemas か環境変数 OOXML_SCHEMA_DIR で明示する。既定の場所
（~/.claude*/plugins/cache）は探さない。場所が無ければ exit 2 で止まり、黙って合格にしない。

検証本体（VALID / INVALID の判定）は lxml を使う。lxml が無い環境では、その検査は skip になる。
"""

import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "xsd_validate_pptx.py"
SLIDE = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
)


def _pptx(path: Path) -> Path:
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("ppt/slides/slide1.xml", SLIDE)
    return path


def _schema_dir(root: Path) -> Path:
    d = root / "ISO-IEC29500-4_2016"
    d.mkdir(parents=True)
    (d / "pml.xsd").write_text(
        '<?xml version="1.0"?><xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"/>', encoding="utf-8"
    )
    return d


def _run(pptx: Path, *args: str, env_extra: dict | None = None, home: Path | None = None):
    env = {k: v for k, v in os.environ.items() if k != "OOXML_SCHEMA_DIR"}
    if home is not None:
        env["HOME"] = str(home)
    env.update(env_extra or {})
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(pptx), *args], capture_output=True, text=True, env=env
    )


def test_default_plugin_cache_is_not_searched(tmp_path: Path) -> None:
    """既定の場所（~/.claude*/plugins/cache/**/skills/pptx/...）に置いてあっても使わない。"""
    home = tmp_path / "home"
    _schema_dir(home / ".claude" / "plugins" / "cache" / "x" / "skills" / "pptx" / "scripts" / "office" / "schemas")
    r = _run(_pptx(tmp_path / "d.pptx"), home=home)
    assert r.returncode == 2, r.stdout + r.stderr
    assert "OOXML_SCHEMA_DIR" in r.stderr


def test_missing_location_fails_with_a_clear_message(tmp_path: Path) -> None:
    """場所の指定が無ければ、何を設定すべきかを示して exit 2。"""
    r = _run(_pptx(tmp_path / "d.pptx"), home=tmp_path / "empty")
    assert r.returncode == 2, r.stdout + r.stderr
    assert "OOXML_SCHEMA_DIR" in r.stderr and "--schemas" in r.stderr


def test_location_without_pml_xsd_fails(tmp_path: Path) -> None:
    """指定された場所に pml.xsd が無ければ、その場所を名指しして exit 2。"""
    empty = tmp_path / "no-schema"
    empty.mkdir()
    r = _run(_pptx(tmp_path / "d.pptx"), env_extra={"OOXML_SCHEMA_DIR": str(empty)})
    assert r.returncode == 2, r.stdout + r.stderr
    assert str(empty) in r.stderr


def test_env_location_is_used_when_given(tmp_path: Path) -> None:
    """場所を正しく渡すと検証まで進む（lxml が無ければ skip）。"""
    pytest.importorskip("lxml")
    schema = _schema_dir(tmp_path / "s")
    r = _run(_pptx(tmp_path / "d.pptx"), env_extra={"OOXML_SCHEMA_DIR": str(schema)})
    assert f"schemas={schema}" in r.stdout, r.stdout + r.stderr


def test_argument_wins_over_env(tmp_path: Path) -> None:
    """--schemas は環境変数より優先される（lxml が無ければ skip）。"""
    pytest.importorskip("lxml")
    good = _schema_dir(tmp_path / "good")
    bad = tmp_path / "bad"
    bad.mkdir()
    r = _run(
        _pptx(tmp_path / "d.pptx"),
        "--schemas",
        str(good),
        env_extra={"OOXML_SCHEMA_DIR": str(bad)},
    )
    assert f"schemas={good}" in r.stdout, r.stdout + r.stderr
