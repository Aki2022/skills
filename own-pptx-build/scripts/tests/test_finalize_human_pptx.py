"""finalize_human_pptx.py: 人間が PowerPoint で編集した pptx の納品前仕上げ。

PowerPoint で保存し直すと、最終更新者名が docProps/core.xml に入り、
レイアウト・マスター・テーマに Arial が戻る（2026-10-19 豊橋デッキで実測）。
提出用コピーを作るとき、最終更新者の除去・フォント統一・検証を 1 コマンドで通す。
"""

import re
import subprocess
import sys
import zipfile
from pathlib import Path

from pptx import Presentation

SCRIPT = Path(__file__).resolve().parents[1] / "finalize_human_pptx.py"


def make_human_pptx(path: Path) -> Path:
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    prs.core_properties.last_modified_by = "Some Person"
    prs.core_properties.title = "Title"
    prs.save(path)
    # テーマに Arial を持たせる（python-pptx 既定テーマは Calibri のため置換して再現）
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/theme/") and item.filename.endswith(".xml"):
                data = re.sub(rb'<a:latin typeface="[^"]*"', b'<a:latin typeface="Arial"', data)
            zout.writestr(item, data)
    tmp.replace(path)
    return path


def run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(SCRIPT), *args], capture_output=True, text=True)


def read(path: Path, name: str) -> str:
    return zipfile.ZipFile(path).read(name).decode("utf-8")


def test_strips_last_modified_by_and_sets_title(tmp_path):
    src = make_human_pptx(tmp_path / "rev.pptx")
    out = tmp_path / "submit.pptx"
    r = run(str(src), str(out), "--title", "応募資料")
    assert r.returncode == 0, r.stdout + r.stderr
    core = read(out, "docProps/core.xml")
    assert "Some Person" not in core
    assert "応募資料" in core


def test_unifies_fonts_in_theme(tmp_path):
    src = make_human_pptx(tmp_path / "rev.pptx")
    assert '<a:latin typeface="Arial"' in read(src, "ppt/theme/theme1.xml")
    out = tmp_path / "submit.pptx"
    assert run(str(src), str(out)).returncode == 0
    theme = read(out, "ppt/theme/theme1.xml")
    # 言語別の <a:font script=...> は対象外。既定の latin が統一される
    assert '<a:latin typeface="Arial"' not in theme
    assert '<a:latin typeface="BIZ UDPGothic"' in theme


def test_source_is_not_modified(tmp_path):
    src = make_human_pptx(tmp_path / "rev.pptx")
    before = src.read_bytes()
    assert run(str(src), str(tmp_path / "submit.pptx")).returncode == 0
    assert src.read_bytes() == before


def test_refuses_to_overwrite_source(tmp_path):
    src = make_human_pptx(tmp_path / "rev.pptx")
    r = run(str(src), str(src))
    assert r.returncode != 0
    assert "same" in (r.stdout + r.stderr).lower() or "同じ" in (r.stdout + r.stderr)


def test_fails_when_validation_fails(tmp_path):
    bad = tmp_path / "bad.pptx"
    with zipfile.ZipFile(bad, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
    r = run(str(bad), str(tmp_path / "out.pptx"))
    assert r.returncode != 0
