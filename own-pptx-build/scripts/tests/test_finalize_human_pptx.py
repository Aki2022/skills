"""finalize_human_pptx.py: 人間が PowerPoint で編集した pptx の納品前仕上げ。

PowerPoint で保存し直すと、最終更新者名が docProps/core.xml に入り、
レイアウト・マスター・テーマに Arial が戻る（2026-10-19 豊橋デッキで実測）。
提出用コピーを作るとき、最終更新者の除去・フォント統一・検証を 1 コマンドで通す。
"""

import os
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


def test_schema_location_is_configured_for_the_xsd_step() -> None:
    """最終段の XSD 検証は OOXML_SCHEMA_DIR の指定を前提にする（既定の場所は探さない）。
    未指定なら、ここで理由を示して落とす（他のテストが曖昧な失敗で落ちる前に分かるように）。"""
    assert os.environ.get("OOXML_SCHEMA_DIR"), (
        "OOXML_SCHEMA_DIR を設定してください（ISO-IEC29500-4_2016 の dir。既定の場所は探しません）"
    )


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
    """XSD だけが落とす違反（段落途中の a:pPr）を入れた pptx は、[xsd] の段で exit 1（INVALID）として
    止まり、出力を残さない。スキーマの場所が無い exit 2 で落ちた場合は、検証の検査にならないので区別する。"""
    src = make_human_pptx(tmp_path / "human.pptx")
    slide = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>X</a:t></a:r>'
        '<a:pPr algn="r"/></a:p></p:txBody></p:sp>'
        '</p:spTree></p:cSld></p:sld>'
    )
    tmp = src.with_suffix(".tmp")
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = slide.encode("utf-8") if item.filename == "ppt/slides/slide1.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    tmp.replace(src)
    out = tmp_path / "out.pptx"
    r = run(str(src), str(out))
    assert "[xsd] failed (exit 1)" in r.stderr, r.stdout + r.stderr
    assert not out.exists()
