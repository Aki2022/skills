"""sanitize_pptx.py の規則ごとの回帰テスト。

各テストは欠陥を 1 つ持つ最小の pptx（zip）を作り、本番と同じ CLI 経路
（--check → 修正 → --check）で検出と修正を確かめる。欠陥が消えたことに加え、
同じ部品の無関係な内容が変わっていないことも見る。

検査は XML の構造（子要素の順序・属性値）に対するもので、PowerPoint では開いていない。
"""

import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "sanitize_pptx.py"

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

DLBLS = (
    '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>'
    '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
)
DPT = (
    '<c:dPt><c:idx val="0"/><c:bubble3D val="0"/>'
    '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>'
)
CAT_VAL = (
    '<c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>'
    '<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/>'
    '<c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>'
)
MARKER = '<c:marker><c:symbol val="circle"/></c:marker>'


def run(pptx: Path, *flags: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(SCRIPT), str(pptx), *flags], capture_output=True, text=True)


def write_pptx(path: Path, parts: dict[str, str]) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, xml.encode("utf-8"))
    return path


def part_bytes(pptx: Path, name: str) -> bytes:
    with zipfile.ZipFile(pptx) as z:
        return z.read(name)


def part_tree(pptx: Path, name: str) -> ET.Element:
    return ET.fromstring(part_bytes(pptx, name))


def local(el: ET.Element) -> str:
    return el.tag.split("}", 1)[1]


def assert_detected_then_fixed(pptx: Path) -> None:
    before = run(pptx, "--check")
    assert before.returncode == 1, f"修正前の --check は NG のはず:\n{before.stdout}{before.stderr}"
    fixed = run(pptx)
    assert fixed.returncode == 0, fixed.stderr
    after = run(pptx, "--check")
    assert after.returncode == 0, f"修正後の --check は OK のはず:\n{after.stdout}{after.stderr}"


def slide(shapes: str, bg: str = "") -> str:
    return (
        XML_DECL
        + f'<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"><p:cSld>{bg}'
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f"<p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"
    )


def sp(shape_id: int, name: str, prst: str, av: str = "", descr: str | None = None, txbody: str = "") -> str:
    d = "" if descr is None else f' descr="{descr}"'
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"{d}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom></p:spPr>{txbody}</p:sp>'
    )


def chart(plot: str) -> str:
    return (
        XML_DECL
        + f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
        f"<c:chart><c:plotArea>{plot}</c:plotArea></c:chart></c:chartSpace>"
    )




# --- 30ba05c: dPt は dLbls より前（全チャート種）---------------------------------

@pytest.mark.parametrize(
    ("plot", "expected_order"),
    [
        pytest.param(
            '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/><c:invertIfNegative val="0"/>{DLBLS}{DPT}{CAT_VAL}</c:ser>'
            '<c:axId val="1"/><c:axId val="2"/></c:barChart>',
            ["idx", "order", "invertIfNegative", "dPt", "dLbls", "cat", "val"],
            id="barChart",
        ),
        pytest.param(
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>{MARKER}{DLBLS}{DPT}{CAT_VAL}<c:smooth val="0"/></c:ser>'
            '<c:axId val="1"/><c:axId val="2"/></c:lineChart>',
            ["idx", "order", "marker", "dPt", "dLbls", "cat", "val", "smooth"],
            id="lineChart",
        ),
        pytest.param(
            '<c:pieChart><c:varyColors val="1"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>{DLBLS}{DPT}{CAT_VAL}</c:ser>'
            '<c:firstSliceAng val="0"/></c:pieChart>',
            ["idx", "order", "dPt", "dLbls", "cat", "val"],
            id="pieChart",
        ),
    ],
)
def test_dpt_is_moved_before_dlbls_in_every_chart_type(tmp_path: Path, plot: str, expected_order: list[str]) -> None:
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/charts/chart1.xml": chart(plot)})
    assert_detected_then_fixed(pptx)
    ser = part_tree(pptx, "ppt/charts/chart1.xml").find(f".//{{{NS_C}}}ser")
    assert [local(c) for c in ser] == expected_order
    assert ser.find(f"{{{NS_C}}}dPt/{{{NS_C}}}spPr/{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr").get("val") == "FF0000"


def test_renamed_chart_part_is_scanned(tmp_path: Path) -> None:
    """inject_template.py は衝突した部品を g_chart1.xml へ改名する。旧フィルタ（ppt/charts/chart*）はこれを見逃した。"""
    plot = (
        '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        f'<c:ser><c:idx val="0"/><c:order val="0"/>{DLBLS}{MARKER}{CAT_VAL}<c:smooth val="0"/></c:ser>'
        '<c:axId val="1"/><c:axId val="2"/></c:lineChart>'
    )
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/charts/g_chart1.xml": chart(plot)})
    assert_detected_then_fixed(pptx)
    ser = part_tree(pptx, "ppt/charts/g_chart1.xml").find(f".//{{{NS_C}}}ser")
    assert [local(c) for c in ser] == ["idx", "order", "marker", "dLbls", "cat", "val", "smooth"]


# --- 限界 1（#58 の既知の限界）: dPt は marker・invertIfNegative より後 -------------

@pytest.mark.parametrize(
    ("plot", "expected_order"),
    [
        pytest.param(
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>{DPT}{MARKER}{DLBLS}{CAT_VAL}<c:smooth val="0"/></c:ser>'
            '<c:axId val="1"/><c:axId val="2"/></c:lineChart>',
            ["idx", "order", "marker", "dPt", "dLbls", "cat", "val", "smooth"],
            id="line-dPt-before-marker",
        ),
        pytest.param(
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>{DLBLS}{DPT}{MARKER}{CAT_VAL}<c:smooth val="0"/></c:ser>'
            '<c:axId val="1"/><c:axId val="2"/></c:lineChart>',
            ["idx", "order", "marker", "dPt", "dLbls", "cat", "val", "smooth"],
            id="line-dLbls-and-dPt-before-marker",
        ),
        pytest.param(
            '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>{DPT}{MARKER}{CAT_VAL}<c:smooth val="0"/></c:ser>'
            '<c:axId val="1"/><c:axId val="2"/></c:lineChart>',
            ["idx", "order", "marker", "dPt", "cat", "val", "smooth"],
            id="line-without-dLbls",
        ),
        pytest.param(
            '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
            f'<c:ser><c:idx val="0"/><c:order val="0"/>{DPT}<c:invertIfNegative val="0"/>{CAT_VAL}</c:ser>'
            '<c:axId val="1"/><c:axId val="2"/></c:barChart>',
            ["idx", "order", "invertIfNegative", "dPt", "cat", "val"],
            id="bar-dPt-before-invertIfNegative",
        ),
    ],
)
def test_dpt_is_placed_after_every_required_predecessor(tmp_path: Path, plot: str, expected_order: list[str]) -> None:
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/charts/chart1.xml": chart(plot)})
    assert_detected_then_fixed(pptx)
    ser = part_tree(pptx, "ppt/charts/chart1.xml").find(f".//{{{NS_C}}}ser")
    assert [local(c) for c in ser] == expected_order


# --- a5e8740: serAx の無い 2D の 3 つ目の axId を落とす --------------------------

@pytest.mark.parametrize("kind", ["barChart", "lineChart", "areaChart", "radarChart"])
def test_extra_third_axid_is_dropped_from_2d_charts_without_series_axis(tmp_path: Path, kind: str) -> None:
    head = {
        "barChart": '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>',
        "lineChart": '<c:grouping val="standard"/><c:varyColors val="0"/>',
        "areaChart": '<c:grouping val="standard"/><c:varyColors val="0"/>',
        "radarChart": '<c:radarStyle val="marker"/><c:varyColors val="0"/>',
    }[kind]
    ser = f'<c:ser><c:idx val="0"/><c:order val="0"/>{CAT_VAL}</c:ser>'
    axids = "".join(f'<c:axId val="{v}"/>' for v in ("101", "102", "103"))
    plot = f"<c:{kind}>{head}{ser}{axids}</c:{kind}>"
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/charts/chart1.xml": chart(plot)})
    assert_detected_then_fixed(pptx)
    chart_el = part_tree(pptx, "ppt/charts/chart1.xml").find(f".//{{{NS_C}}}{kind}")
    assert [a.get("val") for a in chart_el.findall(f"{{{NS_C}}}axId")] == ["101", "102"]
    assert chart_el.find(f"{{{NS_C}}}ser") is not None


def test_axid_on_chart_with_series_axis_is_kept(tmp_path: Path) -> None:
    """serAx を持つ 3D 棒の 3 つ目の axId は正規の系列軸。規則は serAx の無い 2D にだけ効く（ガード）。"""
    plot = (
        '<c:bar3DChart><c:barDir val="col"/><c:grouping val="standard"/><c:varyColors val="0"/>'
        f'<c:ser><c:idx val="0"/><c:order val="0"/>{CAT_VAL}</c:ser>'
        '<c:axId val="101"/><c:axId val="102"/><c:axId val="103"/></c:bar3DChart>'
        '<c:serAx><c:axId val="103"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="102"/></c:serAx>'
    )
    original = chart(plot)
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/charts/chart1.xml": original})
    assert run(pptx, "--check").returncode == 0
    assert run(pptx).returncode == 0
    assert part_bytes(pptx, "ppt/charts/chart1.xml") == original.encode("utf-8")


# --- 15c5a83: スライド内の重複 cNvPr id を振り直す ------------------------------

def test_duplicate_shape_ids_within_a_slide_are_renumbered(tmp_path: Path) -> None:
    table = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/>'
        "</p:nvGraphicFramePr>"
        '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/></a:graphic>'
        "</p:graphicFrame>"
    )
    original = slide(sp(2, "Title", "rect") + table + sp(3, "Body", "rect"))
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/slides/slide1.xml": original})
    assert_detected_then_fixed(pptx)
    tree = part_tree(pptx, "ppt/slides/slide1.xml")
    ids = [c.get("id") for c in tree.iter(f"{{{NS_P}}}cNvPr")]
    assert len(ids) == len(set(ids)) == 4
    # 振り直したのは 2 回目に出た id="2"（表）だけ。他の内容はバイト単位で同じ
    text = part_bytes(pptx, "ppt/slides/slide1.xml").decode("utf-8")
    assert text == original.replace('<p:cNvPr id="2" name="Table"/>', '<p:cNvPr id="4" name="Table"/>')


def test_duplicate_id_is_found_when_name_precedes_id(tmp_path: Path) -> None:
    """限界 2（#58 の既知の限界）: 属性の順序が name→id でも重複を見つけて振り直す（id が先頭という前提を外す）。"""
    table = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr name="Table" id="2"/><p:cNvGraphicFramePr/><p:nvPr/>'
        "</p:nvGraphicFramePr>"
        '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="500000"/></p:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"/></a:graphic>'
        "</p:graphicFrame>"
    )
    original = slide(sp(2, "Title", "rect") + table + sp(3, "Body", "rect"))
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/slides/slide1.xml": original})
    assert_detected_then_fixed(pptx)
    tree = part_tree(pptx, "ppt/slides/slide1.xml")
    ids = [c.get("id") for c in tree.iter(f"{{{NS_P}}}cNvPr")]
    assert len(ids) == len(set(ids)) == 4
    text = part_bytes(pptx, "ppt/slides/slide1.xml").decode("utf-8")
    assert text == original.replace('<p:cNvPr name="Table" id="2"/>', '<p:cNvPr name="Table" id="4"/>')


def shape_gds(tree: ET.Element, shape_name: str) -> list[tuple[str | None, str | None]]:
    for sp_el in tree.iter(f"{{{NS_P}}}sp"):
        if sp_el.find(f".//{{{NS_P}}}cNvPr").get("name") == shape_name:
            return [(g.get("name"), g.get("fmla")) for g in sp_el.iter(f"{{{NS_A}}}gd")]
    raise AssertionError(f"shape {shape_name!r} not found")


# --- 採用（仕様で確認）: round2SameRect の調整名は adj1。roundRect の adj は変えない ---

def test_round2sameRect_adj_is_renamed_to_adj1_and_roundrect_is_kept(tmp_path: Path) -> None:
    shapes = (
        sp(2, "Round2", "round2SameRect", av='<a:gd name="adj" fmla="val 16667"/>')
        + sp(3, "Round", "roundRect", av='<a:gd name="adj" fmla="val 16667"/>')
    )
    pptx = write_pptx(tmp_path / "d.pptx", {"ppt/slides/slide1.xml": slide(shapes)})
    assert_detected_then_fixed(pptx)
    tree = part_tree(pptx, "ppt/slides/slide1.xml")
    assert shape_gds(tree, "Round2") == [("adj1", "val 16667")]
    # roundRect の調整ハンドルは仕様どおり adj（改名しない）
    assert shape_gds(tree, "Round") == [("adj", "val 16667")]
