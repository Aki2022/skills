#!/usr/bin/env python3
"""
sanitize_pptx.py — pptxgenjs 産 pptx の「PowerPoint 修復エラー」要因をビルド後に矯正する。

③ネイティブビルドの必須最終工程（チャートの有無に関わらず実行してよい・冪等）。
根拠（2026-07-13 実測、pptxgenjs 4.0.1）:
  1) lineChart / scatterChart / radarChart の <c:ser> に <c:invertIfNegative>（bar系専用
     要素）が混入する → スキーマ違反で PowerPoint が修復を要求（LibreOffice/python-pptx
     は許容するため④のレンダリング検証では検出できない）
  2) <c:marker> が <c:dLbls> の後に出力される（CT_LineSer 等の子要素順序は marker が先）
  3) defineSlideMaster の slideNumber オプションは idx="4294967295" の不正プレース
     ホルダを吐く（deck_helpers v3 は不使用。検出したら警告のみ＝手動対応）
  4) pptxgenjs はプリセット図形の adjustment 値（吹き出しの尻尾位置等）を公開しないため、
     deck_helpers.speechBubble() 等は shape 名に `<名前>@adj1=..,adj2=..` とエンコードする。
     本スクリプトが <a:avLst> へ <a:gd> を注入して印を除去する（2026-08-27 追加。
     --check は未注入の印を NG として検出する）

使い方:
  python3 sanitize_pptx.py output.pptx          # in-place 修正
  python3 sanitize_pptx.py output.pptx --check  # 修正せず検査のみ（exit 1 = 要修正）
"""
import re
import sys
import shutil
import tempfile
import zipfile

# XXE/billion-laughs 対策: defusedxml があれば優先。無い環境でも下の _reject_dtd ガードで
# DTD/ENTITY を含む入力を拒否する（pptxgenjs 産の正規 pptx は DTD を含まない）。
try:
    import defusedxml.ElementTree as _SAFE_ET  # type: ignore
    import xml.etree.ElementTree as ET  # 出力（tostring/register_namespace）は stdlib

    def _fromstring(data: bytes):
        return _SAFE_ET.fromstring(data)
except ImportError:
    import xml.etree.ElementTree as ET

    def _fromstring(data: bytes):
        return ET.fromstring(data)


def _reject_dtd(data: bytes, name: str) -> None:
    head = data[:4096]
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise SystemExit(f"refusing to parse {name}: DTD/ENTITY declaration found (not a pptxgenjs artifact)")

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = {
    "c": C,
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c16": "http://schemas.microsoft.com/office/drawing/2014/chart",
}
for p, u in NS.items():
    ET.register_namespace(p, u)

SER_PARENTS = [f"{{{C}}}lineChart", f"{{{C}}}scatterChart", f"{{{C}}}radarChart"]

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def fix_negative_extents(data: bytes, name: str):
    """<a:ext> の負の cx/cy を flipH/flipV + 正値へ幾何学的に等価変換する。
    負の extent は OOXML 違反（ST_PositiveCoordinate）で PowerPoint 修復エラーの原因。"""
    if b'cx="-' not in data and b'cy="-' not in data:
        return data, 0
    _reject_dtd(data, name)
    root = _fromstring(data)
    n = 0
    for xfrm in root.iter(f"{{{A}}}xfrm"):
        off = xfrm.find(f"{{{A}}}off")
        ext = xfrm.find(f"{{{A}}}ext")
        if off is None or ext is None:
            continue
        cx, cy = int(ext.get("cx", "0")), int(ext.get("cy", "0"))
        if cx >= 0 and cy >= 0:
            continue
        ox, oy = int(off.get("x", "0")), int(off.get("y", "0"))
        if cx < 0:
            off.set("x", str(ox + cx))
            ext.set("cx", str(-cx))
            xfrm.set("flipH", "0" if xfrm.get("flipH") == "1" else "1")
        if cy < 0:
            off.set("y", str(oy + cy))
            ext.set("cy", str(-cy))
            xfrm.set("flipV", "0" if xfrm.get("flipV") == "1" else "1")
        n += 1
    if n == 0:
        return data, 0
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True), n


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
ET.register_namespace("p", P_NS)
_ADJ_MARK = re.compile(r"^(?P<base>.*)@(?P<adjs>adj\d+=-?\d+(?:,adj\d+=-?\d+)*)$")


def fix_callout_adjustments(data: bytes, name: str):
    """shape 名 `<名前>@adj1=..,adj2=..` の印を <a:prstGeom> の <a:avLst>/<a:gd> に変換する。
    pptxgenjs 4.0.1 はプリセット図形の adjustment（吹き出し尻尾位置等）を公開しないための後処理。
    印は注入後に shape 名から除去する（冪等: 2回目は印が無いので no-op）。"""
    if b"@adj" not in data:
        return data, 0
    _reject_dtd(data, name)
    root = _fromstring(data)
    n = 0
    for sp in root.iter(f"{{{P_NS}}}sp"):
        cnv = sp.find(f".//{{{P_NS}}}cNvPr")
        if cnv is None:
            continue
        m = _ADJ_MARK.match(cnv.get("name", ""))
        if not m:
            continue
        geom = sp.find(f".//{{{A}}}prstGeom")
        if geom is None:
            continue
        av = geom.find(f"{{{A}}}avLst")
        if av is None:
            av = ET.SubElement(geom, f"{{{A}}}avLst")
        else:
            for gd in list(av):
                av.remove(gd)
        for pair in m.group("adjs").split(","):
            key, val = pair.split("=")
            gd = ET.SubElement(av, f"{{{A}}}gd")
            gd.set("name", key)
            gd.set("fmla", f"val {val}")
        cnv.set("name", m.group("base"))
        n += 1
    if n == 0:
        return data, 0
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True), n


def fix_chart_xml(data: bytes, name: str = "chart.xml"):
    """(fixed_bytes, n_fixes) を返す。修正不要なら n_fixes=0。"""
    _reject_dtd(data, name)
    root = _fromstring(data)
    n = 0
    for parent_tag in SER_PARENTS:
        for chart in root.iter(parent_tag):
            for ser in chart.findall(f"{{{C}}}ser"):
                # 1) invertIfNegative は line/scatter/radar 系列では不正
                for bad in ser.findall(f"{{{C}}}invertIfNegative"):
                    ser.remove(bad)
                    n += 1
                # 2) marker は dLbls より前が正
                children = list(ser)
                marker = ser.find(f"{{{C}}}marker")
                dlbls = ser.find(f"{{{C}}}dLbls")
                if marker is not None and dlbls is not None:
                    if children.index(marker) > children.index(dlbls):
                        ser.remove(marker)
                        ser.insert(list(ser).index(dlbls), marker)
                        n += 1
    if n == 0:
        return data, 0
    out = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    return out, n


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    check_only = "--check" in sys.argv
    if len(args) != 1:
        print(__doc__)
        sys.exit(2)
    path = args[0]

    issues = 0
    fixes = {}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for name in names:
            data = z.read(name)
            if name.startswith("ppt/charts/chart") and name.endswith(".xml"):
                fixed, n = fix_chart_xml(data, name)
                if n:
                    issues += n
                    fixes[name] = fixed
                    print(f"chart schema fix: {name} ({n} issues)")
            if name.endswith(".xml") and (
                name.startswith("ppt/slides/")
                or name.startswith("ppt/slideLayouts/")
                or name.startswith("ppt/slideMasters/")
                or name.startswith("ppt/notesSlides/")
            ):
                fixed, n = fix_negative_extents(fixes.get(name, data), name)
                if n:
                    issues += n
                    fixes[name] = fixed
                    print(f"negative-extent fix: {name} ({n} shapes)")
            if name.endswith(".xml") and name.startswith("ppt/slides/"):
                fixed, n = fix_callout_adjustments(fixes.get(name, data), name)
                if n:
                    issues += n
                    fixes[name] = fixed
                    print(f"callout-adj fix: {name} ({n} shapes)")
            if name.endswith(".xml") and b'idx="4294967295"' in data:
                # 情報表示のみ（issues に数えない）: 実ファイルで PowerPoint が dt ph の
                # idx=4294967295 を許容する反例を確認済み（2026-07-13）。修復トリガーとしては
                # 未確定のため、検出しても失敗にはしない。
                print(f"note: {name} has ph idx=4294967295 (PowerPointは許容する例あり・情報のみ)")

    if check_only:
        print("check:", "NG" if issues else "OK")
        sys.exit(1 if issues else 0)

    if not fixes:
        print("sanitize: no changes needed")
        return

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp.close()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = fixes.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    shutil.move(tmp.name, path)
    print(f"sanitized: {path} ({issues} fixes)")


if __name__ == "__main__":
    main()
