#!/usr/bin/env python3
"""
sanitize_pptx.py — pptxgenjs 産 pptx の「PowerPoint 修復エラー」要因をビルド後に矯正する。

③ネイティブビルドの必須最終工程（チャートの有無に関わらず実行してよい・冪等）。
根拠（2026-07-13 実測、pptxgenjs 4.0.1）:
  1) lineChart / scatterChart / radarChart の <c:ser> に <c:invertIfNegative>（bar系専用
     要素）が混入する → スキーマ違反で PowerPoint が修復を要求（LibreOffice/python-pptx
     は許容するため④のレンダリング検証では検出できない）
  2) <c:marker> が <c:dLbls> の後に出力される（CT_LineSer 等の子要素順序は marker が先）
  3) bar 系で chartColors を点ごとに指定すると <c:dPt> が <c:dLbls> の後に出力される（dPt が先・全チャート種）
  4) 2D の bar/line/area/radar にも 3 つ目の <c:axId>（系列軸）が出るが <c:serAx> が無い → 3 つ目を落とす
  5) defineSlideMaster の slideNumber オプションは idx="4294967295" の不正プレース
     ホルダを吐く（deck_helpers v3 は不使用。検出したら警告のみ＝手動対応）
  6) スライド内で <p:cNvPr id> が重複する（表＋図形の組合せで発生）→ 2 回目以降を振り直す
  7) pptxgenjs はプリセット図形の adjustment 値（吹き出しの尻尾位置等）を公開しないため、
     deck_helpers.speechBubble() 等は shape 名に `<名前>@adj1=..,adj2=..` とエンコードする。
     本スクリプトが <a:avLst> へ <a:gd> を注入して印を除去する（2026-08-27 追加。
     --check は未注入の印を NG として検出する）
  8) round2SameRect の調整名 adj → adj1（仕様では adj1/adj2。PowerPoint の既定と違う半径を取りこぼさないため）

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


# <p:cNvPr> の開始タグと、その中の id 属性。属性の並びに依らず id を拾う（name が先に来ても重複を見逃さない）。
_CNV_TAG_RE = re.compile(r"<p:cNvPr\b[^>]*>")
_CNV_ID_RE = re.compile(r'(\s)id="(\d+)"')


def fix_duplicate_shape_ids(data: bytes, name: str):
    """スライド内で重複する <p:cNvPr id> を振り直す（(fixed_bytes, n_fixes)）。
    PowerPoint はスライド内の図形 ID の重複で「修復」を要求する（2026-10-19 実測、pptxgenjs 4.0.1 の表＋図形の組合せで発生）。
    最大 ID + 1 から順に、2 回目以降の出現だけを書き換える。ID は同スライド内でしか参照されない。"""
    _reject_dtd(data, name)
    text = data.decode("utf-8")
    ids = [m.group(2) for tag in _CNV_TAG_RE.finditer(text) for m in _CNV_ID_RE.finditer(tag.group(0))]
    if len(ids) == len(set(ids)):
        return data, 0
    next_id = max(int(i) for i in ids) + 1
    seen = set()
    n = 0

    def _sub(m):
        nonlocal next_id, n
        cur = m.group(2)
        if cur in seen:
            new = str(next_id)
            next_id += 1
            n += 1
            return f'{m.group(1)}id="{new}"'
        seen.add(cur)
        return m.group(0)

    text = _CNV_TAG_RE.sub(lambda t: _CNV_ID_RE.sub(_sub, t.group(0), count=1), text)
    return text.encode("utf-8"), n


# 仕様（presetShapeDefinitions）で調整ハンドルが adj1/adj2 の図形。pptxgenjs は adj の名前で出す。
_ROUND2_SHAPES = {"round2SameRect"}


def fix_round2_adj_names(data: bytes, name: str):
    """round2SameRect の調整ハンドル名 adj を adj1 にする（(fixed_bytes, n_fixes)）。
    仕様の preset 定義では round2SameRect の調整は adj1 と adj2 で、adj は無い（2026-11-06 に LibreOffice 版の
    presetShapeDefinitions で確認）。roundRect の adj は仕様どおりなので触らない。adj2 は補わない（既定値 0）。"""
    _reject_dtd(data, name)
    root = _fromstring(data)
    n = 0
    for geom in root.iter(f"{{{A}}}prstGeom"):
        if geom.get("prst") in _ROUND2_SHAPES:
            for gd in geom.iter(f"{{{A}}}gd"):
                if gd.get("name") == "adj":
                    gd.set("name", "adj1")
                    n += 1
    if n == 0:
        return data, 0
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True), n


# CT_*Ser（ECMA-376 DrawingML Chart）の子要素の順序。同じ rank の要素は、同じ種類の系列には同時に現れない。
SER_CHILD_RANK = {
    "idx": 0, "order": 1, "tx": 2, "spPr": 3,
    "invertIfNegative": 4, "explosion": 4, "pictureOptions": 4, "marker": 4,
    "dPt": 5, "dLbls": 6, "trendline": 7, "errBars": 8,
    "cat": 9, "xVal": 9, "val": 10, "yVal": 10,
    "smooth": 11, "shape": 11, "bubbleSize": 11,
    "bubble3D": 12, "extLst": 13,
}


def _reorder_ser_children(ser) -> int:
    """<c:ser> の子要素を SER_CHILD_RANK の順に並べ替える（並べ替えたら 1、不要なら 0）。
    未知の子要素がある系列は触らない。スキーマに無い要素の位置を推測で動かさないため。"""
    children = list(ser)
    names = [c.tag.rsplit("}", 1)[-1] for c in children]
    if any(name not in SER_CHILD_RANK for name in names):
        return 0
    ordered = sorted(children, key=lambda c: SER_CHILD_RANK[c.tag.rsplit("}", 1)[-1]])
    if all(a is b for a, b in zip(ordered, children)):
        return 0
    for c in children:
        ser.remove(c)
    ser.extend(ordered)
    return 1


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
    # 2) 系列の子要素の順序（marker・dPt は dLbls より前、dLbls は cat/val より前 等）。全チャート種の ser が対象。
    #    dPt の順序違いは PowerPoint が「修復」を要求する（2026-10-19 実測。LibreOffice は黙って描く）。
    #    判定と修正を同じ関数で行うため、--check と修正の結果は食い違わない。
    for ser in root.iter(f"{{{C}}}ser"):
        n += _reorder_ser_children(ser)
    # 3) 2D の bar/line/area/radar に 3 つ目の <c:axId>（系列軸 ID）が出るが、対応する <c:serAx> が無い
    #    （pptxgenjs 4.0.1 は bar3D 以外でも無条件に 3 つ出す。CT_BarChart 等の axId は 2 つ固定で、
    #    PowerPoint が「修復」を要求する。2026-10-19 実測）。serAx が無ければ 3 つ目以降の axId を落とす
    ser_ax_ids = {ax.find(f"{{{C}}}axId").get("val") for ax in root.iter(f"{{{C}}}serAx") if ax.find(f"{{{C}}}axId") is not None}
    for tag in ("barChart", "bar3DChart", "lineChart", "areaChart", "radarChart"):
        for chart in root.iter(f"{{{C}}}{tag}"):
            ax_ids = chart.findall(f"{{{C}}}axId")
            for extra in ax_ids[2:]:
                if extra.get("val") not in ser_ax_ids:
                    chart.remove(extra)
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
            # inject_template.py 後は chart1 が g_chart1.xml 等に改名されるため、charts/ 直下の .xml を全て対象にする
            if name.startswith("ppt/charts/") and "/_rels/" not in name and name.endswith(".xml"):
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
                fixed, n = fix_round2_adj_names(fixes.get(name, data), name)
                if n:
                    issues += n
                    fixes[name] = fixed
                    print(f"round2 adj fix: {name} ({n} shapes)")
                fixed, n = fix_duplicate_shape_ids(fixes.get(name, data), name)
                if n:
                    issues += n
                    fixes[name] = fixed
                    print(f"shape-id fix: {name} ({n} duplicate ids)")
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
