#!/usr/bin/env python3
"""
inject_template.py — 生成pptx（PptxGenJS産）を template_v3.pptx へ移植する（テーマ注入 Phase 1）。

方式: 「生成pptxにテーマを注入」ではなく**「テンプレを土台に生成スライドを移植」**する。
テンプレ側の presentation/master/layouts/theme は PowerPoint 純正の部品なので、
移植後のファイルは (a) 人間が「新しいスライド」で会社レイアウト（11種）を使える
(b) テーマの色/フォントに 44546A・Noto が並ぶ (c) 修復エラーの温床を持たない。

処理:
  1. template_v3.pptx を読み、そのサンプルスライドを全削除
  2. 生成pptxの各スライドを順に移植（画像・チャート・埋込xlsxも再帰的にコピー）
  3. 各スライドのレイアウト参照をテンプレの受け皿レイアウト（既定 16_Blank）へ張り替え
     （notesSlide参照は落とす）
  4. [Content_Types] / presentation.xml / rels を整合
     （PowerPoint純正部品を壊さないため、presentation.xml と [Content_Types] は
      文字列手術でバイト温存し、再シリアライズは rels のみに限定する）

前提（deck_helpers v3 と対で設計）:
  - 日付・ページ番号は生成側がスライド直描き → 移植後も保たれる
  - コピーライトは生成側マスターにのみ存在 → 移植で消え、テンプレのレイアウト側
    コピーライトが表示される（二重にならない）

使い方:
  python3 inject_template.py <generated.pptx> <template_v3.pptx> <out.pptx> [--layout 16_Blank]
実行後は sanitize_pptx.py --check と PowerPoint での repair なし確認を行うこと。
"""
import re
import sys
import posixpath
import zipfile

# XXE/billion-laughs 対策: defusedxml があれば優先。無い環境でも _reject_dtd ガードで
# DTD/ENTITY を含む入力を拒否する（正規の pptx 部品は DTD を含まない）。
import xml.etree.ElementTree as ET

try:
    import defusedxml.ElementTree as _SAFE_ET  # type: ignore

    def _fromstring(data):
        return _SAFE_ET.fromstring(data)
except ImportError:

    def _fromstring(data):
        return ET.fromstring(data)


def _reject_dtd(data: bytes, name: str) -> None:
    head = data[:4096]
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise SystemExit(f"refusing to parse {name}: DTD/ENTITY declaration found")


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", REL_NS)

SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RT_SLIDE = f"{R_NS}/slide"
RT_LAYOUT = f"{R_NS}/slideLayout"
RT_NOTES = f"{R_NS}/notesSlide"


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def resolve(base_part, target):
    # pptxgenjs は Target をパッケージ絶対（先頭 "/"）で書くことがある
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def parse_rels(data, name):
    _reject_dtd(data, name)
    return _fromstring(data)


def tostr(root):
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    layout_name = "16_Blank"
    if "--layout" in sys.argv:
        layout_name = sys.argv[sys.argv.index("--layout") + 1]
    if len(args) != 3:
        print(__doc__)
        sys.exit(2)
    gen_path, tpl_path, out_path = args

    base = read_zip(tpl_path)
    gen = read_zip(gen_path)

    # ---- 1. base のサンプルスライドを全削除 ----
    pres_rels = parse_rels(base["ppt/_rels/presentation.xml.rels"], "base pres rels")
    pres_xml = base["ppt/presentation.xml"].decode("utf-8")
    ct_xml = base["[Content_Types].xml"].decode("utf-8")

    removed_parts = []
    for rel in list(pres_rels):
        if rel.get("Type") == RT_SLIDE:
            removed_parts.append(resolve("ppt/presentation.xml", rel.get("Target")))
            pres_rels.remove(rel)
    # presentation.xml: sldIdLst の中身を空にする（文字列手術・他はバイト温存）
    pres_xml = re.sub(r"<p:sldId [^>]*/>", "", pres_xml)
    for part in removed_parts:
        base.pop(part, None)
        base.pop(rels_path(part), None)
        ct_xml = ct_xml.replace(f'<Override PartName="/{part}" ContentType="{SLIDE_CT}"/>', "")
    print(f"template slides removed: {len(removed_parts)}")

    # 受け皿レイアウトの partname
    blank_layout = None
    for name in sorted(base):
        if name.startswith("ppt/slideLayouts/slideLayout") and name.endswith(".xml"):
            m = re.search(rb'<p:cSld[^>]*name="([^"]*)"', base[name])
            if m and m.group(1).decode("utf-8") == layout_name:
                blank_layout = name
                break
    if not blank_layout:
        raise SystemExit(f"layout {layout_name!r} not found in template")

    # ---- 2. gen のスライドを順に移植 ----
    gen_pres_rels = parse_rels(gen["ppt/_rels/presentation.xml.rels"], "gen pres rels")
    gen_rid2part = {
        rel.get("Id"): resolve("ppt/presentation.xml", rel.get("Target"))
        for rel in gen_pres_rels
        if rel.get("Type") == RT_SLIDE
    }
    order = re.findall(rb'<p:sldId [^>]*r:id="([^"]+)"', gen["ppt/presentation.xml"])
    gen_slides = [gen_rid2part[rid.decode("utf-8")] for rid in order]

    def ensure_default(ext, ctype):
        nonlocal ct_xml
        if f'Extension="{ext}"' not in ct_xml:
            ct_xml = ct_xml.replace(
                "</Types>", f'<Default Extension="{ext}" ContentType="{ctype}"/></Types>'
            )

    def add_override(partname, ctype):
        nonlocal ct_xml
        if f'PartName="/{partname}"' not in ct_xml:
            ct_xml = ct_xml.replace(
                "</Types>",
                f'<Override PartName="/{partname}" ContentType="{ctype}"/></Types>',
            )

    copied = {}  # gen partname -> base partname

    def copy_part(part):
        """gen の部品（media/chart/embedding等）を base へ再帰コピーし、base での名前を返す"""
        if part in copied:
            return copied[part]
        name = part
        if name in base and base[name] != gen[part]:
            d, b = posixpath.split(part)
            i = 0
            name = posixpath.join(d, "g_" + b)
            while name in base:
                i += 1
                name = posixpath.join(d, f"g{i}_" + b)
        base[name] = gen[part]
        copied[part] = name
        if part.startswith("ppt/charts/") and part.endswith(".xml"):
            add_override(name, CHART_CT)
        if part.endswith(".xlsx"):
            ensure_default("xlsx", XLSX_CT)
        rp = rels_path(part)
        if rp in gen:
            rroot = parse_rels(gen[rp], rp)
            for rel in rroot:
                if rel.get("TargetMode") == "External":
                    continue
                tgt = resolve(part, rel.get("Target"))
                new_tgt = copy_part(tgt)
                rel.set("Target", posixpath.relpath(new_tgt, posixpath.dirname(part)))
            base[rels_path(name)] = tostr(rroot)
        return name

    new_sldids = []
    new_rels = []
    for k, gpart in enumerate(gen_slides, 1):
        new_part = f"ppt/slides/slide{k}.xml"
        base[new_part] = gen[gpart]
        add_override(new_part, SLIDE_CT)
        rroot = parse_rels(gen[rels_path(gpart)], rels_path(gpart))
        for rel in list(rroot):
            t = rel.get("Type")
            if t == RT_LAYOUT:
                rel.set("Target", posixpath.relpath(blank_layout, "ppt/slides"))
            elif t == RT_NOTES:
                rroot.remove(rel)
            elif rel.get("TargetMode") != "External":
                tgt = resolve(gpart, rel.get("Target"))
                new_tgt = copy_part(tgt)
                rel.set("Target", posixpath.relpath(new_tgt, "ppt/slides"))
        base[rels_path(new_part)] = tostr(rroot)
        rid = f"rIdGen{k}"
        new_rels.append(f'<Relationship Id="{rid}" Type="{RT_SLIDE}" Target="slides/slide{k}.xml"/>')
        new_sldids.append(f'<p:sldId id="{255 + k}" r:id="{rid}"/>')

    ensure_default("png", "image/png")
    ensure_default("jpeg", "image/jpeg")

    # presentation.xml: sldIdLst へ挿入（空になったリストにも対応）
    if "</p:sldIdLst>" in pres_xml:
        pres_xml = pres_xml.replace("</p:sldIdLst>", "".join(new_sldids) + "</p:sldIdLst>")
    else:
        pres_xml = pres_xml.replace(
            "<p:sldIdLst/>", "<p:sldIdLst>" + "".join(new_sldids) + "</p:sldIdLst>"
        )
    base["ppt/presentation.xml"] = pres_xml.encode("utf-8")
    base["[Content_Types].xml"] = ct_xml.encode("utf-8")

    # base presentation rels: ET再シリアライズを避け、文字列で追記
    rels_xml = base["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    # 既存のslide relを除去
    rels_xml = re.sub(r'<Relationship [^>]*Type="[^"]*/slide"[^>]*/>', "", rels_xml)
    rels_xml = rels_xml.replace("</Relationships>", "".join(new_rels) + "</Relationships>")
    base["ppt/_rels/presentation.xml.rels"] = rels_xml.encode("utf-8")

    # ---- 3. GC: 削除したサンプルスライドだけが参照していた孤児部品を mark&sweep で掃除 ----
    GC_DIRS = ("ppt/charts/", "ppt/embeddings/", "ppt/media/", "ppt/notesSlides/")

    def targets_of(part):
        rp = rels_path(part)
        if rp not in base:
            return []
        out = []
        for m in re.finditer(rb'Target="([^"]+)"', base[rp]):
            tgt = m.group(1).decode("utf-8")
            if not tgt.startswith(("http", "mailto")):
                out.append(resolve(part, tgt))
        return out

    roots = [n for n in base if not n.endswith(".rels") and not n.startswith(GC_DIRS)]
    reachable = set(roots)
    frontier = list(roots)
    while frontier:
        part = frontier.pop()
        for tgt in targets_of(part):
            if tgt in base and tgt not in reachable:
                reachable.add(tgt)
                frontier.append(tgt)
    gc = 0
    for name in list(base):
        part = name
        if name.endswith(".rels"):
            part = posixpath.join(
                posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[: -len(".rels")]
            )
        if part.startswith(GC_DIRS) and part not in reachable:
            del base[name]
            gc += 1
            if not name.endswith(".rels"):
                ct_xml = re.sub(rf'<Override PartName="/{re.escape(part)}"[^>]*/>', "", ct_xml)
    if gc:
        base["[Content_Types].xml"] = ct_xml.encode("utf-8")
        print(f"gc: removed {gc} orphan parts")

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in base.items():
            z.writestr(name, data)
    print(f"injected {len(gen_slides)} slides onto {layout_name}: {out_path}")


if __name__ == "__main__":
    main()
