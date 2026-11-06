#!/usr/bin/env python3
"""pptx の全 XML 部品を ISO/IEC 29500-4 (transitional) の XSD で検証する（納品前の必須ゲート）。

usage: xsd_validate_pptx.py <file.pptx> [--schemas <dir>]
  exit 0 = 全部品 VALID、1 = INVALID あり、2 = スキーマや入力の不備

LibreOffice / python-pptx は多くのスキーマ違反を黙って通すが、PowerPoint は「修復」を要求する
（2026-10-19 実測: 段落途中の a:pPr、3 つ目の axId、dPt の順序、cNvPr 重複 など）。
スキーマの場所は `--schemas` か環境変数 OOXML_SCHEMA_DIR で明示する。既定の場所は探さない。
場所が無ければ exit 2 で止まる（黙って合格にしない）。
"""
import argparse
import os
import sys
import tempfile
import zipfile

NS_TO_XSD = {
    "http://schemas.openxmlformats.org/presentationml/2006/main": "pml.xsd",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "dml-chart.xsd",
    "http://schemas.openxmlformats.org/drawingml/2006/main": "dml-main.xsd",
}


def find_schema_dir(explicit):
    """スキーマの場所を --schemas か環境変数 OOXML_SCHEMA_DIR から決める。既定の場所は探さない。
    (dir, None) か (None, 理由) を返す。"""
    loc = explicit or os.environ.get("OOXML_SCHEMA_DIR")
    if not loc:
        return None, "スキーマの場所が未指定。--schemas か環境変数 OOXML_SCHEMA_DIR で ISO-IEC29500-4_2016 の dir を指定（既定の場所は探さない）"
    if not os.path.isfile(os.path.join(loc, "pml.xsd")):
        return None, f"指定された場所に pml.xsd が無い: {loc}"
    return loc, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--schemas", default=None)
    a = ap.parse_args()
    sdir, why = find_schema_dir(a.schemas)
    if not sdir:
        print(f"xsd_validate: {why}", file=sys.stderr)
        return 2
    try:
        from lxml import etree
    except ImportError:
        print("xsd_validate: lxml が要る（pip install lxml）", file=sys.stderr)
        return 2
    schemas = {}
    checked = invalid = 0
    skipped = []
    with zipfile.ZipFile(a.pptx) as z:
        for name in sorted(z.namelist()):
            if not (name.startswith("ppt/") and name.endswith(".xml")):
                continue
            data = z.read(name)
            try:
                doc = etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            except etree.XMLSyntaxError as e:
                invalid += 1
                print(f"INVALID {name}: XML parse error: {e}")
                continue
            ns = etree.QName(doc).namespace
            xsd = NS_TO_XSD.get(ns)
            if not xsd:
                skipped.append(name)
                continue
            if xsd not in schemas:
                schemas[xsd] = etree.XMLSchema(etree.parse(os.path.join(sdir, xsd)))
            s = schemas[xsd]
            checked += 1
            if not s.validate(doc):
                invalid += 1
                print(f"INVALID {name}")
                for e in list(s.error_log)[:5]:
                    print(f"   line {e.line}: {e.message[:200]}")
    print(f"xsd_validate: checked={checked} invalid={invalid} skipped={len(skipped)} schemas={sdir}")
    if checked == 0:
        # 対象 0 件は合格ではない（壊れた・空のパッケージが緑になる）
        print("xsd_validate: 検証できた部品が 0 件。pptx として読めていない", file=sys.stderr)
        return 2
    return 1 if invalid else 0


if __name__ == "__main__":
    sys.exit(main())
