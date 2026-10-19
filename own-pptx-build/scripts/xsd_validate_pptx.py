#!/usr/bin/env python3
"""pptx の全 XML 部品を ISO/IEC 29500-4 (transitional) の XSD で検証する（納品前の必須ゲート）。

usage: xsd_validate_pptx.py <file.pptx> [--schemas <dir>]
  exit 0 = 全部品 VALID、1 = INVALID あり、2 = スキーマや入力の不備

LibreOffice / python-pptx は多くのスキーマ違反を黙って通すが、PowerPoint は「修復」を要求する
（2026-10-19 実測: 段落途中の a:pPr、3 つ目の axId、dPt の順序、cNvPr 重複 など）。
スキーマは `--schemas` か環境変数 OOXML_SCHEMA_DIR、無ければ Anthropic 製 pptx skill のキャッシュ
（~/.claude*/plugins/cache/**/skills/pptx/scripts/office/schemas/ISO-IEC29500-4_2016）を探す。
"""
import argparse
import glob
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
    cands = []
    if explicit:
        cands.append(explicit)
    if os.environ.get("OOXML_SCHEMA_DIR"):
        cands.append(os.environ["OOXML_SCHEMA_DIR"])
    home = os.path.expanduser("~")
    cands += sorted(glob.glob(os.path.join(home, ".claude*", "plugins", "cache", "**", "skills", "pptx", "scripts", "office", "schemas", "ISO-IEC29500-4_2016"), recursive=True))
    for c in cands:
        if os.path.isfile(os.path.join(c, "pml.xsd")):
            return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--schemas", default=None)
    a = ap.parse_args()
    try:
        from lxml import etree
    except ImportError:
        print("xsd_validate: lxml が要る（pip install lxml）", file=sys.stderr)
        return 2
    sdir = find_schema_dir(a.schemas)
    if not sdir:
        print("xsd_validate: スキーマが見つからない。--schemas か OOXML_SCHEMA_DIR で ISO-IEC29500-4_2016 の dir を指定", file=sys.stderr)
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
    return 1 if invalid else 0


if __name__ == "__main__":
    sys.exit(main())
