"""IA outputs against design's page inventory (ADR-20261115-design-ia-two-skills-shared-base, biz_ops; instead of
merging the two declarations, check that they agree).

A shipped HTML output that design's validation.json pages neither covers nor excludes is a page the design gate
does not check: warn. One design excludes is still checked by IA, and is reported. No pages declared: nothing to compare.
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_check import check, codes, write_repo  # noqa: E402
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)


def pages(root, covered=None, excluded=None):
    path = layer_paths.path(root, 'validation')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({'pages': {
        'sources': ['site/**/*.html'], 'covered': covered or {}, 'excluded': excluded or {}}}), encoding='utf-8')


def test_shipped_output_missing_from_the_page_inventory_warns(tmp_path):
    write_repo(tmp_path)
    pages(tmp_path, covered={'site/other.html': ['html']})
    rc, report = check(tmp_path)
    assert 'IA_PAGES_OUTPUT_NOT_INVENTORIED' in codes(report, 'warn'), report['findings']


def test_covered_output_is_quiet(tmp_path):
    write_repo(tmp_path)
    pages(tmp_path, covered={'site/index.html': ['html']})
    rc, report = check(tmp_path)
    assert not [f for f in report['findings'] if f['code'].startswith('IA_PAGES_')], report['findings']


def test_output_design_excludes_is_reported_but_ia_still_checks_it(tmp_path):
    write_repo(tmp_path)
    pages(tmp_path, excluded={'site/index.html': '外部で描く'})
    rc, report = check(tmp_path)
    assert 'IA_PAGES_OUTPUT_EXCLUDED_BY_DESIGN' in codes(report, 'info'), report['findings']
    assert rc == 0


def test_no_page_inventory_compares_nothing(tmp_path):
    write_repo(tmp_path)
    rc, report = check(tmp_path)
    assert rc == 0 and not [f for f in report['findings'] if f['code'].startswith('IA_PAGES_')]
