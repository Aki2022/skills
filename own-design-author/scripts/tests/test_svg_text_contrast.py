"""SVG chart text is judged by its painted fill against the pixels behind it.

axe-core cannot determine a background inside an <svg> (an "image node") and leaves such text
incomplete, so every chart page used to fail as unverified. The measurement now reads the fill
and the rendered pixels itself, and only hands axe's incomplete nodes over when it measured
at least as many SVG text elements.
"""
import pytest
from test_core import project  # noqa: F401  (pytest fixture)
from test_measure_endpoints import BROWSER, NORMAL, measure, prepare

pytestmark = pytest.mark.skipif(not BROWSER, reason='explicit real browser run required')


def chart(body):
    return NORMAL.replace('CONTENT', '<svg width="240" height="60" role="img" aria-label="chart">' + body + '</svg>')


def ids(result):
    return {row['id'] for row in result['failures']}


def contrast_rows(result):
    return [row for row in result['failures'] if row['id'] == 'contrast']


def test_readable_svg_text_is_measured_instead_of_left_unverified(project):
    root = prepare(project)
    run, result = measure(root, chart('<text x="8" y="36" fill="#000000" font-size="16">Label</text>'), 'product-ui')
    assert 'axe-incomplete' not in ids(result), result['failures']
    assert 'contrast' not in ids(result), result['failures']
    assert run.returncode == 0, run.stdout + run.stderr


def test_svg_text_uses_its_fill_not_the_inherited_css_color(project):
    root = prepare(project)
    # CSS color stays black; only the painted fill is pale.
    run, result = measure(root, chart('<text x="8" y="36" fill="#eeeeee" font-size="16">Label</text>'), 'product-ui')
    assert run.returncode == 1
    assert any(row.get('sample') and any(s.get('measuredBy') == 'svg-fill' for s in row['sample'])
               for row in contrast_rows(result)), result['failures']


def test_svg_text_background_comes_from_the_shapes_behind_it(project):
    root = prepare(project)
    # Black text on a black rectangle inside the chart: the card behind the svg is white.
    body = ('<rect x="0" y="0" width="240" height="60" fill="#000000"/>'
            '<text x="8" y="36" fill="#000000" font-size="16">Label</text>')
    run, result = measure(root, chart(body), 'product-ui')
    assert run.returncode == 1
    assert contrast_rows(result), result['failures']


def test_svg_text_over_a_gradient_stays_unverified_not_green(project):
    root = prepare(project)
    body = ('<defs><linearGradient id="g"><stop offset="0" stop-color="#000000"/>'
            '<stop offset="1" stop-color="#ffffff"/></linearGradient>'
            '<clipPath id="c"><rect x="0" y="0" width="240" height="60" fill="#000000"/></clipPath></defs>'
            '<rect x="0" y="0" width="240" height="60" fill="url(#g)"/>'
            '<text x="8" y="36" fill="#000000" font-size="16">Label</text>')
    run, result = measure(root, chart(body), 'product-ui')
    assert run.returncode == 1
    assert any(s.get('skipped') for row in contrast_rows(result) for s in row['sample']), result['failures']


def test_shapes_inside_clip_paths_are_not_a_background(project):
    root = prepare(project)
    # The black clip rectangle is never painted, so black text sits on the white card.
    body = ('<defs><clipPath id="c"><rect x="0" y="0" width="240" height="60" fill="#000000"/></clipPath></defs>'
            '<text x="8" y="36" fill="#000000" font-size="16">Label</text>')
    run, result = measure(root, chart(body), 'product-ui')
    assert 'contrast' not in ids(result), result['failures']
    assert run.returncode == 0, run.stdout + run.stderr
