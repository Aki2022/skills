"""audit_html's contrast check on a page with a dark `prefers-color-scheme` block.

The check used to collect every `:root{}` variable into one table, so the `:root` nested in
`@media (prefers-color-scheme:dark)` overwrote the base one and only the dark pairs were judged: a low-contrast light theme
passed. Each case here is a page the check must judge in BOTH themes. A page that passes must still be reported with the pairs
it judged (never a silent pass on nothing)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT / 'own-web-build/scripts'))

import audit_html  # noqa: E402

MIN = 4.5


def page(light_fg='#111111', light_bg='#ffffff', dark_fg='#eeeeee', dark_bg='#161616', extra_css='', dark_extra=''):
    return (f'<style>:root{{--fg:{light_fg};--bg:{light_bg}}}'
            f'body{{color:var(--fg);background:var(--bg)}}{extra_css}'
            f'@media (prefers-color-scheme:dark){{:root{{--fg:{dark_fg};--bg:{dark_bg}}}{dark_extra}}}</style>')


def judged(html, minimum=MIN, linked_css=''):
    return audit_html.check_contrast(html, minimum, linked_css)


def failing(results):
    return [label for ok, label in results if not ok]


def test_a_low_contrast_light_theme_fails_while_the_dark_theme_is_fine():
    results = judged(page(light_fg='#8fb0ff', light_bg='#ffffff'))
    assert failing(results), 'a pale blue on white passed: the light theme was not judged'
    assert any('#8fb0ff' in label for label in failing(results))


def test_a_low_contrast_dark_theme_still_fails_while_the_light_theme_is_fine():
    results = judged(page(dark_fg='#333333', dark_bg='#161616'))
    assert failing(results), 'a dark grey on near-black passed: the dark theme was not judged'


def test_a_page_whose_both_themes_have_enough_contrast_passes_and_reports_both_pairs():
    results = judged(page())
    assert results and failing(results) == []
    labels = ' '.join(label for _, label in results)
    assert '#111111 on #ffffff' in labels and '#eeeeee on #161616' in labels


def test_the_dark_variables_resolve_for_rules_that_live_inside_the_dark_media_block():
    # `.note` exists only in dark: its pair uses the dark values and must be judged against them, not the light ones
    html = page(dark_extra='.note{color:var(--fg);background:#2a2a2a}')
    assert failing(judged(html)) == []
    html = page(dark_fg='#3a3a3a', dark_extra='.note{color:var(--fg);background:#2a2a2a}')
    assert any('#3a3a3a on #2a2a2a' in label for label in failing(judged(html)))


def test_a_rule_that_lives_only_in_the_dark_block_is_not_judged_with_the_light_values():
    # `.note` is dark-only; judged with the light fg (#111111) on #2a2a2a it would fail, wrongly
    html = page(dark_extra='.note{color:var(--fg);background:#2a2a2a}')
    assert not any('#111111 on #2a2a2a' in label for _, label in judged(html))


def test_a_base_rule_with_a_variable_background_is_judged_in_each_theme():
    html = page(extra_css='.card{color:#999999;background:var(--bg)}')
    labels = [label for _, label in judged(html)]
    # #999 on white = 2.85 (fails), #999 on #161616 = 6.35 (passes): only the light one fails
    assert any('#999999 on #ffffff' in label for label in failing(judged(html)))
    assert any('#999999 on #161616' in label for label in labels)
    assert not any('#999999 on #161616' in label for label in failing(judged(html)))


def test_a_page_without_a_dark_block_is_judged_as_before():
    html = '<style>:root{--fg:#8fb0ff;--bg:#ffffff}body{color:var(--fg);background:var(--bg)}</style>'
    assert failing(judged(html))
    html = '<style>:root{--fg:#111111;--bg:#ffffff}body{color:var(--fg);background:var(--bg)}</style>'
    results = judged(html)
    assert results and failing(results) == []


def test_a_dark_block_without_root_variables_does_not_hide_the_base_pairs():
    html = ('<style>:root{--fg:#8fb0ff;--bg:#ffffff}body{color:var(--fg);background:var(--bg)}'
            '@media (prefers-color-scheme:dark){.x{color:#fff;background:#000}}</style>')
    assert failing(judged(html))


def test_other_media_blocks_stay_in_the_base_theme():
    html = ('<style>:root{--fg:#111111;--bg:#ffffff}'
            '@media (min-width:600px){.wide{color:var(--fg);background:var(--bg)}}'
            '@media (prefers-color-scheme:dark){:root{--fg:#eeeeee;--bg:#161616}}</style>')
    labels = [label for _, label in judged(html)]
    assert any('#111111 on #ffffff' in label for label in labels)


def test_the_linked_css_and_inline_styles_get_the_same_treatment():
    css = ':root{--fg:#8fb0ff;--bg:#ffffff}.a{color:var(--fg);background:var(--bg)}@media (prefers-color-scheme:dark){:root{--fg:#eee;--bg:#161616}}'
    assert any('#8fb0ff' in label for label in failing(judged('<p>x</p>', linked_css=css)))
    html = page(light_fg='#8fb0ff') + '<p style="color:var(--fg);background:var(--bg)">x</p>'
    inline = [(ok, label) for ok, label in judged(html) if label.startswith('[inline]')]
    assert len(inline) >= 2 and failing(inline), 'the inline pair was judged in one theme only'


def test_the_nested_braces_of_the_dark_block_do_not_swallow_the_rules_after_it():
    html = page(extra_css='') + '<style>.after{color:#8fb0ff;background:#ffffff}</style>'
    assert any('#8fb0ff on #ffffff' in label for label in failing(judged(html)))
    css = page().replace('</style>', '.tail{color:#8fb0ff;background:#ffffff}</style>')
    assert any('#8fb0ff on #ffffff' in label for label in failing(judged(css)))


def test_a_dark_default_with_a_light_media_override_is_judged_in_both_themes():
    # the mirror arrangement: the base :root is the dark theme and `prefers-color-scheme:light` overrides it
    css = ('<style>:root{--fg:#eeeeee;--bg:#161616}body{color:var(--fg);background:var(--bg)}'
           '@media (prefers-color-scheme:light){:root{--fg:#8fb0ff;--bg:#ffffff}}</style>')
    assert any('#8fb0ff on #ffffff' in label for label in failing(judged(css))), 'the light override was not judged'
    ok = css.replace('#8fb0ff', '#111111')
    results = judged(ok)
    labels = ' '.join(label for _, label in results)
    assert failing(results) == [] and '#eeeeee on #161616' in labels and '#111111 on #ffffff' in labels
    low_base = css.replace('#eeeeee', '#333333')
    assert any('#333333 on #161616' in label for label in failing(judged(low_base))), 'the dark base was not judged'


def test_a_page_with_both_a_dark_and_a_light_media_block_judges_each_overlay_on_the_base():
    html = ('<style>:root{--fg:#111111;--bg:#ffffff}body{color:var(--fg);background:var(--bg)}'
            '@media (prefers-color-scheme:dark){:root{--fg:#eeeeee;--bg:#161616}}'
            '@media (prefers-color-scheme:light){:root{--fg:#8fb0ff}}</style>')
    results = judged(html)
    assert any('#8fb0ff on #ffffff' in label for label in failing(results))
    labels = ' '.join(label for _, label in results)
    assert '#111111 on #ffffff' in labels and '#eeeeee on #161616' in labels


def test_the_media_query_may_carry_other_conditions_and_other_spacing():
    html = ('<style>:root{--fg:#111111;--bg:#ffffff}body{color:var(--fg);background:var(--bg)}'
            '@media screen and ( prefers-color-scheme : dark ){:root{--fg:#333333;--bg:#161616}}</style>')
    assert any('#333333 on #161616' in label for label in failing(judged(html)))
