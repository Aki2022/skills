"""Contrast is computed in one place. Before ISSUE-06 the same WCAG luminance formula was written three times
(a static pass in audit_html.py, the rendered measurement, and the component-boundary rules); the static one could only
read literal CSS pairs, so it agreed with the rendered measurement by luck. Now the formula lives in the design skill's
contrast kit, the rendered measurement and the rule collectors use it, and audit_html.py no longer judges contrast."""
import re
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[3]
KIT = 'own-design-author/scripts/contrast_kit.mjs'
CODE = ('.py', '.mjs', '.js', '.cjs')
FORMULA = re.compile(r'0\.2126')


def code_files():
    """Code of every skill (tracked, and new files not yet tracked), tests and fixtures excluded (a test may spell the formula to check a value)."""
    try:
        listing = subprocess.run(['git', '-C', str(SKILLS), 'ls-files', '--cached', '--others', '--exclude-standard'],
                                 capture_output=True, text=True, check=True).stdout.splitlines()
    except (OSError, subprocess.CalledProcessError):  # a copy of the skills without git history
        listing = [str(p.relative_to(SKILLS)) for p in SKILLS.rglob('*') if p.is_file()]
    files = [f for f in listing if f.endswith(CODE)
             and '/tests/' not in f and not Path(f).name.startswith('test_') and 'node_modules' not in f]
    return files


def test_the_luminance_formula_is_written_in_exactly_one_file():
    files = code_files()
    assert len(files) > 100, f'scanned {len(files)} files; the listing is not the repository'
    hits = [f for f in files if FORMULA.search((SKILLS / f).read_text(errors='replace'))]
    print(f'scanned {len(files)} code files, {len(hits)} hold the luminance formula: {hits}')
    assert hits == [KIT], hits


def test_the_rendered_measurement_and_the_rule_collectors_use_the_kit_and_nothing_else():
    for rel in ('own-web-build/scripts/measure_render.mjs', 'own-design-author/scripts/resolved_contract.mjs'):
        text = (SKILLS / rel).read_text()
        assert 'contrast_kit' in text or '__contrastKit' in text, rel + ' does not use the contrast kit'
        assert not re.search(r'\*\*\s*2\.4', text), rel + ' still computes the sRGB curve itself'


def test_audit_html_no_longer_judges_contrast_and_says_it_did_not(tmp_path, capsys):
    """A page whose CSS holds a literal low-contrast pair passes L1 now; L1 must say contrast is only measured after
    rendering, so a pass is never read as a contrast check."""
    sys.path.insert(0, str(SKILLS / 'own-web-build/scripts'))
    import audit_html
    page = tmp_path / 'low.html'
    page.write_text('<!doctype html><html lang="en"><head><meta name="viewport" content="width=device-width">'
                    '<title>t</title><style>:focus-visible{outline:2px solid #000}@media(min-width:1px){main{margin:0}}'
                    'p{color:#777777;background:#ffffff}</style></head>'
                    '<body><main><h1>t</h1><p>low</p></main></body></html>')
    manifest = {'minimums': {'text_contrast': 4.5}, 'coverage': {'uncovered': []}}
    ok = audit_html.audit(str(page), manifest)
    out = capsys.readouterr().out
    assert ok is True, out
    assert not re.search(r'\[(PASS|FAIL)\][^\n]*contrast', out), 'L1 still reports a contrast verdict:\n' + out
    assert re.search(r'未カバー[^\n]*コントラスト', out), 'L1 must list contrast as not covered here:\n' + out
    assert not hasattr(audit_html, 'contrast_ratio') and not hasattr(audit_html, 'check_contrast')
