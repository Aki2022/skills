"""Page inventory: every page source in the repository is either validated or excluded with a reason."""
import json
from pathlib import Path
import subprocess
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from page_inventory import check_pages
from test_core import project, write, SKILL  # noqa: F401  (fixture reuse)


def repo(tmp_path, files):
    root = tmp_path / 'inv'
    root.mkdir()
    subprocess.run(['git', 'init', '-q', str(root)], check=True)
    for name in files:
        path = root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text('x')
    return root


def config(pages, html=('site/index.html',), envs=('DESIGN_WEB_URL',)):
    row = {'schema_version': 1, 'html': list(html), 'urls': [{'env': e} for e in envs],
           'viewports': [375], 'schemes': ['light']}
    if pages is not None:
        row['pages'] = pages
    return row


PAGES = ['app/page.tsx', 'app/login/page.tsx', 'app/report/[id]/page.tsx', 'site/index.html']


def full(**over):
    pages = {'sources': ['app/**/page.tsx', 'site/*.html'],
             'covered': {'app/page.tsx': ['DESIGN_WEB_URL'], 'site/index.html': ['html']},
             'excluded': {'app/login/page.tsx': 'auth screen owned by the identity provider',
                          'app/report/[id]/page.tsx': 'needs a session fixture; tracked separately'}}
    pages.update(over)
    return pages


def test_complete_inventory_passes(tmp_path):
    root = repo(tmp_path, PAGES)
    assert check_pages(root, config(full())) == {'errors': [], 'warnings': [], 'pages': 4}


def test_new_page_not_declared_is_a_failure_naming_the_file(tmp_path):
    root = repo(tmp_path, PAGES + ['app/settings/page.tsx'])
    result = check_pages(root, config(full()))
    assert result['errors'] == ['page not validated or excluded: app/settings/page.tsx']


def test_untracked_new_page_is_seen_but_ignored_files_are_not(tmp_path):
    root = repo(tmp_path, PAGES + ['app/draft/page.tsx', 'node_modules/x/app/page.tsx'])
    (root / '.gitignore').write_text('node_modules/\n')
    result = check_pages(root, config(full()))
    assert result['errors'] == ['page not validated or excluded: app/draft/page.tsx']


def test_missing_inventory_is_a_visible_warning(tmp_path):
    root = repo(tmp_path, PAGES)
    result = check_pages(root, config(None))
    assert result['errors'] == [] and result['pages'] == 0
    assert result['warnings'] == ['page inventory not declared; pages outside validation.json are not detected']


@pytest.mark.parametrize('pages,message', [
    ({'sources': [], 'covered': {}, 'excluded': {}}, 'page inventory needs at least one source pattern'),
    (full(sources=['app/**/page.tsx', 'site/*.html', 'pages/**/*.vue']), 'page source matches no files: pages/**/*.vue'),
    (full(excluded={'app/login/page.tsx': '  ', 'app/report/[id]/page.tsx': 'r'}), 'excluded page needs a reason: app/login/page.tsx'),
    (full(covered={'app/page.tsx': ['DESIGN_WEB_URL'], 'site/index.html': ['html'], 'app/login/page.tsx': ['DESIGN_WEB_URL']}),
     'validation target covers more than one page: DESIGN_WEB_URL'),
    (full(covered={'app/page.tsx': ['DESIGN_OTHER_URL'], 'site/index.html': ['html']}),
     'covered page names an undeclared target: app/page.tsx'),
    (full(covered={'app/page.tsx': ['DESIGN_WEB_URL'], 'site/index.html': ['html'], 'app/gone/page.tsx': ['html']}),
     'declared page does not exist or matches no source: app/gone/page.tsx'),
    (full(covered={'app/page.tsx': ['html'], 'site/index.html': ['html'], 'app/x': []}), None),
])
def test_inventory_declaration_errors(tmp_path, pages, message):
    root = repo(tmp_path, PAGES)
    errors = check_pages(root, config(pages))['errors']
    if message:
        assert message in errors, errors
    else:
        assert errors


def test_every_validation_target_must_be_attributed_to_a_page(tmp_path):
    root = repo(tmp_path, PAGES)
    errors = check_pages(root, config(full(), envs=('DESIGN_WEB_URL', 'DESIGN_LP_URL')))['errors']
    assert 'validation target is not attributed to a page: DESIGN_LP_URL' in errors


def test_page_both_covered_and_excluded_fails(tmp_path):
    root = repo(tmp_path, PAGES)
    pages = full()
    pages['excluded']['app/page.tsx'] = 'contradiction'
    assert 'page is both covered and excluded: app/page.tsx' in check_pages(root, config(pages))['errors']


def test_web_check_stops_before_l1_on_undeclared_page(project):
    root, catalogs = project
    (root / 'index.html').write_text('<html></html>')
    (root / 'extra.html').write_text('<html></html>')
    pages = {'sources': ['*.html'], 'covered': {'index.html': ['html', 'DESIGN_FIXTURE_URL']}, 'excluded': {}}
    write(root / 'design/validation.json', config(pages, html=('index.html',), envs=('DESIGN_FIXTURE_URL',)))
    argv = [sys.executable, str(SKILL / 'scripts/web_check.py'), '--repo', str(root), '--catalog-dir', str(catalogs)]
    run = subprocess.run(argv, capture_output=True, text=True)
    result = json.loads(run.stdout)
    assert run.returncode == 2, result
    assert [s['stage'] for s in result['stages']] == ['L0']
    assert 'page not validated or excluded: extra.html' in result['errors']
