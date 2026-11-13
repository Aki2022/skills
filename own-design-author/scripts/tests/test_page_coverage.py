"""Every page is looked at: shipping needs a page inventory, an excluded page needs the owner's approval, and
each rendered URL must be the route of the page file it claims to cover."""
import json
from pathlib import Path
import subprocess
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from page_inventory import check_pages, derive_route, route_matches  # noqa: E402
from test_core import project, write, check, approve, approve_rows  # noqa: F401  (fixture reuse)
from test_page_inventory import repo, config  # noqa: E402


@pytest.mark.parametrize('page,route', [
    ('app/page.tsx', '/'),
    ('apps/web/src/app/diagnosis/page.tsx', '/diagnosis'),
    ('app/(marketing)/about/page.jsx', '/about'),                 # route groups do not appear in the URL
    ('app/users/[id]/page.tsx', '/users/[id]'),
    ('app/docs/[...slug]/page.mdx', '/docs/[...slug]'),
    ('app/@modal/login/page.tsx', '/login'),                     # parallel-route slots do not appear either
    ('pages/index.tsx', '/'),
    ('src/pages/blog/[slug].tsx', '/blog/[slug]'),
    ('site/index.html', None),                                   # not a framework page: needs an explicit route
    ('components/Button.tsx', None),
])
def test_route_is_derived_from_the_page_file(page, route):
    assert derive_route(page) == route


@pytest.mark.parametrize('pattern,path,ok', [
    ('/', '/', True), ('/', '/about', False),
    ('/diagnosis', '/diagnosis', True), ('/diagnosis', '/diagnosis/', True), ('/diagnosis', '/diagnosisx', False),
    ('/users/[id]', '/users/sample', True), ('/users/[id]', '/users', False), ('/users/[id]', '/users/a/b', False),
    ('/docs/[...slug]', '/docs/a/b', True), ('/docs/[...slug]', '/docs', False),
    ('/docs/[[...slug]]', '/docs', True),
    ('/a', '/a?tab=1', True),                                    # the query does not choose the page
])
def test_route_pattern_matches_a_concrete_path(pattern, path, ok):
    assert route_matches(pattern, path) is ok


PAGES = ['app/page.tsx', 'app/users/[id]/page.tsx', 'app/login/page.tsx']


def inventory(**over):
    pages = {'sources': ['app/**/page.tsx'], 'covered': {'app/page.tsx': ['DESIGN_WEB_URL'],
             'app/users/[id]/page.tsx': ['DESIGN_USER_URL']}, 'excluded': {'app/login/page.tsx': 'identity provider screen'}}
    pages.update(over)
    return pages


def test_rendered_url_must_be_the_route_of_the_page_it_covers(tmp_path):
    root = repo(tmp_path, PAGES)
    row = config(inventory(), html=(), envs=('DESIGN_WEB_URL', 'DESIGN_USER_URL'))
    assert check_pages(root, row, paths={'DESIGN_WEB_URL': '/', 'DESIGN_USER_URL': '/users/sample'})['errors'] == []
    wrong = check_pages(root, row, paths={'DESIGN_WEB_URL': '/', 'DESIGN_USER_URL': '/users'})['errors']
    assert wrong == ['DESIGN_USER_URL renders /users, which is not the route /users/[id] of app/users/[id]/page.tsx']


def test_page_without_a_derivable_route_needs_one_declared(tmp_path):
    root = repo(tmp_path, ['site/index.html', 'site/about.html'])
    pages = {'sources': ['site/*.html'], 'covered': {'site/index.html': ['html'], 'site/about.html': ['DESIGN_WEB_URL']},
             'excluded': {}}
    row = config(pages, html=('site/index.html',))
    assert check_pages(root, row, paths={'DESIGN_WEB_URL': '/about.html'})['errors'] == [
        'cannot derive the route of site/about.html; declare it in pages.routes']
    pages['routes'] = {'site/about.html': '/about.html'}
    assert check_pages(root, config(pages, html=('site/index.html',)), paths={'DESIGN_WEB_URL': '/about.html'})['errors'] == []


def test_render_server_without_a_path_cannot_claim_a_page(tmp_path):
    root = repo(tmp_path, PAGES)
    row = config(inventory(), html=(), envs=('DESIGN_WEB_URL', 'DESIGN_USER_URL'))
    assert check_pages(root, row, paths={'DESIGN_WEB_URL': '/'})['errors'] == [
        'DESIGN_USER_URL has no rendered path to compare with app/users/[id]/page.tsx']


# --- L0 ---------------------------------------------------------------------------------------------

def web_repo(project, pages):
    root, _ = project
    (root / 'index.html').write_text('<html></html>')
    (root / 'app').mkdir(exist_ok=True)
    for name in ('page.tsx', 'login/page.tsx'):
        (root / 'app' / name).parent.mkdir(parents=True, exist_ok=True)
        (root / 'app' / name).write_text('x')
    validation = {'schema_version': 1, 'html': ['index.html'], 'urls': [{'env': 'DESIGN_WEB_URL'}],
                  'viewports': [375], 'schemes': ['light']}
    if pages is not None:
        validation['pages'] = pages
    write(root / 'design/validation.json', validation)
    config_ = json.loads((root / 'design/design.yaml').read_text())
    config_['integration_gate'] = {'media': 'web', 'merge_base': 'origin/main', 'render_servers': {
        'DESIGN_WEB_URL': {'command': ['python3', '-m', 'http.server', '{port}'], 'path': '/'}}}
    write(root / 'design/design.yaml', config_)


L0_PAGES = {'sources': ['app/**/page.tsx', 'index.html'], 'covered': {'app/page.tsx': ['DESIGN_WEB_URL'], 'index.html': ['html']},
            'excluded': {'app/login/page.tsx': 'identity provider screen'}}


def test_ship_needs_a_page_inventory(project):
    web_repo(project, None)
    approve(project)
    assert not any('page inventory' in e for e in check(project)['errors'])   # design work is not blocked
    shipped = check(project, ship=True)
    assert shipped['exit_code'] == 2 and any('page inventory required' in e for e in shipped['errors']), shipped['errors']


def approve_with_pages(project):
    targets = check(project)['approval_targets']
    approve_rows(project, [{'key': k, 'value_hash': v} for k, v in targets.items() if k.endswith('color.primary') or ':page:' in k])


def test_excluded_page_needs_the_owners_approval_bound_to_its_reason(project):
    web_repo(project, L0_PAGES)
    approve(project)
    result = check(project, ship=True)
    assert any('unapproved page exclusions 1' in e for e in result['errors']), result['errors']
    assert 'ds:web:page:app/login/page.tsx' in check(project)['approval_targets']
    approve_with_pages(project)
    shipped = check(project, ship=True)
    assert shipped['exit_code'] == 0, shipped['errors']
    assert shipped['counts']['pages'] == {'found': 3, 'covered': 2, 'excluded': 1}
    web_repo(project, {**L0_PAGES, 'excluded': {'app/login/page.tsx': 'changed reason'}})
    assert any('unapproved page exclusions 1' in e for e in check(project, ship=True)['errors'])


def test_l0_checks_page_routes_against_the_declared_render_paths(project):
    root, _ = project
    web_repo(project, L0_PAGES)
    config_ = json.loads((root / 'design/design.yaml').read_text())
    config_['integration_gate']['render_servers']['DESIGN_WEB_URL']['path'] = '/login'
    write(root / 'design/design.yaml', config_)
    errors = check(project)['errors']
    assert any('DESIGN_WEB_URL renders /login, which is not the route / of app/page.tsx' in e for e in errors), errors


def test_undeclared_page_stops_l0_before_any_rendering(project):
    root, _ = project
    web_repo(project, L0_PAGES)
    (root / 'app/settings').mkdir()
    (root / 'app/settings/page.tsx').write_text('x')
    assert any('page not validated or excluded: app/settings/page.tsx' in e for e in check(project)['errors'])


# --- several pages from one server -------------------------------------------------------------------

def test_render_server_can_share_another_servers_process(tmp_path):
    from ship_contract import validate_integration_gate
    from ship import render_servers, hook_free_env
    base = {'media': 'web', 'merge_base': 'origin/main'}
    servers = {'DESIGN_WEB_URL': {'command': [sys.executable, '-m', 'http.server', '{port}', '--bind', '127.0.0.1'], 'path': '/'},
               'DESIGN_USER_URL': {'same_as': 'DESIGN_WEB_URL', 'path': '/users/'}}
    validate_integration_gate({'integration_gate': {**base, 'render_servers': servers}})
    for broken in ({'same_as': 'DESIGN_NOPE_URL', 'path': '/'}, {'same_as': 'DESIGN_USER_URL', 'path': '/'},
                   {'same_as': 'DESIGN_WEB_URL', 'path': '/', 'command': ['x']}):
        with pytest.raises(ValueError):
            validate_integration_gate({'integration_gate': {**base, 'render_servers': {**servers, 'DESIGN_OTHER_URL': broken}}})
    (tmp_path / 'users').mkdir()
    (tmp_path / 'users/index.html').write_text('ok')
    with render_servers(tmp_path, servers, hook_free_env()) as (_env, urls, processes):
        assert len(processes) == 1
        assert urls['DESIGN_USER_URL'] == urls['DESIGN_WEB_URL'].rstrip('/') + '/users/'
