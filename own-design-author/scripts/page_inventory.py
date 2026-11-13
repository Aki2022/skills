"""Page inventory: detect pages that neither L1/L2 validates nor the owner excluded with a reason.

validation.json may declare `pages`:
  sources  - repo-relative git pathspec globs that find every page source (e.g. app/**/page.tsx)
  covered  - {page path: [evidence]}; evidence is "html" (the path is an L1 target) or a URL env of `urls`
  excluded - {page path: reason}; shipping also needs the owner's approval of each exclusion (L0)
  routes   - optional {page path: route pattern} for page files whose route cannot be derived
When the caller passes the rendered path of each URL env, every page covered by a URL must be rendered at its
own route: Next.js app/ and pages/ files derive it from the file path; other files declare it in routes.
Files are listed from git (tracked plus untracked, honoring .gitignore), so dependencies and build output
stay out. Without `pages` the gate keeps working but reports that new pages cannot be detected.
"""
import re
import subprocess
from urllib.parse import urlsplit

MISSING = 'page inventory not declared; pages outside validation.json are not detected'


def _files(root, pattern):
    out = subprocess.run(['git', '-C', str(root), 'ls-files', '--cached', '--others', '--exclude-standard',
                          '-z', '--', ':(glob)' + pattern], capture_output=True, text=True, check=True).stdout
    return {name for name in out.split('\0') if name}


PAGE_FILE = re.compile(r'\.(tsx|jsx|ts|js|mdx)$')


def derive_route(page):
    """The URL route a Next.js page file serves, as a pattern with [param] segments; None when not derivable."""
    parts = page.split('/')
    if 'app' in parts and re.fullmatch(r'page\.(tsx|jsx|ts|js|mdx)', parts[-1]):
        segments = parts[len(parts) - parts[::-1].index('app'):-1]
        segments = [x for x in segments if not (x.startswith('(') and x.endswith(')')) and not x.startswith('@')]
        return '/' + '/'.join(segments)
    if 'pages' in parts and PAGE_FILE.search(parts[-1]) and not parts[-1].startswith('_'):
        segments = parts[len(parts) - parts[::-1].index('pages'):]
        segments[-1] = PAGE_FILE.sub('', segments[-1])
        if segments[-1] == 'index':
            segments = segments[:-1]
        return '/' + '/'.join(segments)
    return None


def route_matches(pattern, path):
    """Whether a concrete rendered path is an instance of a route pattern ([x], [...x], [[...x]])."""
    want = [x for x in pattern.split('/') if x]
    have = [x for x in urlsplit(path).path.split('/') if x]

    def walk(i, j):
        if i == len(want):
            return j == len(have)
        part = want[i]
        if part.startswith('[[...') and part.endswith(']]'):
            return any(walk(i + 1, k) for k in range(j, len(have) + 1))
        if part.startswith('[...') and part.endswith(']'):
            return any(walk(i + 1, k) for k in range(j + 1, len(have) + 1))
        if j == len(have):
            return False
        if part.startswith('[') and part.endswith(']'):
            return walk(i + 1, j + 1)
        return part == have[j] and walk(i + 1, j + 1)
    return walk(0, 0)


def check_pages(root, config, paths=None):
    pages = config.get('pages')
    if pages is None:
        return {'errors': [], 'warnings': [MISSING], 'pages': 0}
    if not isinstance(pages, dict) or not {'sources', 'covered', 'excluded'} <= set(pages) <= {'sources', 'covered', 'excluded', 'routes'}:
        return {'errors': ['page inventory must declare sources, covered and excluded (and optionally routes)'], 'warnings': [], 'pages': 0}
    routes = pages.get('routes', {})
    if not isinstance(routes, dict) or any(not isinstance(v, str) or not v.startswith('/') for v in routes.values()):
        return {'errors': ['pages.routes must map page paths to routes starting with /'], 'warnings': [], 'pages': 0}
    sources, covered, excluded = pages['sources'], pages['covered'], pages['excluded']
    errors = []
    if not isinstance(sources, list) or not sources or not all(isinstance(s, str) and s.strip() for s in sources):
        return {'errors': ['page inventory needs at least one source pattern'], 'warnings': [], 'pages': 0}
    if not isinstance(covered, dict) or not isinstance(excluded, dict):
        return {'errors': ['covered and excluded must map page paths'], 'warnings': [], 'pages': 0}
    found = set()
    for pattern in sources:
        if pattern.startswith('/') or '..' in pattern.split('/'):
            errors.append('page source must stay inside the repository: ' + pattern)
            continue
        matched = _files(root, pattern)
        if not matched:
            errors.append('page source matches no files: ' + pattern)
        found |= matched
    html = set(config.get('html') or [])
    envs = {row.get('env') for row in config.get('urls') or [] if isinstance(row, dict)}
    claimed = {}
    for page, evidence in covered.items():
        if not isinstance(evidence, list) or not evidence:
            errors.append('covered page needs validation evidence: ' + page)
            continue
        for target in evidence:
            if target == 'html':
                if page not in html:
                    errors.append('covered page is not an L1 html target: ' + page)
            elif target in envs:
                claimed.setdefault(target, []).append(page)
            else:
                errors.append('covered page names an undeclared target: ' + page)
    for page, reason in excluded.items():
        if not isinstance(reason, str) or not reason.strip():
            errors.append('excluded page needs a reason: ' + page)
    for page in sorted(set(covered) & set(excluded)):
        errors.append('page is both covered and excluded: ' + page)
    for page in sorted((set(covered) | set(excluded)) - found):
        errors.append('declared page does not exist or matches no source: ' + page)
    for env in sorted(env for env, owners in claimed.items() if len(owners) > 1):
        errors.append('validation target covers more than one page: ' + env)
    for env in sorted(e for e in envs - set(claimed) if isinstance(e, str)):
        errors.append('validation target is not attributed to a page: ' + env)
    for page in sorted(found - set(covered) - set(excluded)):
        errors.append('page not validated or excluded: ' + page)
    for page in sorted(set(routes) - set(covered)):
        errors.append('route declared for a page no URL covers: ' + page)
    if paths is not None:
        for env, owners in sorted(claimed.items()):
            if len(owners) != 1:
                continue
            page = owners[0]
            route = routes.get(page) or derive_route(page)
            if route is None:
                errors.append(f'cannot derive the route of {page}; declare it in pages.routes')
            elif env not in paths:
                errors.append(f'{env} has no rendered path to compare with {page}')
            elif not route_matches(route, paths[env]):
                errors.append(f'{env} renders {paths[env]}, which is not the route {route} of {page}')
    return {'errors': errors, 'warnings': [], 'pages': len(found)}
