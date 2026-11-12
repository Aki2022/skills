"""Page inventory: detect pages that neither L1/L2 validates nor the owner excluded with a reason.

validation.json may declare `pages`:
  sources  - repo-relative git pathspec globs that find every page source (e.g. app/**/page.tsx)
  covered  - {page path: [evidence]}; evidence is "html" (the path is an L1 target) or a URL env of `urls`
  excluded - {page path: reason}
Files are listed from git (tracked plus untracked, honoring .gitignore), so dependencies and build output
stay out. Without `pages` the gate keeps working but reports that new pages cannot be detected.
"""
import subprocess

MISSING = 'page inventory not declared; pages outside validation.json are not detected'


def _files(root, pattern):
    out = subprocess.run(['git', '-C', str(root), 'ls-files', '--cached', '--others', '--exclude-standard',
                          '-z', '--', ':(glob)' + pattern], capture_output=True, text=True, check=True).stdout
    return {name for name in out.split('\0') if name}


def check_pages(root, config):
    pages = config.get('pages')
    if pages is None:
        return {'errors': [], 'warnings': [MISSING], 'pages': 0}
    if not isinstance(pages, dict) or set(pages) != {'sources', 'covered', 'excluded'}:
        return {'errors': ['page inventory must declare exactly sources, covered and excluded'], 'warnings': [], 'pages': 0}
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
    return {'errors': errors, 'warnings': [], 'pages': len(found)}
