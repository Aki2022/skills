"""The rule for where the layers live sits in one place (references/layer-paths.json plus the two small loaders).

Nothing else in the skills' code may spell the repository's layer root: a code path that does would keep reading
the old location after a move, and silently so. This test counts the files it looked at, proves the scanner sees a
planted literal, and then requires zero literals outside the rule."""
import ast
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

SKILL = Path(__file__).resolve().parents[2]
SKILLS = SKILL.parent
LAYOUT = json.loads((SKILL / 'references/layer-paths.json').read_text())
ROOT = LAYOUT['entries']['layers']
RULE_FILES = {
    SKILL / 'references/layer-paths.json',
    SKILL / 'scripts/layer_paths.py',
    SKILL / 'scripts/layer_paths.mjs',
}
CODE_SUFFIXES = {'.py', '.mjs', '.js', '.cjs', '.sh'}
# Names the layout JSON owns. A string literal that is, starts with, or contains "<root>/" is a layer path;
# a string that is exactly the root name is the path component of one.
PATH_RE = re.compile(r'(?<![A-Za-z0-9_.-])' + re.escape(ROOT) + r'/')


def _hit(value):
    return value == ROOT or bool(PATH_RE.search(value))


def python_literals(source):
    """Executable string constants: docstrings and comments are not code and are not returned."""
    tree = ast.parse(source)
    docstrings = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = node.body
            if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) \
                    and isinstance(body[0].value.value, str):
                docstrings.add(id(body[0].value))
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in docstrings:
            yield node.value


def js_literals(source):
    """String and template literals of a JavaScript source, skipping comments (a small scanner, not a parser)."""
    i, n = 0, len(source)
    while i < n:
        ch = source[i]
        if source.startswith('//', i):
            i = source.find('\n', i)
            i = n if i < 0 else i
        elif source.startswith('/*', i):
            j = source.find('*/', i + 2)
            i = n if j < 0 else j + 2
        elif ch in '\'"`':
            quote, j, buf = ch, i + 1, []
            while j < n and source[j] != quote:
                if source[j] == '\\' and j + 1 < n:
                    buf.append(source[j + 1])
                    j += 2
                    continue
                buf.append(source[j])
                j += 1
            yield ''.join(buf)
            i = j + 1
        else:
            i += 1


def shell_literals(source):
    for line in source.splitlines():
        if not line.lstrip().startswith('#'):
            yield line


def literals_of(path):
    text = path.read_text(encoding='utf-8')
    if path.suffix == '.py':
        return list(python_literals(text))
    if path.suffix in {'.mjs', '.js', '.cjs'}:
        return list(js_literals(text))
    if path.suffix == '.json':
        found = []

        def walk(value):
            if isinstance(value, str):
                found.append(value)
            elif isinstance(value, list):
                [walk(v) for v in value]
            elif isinstance(value, dict):
                [walk(v) for v in value.values()]
        walk(json.loads(text))
        return [v for v in found if v == ROOT or v.startswith(ROOT + '/')]  # data values that are paths
    return list(shell_literals(text))


def is_test(path):
    parts = path.relative_to(SKILLS).parts
    return 'tests' in parts or path.name.startswith('test_') or 'fixtures' in parts


def is_extensionless_script(path):
    if path.suffix or not path.is_file() or path.parent.name != 'scripts':
        return False
    return path.read_bytes()[:2] == b'#!'


def scanned_files():
    files = []
    for path in sorted(SKILLS.rglob('*')):
        rel = path.relative_to(SKILLS).parts
        if not path.is_file() or path.is_symlink() or '.git' in rel or 'node_modules' in rel or is_test(path):
            continue
        if path in RULE_FILES:
            continue
        in_scripts = 'scripts' in rel
        in_refs = 'references' in rel and path.suffix == '.json'
        if (in_scripts and (path.suffix in CODE_SUFFIXES or is_extensionless_script(path))) or in_refs:
            files.append(path)
    return files


def violations():
    found = []
    for path in scanned_files():
        for value in literals_of(path):
            if _hit(value) if path.suffix != '.json' else True:
                found.append((str(path.relative_to(SKILLS)), value[:80]))
    return found


def test_scanner_sees_planted_literals_and_ignores_prose(tmp_path):
    planted_py = tmp_path / 'a.py'
    planted_py.write_text('"""doc mentions %s/ only"""\n# %s/ in a comment\nx = root / "%s" / "y"\nz = "%s/approvals.yaml"\n'
                          % (ROOT, ROOT, ROOT, ROOT))
    assert sorted(v for v in literals_of(planted_py) if _hit(v)) == sorted([ROOT, ROOT + '/approvals.yaml'])
    planted_js = tmp_path / 'a.mjs'
    planted_js.write_text('// %s/ in a comment\n/* %s/ */\nconst a = "https://x.test/%s/q";\nconst b = join(r, \'%s\');\n'
                          % (ROOT, ROOT, ROOT, ROOT))
    assert [v for v in literals_of(planted_js) if _hit(v)] == ['https://x.test/%s/q' % ROOT, ROOT]
    clean = tmp_path / 'b.py'
    clean.write_text('x = "redesign/"\ny = "designs"\n')
    assert [v for v in literals_of(clean) if _hit(v)] == []


def test_the_scan_looked_at_real_code_in_every_language():
    files = scanned_files()
    by_suffix = {}
    for path in files:
        by_suffix[path.suffix] = by_suffix.get(path.suffix, 0) + 1
    assert len(files) >= 60, 'a scan of %d files is not a scan (0 files must not pass)' % len(files)
    assert by_suffix.get('.py', 0) >= 40 and by_suffix.get('.mjs', 0) >= 5, by_suffix
    assert any(p.name == 'design-author' for p in files), 'extensionless entry points must be scanned too'


def test_no_code_outside_the_rule_spells_the_layer_root():
    found = violations()
    assert found == [], '%d literal(s) spell the layer root outside the rule:\n%s' % (
        len(found), '\n'.join('%s: %r' % item for item in found[:40]))


def test_the_rule_files_are_the_only_ones_exempt():
    for path in RULE_FILES:
        assert path.is_file(), path
    assert any(_hit(v) or v == ROOT for v in json.loads((SKILL / 'references/layer-paths.json').read_text())['entries'].values())
