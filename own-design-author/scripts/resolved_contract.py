"""Read and validate an L0 output; never resolve DTCG in a consumer."""
import hashlib
import json
import math
from pathlib import Path
from tailwind_contract import tailwind_artifact


def load_manifest(path):
    original = Path(path)
    if original.is_symlink() or original.parent.is_symlink() or original.parent.parent.is_symlink():
        raise ValueError("symlinked manifest directory is unsupported")
    path = original.resolve()
    if path.parent.name != '.resolved' or path.parent.parent.name != 'design':
        raise ValueError('manifest must be design/.resolved/<media>.json')
    design = path.parent.parent
    data = json.loads(path.read_text())
    if data.get('schema_version') != 1 or data.get('media') != 'web':
        raise ValueError('unsupported manifest schema/media')
    if not data.get('tokens') or not data.get('input_hashes'):
        raise ValueError('empty manifest')
    expected = data['input_hashes']
    actual = {}
    for source in sorted(design.rglob('*')):
        rel = source.relative_to(design)
        if '.resolved' in rel.parts or source.name in ('assets.json', 'migration.json'):
            continue
        if source.is_file():
            if not source.resolve().is_relative_to(design):
                raise ValueError('design input escapes repository')
            actual[rel.as_posix()] = hashlib.sha256(source.read_bytes()).hexdigest()
    if actual != expected:
        raise ValueError('stale manifest; rerun L0')
    for key in ('text_contrast', 'large_text_contrast', 'ui_contrast', 'target_size_px',
                'focus_width_px', 'focus_contrast'):
        value = data.get('minimums', {}).get(key)
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value <= 0:
            raise ValueError('invalid manifest minimums')
    return data


def check_tailwind(path, manifest):
    artifact = manifest.get('artifacts', {}).get('tailwind.config.cjs')
    if artifact != tailwind_artifact(manifest['tokens']):
        raise ValueError('manifest has no generated Tailwind contract')
    content = artifact['content'].encode()
    digest = hashlib.sha256(content).hexdigest()
    if digest != artifact['sha256'] or Path(path).read_bytes() != content:
        raise ValueError('Tailwind config differs from generated contract')
    return digest
