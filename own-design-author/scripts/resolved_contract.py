"""Read and validate an L0 output; never resolve DTCG in a consumer."""
import hashlib
import json
import math
import os
from pathlib import Path
from tailwind_contract import build_artifacts


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
    if data.get('approval_scheme')!='ssh-jcs-v1':
        raise ValueError('unsupported approval scheme; rerun L0')
    context=data.get('approval_context')
    if context is None and (any(t.get('state')=='approved' for t in data['tokens'].values()) or any(i.get('approval_state')=='approved' for i in data.get('items',{}).values())):
        raise ValueError('approved manifest requires trust context')
    if context is not None:
        from approvals import protected_store,repository_id
        if not isinstance(context,dict) or set(context)!={'repo_id','trust_sha256'}:
            raise ValueError('invalid approval trust context')
        trust=protected_store(os.environ.get('DESIGN_ALLOWED_SIGNERS'),design.parent)
        if hashlib.sha256(trust.read_bytes()).hexdigest()!=context['trust_sha256'] or repository_id(design.parent)!=context['repo_id']:
            raise ValueError('approval trust changed; rerun L0')
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


def check_tailwind(path, manifest, repo=None):
    expected = build_artifacts(manifest['tokens'],manifest.get('tailwind_templates'))
    target = 'tailwind.config.cjs'
    if manifest.get('tailwind_templates') is not None:
        if repo is None:
            raise ValueError('repository root required for project Tailwind contract')
        try:
            target = Path(path).resolve().relative_to(Path(repo).resolve()).as_posix()
        except ValueError:
            raise ValueError('Tailwind config escapes repository')
    artifact = manifest.get('artifacts', {}).get(target)
    if not artifact or artifact != expected.get(target):
        raise ValueError('manifest has no matching generated Tailwind contract')
    content = artifact['content'].encode()
    digest = hashlib.sha256(content).hexdigest()
    if digest != artifact['sha256'] or Path(path).read_bytes() != content:
        raise ValueError('Tailwind config differs from generated contract')
    return digest


def check_asset_exports(repo, manifest):
    root=Path(repo).resolve()
    for row in manifest.get('asset_exports',[]):
        source=root/'design'/row['source']
        if not source.resolve().is_relative_to(root/'design/assets') or hashlib.sha256(source.read_bytes()).hexdigest() != row['sha256']:
            raise ValueError('asset source differs from manifest')
        for target in row['targets']:
            path=root/target
            if not path.resolve().is_relative_to(root) or hashlib.sha256(path.read_bytes()).hexdigest() != row['sha256']:
                raise ValueError('deployed asset differs from design source')


if __name__=='__main__':
    import sys
    try:
        if len(sys.argv)!=2:raise ValueError('one manifest required')
        load_manifest(sys.argv[1])
    except (OSError,ValueError,KeyError,TypeError):
        print('manifest or approval trust unavailable, invalid or stale',file=sys.stderr)
        raise SystemExit(2)
