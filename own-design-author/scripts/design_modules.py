"""Fixed trusted IA extension point; repository configuration never names code."""
import importlib.util
import hashlib
import re
from reporting import stage
from pathlib import Path, PurePosixPath
import layer_paths


IA_GIT_CONTRACT = 'ia-git-history-v4'
IA_VERSION = '0.4.0'


# The IA checker is more than its own skill: the positioning layer (own-positioning-author) is imported by it, and the
# docs dictionary reader (own-doc-update's dictionary.py) reads the fixed values and official names it checks against.
# What the resolved result trusts is bound together; a change to any of them invalidates it (rerun L0).
IA_TRUSTED_SKILLS = ('own-positioning-author',)
IA_TRUSTED_FILES = ('own-doc-update/scripts/dictionary.py',)
# What the checker reads outside the infoarch layer (layer_paths 'infoarch'): the docs dictionary, one `docs/dictionary/<domain>.json` per domain
# (the same folder and glob as ia_module.DICTIONARY_DIR; resolved_contract.mjs mirrors it).
IA_DICTIONARY_DIR = 'docs/dictionary'


def _ia_input_key_ok(key):
    if not isinstance(key, str):
        return False
    parts = Path(key).parts
    if '..' in parts or key != '/'.join(parts):
        return False
    for key_name in ('infoarch', 'positioning_dir'):  # the infoarch tree and the positioning tree are the IA's inputs
        layer = PurePosixPath(layer_paths.rel(key_name)).parts
        if parts[:len(layer)] == layer and len(parts) > len(layer):
            return True
    return len(parts) == 3 and parts[:2] == tuple(IA_DICTIONARY_DIR.split('/')) and parts[2].endswith('.json')


def check_ia_dictionary_current(modules, root):
    """A resolved manifest whose IA inputs include the docs dictionary is stale once a dictionary file changed,
    appeared or went away (ISSUE-20261121). A manifest that recorded none is stale only if the repository has some."""
    governance = (modules.get('ia') or {}).get('resolved', {}).get('governance')
    if not isinstance(governance, dict):
        return
    prefix = IA_DICTIONARY_DIR + '/'
    expected = {k: v for k, v in governance.get('input_hashes', {}).items() if k.startswith(prefix)}
    root = Path(root)
    folder = root / IA_DICTIONARY_DIR
    actual = {}
    if folder.exists() or folder.is_symlink():
        for path in [root / 'docs', folder]:
            if path.is_symlink() or not path.resolve().is_relative_to(root.resolve()):
                raise ValueError('IA input symlink or escape')
        for path in sorted(folder.glob('*.json')):
            if path.is_symlink() or not path.resolve().is_relative_to(root.resolve()):
                raise ValueError('IA input symlink or escape')
            if path.is_file():
                actual[path.relative_to(root).as_posix()] = hashlib.sha256(path.read_bytes()).hexdigest()
    if actual != expected:
        raise ValueError('stale manifest; rerun L0')


def _digest(path, full=False):
    if not full:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    skill = path.parent.parent
    root = skill.parent
    inputs = {}
    for owner in (skill, *(root / name for name in IA_TRUSTED_SKILLS)):
        for sub in ('scripts', 'references'):
            directory = owner / sub
            if directory.is_symlink() or not directory.is_dir():
                raise ValueError('IA trusted dependencies are missing')
            for candidate in sorted(directory.rglob('*')):
                relative = candidate.relative_to(owner)
                if 'tests' in relative.parts or '__pycache__' in relative.parts or candidate.suffix == '.pyc':
                    continue
                if candidate.is_symlink() or not candidate.resolve().is_relative_to(owner):
                    raise ValueError('IA trusted dependency symlink or escape')
                # Prose under references/ explains; it is not executed or read by the checker. Editing it must not
                # invalidate resolved results (ADR-20261115-design-ia-two-skills-shared-base, biz_ops). Files only:
                # a directory named *.md still has its contents bound.
                if candidate.is_file() and sub == 'references' and candidate.suffix == '.md':
                    continue
                if candidate.is_file():
                    inputs[(owner.name / relative).as_posix()] = hashlib.sha256(candidate.read_bytes()).hexdigest()
    for name in IA_TRUSTED_FILES:
        file = root / name
        if file.is_symlink() or not file.is_file() or not file.resolve().is_relative_to(root):
            raise ValueError('IA trusted dependencies are missing')
        inputs[name] = hashlib.sha256(file.read_bytes()).hexdigest()
    if not inputs:
        raise ValueError('zero IA trusted dependencies')
    import json
    return hashlib.sha256(json.dumps(dict(sorted(inputs.items(), key=lambda row: row[0].encode('utf-8'))), ensure_ascii=False, separators=(',', ':')).encode()).hexdigest()


def _zero_targets(targets):
    sides = [[targets[k] for k in ('channels', 'outlines', 'sections')], [targets[k] for k in ('screens', 'items')]]
    return any(any(side) and not all(side) for side in sides) or not any(all(side) for side in sides)


def _git_contract(result):
    expected = {'contract', 'governance', 'targets', 'media', 'approval_hashes', 'must_approval_keys', 'errors', 'warnings', 'report'}
    # IA v6: the owner's approvals of positioning atoms and arrows are approval targets (ia:positioning:<key>);
    # their shape is validated with every other module's approval hashes in resolve_ia.
    if set(result) != expected or not isinstance(result['approval_hashes'], dict) or not isinstance(result['must_approval_keys'], list):
        raise ValueError('invalid IA git-history contract')
    targets = result['targets']
    # Five fixed keys (mirrors ia_module.TARGET_KEYS): the outline side and the blueprint side. A side is all or nothing and
    # one whole side is enough; a check that counted nothing, or half a side, is accepted only when errors explain it.
    if not isinstance(targets, dict) or set(targets) != {'channels', 'outlines', 'sections', 'screens', 'items'} or any(
            type(n) is not int or n < 0 for n in targets.values()) or (_zero_targets(targets) and not result['errors']):
        raise ValueError('zero or invalid IA check targets')
    governance = result['governance']
    if (not isinstance(governance, dict) or set(governance) != {'contract', 'method', 'schema_version', 'commit', 'input_hashes'}
            or governance['contract'] != IA_GIT_CONTRACT or governance['method'] != 'git-history' or governance['schema_version'] != 4
            or (governance['commit'] is not None and (not isinstance(governance['commit'], str) or not re.fullmatch(r'[0-9a-f]{40}|[0-9a-f]{64}', governance['commit'])))
            or not isinstance(governance['input_hashes'], dict) or not governance['input_hashes']):
        raise ValueError('invalid IA git-history provenance')
    for key, digest in governance['input_hashes'].items():
        if (not _ia_input_key_ok(key)
                or not isinstance(digest, str) or not re.fullmatch(r'[0-9a-f]{64}', digest)):
            raise ValueError('invalid IA input hash')
    if not isinstance(result['media'], dict):
        raise ValueError('invalid IA git-history media')
    for media in result['media'].values():
        if media != {'governance': governance, 'targets': targets}:
            raise ValueError('invalid IA git-history media')


def resolve_ia(root,config,approvals,profiles,report):
    declared=config.get('modules',{})
    if not isinstance(declared,dict) or set(declared)-{'ia'}:
        raise ValueError('unknown design module declaration')
    present=layer_paths.has_ia(root)
    if not declared and not present:
        report['skipped'].append({'id':'ia','reason':'IA module not declared and no IA inputs'})
        return None
    if not isinstance(declared.get('ia'),str) or not declared['ia']:
        report['missing'].append({'id':'ia','reason':'IA inputs require a pinned module version'})
        raise ValueError('IA module version is missing')
    implementation=Path(__file__).resolve().parents[2]/'own-infoarch-author/scripts/ia_module.py'
    if not implementation.is_file() or implementation.is_symlink() or implementation.parent.is_symlink() or implementation.parent.parent.is_symlink():
        report['missing'].append({'id':'ia','reason':'trusted IA resolver is not installed'})
        raise ValueError('IA module is unavailable')
    implementation_sha256=_digest(implementation, declared['ia']==IA_VERSION)
    spec=importlib.util.spec_from_file_location('design_ia_module',implementation)
    module=importlib.util.module_from_spec(spec)
    try:spec.loader.exec_module(module)
    except layer_paths.LayerLayoutError:raise
    except Exception as exc:raise ValueError('IA module could not load') from exc
    if getattr(module,'VERSION',None)!=declared['ia'] or not callable(getattr(module,'resolve',None)):
        raise ValueError('IA module version/interface mismatch')
    try:result=module.resolve(root=root,approvals=approvals,profiles=profiles)
    except layer_paths.LayerLayoutError:raise
    except Exception as exc:raise ValueError('IA module resolution failed') from exc
    if not isinstance(result,dict):
        raise ValueError('invalid IA resolution contract')
    git_contract=result.get('contract')==IA_GIT_CONTRACT
    if git_contract:
        if declared['ia']!=IA_VERSION:raise ValueError('IA git-history version mismatch')
        _git_contract(result)
    elif set(result)!={'media','approval_hashes','must_approval_keys','errors','warnings','report'}:
        raise ValueError('invalid IA resolution contract')
    if not isinstance(result['media'],dict) or set(result['media'])!=set(profiles) or any(not isinstance(value,dict) for value in result['media'].values()):
        raise ValueError('IA media resolution mismatch')
    if not isinstance(result['approval_hashes'],dict) or any(not isinstance(k,str) or not re.fullmatch(r'ia:[a-z][a-z0-9-]*:[A-Za-z0-9_.:/-]+',k) or not isinstance(v,str) or len(v)!=64 or any(c not in '0123456789abcdef' for c in v) for k,v in result['approval_hashes'].items()):
        raise ValueError('invalid IA approval hash contract')
    if not git_contract and not result['approval_hashes'] and not result['errors']:
        raise ValueError('zero IA approval targets')
    if not isinstance(result['must_approval_keys'],list) or any(key not in result['approval_hashes'] for key in result['must_approval_keys']):
        raise ValueError('invalid IA required approval keys')
    if any(not isinstance(result[k],list) or any(not isinstance(v,str) for v in result[k]) for k in ('errors','warnings')):
        raise ValueError('invalid IA diagnostics')
    if not isinstance(result['report'],dict) or set(result['report'])!=set(report) or any(not isinstance(result['report'][k],list) for k in report):
        raise ValueError('invalid IA reporting contract')
    for key in report:
        if any(not isinstance(row,dict) or not isinstance(row.get('id'),str) or not row['id'] for row in result['report'][key]):
            raise ValueError('invalid IA reporting row')
        report[key].extend(result['report'][key])
    stage(report,'ia',2 if result['errors'] else (1 if result['warnings'] else 0),result['targets']['sections']+result['targets']['items'] if git_contract else len(result['approval_hashes']))
    if _digest(implementation,git_contract)!=implementation_sha256:
        raise ValueError('IA resolver changed during resolution')
    result['_implementation_sha256']=implementation_sha256
    return result


def validate_module_state(modules):
    if not isinstance(modules,dict) or set(modules)-{'ia'}:
        raise ValueError('unsupported resolved module')
    if 'ia' not in modules:return
    row=modules['ia']
    if not isinstance(row,dict) or set(row)!={'version','implementation_sha256','resolved'}:
        raise ValueError('invalid resolved IA module')
    path=Path(__file__).resolve().parents[2]/'own-infoarch-author/scripts/ia_module.py'
    if not path.is_file() or path.resolve()!=path or _digest(path,row.get('resolved',{}).get('governance',{}).get('contract')==IA_GIT_CONTRACT)!=row['implementation_sha256']:
        raise ValueError('IA resolver changed; rerun L0')
