"""Fixed trusted IA extension point; repository configuration never names code."""
import importlib.util
import hashlib
import re
from reporting import stage
from pathlib import Path


def resolve_ia(root,config,approvals,profiles,report):
    declared=config.get('modules',{})
    if not isinstance(declared,dict) or set(declared)-{'ia'}:
        raise ValueError('unknown design module declaration')
    present=(root/'design/info_arch').exists()
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
    implementation_sha256=hashlib.sha256(implementation.read_bytes()).hexdigest()
    spec=importlib.util.spec_from_file_location('design_ia_module',implementation)
    module=importlib.util.module_from_spec(spec)
    try:spec.loader.exec_module(module)
    except Exception as exc:raise ValueError('IA module could not load') from exc
    if getattr(module,'VERSION',None)!=declared['ia'] or not callable(getattr(module,'resolve',None)):
        raise ValueError('IA module version/interface mismatch')
    try:result=module.resolve(root=root,approvals=approvals,profiles=profiles)
    except Exception as exc:raise ValueError('IA module resolution failed') from exc
    if not isinstance(result,dict) or set(result)!={'media','approval_hashes','must_approval_keys','errors','warnings','report'}:
        raise ValueError('invalid IA resolution contract')
    if not isinstance(result['media'],dict) or set(result['media'])!=set(profiles) or any(not isinstance(value,dict) for value in result['media'].values()):
        raise ValueError('IA media resolution mismatch')
    if not isinstance(result['approval_hashes'],dict) or any(not isinstance(k,str) or not re.fullmatch(r'ia:[a-z][a-z0-9-]*:[A-Za-z0-9_.:/-]+',k) or not isinstance(v,str) or len(v)!=64 or any(c not in '0123456789abcdef' for c in v) for k,v in result['approval_hashes'].items()):
        raise ValueError('invalid IA approval hash contract')
    if not result['approval_hashes'] and not result['errors']:
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
    stage(report,'ia',2 if result['errors'] else (1 if result['warnings'] else 0),len(result['approval_hashes']))
    if hashlib.sha256(implementation.read_bytes()).hexdigest()!=implementation_sha256:
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
    if not path.is_file() or path.resolve()!=path or hashlib.sha256(path.read_bytes()).hexdigest()!=row['implementation_sha256']:
        raise ValueError('IA resolver changed; rerun L0')
