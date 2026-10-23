"""Declarative ship channels and strict evidence binding; never deploys by itself."""
import hashlib
import json
from pathlib import Path, PurePosixPath
import re


def _relative(value):
    if not isinstance(value,str) or not value or '\\' in value or value.startswith('/') or any(p in ('','.','..') for p in value.split('/')):
        raise ValueError('ship artifact must be a repository-relative path')
    if PurePosixPath(value).parts[0]=='.git':
        raise ValueError('git metadata is not a ship artifact')
    return value


def _argv(value, *, port=False):
    if not isinstance(value,list) or not value or any(not isinstance(v,str) or not v or '\0' in v for v in value):
        raise ValueError('ship commands require a nonempty argv array')
    if port and not any('{port}' in v for v in value):
        raise ValueError('render server command must bind the assigned {port}')
    return value


def validate_channels(config):
    """Validate config's optional channels, returning only explicitly shipped channels."""
    rows=config.get('channels',{})
    if not isinstance(rows,dict):raise ValueError('channels must be an object')
    selected={}
    fields={'ship','media','ship_command','artifacts','merge_base','build_command','render_servers'}
    for name,row in rows.items():
        if not isinstance(name,str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]*',name) or not isinstance(row,dict) or set(row)-fields or type(row.get('ship')) is not bool:
            raise ValueError('invalid ship channel')
        if not row['ship']:continue
        if row.get('media')!='web':raise ValueError('only web ship channels are implemented')
        _argv(row.get('ship_command'))
        if 'build_command' in row:_argv(row['build_command'])
        base=row.get('merge_base')
        if not isinstance(base,str) or not base or base.startswith('-') or any(c.isspace() for c in base):
            raise ValueError('ship channel requires an explicit merge_base ref')
        artifacts=row.get('artifacts')
        if not isinstance(artifacts,list) or not artifacts or len(set(artifacts))!=len(artifacts):
            raise ValueError('ship channel requires nonempty unique artifact paths')
        for path in artifacts:_relative(path)
        servers=row.get('render_servers')
        if not isinstance(servers,dict) or not servers:raise ValueError('ship channel requires bound render servers')
        for env,server in servers.items():
            if not isinstance(env,str) or not re.fullmatch(r'[A-Z][A-Z0-9_]*',env) or not isinstance(server,dict) or set(server)!={'command','path'}:
                raise ValueError('invalid render server')
            _argv(server['command'],port=True)
            path=server['path']
            if not isinstance(path,str) or not path.startswith('/') or path.startswith('//') or any(c in path for c in ('#','\\','\0','\n','\r')):
                raise ValueError('render server path must be local to the assigned origin')
        selected[name]=row
    return selected


def safe_path(root,relative):
    _relative(relative)
    root=Path(root).resolve();path=root/relative
    if not path.resolve().is_relative_to(root):raise ValueError('ship artifact escapes repository')
    cursor=path
    while cursor!=root:
        if cursor.is_symlink():raise ValueError('symlinked ship artifacts are unsupported')
        cursor=cursor.parent
    return path


def artifact_hashes(root,paths):
    """Expand configured output directories into the exact nonempty file/hash set."""
    root=Path(root).resolve()
    files={}
    for name in paths:
        path=safe_path(root,name)
        if not path.exists():raise ValueError('ship artifact missing: '+name)
        candidates=[path] if path.is_file() else sorted(path.rglob('*'))
        found=0
        for candidate in candidates:
            rel=candidate.relative_to(root).as_posix();safe_path(root,rel)
            if candidate.is_dir():continue
            if not candidate.is_file():raise ValueError('unsupported artifact type')
            content=candidate.read_bytes()
            files[rel]=hashlib.sha256(content).hexdigest();found+=1
        if not found:raise ValueError('zero ship artifact files: '+name)
    if not files:raise ValueError('zero ship artifact files')
    return dict(sorted(files.items()))


def read_json(path):
    return json.loads(Path(path).read_text())


def validate_web_result(root,result,urls):
    """Reject a zero/partial result even when an orchestrator exits successfully."""
    root=Path(root)
    if result.get('exit_code') not in (0,1) or result.get('errors'):
        raise ValueError('L0/L1/L2 result failed')
    config=read_json(root/'design/validation.json')
    expected=[row['env'] for row in config['urls']]
    if not expected or len(set(expected))!=len(expected) or set(expected)!=set(urls):
        raise ValueError('render servers must cover exactly all validation endpoints')
    stages=result.get('stages',[])
    if [r.get('stage') for r in stages]!=['L0','L1']+['L2']*len(expected):
        raise ValueError('incomplete L0/L1/L2 stage evidence')
    if stages[0].get('exit_code') not in (0,1) or stages[0].get('counts',{}).get('items',0)<=0:
        raise ValueError('zero or failed L0 stage')
    if stages[1].get('exit_code')!=0 or stages[1].get('targets')!=len(config['html']) or not config['html']:
        raise ValueError('zero, partial or failed L1 stage')
    environment=root/'design/render-environment.json';fixed=read_json(environment)
    environment_hash=hashlib.sha256(environment.read_bytes()).hexdigest()
    evidence={}
    for i,env in enumerate(expected):
        stage=stages[i+2]
        if stage.get('exit_code')!=0 or stage.get('target_env')!=env:
            raise ValueError('partial or failed L2 stage')
        path=safe_path(root,'design/.resolved/l2-'+str(i)+'.json');measured=read_json(path)
        actual=measured.get('environment',{})
        if measured.get('url')!=urls[env] or actual.get('repeat_equal') is not True or actual.get('runtime')!=fixed['runtime'] or actual.get('config_sha256')!=environment_hash or measured.get('failures'):
            raise ValueError('unbound or failed fixed rendering evidence')
        runs=measured.get('runs',[])
        expected_runs=[(repeat,scheme,width) for repeat in (0,1) for scheme in config['schemes'] for width in config['viewports']]
        if not expected_runs or [(r.get('repeat'),r.get('scheme'),r.get('width')) for r in runs]!=expected_runs:
            raise ValueError('zero or partial fixed rendering runs')
        half=len(runs)//2
        for first,second in zip(runs[:half],runs[half:]):
            a={k:v for k,v in first.items() if k!='repeat'};b={k:v for k,v in second.items() if k!='repeat'}
            if a!=b or not re.fullmatch(r'[a-f0-9]{64}',a.get('render_evidence',{}).get('screenshot_sha256','')):
                raise ValueError('fixed rendering repeats differ')
        evidence['design/.resolved/l2-'+str(i)+'.json']=hashlib.sha256(path.read_bytes()).hexdigest()
    return evidence
