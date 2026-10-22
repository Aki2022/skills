"""Prepare review requests; only an explicit owner command invokes SSH signing."""
from datetime import date
import os
from pathlib import Path
import subprocess
import re
import tempfile
import fcntl
from approvals import protected_store,repository_id,payload,verify_rows,namespace
from jcs import canonicalize_batch


def _snapshot(root,catalog_dir):
    from design_author import check_design
    result=check_design(root,catalog_dir=catalog_dir)
    allowed=('must n-a approval hash missing/mismatch:','approval hash mismatch:')
    if any(not error.startswith(allowed) for error in result['errors']) or not result['manifest']:
        raise ValueError('repair structural or signature errors before preparing approval')
    values,must={},[]
    for media,manifest in result['manifest'].items():
        for kind,rows in [('token',manifest['tokens']),('item',manifest['items'])]:
            for target,row in rows.items():
                if 'value_hash' not in row:continue
                key=f'ds:{media}:{kind}:{target}'
                values[key]=row['value_hash']
                if kind=='item' and row['must']:must.append(key)
    return values,sorted(must)


def prepare(root,catalog_dir,targets=None,all_must=False,when=None):
    when=when or date.today().isoformat()
    if date.fromisoformat(when).isoformat()!=when:raise ValueError('invalid approval date')
    values,must=_snapshot(root,catalog_dir)
    if all_must and targets:raise ValueError('select either must items or explicit targets')
    targets=must if all_must else targets
    if not targets or len(set(targets))!=len(targets) or any(key not in values for key in targets):
        raise ValueError('zero, duplicate or unknown approval targets')
    return {'schema_version':1,'repo_id':repository_id(root),
        'rows':[{'key':key,'value_hash':values[key],'date':when} for key in sorted(targets)]}


def _sign(content,public_key,scope):
    try:
        run=subprocess.run(['ssh-keygen','-Y','sign','-f',str(public_key),'-n',scope],
            input=content,capture_output=True,timeout=60)
    except (OSError,subprocess.TimeoutExpired) as exc:
        raise ValueError('owner signing failed or timed out') from exc
    if run.returncode:raise ValueError('owner signing failed; no approval rows written')
    return run.stdout.decode()


def _approve_request(root,catalog_dir,request,public_key,approver):
    from design_author import read_data,check_design
    if not isinstance(request,dict) or set(request)!={'schema_version','repo_id','rows'} or request['schema_version']!=1:
        raise ValueError('invalid approval request')
    rows=request['rows']
    if not isinstance(rows,list) or not rows or any(not isinstance(row,dict) or set(row)!={'key','value_hash','date'} for row in rows):
        raise ValueError('invalid approval request rows')
    for row in rows:
        if not isinstance(row['date'],str) or date.fromisoformat(row['date']).isoformat()!=row['date']:
            raise ValueError('invalid approval request date')
    keys=[row['key'] for row in rows]
    if len(set(keys))!=len(keys):raise ValueError('duplicate approval request target')
    identity=repository_id(root)
    if request['repo_id']!=identity:raise ValueError('approval request belongs to another repository')
    values,_=_snapshot(root,catalog_dir)
    if any(values.get(row['key'])!=row['value_hash'] for row in rows):
        raise ValueError('reviewed inputs changed; regenerate approval request')
    protected_store(os.environ.get('DESIGN_ALLOWED_SIGNERS'),root)
    if not isinstance(approver,str) or not re.fullmatch(r'[A-Za-z0-9_.@+-]+',approver):
        raise ValueError('invalid owner principal')
    key=Path(public_key)
    try:key_text=key.read_text()
    except OSError as exc:raise ValueError('owner public key unavailable') from exc
    if key.suffix!='.pub' or not key_text.split() or key_text.split()[0] not in ('ssh-ed25519','sk-ssh-ed25519@openssh.com'):
        raise ValueError('supply an Ed25519 public key backed by the owner agent')
    signed=[]
    for row in rows:
        record={**row,'approver':approver}
        record['signature']=_sign(payload(identity,record),key,namespace(record['key']))
        signed.append(record)
    verify_rows(signed,root)
    current,_=_snapshot(root,catalog_dir)
    if repository_id(root)!=identity or any(current.get(row['key'])!=row['value_hash'] for row in rows):
        raise ValueError('reviewed inputs changed during signing; no rows written')
    ledger=read_data(Path(root)/'design/approvals.yaml')
    # Preserve unrelated owner approvals. No partial writes if signing/verification fails.
    combined=[row for row in ledger if row['key'] not in keys]+signed
    verify_rows(combined,root)
    target=Path(root)/'design/approvals.yaml'
    temporary=None
    try:
        with tempfile.NamedTemporaryFile(mode='w',encoding='utf-8',dir=target.parent,delete=False) as stream:
            import json
            temporary=Path(stream.name)
            json.dump(combined,stream,ensure_ascii=False,indent=2,allow_nan=False)
            stream.write('\n');stream.flush();os.fsync(stream.fileno())
        temporary.replace(target)
    finally:
        if temporary is not None and temporary.exists():temporary.unlink()
    result=check_design(root,catalog_dir=catalog_dir)
    return {'status':'recorded','signed_rows':len(signed),'exit_code':result['exit_code'],
        'errors':result['errors'],'warnings':result['warnings']}


def approve_request(root,catalog_dir,request,public_key,approver):
    root=Path(root)
    if (root/'design/approvals.yaml').is_symlink():raise ValueError('symlinked ledger is unsupported')
    directory=root/'design/.resolved'
    if directory.is_symlink():raise ValueError('symlinked output directory is unsupported')
    directory.mkdir(parents=True,exist_ok=True)
    with (directory/'approval.lock').open('a') as lock:
        try:fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        except BlockingIOError as exc:raise ValueError('another approval command is active') from exc
        return _approve_request(root,catalog_dir,request,public_key,approver)
