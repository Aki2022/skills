"""Prepare review requests; record only after explicit human review."""
from datetime import date
import os
from pathlib import Path
import tempfile
import fcntl
from approvals import SCHEME,repository_id,verify_rows


def _snapshot(root,catalog_dir):
    from design_author import check_design
    result=check_design(root,catalog_dir=catalog_dir)
    allowed=('must n-a approval hash missing/mismatch:','approval hash mismatch:')
    if any(not error.startswith(allowed) for error in result['errors']) or not result['manifest']:
        raise ValueError('repair structural or approval errors before preparing approval')
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
    return {'schema_version':2,'approval_scheme':SCHEME,'repo_id':repository_id(root),
        'rows':[{'key':key,'value_hash':values[key],'date':when} for key in sorted(targets)]}


def _approve_request(root,catalog_dir,request,approver):
    from design_author import read_data,check_design
    if not isinstance(request,dict) or set(request)!={'schema_version','approval_scheme','repo_id','rows'} or request['schema_version']!=2 or request['approval_scheme']!=SCHEME:
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
    recorded=[{**row,'approver':approver} for row in rows]
    verify_rows(recorded,root)
    current,_=_snapshot(root,catalog_dir)
    if repository_id(root)!=identity or any(current.get(row['key'])!=row['value_hash'] for row in rows):
        raise ValueError('reviewed inputs changed during recording; no rows written')
    ledger=read_data(Path(root)/'design/approvals.yaml')
    # Preserve unrelated owner approvals. No partial writes if validation fails.
    combined=[row for row in ledger if row['key'] not in keys]+recorded
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
    return {'status':'recorded','recorded_rows':len(recorded),'exit_code':result['exit_code'],
        'errors':result['errors'],'warnings':result['warnings']}


def approve_request(root,catalog_dir,request,approver):
    root=Path(root)
    if (root/'design/approvals.yaml').is_symlink():raise ValueError('symlinked ledger is unsupported')
    directory=root/'design/.resolved'
    if directory.is_symlink():raise ValueError('symlinked output directory is unsupported')
    directory.mkdir(parents=True,exist_ok=True)
    with (directory/'approval.lock').open('a') as lock:
        try:fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        except BlockingIOError as exc:raise ValueError('another approval command is active') from exc
        return _approve_request(root,catalog_dir,request,approver)
