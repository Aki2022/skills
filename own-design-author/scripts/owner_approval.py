"""Prepare review requests; the AI records only what the owner approved in their own words (quoted, dated)."""
from datetime import date
import os
from pathlib import Path
import tempfile
import fcntl
from approvals import SCHEME,repository_id,verify_rows


# Approval gaps, not structural errors: preparing that very approval must not be refused because of them.
# IA_POSITIONING_UNAPPROVED (a value path without the owner's current approval) and IA_OUTLINE_VALUE_UNAPPROVED (a
# shipped message pointing at such a value) are approval gaps like the two design ones (IA v6, biz_ops WS-20261114;
# ADR-20261118-ia-approval-per-value-path). IA_OUTLINE_COMPARE_UNAPPROVED no longer exists: a compared alternative
# is approved with the path it belongs to. Positioning is approved before
# any outline is written, so "no outlines yet" (IA_NO_OUTLINES, zero IA check targets) is the expected state of a
# positioning approval, not damage (ADR-20261118-ai-records-owner-approvals, biz_ops).
ALLOWED=('must n-a approval hash missing/mismatch:','approval hash mismatch:','IA_POSITIONING_UNAPPROVED:',
         'IA_OUTLINE_VALUE_UNAPPROVED:','IA_OUTLINE_COMPARE_UNAPPROVED:','IA_NO_OUTLINES:','zero IA check targets',
         # A disqualified path that a shipped outline points at fails; the owner may still be asked to approve it (the
         # request page lists what keeps it from passing). Damage to the data itself (EDGE_SHAPE, EDGE_DUPLICATE,
         # CROSS_SEGMENT, COPY_SOURCE, references, schema, COMPARE_OUTSIDE_SET) still refuses.
         'IA_POSITIONING_ISOLATED_ATOM:','IA_POSITIONING_JOB_WITHOUT_NEED:','IA_POSITIONING_STRENGTH_UNCOMPARED:',
         'IA_POSITIONING_INCOMPLETE:','IA_POSITIONING_CATEGORY_COUNT:','IA_POSITIONING_NO_ALTERNATIVES:',
         'IA_POSITIONING_NO_CATEGORY_ALTERNATIVES:','IA_POSITIONING_FULFILLMENT_MISSING:','IA_POSITIONING_VALUE_UNANCHORED:')


def _ia_only(root):
    """A repository with design/info_arch and no design/design.yaml: the IA resolver alone gives the targets."""
    return not (Path(root)/'design/design.yaml').exists() and (Path(root)/'design/info_arch').is_dir()


def _ia_resolver():
    import importlib.util
    implementation=Path(__file__).resolve().parents[2]/'own-infoarch-author/scripts/ia_module.py'
    if (not implementation.is_file() or implementation.is_symlink() or implementation.parent.is_symlink()
            or implementation.parent.parent.is_symlink()):
        raise ValueError('IA module is unavailable')
    spec=importlib.util.spec_from_file_location('approval_ia_module',implementation)
    module=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _snapshot(root,catalog_dir):
    if _ia_only(root):
        result=_ia_resolver().resolve(root=root,approvals={},profiles={})
        if any(not error.startswith(ALLOWED) for error in result['errors']):
            raise ValueError('repair structural or approval errors before preparing approval')
        return result['approval_hashes'],sorted(set(result['must_approval_keys']))
    from design_author import check_design
    result=check_design(root,catalog_dir=catalog_dir)
    if any(not error.startswith(ALLOWED) for error in result['errors']) or not result['manifest']:
        raise ValueError('repair structural or approval errors before preparing approval')
    return result['approval_targets'],sorted(set(result['approval_required']))


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


def _owner_approval(owner_words,approved_on):
    """The AI records only what the owner approved in their own words; the words and date go into the commit."""
    if not isinstance(owner_words,str) or not owner_words.strip():
        raise ValueError("recording needs the owner's approval words (quote them exactly)")
    try:ok=isinstance(approved_on,str) and date.fromisoformat(approved_on).isoformat()==approved_on
    except ValueError:ok=False
    if not ok:raise ValueError("recording needs the ISO date of the owner's approval")
    return {'words':owner_words.strip(),'date':approved_on}


def _approve_request(root,catalog_dir,request,approver,owner_words,approved_on):
    from design_author import read_data,check_design
    owner=_owner_approval(owner_words,approved_on)
    if not isinstance(request,dict) or set(request)!={'schema_version','approval_scheme','repo_id','rows'} or request['schema_version']!=2 or request['approval_scheme']!=SCHEME:
        raise ValueError('invalid approval request')
    rows=request['rows']
    if not isinstance(rows,list) or not rows or any(not isinstance(row,dict) or set(row)!={'key','value_hash','date'} for row in rows):
        raise ValueError('invalid approval request rows')
    for row in rows:
        if not isinstance(row['date'],str) or date.fromisoformat(row['date']).isoformat()!=row['date']:
            raise ValueError('invalid approval request date')
    if any(row['date']!=owner['date'] for row in rows):
        raise ValueError("the request's approval date differs from the day the owner approved (prepare it again)")
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
    path=Path(root)/'design/approvals.yaml'
    ledger=(read_data(path) or []) if path.is_file() else []
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
    if _ia_only(root):
        ia=_ia_resolver().resolve(root=root,approvals={},profiles={})
        result={'exit_code':2 if ia['errors'] else (1 if ia['warnings'] else 0),'errors':ia['errors'],'warnings':ia['warnings']}
    else:
        result=check_design(root,catalog_dir=catalog_dir)
    return {'status':'recorded','recorded_rows':len(recorded),'exit_code':result['exit_code'],
        'errors':result['errors'],'warnings':result['warnings'],'owner_approval':owner,
        'commit_trailer':f"Owner-Approval: 「{owner['words']}」 ({owner['date']})"}


def approve_request(root,catalog_dir,request,approver,*,owner_words=None,approved_on=None):
    root=Path(root)
    _owner_approval(owner_words,approved_on)
    if (root/'design/approvals.yaml').is_symlink():raise ValueError('symlinked ledger is unsupported')
    directory=root/'design/.resolved'
    if directory.is_symlink():raise ValueError('symlinked output directory is unsupported')
    directory.mkdir(parents=True,exist_ok=True)
    with (directory/'approval.lock').open('a') as lock:
        try:fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        except BlockingIOError as exc:raise ValueError('another approval command is active') from exc
        return _approve_request(root,catalog_dir,request,approver,owner_words,approved_on)
