"""Offline judgment reader. Model execution is deliberately outside validation."""
import argparse
import copy
import fcntl
import hashlib
import json
import math
from pathlib import Path

POLICY = Path(__file__).resolve().parents[1]/'references/judgment-policy.json'


def digest(value):
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,
        separators=(',',':'),allow_nan=False).encode()).hexdigest()


def probability(value):
    if type(value) not in (int,float) or not math.isfinite(value) or not 0<=value<=1:
        raise ValueError('invalid judgment probability')
    return value


def validate_policy(policy):
    if not isinstance(policy,dict) or set(policy)!={'schema_version','model','theta','prompt_version','prompts'}:
        raise ValueError('invalid judgment policy')
    if type(policy['schema_version']) is not int or policy['schema_version']!=1 or not all(isinstance(policy[k],str) and policy[k].strip() for k in ('model','prompt_version')):
        raise ValueError('invalid judgment policy version/model')
    probability(policy['theta'])
    if not isinstance(policy['prompts'],dict) or set(policy['prompts'])!={'doc-relevance','na-rationale'} or any(not isinstance(v,str) or not v.strip() for v in policy['prompts'].values()):
        raise ValueError('invalid judgment prompts')
    return policy


def read_policy():
    raw=POLICY.read_bytes()
    return validate_policy(json.loads(raw)),hashlib.sha256(raw).hexdigest()


def load_policy():
    return read_policy()[0]


def policy_digest(policy):
    return digest(validate_policy(policy))


def candidate(kind,ident,text,context,*,policy=None):
    policy=validate_policy(load_policy() if policy is None else policy)
    if kind not in policy['prompts'] or not isinstance(text,str) or not text.strip():
        raise ValueError('invalid judgment candidate')
    text=' '.join(text.split())
    span_hash=hashlib.sha256(text.encode()).hexdigest()
    # Context matters: identical prose for a different catalog item must not reuse a judgment.
    inputs={'span_hash':span_hash,'context_hash':digest({'kind':kind,'id':ident,'context':context}),
        'model_digest':digest({'provider':'typesafe','model':policy['model']}),
        'prompt_version':policy['prompt_version'],'prompt_digest':digest(policy['prompts']),
        'theta':policy['theta']}
    return {'id':kind+':'+ident,'judgment_hash':digest(inputs),**inputs,
        'input':{'kind':kind,'catalog_id':ident,'text':text,'context':context}}


def holdout_metrics(holdout,theta):
    if not isinstance(holdout,dict) or set(holdout)!={'source','cases'} or holdout['source']!='owner':
        raise ValueError('owner-labeled holdout is required')
    cases=holdout['cases']
    if not isinstance(cases,list) or not cases:raise ValueError('empty holdout')
    seen=set();positives=tp=fp=0
    for row in cases:
        if not isinstance(row,dict) or set(row)!={'id','positive','score'} or type(row['positive']) is not bool:
            raise ValueError('invalid holdout case')
        if not isinstance(row['id'],str) or not row['id'].strip() or row['id'] in seen:
            raise ValueError('duplicate/invalid holdout identity')
        seen.add(row['id']);flag=probability(row['score'])>=theta
        positives+=row['positive'];tp+=flag and row['positive'];fp+=flag and not row['positive']
    if positives==0 or positives==len(cases):raise ValueError('holdout needs positive and negative cases')
    return {'total':len(cases),'positives':positives,'true_positives':tp,
        'false_negatives':positives-tp,'false_positives':fp,'recall':tp/positives}


def evaluation_metrics(holdout,theta):
    """An explicit absence of labels is not a measured zero or a synthetic recall."""
    if holdout=={'source':'unmeasured'}:
        return {'status':'unmeasured'}
    return {'status':'owner-labeled',**holdout_metrics(holdout,theta)}


def _candidate_row(row,policy):
    if not isinstance(row,dict) or set(row)!={'candidate','score'}:
        raise ValueError('invalid judgment candidate row')
    target=row['candidate']
    try:
        data=target['input']
        if (not isinstance(data['catalog_id'],str) or not data['catalog_id'].strip()
                or not isinstance(data['context'],dict)):
            raise ValueError('invalid candidate identity/context')
        expected=candidate(data['kind'],data['catalog_id'],data['text'],data['context'],policy=policy)
        if digest(target)!=digest(expected):raise ValueError('judgment candidate identity mismatch')
    except (KeyError,TypeError,AttributeError):
        raise ValueError('invalid judgment candidate') from None
    return target,probability(row['score'])


def _ledger_rows(ledger,policy):
    if (not isinstance(ledger,dict) or set(ledger)-{'schema_version','policy','rows','holdout','migration','refresh'}
            or not {'schema_version','policy','rows','holdout'}<=set(ledger)):
        raise ValueError('invalid judgment ledger')
    version=ledger['schema_version']
    if type(version) is not int or version not in (1,2):raise ValueError('invalid judgment ledger version')
    if policy_digest(ledger['policy'])!=policy_digest(policy):
        raise ValueError('judgment policy changed; migrate with a same input comparison')
    # Schema 1 retains its original owner-labeled contract. Schema 2 permits an
    # explicit unmeasured evaluation and binds each score to its full input.
    metrics=({'status':'owner-labeled',**holdout_metrics(ledger['holdout'],policy['theta'])}
        if version==1 else evaluation_metrics(ledger['holdout'],policy['theta']))
    if not isinstance(ledger['rows'],list):raise ValueError('invalid judgment rows')
    rows={};identities=set()
    for row in ledger['rows']:
        if version==1:
            if (not isinstance(row,dict) or set(row)!={'judgment_hash','score'}
                    or not isinstance(row['judgment_hash'],str) or len(row['judgment_hash'])!=64
                    or any(c not in '0123456789abcdef' for c in row['judgment_hash'])):
                raise ValueError('invalid judgment key/row')
            hashed=row['judgment_hash'];score=probability(row['score'])
        else:
            target,score=_candidate_row(row,policy);hashed=target['judgment_hash']
            ident=(target['id'],digest(target['input']))
            if ident in identities:raise ValueError('duplicate judgment input identity')
            identities.add(ident)
        if hashed in rows:raise ValueError('duplicate judgment key')
        rows[hashed]=score
    return rows,metrics


def _comparison_rows(ledger,other):
    if ledger['schema_version']==2:
        return {(r['candidate']['id'],digest(r['candidate']['input'])):
            {'judgment_hash':r['candidate']['judgment_hash'],'score':r['score']} for r in ledger['rows']}
    # Legacy rows lack input identities. They may be upgraded under the same
    # policy using replacement candidates that reproduce the exact old hashes.
    # A policy migration needs that upgrade first; guessing identities is unsafe.
    if other['schema_version']==2:
        if policy_digest(ledger['policy'])!=policy_digest(other['policy']):
            raise ValueError('same input proof requires upgrading legacy rows before policy changes')
        by_hash={r['candidate']['judgment_hash']:r['candidate'] for r in other['rows']}
        if set(by_hash)!={r['judgment_hash'] for r in ledger['rows']}:
            raise ValueError('migration must compare the same input set')
        return {(by_hash[r['judgment_hash']]['id'],digest(by_hash[r['judgment_hash']]['input'])):r
            for r in ledger['rows']}
    if policy_digest(ledger['policy'])!=policy_digest(other['policy']):
        raise ValueError('same input proof requires upgrading legacy rows before policy changes')
    return {(r['judgment_hash'],r['judgment_hash']):r for r in ledger['rows']}


def migration_evidence(before,after):
    _,before_metrics=_ledger_rows(before,before['policy'])
    _,after_metrics=_ledger_rows(after,after['policy'])
    old=_comparison_rows(before,after);new=_comparison_rows(after,before)
    if set(old)!=set(new):raise ValueError('migration must compare the same input set')
    if policy_digest(before['policy'])==policy_digest(after['policy']):
        if any(digest(old[key]['score'])!=digest(new[key]['score']) for key in old):
            raise ValueError('same policy migration must retain unchanged judgment scores')
        if digest(before['holdout'])!=digest(after['holdout']):
            raise ValueError('same policy migration must retain unchanged holdout')
    inputs=[]
    for ident,input_hash in sorted(old):
        a=old[(ident,input_hash)];b=new[(ident,input_hash)]
        inputs.append({'id':ident,'input_hash':input_hash,
            'before_judgment_hash':a['judgment_hash'],'after_judgment_hash':b['judgment_hash'],
            'before_score':a['score'],'after_score':b['score'],
            'before_flag':a['score']>=before['policy']['theta'],
            'after_flag':b['score']>=after['policy']['theta']})
    # Existing real owner labels remain comparable. Never manufacture labels or
    # metrics when either side is explicitly unmeasured.
    changed=[]
    if before['holdout'].get('source')=='owner' and after['holdout'].get('source')!='owner':
        raise ValueError('migration must retain existing owner-labeled holdout')
    if before['holdout'].get('source')=='owner' and after['holdout'].get('source')=='owner':
        old_labels={r['id']:r for r in before['holdout']['cases']}
        new_labels={r['id']:r for r in after['holdout']['cases']}
        if set(old_labels)!=set(new_labels) or any(old_labels[k]['positive']!=new_labels[k]['positive'] for k in old_labels):
            raise ValueError('migration must compare the same labeled holdout')
        changed=sorted(k for k in old_labels if (old_labels[k]['score']>=before['policy']['theta'])!=(new_labels[k]['score']>=after['policy']['theta']))
    return {'previous_schema_version':before['schema_version'],'previous_policy':before['policy'],
        'previous_rows':before['rows'],'previous_holdout':before['holdout'],
        'before':before_metrics,'after':after_metrics,'inputs':inputs,
        'changed_inputs':[{'id':r['id'],'input_hash':r['input_hash']} for r in inputs if r['before_flag']!=r['after_flag']],
        'changed':changed}


def _legacy_migration_evidence(proof,ledger):
    """Verify an already recorded schema-1 proof under its original contract.

    This compatibility path is read-only: new migrations use input-bound proof.
    Legacy proofs never contained the old candidate rows and therefore cannot
    claim that the old/new current input sets were compared.
    """
    previous_policy=validate_policy(proof['previous_policy'])
    before=holdout_metrics(proof['previous_holdout'],previous_policy['theta'])
    after=holdout_metrics(ledger['holdout'],ledger['policy']['theta'])
    old={r['id']:r for r in proof['previous_holdout']['cases']}
    new={r['id']:r for r in ledger['holdout']['cases']}
    if set(old)!=set(new) or any(old[k]['positive']!=new[k]['positive'] for k in old):
        raise ValueError('migration must compare the same labeled holdout')
    return {'previous_policy':previous_policy,'previous_holdout':proof['previous_holdout'],
        'before':before,'after':after,'changed':sorted(k for k in old if
            (old[k]['score']>=previous_policy['theta'])!=(new[k]['score']>=ledger['policy']['theta']))}


def refresh_evidence(before,after):
    old,_=_ledger_rows(before,before['policy']);new,_=_ledger_rows(after,after['policy'])
    if before['schema_version']!=2 or after['schema_version']!=2:
        raise ValueError('refresh requires schema 2 input identities')
    if policy_digest(before['policy'])!=policy_digest(after['policy']):
        raise ValueError('refresh cannot change judgment policy; use migrate')
    retained=sorted(set(old)&set(new))
    if any(digest(old[key])!=digest(new[key]) for key in retained):
        raise ValueError('refresh cannot rescore unchanged judgment inputs')
    if digest(before['holdout'])!=digest(after['holdout']):
        raise ValueError('refresh must retain the existing holdout unchanged')
    return {'previous_policy':before['policy'],'previous_rows':before['rows'],
        'previous_holdout':before['holdout'],'retained':retained,
        'added':sorted(set(new)-set(old)),'removed':sorted(set(old)-set(new))}


def validate_ledger(ledger,policy):
    rows,metrics=_ledger_rows(ledger,policy)
    if 'migration' in ledger and 'refresh' in ledger:
        raise ValueError('ambiguous judgment transition evidence')
    if 'migration' in ledger:
        proof=ledger['migration']
        required={'previous_schema_version','previous_policy','previous_rows','previous_holdout',
            'before','after','inputs','changed_inputs','changed'}
        legacy={'previous_policy','previous_holdout','before','after','changed'}
        if isinstance(proof,dict) and ledger['schema_version']==1 and set(proof)==legacy:
            expected=_legacy_migration_evidence(proof,ledger)
        else:
            if not isinstance(proof,dict) or set(proof)!=required:
                raise ValueError('invalid judgment migration evidence')
            previous={'schema_version':proof['previous_schema_version'],'policy':proof['previous_policy'],
                'rows':proof['previous_rows'],'holdout':proof['previous_holdout']}
            expected=migration_evidence(previous,ledger)
        if digest(proof)!=digest(expected):raise ValueError('judgment migration evidence mismatch')
    if 'refresh' in ledger:
        proof=ledger['refresh']
        required={'previous_policy','previous_rows','previous_holdout','retained','added','removed'}
        if not isinstance(proof,dict) or set(proof)!=required:
            raise ValueError('invalid judgment refresh evidence')
        previous={'schema_version':2,'policy':proof['previous_policy'],
            'rows':proof['previous_rows'],'holdout':proof['previous_holdout']}
        if digest(proof)!=digest(refresh_evidence(previous,ledger)):
            raise ValueError('judgment refresh evidence mismatch')
    return rows,metrics


def evaluate(candidates,ledger,*,policy=None):
    policy=validate_policy(load_policy() if policy is None else policy)
    rows,metrics=validate_ledger(ledger,policy) if ledger is not None else ({},{'status':'unmeasured'})
    result={'errors':[],'warnings':[],'uncovered':[],'screened':[]}
    for row in candidates:
        if row['judgment_hash'] not in rows:
            result['uncovered'].append({'id':row['id'],'judgment_hash':row['judgment_hash'],'reason':'judgment not recorded for current input/policy'})
        else:
            score=rows[row['judgment_hash']];flagged=score>=policy['theta']
            screened={'id':row['id'],'judgment_hash':row['judgment_hash'],'status':'screened',
                'score':score,'theta':policy['theta'],'flagged_advisory':flagged,
                'semantic_guarantee':False,'evaluation':metrics}
            if metrics['status']=='owner-labeled':
                screened['holdout']={k:v for k,v in metrics.items() if k!='status'}
            result['screened'].append(screened)
            if flagged:result['warnings'].append({'id':row['id'],'judgment_hash':row['judgment_hash'],
                'score':score,'theta':policy['theta'],'reason':'semantic evidence requires source review; advisory only'})
    return result


def migrate_ledger(previous,replacement,policy):
    validate_ledger(previous,previous['policy'])
    validate_ledger(replacement,policy)
    updated=copy.deepcopy(replacement)
    updated.pop('refresh',None)
    updated['migration']=migration_evidence(previous,replacement)
    validate_ledger(updated,policy)
    return updated


def refresh_ledger(previous,replacement,policy):
    validate_ledger(previous,previous['policy'])
    validate_ledger(replacement,policy)
    updated=copy.deepcopy(replacement)
    updated.pop('migration',None)
    updated['refresh']=refresh_evidence(previous,replacement)
    validate_ledger(updated,policy)
    return updated


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command',choices=['record','migrate','refresh'])
    parser.add_argument('--repo',default='.')
    parser.add_argument('--replacement',type=Path,required=True)
    parser.add_argument('--write',action='store_true')
    args=parser.parse_args()
    from design_author import git_root,inside,read_data,write_data
    try:
        design=git_root(args.repo)/'design';path=design/'judgments.json'
        directory=design/'.resolved'
        if design.is_symlink() or path.is_symlink() or directory.is_symlink():
            raise ValueError('symlinked judgment input/output')
        directory.mkdir(parents=True,exist_ok=True)
        with (directory/'judgments.lock').open('a') as lock:
            try:fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
            except BlockingIOError as exc:raise ValueError('judgment writer is active') from exc
            updated=read_data(args.replacement);policy,policy_sha256=read_policy()
            if args.command=='migrate':updated=migrate_ledger(read_data(path),updated,policy)
            elif args.command=='refresh':updated=refresh_ledger(read_data(path),updated,policy)
            else:
                if path.exists():raise ValueError('existing ledger requires refresh or migrate')
                validate_ledger(updated,policy)
            if hashlib.sha256(POLICY.read_bytes()).hexdigest()!=policy_sha256:
                raise ValueError('judgment policy changed during recording')
            if args.write:write_data(path,updated)
        report={'status':'recorded' if args.write else 'preview','rows':len(updated['rows']),
            'evaluation':evaluation_metrics(updated['holdout'],policy['theta'])}
        if 'migration' in updated:report['migration']=updated['migration']
        if 'refresh' in updated:report['refresh']=updated['refresh']
        print(json.dumps(report,ensure_ascii=False))
        return 0
    except (OSError,ValueError,KeyError,TypeError):
        print(json.dumps({'status':'rejected','error':'judgment record/refresh/migration requires valid current policy and matching input evidence'}))
        return 2


if __name__=='__main__':raise SystemExit(main())
