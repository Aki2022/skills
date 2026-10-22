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
    if policy['schema_version']!=1 or not all(isinstance(policy[k],str) and policy[k].strip() for k in ('model','prompt_version')):
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


def migration_evidence(before,after):
    old={r['id']:r for r in before['holdout']['cases']}
    new={r['id']:r for r in after['holdout']['cases']}
    if set(old)!=set(new) or any(old[k]['positive']!=new[k]['positive'] for k in old):
        raise ValueError('migration must compare the same labeled holdout')
    return {'previous_policy':before['policy'],'previous_holdout':before['holdout'],
        'before':holdout_metrics(before['holdout'],before['policy']['theta']),
        'after':holdout_metrics(after['holdout'],after['policy']['theta']),
        'changed':sorted(k for k in old if (old[k]['score']>=before['policy']['theta'])!=(new[k]['score']>=after['policy']['theta']))}


def validate_ledger(ledger,policy):
    if not isinstance(ledger,dict) or set(ledger)-{'schema_version','policy','rows','holdout','migration'} or not {'schema_version','policy','rows','holdout'}<=set(ledger):
        raise ValueError('invalid judgment ledger')
    if ledger['schema_version']!=1 or policy_digest(ledger['policy'])!=policy_digest(policy):
        raise ValueError('judgment policy changed; migrate with a holdout comparison')
    metrics=holdout_metrics(ledger['holdout'],policy['theta'])
    if not isinstance(ledger['rows'],list):raise ValueError('invalid judgment rows')
    rows={}
    for row in ledger['rows']:
        if not isinstance(row,dict) or set(row)!={'judgment_hash','score'} or not isinstance(row['judgment_hash'],str) or len(row['judgment_hash'])!=64 or any(c not in '0123456789abcdef' for c in row['judgment_hash']):
            raise ValueError('invalid judgment key/row')
        if row['judgment_hash'] in rows:raise ValueError('duplicate judgment key')
        rows[row['judgment_hash']]=probability(row['score'])
    if 'migration' in ledger:
        proof=ledger['migration']
        if not isinstance(proof,dict) or 'previous_policy' not in proof or 'previous_holdout' not in proof:
            raise ValueError('invalid judgment migration evidence')
        validate_policy(proof['previous_policy'])
        expected=migration_evidence({'policy':proof['previous_policy'],'holdout':proof['previous_holdout']},ledger)
        if proof!=expected:raise ValueError('judgment migration evidence mismatch')
    return rows,metrics


def evaluate(candidates,ledger,*,policy=None):
    policy=validate_policy(load_policy() if policy is None else policy)
    rows,metrics=validate_ledger(ledger,policy) if ledger is not None else ({},None)
    result={'errors':[],'uncovered':[],'screened':[]}
    for row in candidates:
        if row['judgment_hash'] not in rows:
            result['uncovered'].append({'id':row['id'],'judgment_hash':row['judgment_hash'],'reason':'judgment not recorded for current input/policy'})
        elif rows[row['judgment_hash']]>=policy['theta']:
            result['errors'].append('judgment flags evidence: '+row['id'])
        else:
            result['screened'].append({'id':row['id'],'judgment_hash':row['judgment_hash'],'status':'screened',
                'score':rows[row['judgment_hash']],'theta':policy['theta'],'holdout':metrics})
    return result


def migrate_ledger(previous,replacement,policy):
    validate_ledger(previous,previous['policy'])
    validate_ledger(replacement,policy)
    updated=copy.deepcopy(replacement)
    updated['migration']=migration_evidence(previous,replacement)
    validate_ledger(updated,policy)
    return updated


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('command',choices=['record','migrate'])
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
            else:
                if path.exists():raise ValueError('existing ledger requires migrate')
                validate_ledger(updated,policy)
            if hashlib.sha256(POLICY.read_bytes()).hexdigest()!=policy_sha256:
                raise ValueError('judgment policy changed during recording')
            if args.write:write_data(path,updated)
        report={'status':'recorded' if args.write else 'preview','rows':len(updated['rows']),
            'holdout':holdout_metrics(updated['holdout'],policy['theta'])}
        if 'migration' in updated:report['migration']=updated['migration']
        print(json.dumps(report,ensure_ascii=False))
        return 0
    except (OSError,ValueError,KeyError,TypeError):
        print(json.dumps({'status':'rejected','error':'judgment record/migration requires valid current policy and matching labeled holdout'}))
        return 2


if __name__=='__main__':raise SystemExit(main())
