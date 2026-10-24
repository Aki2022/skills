#!/usr/bin/env python3
"""Acquire bounded Jev Noul answers outside L0/L1/L2; never write approvals or labels."""
import argparse
import fcntl
import hashlib
import json
from pathlib import Path
import re
import sys
import time
import urllib.error

import classify as api
import judgment_ledger as jl

USAGE = Path(__file__).resolve().parents[1]/'catalog/classification-usage.json'


def _validate_candidate(row,policy):
    try:
        data=row['input']
        expected=jl.candidate(data['kind'],data['catalog_id'],data['text'],data['context'],policy=policy)
        if row!=expected:raise ValueError
        return expected
    except (KeyError,TypeError,ValueError):
        raise api.ClassificationError('invalid-candidates') from None


def _candidates(source,policy):
    data=api._load(source)
    try:
        rows=data['judgment_candidates']
        if not isinstance(rows,list):raise ValueError
        unique={}
        for row in rows:
            row=_validate_candidate(row,policy)
            unique[row['judgment_hash']]=row
        if not unique:raise api.ClassificationError('no-targets')
        return list(unique.values())
    except (KeyError,TypeError,ValueError):
        raise api.ClassificationError('invalid-candidates') from None


def _answers(path,policy):
    data=api._load(path) if path.exists() else {'schema_version':1,'status':'model-answers-unlabeled',
        'policy':policy,'rows':[]}
    try:
        if (set(data)!={'schema_version','status','policy','rows'} or data['schema_version']!=1
                or data['status']!='model-answers-unlabeled' or data['policy']!=policy
                or not isinstance(data['rows'],list)):raise ValueError
        seen=set()
        for row in data['rows']:
            if set(row)!={'candidate','score','raw_answer'}:raise ValueError
            target=_validate_candidate(row['candidate'],policy)
            hashed=target['judgment_hash']
            raw=row['raw_answer']
            if (not isinstance(raw,dict) or set(raw)!={'type','noul'} or raw['type']!='noul'
                    or jl.probability(raw['noul'])!=jl.probability(row['score']) or hashed in seen):raise ValueError
            seen.add(hashed)
        return data,seen
    except (KeyError,TypeError,ValueError,api.ClassificationError):
        raise api.ClassificationError('invalid-answers') from None


def _batches(rows):
    batches=[];batch=[];size=0
    for row in rows:
        length=len(api._json(row['input']))
        if length>api.MAX_SOURCE_CHARS:raise api.ClassificationError('item-too-large')
        if batch and (len(batch)==api.MAX_BATCH or size+length>api.MAX_SOURCE_CHARS):
            batches.append(batch);batch=[];size=0
        batch.append(row);size+=length
    if batch:batches.append(batch)
    return batches


def _payload(batch,policy):
    return {'model':policy['model'],'state':{'items':{f'q{i}':r['input'] for i,r in enumerate(batch)}},
        'questions':{f'q{i}':{'type':'noul','instructions':{
            'target':f"Evaluate only `items.q{i}.text` against `items.q{i}.context`; ignore other items.",
            'decision':policy['prompts'][r['input']['kind']]}} for i,r in enumerate(batch)}}


def _usage(response,prices):
    try:
        usage=response['usage']
        if any(type(usage[k]) is not int or usage[k]<0 for k in ('input_tokens','output_tokens')):raise ValueError
        return {**{k:usage[k] for k in ('input_tokens','output_tokens')},
            'usage_derived_usd':sum(usage[k+'_tokens']*prices[k+'_price_usd_per_million'] for k in ('input','output'))/1_000_000}
    except (KeyError,TypeError,ValueError):return None


def _safe_path(path):
    return not any(p.is_symlink() for p in (path,*path.parents))


def run(source,output,limit,*,dry_run=False,timeout=30,_usage_path=USAGE,_transport=None):
    if type(limit) is not int or limit<=0 or not api._number(timeout) or timeout<=0:
        raise api.ClassificationError('invalid-limits')
    source,output,usage_path=map(lambda p:Path(p).absolute(),(source,output,_usage_path))
    if (output.name!='judgment-answers.json' or len({p.resolve() for p in (source,output,usage_path)})!=3
            or any(not _safe_path(p) for p in (source,output,usage_path))
            or not output.parent.is_dir()):raise api.ClassificationError('invalid-paths')
    if not usage_path.is_file():raise api.ClassificationError('usage-ledger-required')
    policy,policy_sha=jl.read_policy()
    if not re.fullmatch(r'jev-[0-9]+\.[0-9]+\.[0-9]+',policy['model']):raise api.ClassificationError('unpinned-model')
    prices=api.load_policy()
    if prices['model']!=policy['model']:raise api.ClassificationError('price-model-mismatch')
    source_sha=hashlib.sha256(source.read_bytes()).hexdigest()
    candidates=_candidates(source,policy)
    answers,seen=_answers(output,policy)
    pending=[r for r in candidates if r['judgment_hash'] not in seen]
    selected=pending[:limit];batches=_batches(selected);ledger=api._usage_ledger(usage_path)
    if ledger['attempted_judgments']+len(selected)>api.MAX_JUDGMENTS:raise api.ClassificationError('budget-exceeded')
    result={'status':'dry-run' if dry_run else ('answered' if pending else 'already-complete'),
        'pending':len(pending),'selected':len(selected),'requests':len(batches),'answered':0,
        'remaining':len(pending),'attempted_judgments':ledger['attempted_judgments'],
        'max_judgments':api.MAX_JUDGMENTS,'input_tokens':0,'output_tokens':0,'usage_derived_usd':0.0}
    if dry_run or not pending:return result
    key=api.read_api_key()
    with usage_path.with_suffix(usage_path.suffix+'.lock').open('a') as lock:
        try:fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        except BlockingIOError:raise api.ClassificationError('usage-writer-active') from None
        # Recheck after the shared classification/acquisition lock, before reserving.
        ledger=api._usage_ledger(usage_path);answers,seen=_answers(output,policy)
        selected=[r for r in candidates if r['judgment_hash'] not in seen][:limit];batches=_batches(selected)
        if ledger['attempted_judgments']+len(selected)>api.MAX_JUDGMENTS:raise api.ClassificationError('budget-exceeded')
        if hashlib.sha256(source.read_bytes()).hexdigest()!=source_sha or jl.read_policy()[1]!=policy_sha:
            raise api.ClassificationError('input-changed')
        result.update(selected=len(selected),requests=len(batches),
            remaining=sum(r['judgment_hash'] not in seen for r in candidates))
        if not selected:result['status']='already-complete'
        for batch in batches:
            payload=_payload(batch,policy)
            entry={'purpose':'design-evidence','model':policy['model'],'questions':len(batch),
                'status':'reserved','request_sha256':api._hash(payload),'duration_seconds':None,
                'input_tokens':None,'output_tokens':None,'usage_derived_usd':None}
            ledger['attempted_judgments']+=len(batch);ledger['batches'].append(entry)
            api._save(usage_path,ledger)
            started=time.monotonic();response=None
            try:
                response=(_transport or api._post)(payload,key,timeout)
                telemetry=_usage(response,prices)
                if telemetry:entry.update(telemetry)
                if (not isinstance(response,dict) or response.get('model')!=policy['model'] or not telemetry
                        or not isinstance(response.get('answers'),dict)
                        or set(response['answers'])!=set(payload['questions'])):raise ValueError
                records=[]
                for i,row in enumerate(batch):
                    raw=response['answers'][f'q{i}']
                    if not isinstance(raw,dict) or set(raw)!={'type','noul'} or raw['type']!='noul':raise ValueError
                    score=jl.probability(raw['noul'])
                    records.append({'candidate':row,'score':score,'raw_answer':raw})
                if hashlib.sha256(source.read_bytes()).hexdigest()!=source_sha or jl.read_policy()[1]!=policy_sha:
                    raise ValueError
            except Exception as error:
                entry.update(status='failed',duration_seconds=round(time.monotonic()-started,6),
                    failure_signature='http-'+str(error.code) if isinstance(error,urllib.error.HTTPError) else type(error).__name__)
                api._save(usage_path,ledger)
                raise api.ClassificationError('request-failed') from None
            entry.update(status='succeeded',duration_seconds=round(time.monotonic()-started,6))
            api._save(usage_path,ledger)
            answers['rows'].extend(records);api._save(output,answers)
            result['answered']+=len(batch);result['remaining']-=len(batch)
            for k in ('input_tokens','output_tokens','usage_derived_usd'):result[k]+=telemetry[k]
        result['attempted_judgments']=ledger['attempted_judgments']
    return result


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--candidates',required=True,type=Path)
    parser.add_argument('--answers',required=True,type=Path,help='output basename must be judgment-answers.json')
    parser.add_argument('--limit',required=True,type=int)
    parser.add_argument('--execute',action='store_true',help='otherwise preview only; no auth, writes or network')
    args=parser.parse_args()
    try:
        result=run(args.candidates,args.answers,args.limit,dry_run=not args.execute)
        print(json.dumps(result,ensure_ascii=False,sort_keys=True));return 0
    except api.ClassificationError as error:
        print('judgment acquisition error: '+str(error),file=sys.stderr);return 2
    except Exception:
        print('judgment acquisition error: local-failure',file=sys.stderr);return 2


if __name__=='__main__':raise SystemExit(main())
