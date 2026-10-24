"""Acquisition tests use a transport double, never claim live-model accuracy."""
import copy
import json
import pytest
from test_core import SKILL
from judgment_ledger import candidate,load_policy
import classify
import judgment_run as m


def setup(tmp_path,count=12):
    source=tmp_path/'candidates.json';output=tmp_path/'judgment-answers.json';usage=tmp_path/'usage.json'
    rows=[candidate('doc-relevance',f'dads-{i:06d}',f'Policy {i}',{'title':'Brand'}) for i in range(count)]
    source.write_text(json.dumps({'judgment_candidates':rows}))
    usage.write_text(json.dumps({'schema_version':1,'attempted_judgments':610,'batches':[]}))
    return source,output,usage,rows


def transport(calls,usage):
    def respond(payload,*args):
        saved=json.loads(usage.read_text())
        assert saved['batches'][-1]['status']=='reserved'
        assert saved['batches'][-1]['purpose']=='design-evidence'
        calls.append(copy.deepcopy(payload))
        return {'model':payload['model'],'answers':{qid:{'type':'noul','noul':0.1} for qid in payload['questions']},
            'usage':{'input_tokens':100,'output_tokens':10}}
    return respond


def run(paths,limit=50,**kwargs):
    return m.run(paths[0],paths[1],limit,_usage_path=paths[2],**kwargs)


def test_dry_run_has_no_auth_network_or_writes(tmp_path,monkeypatch):
    paths=setup(tmp_path);before=paths[2].read_bytes()
    monkeypatch.setattr(classify,'read_api_key',lambda:pytest.fail('auth during dry run'))
    result=run(paths,dry_run=True,_transport=lambda *a:pytest.fail('network'))
    assert result['selected']==12 and result['requests']==2
    assert not paths[1].exists() and paths[2].read_bytes()==before


def test_reserve_shared_budget_resume_and_preserve_unlabeled_answers(tmp_path,monkeypatch):
    paths=setup(tmp_path);monkeypatch.setattr(classify,'read_api_key',lambda:'test-secret');calls=[]
    result=run(paths,limit=10,_transport=transport(calls,paths[2]))
    assert result['answered']==10 and result['attempted_judgments']==620
    assert result['usage_derived_usd']==pytest.approx(100*0.042/1_000_000)
    saved=json.loads(paths[1].read_text())
    assert saved['status']=='model-answers-unlabeled' and 'holdout' not in saved
    assert saved['policy']==load_policy() and len(saved['rows'])==10
    assert saved['rows'][0]['candidate']==paths[3][0]
    assert saved['rows'][0]['raw_answer']=={'type':'noul','noul':0.1}
    assert 'test-secret' not in paths[1].read_text()+paths[2].read_text()
    assert run(paths,_transport=transport(calls,paths[2]))['answered']==2
    assert run(paths,_transport=transport(calls,paths[2]))['status']=='already-complete'
    assert len(calls)==2 and json.loads(paths[2].read_text())['attempted_judgments']==622
    assert calls[0]['questions']['q0']['type']=='noul'
    assert calls[0]['state']['items']['q0']==paths[3][0]['input']
    assert calls[0]['questions']['q0']['instructions']['decision']==load_policy()['prompts']['doc-relevance']


def test_budget_and_missing_usage_fail_before_auth(tmp_path,monkeypatch):
    paths=setup(tmp_path);monkeypatch.setattr(classify,'read_api_key',lambda:pytest.fail('auth'))
    paths[2].write_text(json.dumps({'schema_version':1,'attempted_judgments':1999,'batches':[]}))
    with pytest.raises(classify.ClassificationError,match='budget'):run(paths)
    paths[2].unlink()
    with pytest.raises(classify.ClassificationError,match='usage-ledger-required'):run(paths,dry_run=True)


@pytest.mark.parametrize('kind',['wrong-model','missing-answer','bad-score','bad-usage','network'])
def test_failure_is_charged_sanitized_atomic_and_not_retried(tmp_path,monkeypatch,kind):
    paths=setup(tmp_path);monkeypatch.setattr(classify,'read_api_key',lambda:'test-secret');calls=[]
    def broken(payload,*args):
        response=transport(calls,paths[2])(payload,*args)
        if kind=='wrong-model':response['model']='private-response'
        if kind=='missing-answer':response['answers'].pop('q0')
        if kind=='bad-score':response['answers']['q0']['noul']=True
        if kind=='bad-usage':response['usage']['input_tokens']='private-response'
        if kind=='network':raise RuntimeError('test-secret private-response')
        return response
    with pytest.raises(classify.ClassificationError,match='request-failed'):run(paths,_transport=broken)
    saved=json.loads(paths[2].read_text())
    assert saved['attempted_judgments']==620 and saved['batches'][-1]['status']=='failed'
    assert len(calls)==1 and not paths[1].exists()
    assert 'private-response' not in paths[2].read_text() and 'test-secret' not in paths[2].read_text()
    if kind!='network' and kind!='bad-usage':assert saved['batches'][-1]['input_tokens']==100


def test_candidate_tampering_duplicates_and_policy_changes_fail(tmp_path):
    paths=setup(tmp_path,2);rows=paths[3]
    paths[0].write_text(json.dumps({'judgment_candidates':rows+[rows[0]]}))
    assert run(paths,dry_run=True)['selected']==2
    rows[0]['input']['text']='changed without updating digest'
    paths[0].write_text(json.dumps({'judgment_candidates':rows}))
    with pytest.raises(classify.ClassificationError,match='invalid-candidates'):run(paths,dry_run=True)


def test_successful_batch_survives_later_failure_and_tampered_output_rejected(tmp_path,monkeypatch):
    paths=setup(tmp_path);monkeypatch.setattr(classify,'read_api_key',lambda:'test-secret');calls=[]
    def once(payload,*args):
        if calls:raise RuntimeError('unavailable')
        return transport(calls,paths[2])(payload,*args)
    with pytest.raises(classify.ClassificationError):run(paths,_transport=once)
    assert len(json.loads(paths[1].read_text())['rows'])==10
    assert json.loads(paths[2].read_text())['attempted_judgments']==622
    saved=json.loads(paths[1].read_text());saved['rows'][0]['score']=0.9;paths[1].write_text(json.dumps(saved))
    with pytest.raises(classify.ClassificationError,match='invalid-answers'):run(paths,dry_run=True)


def test_empty_oversize_symlink_and_protected_output_rejected(tmp_path):
    paths=setup(tmp_path,0)
    with pytest.raises(classify.ClassificationError,match='no-targets'):run(paths,dry_run=True)
    paths=setup(tmp_path,1)
    with pytest.raises(classify.ClassificationError,match='invalid-paths'):
        m.run(paths[0],tmp_path/'judgments.json',1,dry_run=True,_usage_path=paths[2])
    linked=tmp_path/'link.json';linked.symlink_to(paths[0])
    with pytest.raises(classify.ClassificationError,match='invalid-paths'):
        m.run(linked,paths[1],1,dry_run=True,_usage_path=paths[2])
    row=candidate('doc-relevance','dads-000001','x'*16001,{'title':'Brand'})
    paths[0].write_text(json.dumps({'judgment_candidates':[row]}))
    with pytest.raises(classify.ClassificationError,match='item-too-large'):run(paths,dry_run=True)


def test_policy_mutation_and_malformed_raw_answer_cannot_be_reused(tmp_path,monkeypatch):
    import hashlib
    paths=setup(tmp_path,1);monkeypatch.setattr(classify,'read_api_key',lambda:'test-secret');calls=[]
    original=m.jl.read_policy
    def changed(payload,*args):
        response=transport(calls,paths[2])(payload,*args)
        policy,sha=original()
        monkeypatch.setattr(m.jl,'read_policy',lambda:(policy,'changed'))
        return response
    with pytest.raises(classify.ClassificationError,match='request-failed'):run(paths,_transport=changed)
    assert not paths[1].exists() and json.loads(paths[2].read_text())['attempted_judgments']==611
    monkeypatch.setattr(m.jl,'read_policy',original)
    run(paths,_transport=transport([],paths[2]))
    data=json.loads(paths[1].read_text());data['rows'][0].update(score=0,raw_answer={'type':'noul','noul':False})
    paths[1].write_text(json.dumps(data))
    with pytest.raises(classify.ClassificationError,match='invalid-answers'):run(paths,dry_run=True)


def test_real_cli_preview_and_default_transport_selection(tmp_path,monkeypatch):
    import subprocess,sys
    paths=setup(tmp_path,1)
    preview=subprocess.run([sys.executable,str(SKILL/'scripts/judgment_run.py'),'--candidates',str(paths[0]),
        '--answers',str(paths[1]),'--limit','1'],capture_output=True,text=True)
    assert preview.returncode==0,preview.stderr
    assert json.loads(preview.stdout)['status']=='dry-run' and not paths[1].exists()
    calls=[];monkeypatch.setattr(classify,'read_api_key',lambda:'test-secret')
    monkeypatch.setattr(classify,'_post',transport(calls,paths[2]))
    assert run(paths)['answered']==1 and len(calls)==1
