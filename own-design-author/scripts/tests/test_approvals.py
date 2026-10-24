"""Hash-bound human approval records, exercised through real CLI/consumers."""
import json
import os
import subprocess
import sys
import pytest
from test_core import project,check,approve,approve_rows,write,SKILL
from resolved_contract import load_manifest


def cli(root,catalogs,*args):
    return subprocess.run([sys.executable,str(SKILL/'scripts/design_author.py'),'approve',
        '--repo',str(root),'--catalog-dir',str(catalogs),*args],capture_output=True,text=True)


def test_unsigned_approval_needs_no_key_agent_or_trust(project,monkeypatch):
    for name in ['SSH_AUTH_SOCK','DESIGN_ALLOWED_SIGNERS']:
        monkeypatch.delenv(name,raising=False)
    approve(project)
    assert check(project,ship=True)['exit_code']==0
    data=load_manifest(project[0]/'design/.resolved/web.json')
    assert data['approval_scheme']=='hash-jcs-v1'
    assert 'approval_context' not in data


@pytest.mark.parametrize('change',[
    {'approver':''},{'date':'yesterday'},{'value_hash':'bad'},
    {'key':'unknown'},{'signature':'not supported'}])
def test_invalid_approval_metadata_is_rejected(project,change):
    approve(project)
    p=project[0]/'design/approvals.yaml';rows=json.loads(p.read_text())
    rows[0].update(change);write(p,rows)
    assert check(project)['exit_code']==2


def test_alias_target_change_invalidates_approval(project):
    root,_=project
    p=root/'design/tokens/brand.tokens.json';tokens=json.loads(p.read_text())
    tokens['color']['alias']={'$value':'{color.primary}','$extensions':{'org.own.design':{'origin':'proposal'}}}
    write(p,tokens)
    value=check(project)['manifest']['web']['tokens']['color.alias']['value_hash']
    approve_rows(project,[{'key':'ds:web:token:color.alias','value_hash':value}])
    assert check(project)['exit_code']==1
    tokens['color']['primary']['$value']={'colorSpace':'srgb','components':[1,1,1],'hex':'#ffffff'}
    write(p,tokens)
    assert check(project)['exit_code']==2


def test_prepare_and_record_cli_keeps_review_boundary(project,tmp_path):
    root,catalogs=project
    before=(root/'design/approvals.yaml').read_bytes()
    prepared=cli(root,catalogs,'--prepare','--must')
    assert prepared.returncode==0,prepared.stdout+prepared.stderr
    request=json.loads(prepared.stdout)
    assert request['approval_scheme']=='hash-jcs-v1' and len(request['rows'])==1
    assert (root/'design/approvals.yaml').read_bytes()==before
    path=tmp_path/'review.json';write(path,request)
    recorded=cli(root,catalogs,'--request',str(path),'--approver','fixture-owner')
    assert recorded.returncode==0,recorded.stdout+recorded.stderr
    assert json.loads(recorded.stdout)['recorded_rows']==1
    assert check(project,ship=True)['exit_code']==0
    rows=json.loads((root/'design/approvals.yaml').read_text())
    assert set(rows[0])=={'key','value_hash','approver','date'}


def test_reviewed_change_and_wrong_repo_are_rejected_before_write(project,tmp_path):
    from owner_approval import prepare,approve_request
    root,catalogs=project
    request=prepare(root,catalogs,all_must=True)
    with pytest.raises(ValueError,match='another repository'):
        approve_request(root,catalogs,{**request,'repo_id':'wrong'},'fixture-owner')
    data=json.loads((root/'design/tokens/brand.tokens.json').read_text())
    data['color']['primary']['$value']={'colorSpace':'srgb','components':[1,1,1],'hex':'#ffffff'}
    write(root/'design/tokens/brand.tokens.json',data)
    with pytest.raises(ValueError,match='reviewed inputs changed'):
        approve_request(root,catalogs,request,'fixture-owner')
    assert json.loads((root/'design/approvals.yaml').read_text())==[]


def test_recording_race_preserves_ledger(project,monkeypatch):
    import owner_approval
    root,catalogs=project
    request=owner_approval.prepare(root,catalogs,all_must=True)
    snapshot=owner_approval._snapshot
    calls=[]
    def changed(*args):
        calls.append(1)
        if len(calls)==2:
            p=root/'design/tokens/brand.tokens.json';data=json.loads(p.read_text())
            data['color']['primary']['$value']={'colorSpace':'srgb','components':[1,1,1],'hex':'#ffffff'}
            write(p,data)
        return snapshot(*args)
    monkeypatch.setattr(owner_approval,'_snapshot',changed)
    with pytest.raises(ValueError,match='changed during recording'):
        owner_approval.approve_request(root,catalogs,request,'fixture-owner')
    assert json.loads((root/'design/approvals.yaml').read_text())==[]


@pytest.mark.parametrize('scheme',[None,'ssh-jcs-v1'])
def test_consumers_reject_previous_manifest_schemes(project,scheme):
    root,_=project;check(project)
    p=root/'design/.resolved/web.json';data=json.loads(p.read_text())
    if scheme is None:data.pop('approval_scheme')
    else:data['approval_scheme']=scheme
    write(p,data)
    with pytest.raises(ValueError):load_manifest(p)


def test_javascript_consumer_reads_approved_manifest_and_rejects_stale_ledger(project):
    root,_=project;approve(project);assert check(project,ship=True)['exit_code']==0
    module=(SKILL/'scripts/resolved_contract.mjs').as_uri()
    script='import {loadManifest} from '+json.dumps(module)+'; loadManifest(process.argv.at(-1));'
    args=[os.environ.get('DESIGN_TEST_NODE','node'),'--input-type=module','-e',script,str(root/'design/.resolved/web.json')]
    assert subprocess.run(args,capture_output=True).returncode==0
    write(root/'design/approvals.yaml',[])
    assert subprocess.run(args,capture_output=True).returncode!=0
    with pytest.raises(ValueError,match='stale manifest'):load_manifest(root/'design/.resolved/web.json')
