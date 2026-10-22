"""Real synthetic SSH signatures; OS protection is tested separately from unit trust."""
import copy
import json
from pathlib import Path
import subprocess
import sys
import pytest
from test_core import project,check,approve,approve_rows,write,SKILL
from approvals import protected_store,verify_signature,repository_id


def test_real_signature_rejects_payload_repo_namespace_and_unknown_key(project,tmp_path):
    root,_=project
    approve(project)
    row=json.loads((root/'design/approvals.yaml').read_text())[0]
    store=root.parent/'fixture-signers'
    verify_signature(repository_id(root),row,store)
    for field,value in [('date','2026-10-21'),('key','ia:web:token:color.primary'),('value_hash','0'*64)]:
        changed={**row,field:value}
        with pytest.raises(ValueError):verify_signature(repository_id(root),changed,store)
    with pytest.raises(ValueError):verify_signature('other-repository',row,store)
    key=tmp_path/'unknown'
    subprocess.run(['ssh-keygen','-q','-t','ed25519','-N','','-C','fixture','-f',str(key)],check=True)
    store.write_text('fixture-owner '+key.with_suffix('.pub').read_text())
    with pytest.raises(ValueError):verify_signature(repository_id(root),row,store)
    assert check(project,ship=True)['exit_code']==2


def test_os_trust_boundary_rejects_same_account_even_readonly(project):
    root,_=project
    store=root.parent/'fixture-signers'
    store.chmod(0o444)
    try:
        with pytest.raises(ValueError,match='agent account'):protected_store(store,root)
        with pytest.raises(ValueError):protected_store(root/'design/approvals.yaml',root)
        with pytest.raises(ValueError):protected_store(None,root)
    finally:store.chmod(0o644)


def test_alias_target_change_and_repo_replay_invalidate_signed_rows(project):
    root,_=project
    p=root/'design/tokens/brand.tokens.json';tokens=json.loads(p.read_text())
    tokens['color']['alias']={'$value':'{color.primary}','$extensions':{'org.own.design':{'origin':'proposal'}}}
    write(p,tokens)
    result=check(project)
    approve_rows(project,[{'key':'ds:web:token:color.alias','value_hash':result['manifest']['web']['tokens']['color.alias']['value_hash']}])
    assert check(project)['exit_code']==1
    tokens['color']['primary']['$value']={'colorSpace':'srgb','components':[1,1,1],'hex':'#ffffff'}
    write(p,tokens)
    assert check(project)['exit_code']==2
    subprocess.run(['git','-C',str(root),'remote','set-url','origin','https://example.org/other.git'],check=True)
    assert 'signature' in str(check(project)['errors'])


def test_prepare_cli_and_owner_signing_detect_review_sign_race(project,monkeypatch):
    import owner_approval
    from approvals import payload,namespace
    root,catalogs=project
    original=(root/'design/approvals.yaml').read_bytes()
    run=subprocess.run([sys.executable,str(SKILL/'scripts/design_author.py'),'approve','--prepare','--must',
        '--repo',str(root),'--catalog-dir',str(catalogs)],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    request=json.loads(run.stdout)
    assert len(request['rows'])==1
    assert (root/'design/approvals.yaml').read_bytes()==original
    def fixture_sign(content,key,scope):
        # Synthetic private key only; the actual owner agent/OS boundary is not claimed here.
        return subprocess.run(['ssh-keygen','-Y','sign','-f',str(root.parent/'fixture-key'),'-n',scope],
            input=content,capture_output=True,check=True).stdout.decode()
    monkeypatch.setattr(owner_approval,'_sign',fixture_sign)
    public=root.parent/'fixture-key.pub'
    result=owner_approval.approve_request(root,catalogs,request,public,'fixture-owner')
    assert result['exit_code']==0
    assert check(project,ship=True)['exit_code']==0
    write(root/'design/approvals.yaml',[])
    def raced_sign(content,key,scope):
        signature=fixture_sign(content,key,scope)
        data=json.loads((root/'design/tokens/brand.tokens.json').read_text())
        data['color']['primary']['$value']={'colorSpace':'srgb','components':[1,1,1],'hex':'#ffffff'}
        write(root/'design/tokens/brand.tokens.json',data)
        return signature
    monkeypatch.setattr(owner_approval,'_sign',raced_sign)
    with pytest.raises(ValueError,match='changed during signing'):
        owner_approval.approve_request(root,catalogs,request,public,'fixture-owner')
    assert json.loads((root/'design/approvals.yaml').read_text())==[]


def test_signed_manifest_is_invalidated_by_trust_rotation(project):
    from resolved_contract import load_manifest
    root,_=project
    approve(project);assert check(project,ship=True)['exit_code']==0
    manifest=root/'design/.resolved/web.json'
    load_manifest(manifest)
    (root.parent/'fixture-signers').write_text('# owner trust rotated\n')
    with pytest.raises(ValueError):load_manifest(manifest)


def test_consumers_reject_legacy_manifest_scheme(project):
    from resolved_contract import load_manifest
    root,_=project;check(project)
    p=root/'design/.resolved/web.json';data=json.loads(p.read_text());data.pop('approval_scheme');write(p,data)
    with pytest.raises(ValueError):load_manifest(p)


def test_javascript_consumer_enforces_external_trust(project):
    import os
    root,_=project;approve(project);check(project,ship=True)
    module=(SKILL/'scripts/resolved_contract.mjs').as_uri()
    script='import {loadManifest} from '+json.dumps(module)+'; loadManifest(process.argv.at(-1));'
    run=subprocess.run([os.environ.get('DESIGN_TEST_NODE','node'),'--input-type=module','-e',script,
        str(root/'design/.resolved/web.json')],capture_output=True,text=True,
        env=dict(os.environ,DESIGN_AUTHOR_PYTHON=sys.executable))
    assert run.returncode!=0
