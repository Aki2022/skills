"""Local integration checks cannot claim shipment or bypass the real rendering gate."""
import copy
import json
import os
import subprocess
import sys

import pytest
from test_core import SKILL, project, approve, write
from test_ship_contract import channel
from test_ship_integration import browser_env, browser_project, commit, run_git, BROWSER
from ship_contract import validate_integration_gate


def local_gate():
    row=channel()
    return {key:value for key,value in row.items() if key in ('media','merge_base','render_servers')}


def test_local_contract_is_explicit_and_cannot_include_shipment_fields():
    assert validate_integration_gate({}) is None
    row=local_gate()
    assert validate_integration_gate({'integration_gate':row})==row
    for key in ['media','merge_base','render_servers']:
        broken=copy.deepcopy(row);broken.pop(key)
        with pytest.raises(ValueError):validate_integration_gate({'integration_gate':broken})
    for key,value in [('ship_command',['true']),('artifacts',['dist']),('ship',True),('prepare_command','npm ci')]:
        with pytest.raises(ValueError):validate_integration_gate({'integration_gate':{**row,key:value}})


def test_no_declared_local_gate_cannot_unlock_pre_push(project,tmp_path):
    root,catalogs=project
    env=browser_env()
    run_git(root,env,'config','user.name','Fixture Owner');run_git(root,env,'config','user.email','fixture@example.org')
    (root/'.gitignore').write_text('design/.resolved/\n')
    approve(project);commit(root,env,'fixture committed source')
    sha=run_git(root,env,'rev-parse','HEAD').stdout.strip()
    refs=tmp_path/'refs';refs.write_text('refs/heads/fixture '+sha+' refs/heads/fixture '+'0'*40+'\n')
    run=subprocess.run([str(SKILL/'scripts/design-ship'),'--repo',str(root),'--catalog-dir',str(catalogs),
        '--pre-push','--refs-file',str(refs)],cwd=root,env=env,capture_output=True,text=True)
    assert run.returncode==2,run.stdout+run.stderr


@pytest.mark.skipif(not BROWSER,reason='explicit real browser run required')
def test_real_local_gate_binds_predicted_tree_and_rejects_source_endpoint_and_prepare_mutations(project,tmp_path):
    env=browser_env();root=browser_project(project,env);_,catalogs=project
    run_git(root,env,'config','user.name','Fixture Owner');run_git(root,env,'config','user.email','fixture@example.org')
    config=json.loads((root/'design/design.yaml').read_text());config.pop('channels')
    config['integration_gate']={**local_gate(),'prepare_command':['python3','build.py']}
    write(root/'design/design.yaml',config)
    validation=json.loads((root/'design/validation.json').read_text())
    validation['urls']=[{'env':'DESIGN_WEB_URL','profile':'product-ui'}]
    write(root/'design/validation.json',validation)
    approve(project);commit(root,env,'fixture local integration')
    source=run_git(root,env,'rev-parse','HEAD').stdout.strip()
    run_git(root,env,'update-ref','refs/remotes/origin/main',source)
    argv=[str(SKILL/'scripts/design-ship'),'--repo',str(root),'--catalog-dir',str(catalogs)]
    def call(*extra):return subprocess.run(argv+list(extra),cwd=root,env=env,capture_output=True,text=True)
    # Declaring a local gate does not authorize or substitute for formal shipment.
    formal=call();assert formal.returncode==2 and 'no ship:true' in formal.stdout
    env['DESIGN_WEB_URL']='http://127.0.0.1:9/'
    direct=call('--integration');assert direct.returncode==0,direct.stdout+direct.stderr
    receipt=json.loads(direct.stdout)
    assert receipt['scope']=='local-integration' and receipt['production_protected'] is False
    assert receipt['deployment']=='not-applicable' and 'artifacts' not in receipt
    assert receipt['commit']==source and receipt['tree']==run_git(root,env,'rev-parse','HEAD^{tree}').stdout.strip()
    assert len(receipt['render_evidence'])==1 and not (root/'.shipment.json').exists()
    refs=tmp_path/'refs'
    def push():
        sha=run_git(root,env,'rev-parse','HEAD').stdout.strip()
        refs.write_text('refs/heads/fixture '+sha+' refs/heads/fixture '+'0'*40+'\n')
        return call('--pre-push','--refs-file',str(refs))
    pushed=push();assert pushed.returncode==0,pushed.stdout+pushed.stderr
    predicted=json.loads(pushed.stdout)['checks'][0]
    assert predicted['scope']=='local-integration' and predicted['predicted_merge'] is True
    assert predicted['source_commit']==source and predicted['production_protected'] is False
    # Actual L1 must reject a committed invalid source in the predicted tree.
    html=root/'index.html';normal=html.read_text();html.write_text(normal.replace(' lang="en"',''))
    commit(root,env,'fixture broken source');rejected=push()
    assert rejected.returncode==2 and 'L0/L1/L2' in rejected.stdout,rejected.stdout+rejected.stderr
    html.write_text(normal)
    # A missing endpoint is rejected even if an inherited URL points at another app.
    config['integration_gate']['render_servers']['OTHER_URL']=config['integration_gate']['render_servers'].pop('DESIGN_WEB_URL')
    write(root/'design/design.yaml',config);commit(root,env,'fixture missing endpoint')
    env['DESIGN_WEB_URL']='http://127.0.0.1:9/'
    rejected=call('--integration');assert rejected.returncode==2 and 'coverage' in rejected.stdout,rejected.stdout
    config['integration_gate']['render_servers']=local_gate()['render_servers']
    config['integration_gate']['prepare_command']=['python3','-c','from pathlib import Path;Path("index.html").write_text("changed")']
    write(root/'design/design.yaml',config);commit(root,env,'fixture source mutating preparation')
    head=run_git(root,env,'rev-parse','HEAD').stdout.strip();rejected=push()
    assert rejected.returncode==2 and 'clean committed source' in rejected.stdout,rejected.stdout
    assert run_git(root,env,'rev-parse','HEAD').stdout.strip()==head
    assert len(run_git(root,env,'worktree','list','--porcelain').stdout.split('worktree '))==2


def test_predicted_tree_cannot_add_an_unchecked_formal_channel(project,tmp_path):
    root,catalogs=project;env=browser_env()
    run_git(root,env,'config','user.name','Fixture Owner');run_git(root,env,'config','user.email','fixture@example.org')
    (root/'.gitignore').write_text('design/.resolved/\n')
    config=json.loads((root/'design/design.yaml').read_text());config['integration_gate']=local_gate()
    write(root/'design/design.yaml',config);commit(root,env,'fixture common local gate')
    base=run_git(root,env,'rev-parse','HEAD').stdout.strip()
    config['channels']={'web':channel()};write(root/'design/design.yaml',config)
    commit(root,env,'fixture upstream formal channel')
    run_git(root,env,'update-ref','refs/remotes/origin/main','HEAD')
    run_git(root,env,'checkout','-b','fixture-local',base)
    refs=tmp_path/'refs';refs.write_text('refs/heads/fixture-local '+base+' refs/heads/fixture-local '+'0'*40+'\n')
    run=subprocess.run([str(SKILL/'scripts/design-ship'),'--repo',str(root),'--catalog-dir',str(catalogs),
        '--pre-push','--refs-file',str(refs)],cwd=root,env=env,capture_output=True,text=True)
    assert run.returncode==2 and 'changed gate coverage' in run.stdout,run.stdout+run.stderr
    assert run_git(root,env,'rev-parse','HEAD').stdout.strip()==base


def test_bound_server_does_not_inherit_another_checkout_url(tmp_path):
    from ship import render_servers
    env={**os.environ,'DESIGN_WEB_URL':'http://127.0.0.1:9/'}
    script=('import os,sys,http.server;assert "DESIGN_WEB_URL" not in os.environ;'
            'http.server.ThreadingHTTPServer(("127.0.0.1",int(sys.argv[1])),http.server.SimpleHTTPRequestHandler).serve_forever()')
    servers={'DESIGN_WEB_URL':{'command':[sys.executable,'-c',script,'{port}'],'path':'/'}}
    with render_servers(tmp_path,servers,env) as (bound,urls,processes):
        assert bound['DESIGN_WEB_URL']==urls['DESIGN_WEB_URL']
        assert bound['DESIGN_WEB_URL']!=env['DESIGN_WEB_URL']
        assert processes[0].poll() is None
    assert processes[0].poll() is not None
