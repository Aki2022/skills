"""Actual Chromium, git push, scanner and mise paths; no mock is ship evidence."""
import hashlib
import json
import os
from pathlib import Path
import shlex
import shutil
import subprocess
import sys

import pytest
from test_core import SKILL, project, approve, check, write
from test_fixed_render import pin
from test_ship_contract import channel

BROWSER=os.environ.get('DESIGN_BROWSER_TEST')=='1'


def run_git(root,env,*args,allowed=(0,)):
    run=subprocess.run(['git',*args],cwd=root,env=env,capture_output=True,text=True)
    assert run.returncode in allowed,run.stdout+run.stderr
    return run


def browser_project(project,env):
    root,_=project
    run_git(root,env,'checkout','-b','fixture-design')
    tokens={'color':{'$type':'color'},'font':{'$type':'fontFamily','$value':['Inter','sans-serif'],
        '$extensions':{'org.own.design':{'origin':'proposal'}}},
        'radius':{'$type':'dimension','$value':{'value':0,'unit':'px'},'$extensions':{'org.own.design':{'origin':'proposal'}}}}
    for name,components,alpha in [('primary',[0,0,0],1),('white',[1,1,1],1),('clear',[0,0,0],0)]:
        tokens['color'][name]={'$value':{'colorSpace':'srgb','components':components,'hex':'#ffffff' if name=='white' else '#000000','alpha':alpha},
            '$extensions':{'org.own.design':{'origin':'proposal'}}}
    write(root/'design/tokens/brand.tokens.json',tokens)
    write(root/'design/render-environment.json',pin(root,env))
    (root/'index.html').write_text('<html lang="en"><head><title>Fixture</title><meta name="viewport" content="width=device-width,initial-scale=1"><style>'
        '@font-face{font-family:Inter;src:url(/fixture.woff2)}'
        'body{font:16px/1.5 Inter,sans-serif;background:white;color:black}'
        '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
        ':focus-visible{outline:2px solid black}@media(min-width:375px){body{max-width:100%}}'
        '</style></head><body><main><h1>Fixture</h1><p>Stable text.</p></main></body></html>')
    (root/'build.py').write_text('from pathlib import Path\nimport shutil\np=Path("dist");p.mkdir(exist_ok=True)\n'
        'shutil.copyfile("index.html",p/"index.html")\nshutil.copyfile("design/assets/fixture.woff2",p/"fixture.woff2")\n')
    (root/'deploy.py').write_text('import json,os\nfrom pathlib import Path\n'
        'Path(".shipment.json").write_text(json.dumps({k:os.environ[k] for k in '
        '["DESIGN_SHIP_COMMIT","DESIGN_SHIP_TREE","DESIGN_SHIP_ARTIFACTS_JSON"]}))\n')
    (root/'.gitignore').write_text('design/.resolved/\ndesign/assets.json\ndist/\n.shipment.json\n')
    cfg=json.loads((root/'design/design.yaml').read_text());row=channel()
    row['build_command']=['python3','build.py'];row['ship_command']=['python3','deploy.py']
    cfg['channels']={'web':row};write(root/'design/design.yaml',cfg)
    write(root/'design/validation.json',{'schema_version':1,'html':['index.html'],
        'urls':[{'env':'DESIGN_WEB_URL'}],'viewports':[375],'schemes':['light']})
    return root


def browser_env():
    env=dict(os.environ,DESIGN_AUTHOR_PYTHON=sys.executable)
    env['PATH']=str(Path(sys.executable).parent)+os.pathsep+str(SKILL/'scripts')+os.pathsep+env['PATH']
    if env.get('DESIGN_TEST_NODE_MODULES'):env['NODE_PATH']=env['DESIGN_TEST_NODE_MODULES']
    if env.get('DESIGN_TEST_NODE'):env['DESIGN_AUDIT_NODE']=env['DESIGN_TEST_NODE']
    return env


def commit(root,env,message):
    run_git(root,env,'add','.')
    run_git(root,env,'commit','-qm',message)


@pytest.mark.skipif(not BROWSER,reason='explicit real browser run required')
def test_real_ship_channel_binds_commit_artifacts_and_stops_l1_l2(project):
    env=browser_env();root=browser_project(project,env);_,catalogs=project
    run_git(root,env,'config','user.name','Fixture Owner');run_git(root,env,'config','user.email','fixture@example.org')
    approve(project);commit(root,env,'fixture approved web')
    def ship(*extra):
        return subprocess.run([str(SKILL/'scripts/design-ship'),'--repo',str(root),'--catalog-dir',str(catalogs),*extra],
            cwd=root,env=env,capture_output=True,text=True)
    run=ship();assert run.returncode==0,run.stdout+run.stderr
    result=json.loads(run.stdout)['checks'][0]
    assert result['commit']==run_git(root,env,'rev-parse','HEAD').stdout.strip()
    assert set(result['artifacts'])=={'dist/index.html','dist/fixture.woff2'}
    assert len(result['render_evidence'])==1 and not (root/'.shipment.json').exists()
    run=ship('--channel','web');assert run.returncode==0,run.stdout+run.stderr
    deployed=json.loads((root/'.shipment.json').read_text())
    assert deployed['DESIGN_SHIP_COMMIT']==result['commit']
    assert json.loads(deployed['DESIGN_SHIP_ARTIFACTS_JSON'])==result['artifacts']
    # Broken source is committed, so the failure comes from actual L1, not dirty-tree rejection.
    html=root/'index.html';original=html.read_text();html.write_text(original.replace(' lang="en"',''))
    commit(root,env,'fixture invalid source');run=ship('--channel','web')
    assert run.returncode==2 and 'L0/L1/L2' in run.stdout,run.stdout
    assert json.loads((root/'.shipment.json').read_text())==deployed
    # Valid source with undeclared fallback reaches real L2 and fails there.
    html.write_text(original.replace('font:16px/1.5 Inter,sans-serif','font:16px/1.5 serif'))
    commit(root,env,'fixture invalid rendered font');run=ship('--channel','web')
    assert run.returncode==2,run.stdout+run.stderr
    assert json.loads((root/'.shipment.json').read_text())==deployed


@pytest.mark.skipif(not BROWSER or not os.environ.get('DESIGN_TEST_VIBE_ROOT'),reason='real browser and vibe-guard fixture run required')
def test_push_and_mise_share_ship_gate_without_bypassing_scanner(project,tmp_path):
    env=browser_env();root=browser_project(project,env);_,catalogs=project
    guard=Path(os.environ['DESIGN_TEST_VIBE_ROOT']).resolve()
    global_config=Path(os.environ.get('GIT_CONFIG_GLOBAL',str(Path.home()/'.gitconfig')))
    before=hashlib.sha256(global_config.read_bytes()).hexdigest()
    harness=tmp_path/'guard';(harness/'bin').mkdir(parents=True)
    shutil.copytree(guard/'git-hooks',harness/'git-hooks')
    hook=harness/'git-hooks/pre-push';lines=hook.read_text().splitlines()
    assert any(line.startswith('ROOT=') for line in lines)
    hook.write_text('\n'.join('ROOT='+shlex.quote(str(harness)) if line.startswith('ROOT=') else line for line in lines)+'\n')
    scanner_marker=tmp_path/'scanner-called';wrapper=harness/'bin/vibe-guard'
    wrapper.write_text('#!/bin/sh\nset -eu\nprintf called >> '+shlex.quote(str(scanner_marker))+
        '\nexec '+shlex.quote(str(guard/'bin/vibe-guard'))+' "$@"\n');wrapper.chmod(0o755)
    extension=harness/'git-hooks/pre-push.d/60-design-ship'
    shutil.copy2(SKILL/'hooks/pre-push.d/60-design-ship',extension)
    config=tmp_path/'gitconfig'
    config.write_text('[include]\n path = '+str(global_config)+'\n[core]\n hooksPath = '+str(harness/'git-hooks')+'\n')
    # Only fixture global config is changed; original scanners/extensions remain installed.
    env.update(GIT_CONFIG_GLOBAL=str(config),VIBE_GUARD_ROOT=str(guard))
    fixture_ship=tmp_path/'fixture-ship'
    fixture_ship.write_text('#!'+sys.executable+'\nimport subprocess,sys\nraise SystemExit(subprocess.call('+
        repr([str(SKILL/'scripts/design-ship')])+'+sys.argv[1:]+'+repr(['--catalog-dir',str(catalogs)])+'))\n')
    fixture_ship.chmod(0o755);env['DESIGN_SHIP_COMMAND']=str(fixture_ship)
    run_git(root,env,'config','user.name','Fixture Owner');run_git(root,env,'config','user.email','fixture@example.org')
    (root/'.vibe-guard').mkdir()
    shutil.copy2(SKILL/'templates/design-ship.enabled',root/'.vibe-guard/design-ship.enabled')
    shutil.copy2(SKILL/'templates/mise.design.toml',root/'.mise.toml')
    commit(root,env,'fixture proposed web')
    run_git(root,env,'update-ref','refs/remotes/origin/main','HEAD')
    remote=tmp_path/'remote.git';subprocess.run(['git','init','--bare','-q',str(remote)],check=True,env=env)
    run_git(root,env,'remote','add','fixture',str(remote))
    blocked=run_git(root,env,'push','fixture','HEAD:refs/heads/main',allowed=(1,))
    assert 'ship gate' in blocked.stdout+blocked.stderr or 'predicted' in blocked.stdout+blocked.stderr
    assert scanner_marker.exists() and not (remote/'refs/heads/main').exists()
    mise=os.environ.get('DESIGN_TEST_MISE')
    assert mise,'real mise executable required for this acceptance scenario'
    mise_env=dict(env,MISE_TRUSTED_CONFIG_PATHS=str(root),DESIGN_AUTHOR_COMMAND=str(SKILL/'scripts/design-author'))
    # mise's exact template invocation uses catalog path supplied through a single CLI wrapper.
    catalog_l0=tmp_path/'fixture-l0'
    catalog_l0.write_text('#!'+sys.executable+'\nimport subprocess,sys\nraise SystemExit(subprocess.call('+
        repr([str(SKILL/'scripts/design-author')])+'+sys.argv[1:]+'+repr(['--catalog-dir',str(catalogs)])+'))\n')
    catalog_l0.chmod(0o755);mise_env['DESIGN_AUTHOR_COMMAND']=str(catalog_l0)
    catalog_web=tmp_path/'fixture-web'
    catalog_web.write_text('#!'+sys.executable+'\nimport subprocess,sys\nraise SystemExit(subprocess.call('+
        repr([str(SKILL/'scripts/design-check')])+'+sys.argv[1:]+'+repr(['--catalog-dir',str(catalogs)])+'))\n')
    catalog_web.chmod(0o755);mise_env['DESIGN_CHECK_COMMAND']=str(catalog_web)
    run=subprocess.run([mise,'run','design:ship'],cwd=root,env=mise_env,capture_output=True,text=True)
    assert run.returncode!=0
    approve(project);commit(root,env,'fixture reviewed web values')
    run_git(root,env,'push','fixture','HEAD:refs/heads/main')
    assert (remote/'refs/heads/main').read_text().strip()==run_git(root,env,'rev-parse','HEAD').stdout.strip()
    run=subprocess.run([mise,'run','design:ship'],cwd=root,env=mise_env,capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    assert scanner_marker.read_text().count('called')>=2
    assert hashlib.sha256(global_config.read_bytes()).hexdigest()==before
    assert not run_git(root,env,'config','--local','--get','core.hooksPath',allowed=(1,)).stdout
    doctor=subprocess.run([str(SKILL/'scripts/design-ship'),'--repo',str(root),'--doctor'],env=env,capture_output=True,text=True)
    assert doctor.returncode==0,doctor.stdout
    extension.unlink()
    doctor=subprocess.run([str(SKILL/'scripts/design-ship'),'--repo',str(root),'--doctor'],env=env,capture_output=True,text=True)
    assert doctor.returncode==2


@pytest.mark.skipif(not BROWSER, reason='explicit real browser run required')
def test_two_green_branches_fail_on_the_predicted_merge_tree(project, tmp_path):
    """One branch changes a nice value; the other approves its old value."""
    env=browser_env();root=browser_project(project,env);_,catalogs=project
    run_git(root,env,'config','user.name','Fixture Owner')
    run_git(root,env,'config','user.email','fixture@example.org')
    token_file=root/'design/tokens/brand.tokens.json'
    data=json.loads(token_file.read_text())
    data['color']['optional']={'$value':{'colorSpace':'srgb','components':[0,0,0],'hex':'#000000'},
        '$extensions':{'org.own.design':{'origin':'proposal'}}}
    write(token_file,data);approve(project);commit(root,env,'fixture common base')
    base=run_git(root,env,'rev-parse','HEAD').stdout.strip()
    old_hash=check(project)['manifest']['web']['tokens']['color.optional']['value_hash']
    run_git(root,env,'checkout','-b','fixture-token')
    data['color']['optional']['$value'].update(components=[1,1,1],hex='#ffffff')
    write(token_file,data);commit(root,env,'fixture nice value changed')
    token_sha=run_git(root,env,'rev-parse','HEAD').stdout.strip()
    argv=[str(SKILL/'scripts/design-ship'),'--repo',str(root),'--catalog-dir',str(catalogs)]
    first=subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True)
    assert first.returncode==0,first.stdout+first.stderr
    run_git(root,env,'checkout','-b','fixture-approval',base)
    ledger=json.loads((root/'design/approvals.yaml').read_text())
    ledger.append({'key':'ds:web:token:color.optional','value_hash':old_hash,
                   'approver':'fixture-owner','date':'2026-10-23'})
    write(root/'design/approvals.yaml',ledger);commit(root,env,'fixture approve old nice value')
    second=subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True)
    assert second.returncode==0,second.stdout+second.stderr
    run_git(root,env,'update-ref','refs/remotes/origin/main','HEAD')
    run_git(root,env,'checkout','fixture-token')
    refs=tmp_path/'push-refs'
    refs.write_text('refs/heads/fixture-token '+token_sha+' refs/heads/fixture-token '+'0'*40+'\n')
    merged=subprocess.run(argv+['--pre-push','--refs-file',str(refs)],cwd=root,env=env,capture_output=True,text=True)
    assert merged.returncode==2,merged.stdout+merged.stderr
    result=json.loads(merged.stdout)
    assert 'predicted merge ship gate failed' in result['errors']
    assert any('approval hash mismatch' in error for error in result['checks'][0]['errors'])
    assert run_git(root,env,'rev-parse','HEAD').stdout.strip()==token_sha
    assert len(run_git(root,env,'worktree','list','--porcelain').stdout.split('worktree '))==2
