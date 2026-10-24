"""Ship contract rejects partial checks, unbound endpoints and changing artifacts."""
import copy
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

import pytest
from test_core import project, approve, write, SKILL
from ship_contract import validate_channels, artifact_hashes, validate_web_result


def channel():
    return {'ship':True,'media':'web','ship_command':['true'], 'artifacts':['dist'],
            'merge_base':'origin/main','render_servers':{
                'DESIGN_WEB_URL':{'command':['python3','-m','http.server','{port}','--bind','127.0.0.1','--directory','dist'],'path':'/'}}}


def test_channels_require_explicit_command_and_bound_server():
    row=channel()
    assert validate_channels({'channels':{'web':row}})['web']==row
    for key in ['ship_command','render_servers','artifacts','merge_base']:
        broken=copy.deepcopy(row);broken.pop(key)
        with pytest.raises(ValueError):validate_channels({'channels':{'web':broken}})
    for bad in [[],{},[''], 'deploy --prod']:
        broken=copy.deepcopy(row);broken['ship_command']=bad
        with pytest.raises(ValueError):validate_channels({'channels':{'web':broken}})


def test_channel_rejects_escape_and_inherited_url():
    for path in ['/outside','../outside','dist/../private']:
        row=channel();row['artifacts']=[path]
        with pytest.raises(ValueError):validate_channels({'channels':{'web':row}})
    row=channel();row['render_servers']['DESIGN_WEB_URL']['path']='https://example.org/'
    with pytest.raises(ValueError):validate_channels({'channels':{'web':row}})
    row=channel();row['render_servers']['DESIGN_WEB_URL']['command']=['already-running']
    with pytest.raises(ValueError):validate_channels({'channels':{'web':row}})


def test_artifacts_nonempty_exact_and_no_symlinks(tmp_path):
    (tmp_path/'dist').mkdir();(tmp_path/'dist/index.html').write_text('first')
    assert artifact_hashes(tmp_path,['dist'])=={'dist/index.html':hashlib.sha256(b'first').hexdigest()}
    (tmp_path/'dist/alias').symlink_to(tmp_path/'dist/index.html')
    with pytest.raises(ValueError):artifact_hashes(tmp_path,['dist'])
    with pytest.raises(ValueError):artifact_hashes(tmp_path,['missing'])


def test_artifact_root_alias_is_canonicalized(tmp_path):
    root=tmp_path/'real';root.mkdir();(root/'output').write_text('built artifact')
    alias=tmp_path/'alias';alias.symlink_to(root,target_is_directory=True)
    assert artifact_hashes(alias,['output'])==artifact_hashes(root,['output'])


def result_fixture(tmp_path):
    write(tmp_path/'design/validation.json',{'schema_version':1,'html':['dist/index.html'],
        'urls':[{'env':'DESIGN_WEB_URL'}], 'viewports':[375],'schemes':['light']})
    env={'runtime':{'browser':'fixture'},'viewports':[{'width':375,'height':812}], 'schemes':['light']}
    write(tmp_path/'design/render-environment.json',env)
    digest=hashlib.sha256((tmp_path/'design/render-environment.json').read_bytes()).hexdigest()
    row={'scheme':'light','width':375,'render_evidence':{'screenshot_sha256':'a'*64}}
    write(tmp_path/'design/.resolved/l2-0.json',{'url':'http://127.0.0.1:12345/',
        'environment':{'runtime':env['runtime'],'config_sha256':digest,'repeat_equal':True},
        'runs':[{**row,'repeat':0},{**row,'repeat':1}], 'failures':[], 'uncovered':[]})
    result={'exit_code':0,'errors':[],'stages':[{'stage':'L0','exit_code':0,'counts':{'items':1}},
        {'stage':'L1','exit_code':0,'targets':1},{'stage':'L2','exit_code':0,'target_env':'DESIGN_WEB_URL'}]}
    return result


def test_full_check_requires_every_stage_and_exact_render_evidence(tmp_path):
    result=result_fixture(tmp_path)
    assert validate_web_result(tmp_path,result,{'DESIGN_WEB_URL':'http://127.0.0.1:12345/'})
    for stage in ['L0','L1','L2']:
        broken=copy.deepcopy(result);broken['stages']=[s for s in broken['stages'] if s['stage']!=stage]
        with pytest.raises(ValueError):validate_web_result(tmp_path,broken,{'DESIGN_WEB_URL':'http://127.0.0.1:12345/'})
    with pytest.raises(ValueError):validate_web_result(tmp_path,result,{'DESIGN_WEB_URL':'http://127.0.0.1:9999/'})
    evidence=tmp_path/'design/.resolved/l2-0.json';data=json.loads(evidence.read_text())
    data['runs'][1]['render_evidence']['screenshot_sha256']='b'*64;write(evidence,data)
    with pytest.raises(ValueError):validate_web_result(tmp_path,result,{'DESIGN_WEB_URL':'http://127.0.0.1:12345/'})


def test_ship_stops_l0_before_server_or_deploy(project):
    root,catalogs=project
    config=json.loads((root/'design/design.yaml').read_text());config['channels']={'web':channel()}
    config['channels']['web']['ship_command']=[sys.executable,'-c','open("DEPLOYED","w").write("bad")']
    write(root/'design/design.yaml',config)
    run=subprocess.run([str(SKILL/'scripts/design-ship'),'--repo',str(root),'--channel','web',
        '--catalog-dir',str(catalogs)],env=dict(os.environ,DESIGN_AUTHOR_PYTHON=sys.executable),capture_output=True,text=True)
    assert run.returncode==2
    assert not (root/'DEPLOYED').exists()
