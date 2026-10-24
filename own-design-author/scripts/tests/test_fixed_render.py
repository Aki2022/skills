"""Fixed rendering uses real Chromium and an existing caller-supplied font."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import pytest
from test_core import SKILL,project,check,write,approve

NODE=os.environ.get('DESIGN_TEST_NODE','node')
HELPER=SKILL/'scripts/fixed_render.mjs'


def pin(root,env):
    source=os.environ.get('DESIGN_TEST_FONT')
    if not source:pytest.skip('existing font fixture not supplied')
    import hashlib
    font=root/'design/assets/fixture.woff2';font.parent.mkdir(exist_ok=True)
    shutil.copyfile(source,font)
    js=('import {probe} from '+json.dumps(HELPER.as_uri())+';'
        'import {createRequire} from "node:module";const require=createRequire(process.cwd()+"/");'
        'const p=require("playwright");console.log(JSON.stringify(await probe(p.chromium,require("playwright/package.json").version)));')
    run=subprocess.run([NODE,'--input-type=module','-e',js],cwd=root,env=env,capture_output=True,text=True)
    assert run.returncode==0,run.stderr
    return {'schema_version':1,'runtime':json.loads(run.stdout),'clock':'2020-01-01T00:00:00.000Z',
        'locale':'en-US','timezone':'UTC','device_scale_factor':1,'reduced_motion':'reduce',
        'viewports':[{'width':375,'height':812}],'schemes':['light'],
        'fonts':[{'path':'design/assets/fixture.woff2','url_path':'/fixture.woff2','family':'Inter',
            'sha256':hashlib.sha256(font.read_bytes()).hexdigest()}]}


def test_fixed_module_rejects_invalid_envelope_without_launch(tmp_path):
    js=('import {loadEnvironment} from '+json.dumps(HELPER.as_uri())+';'
        'try {loadEnvironment(process.argv.at(-1),process.cwd(),[375],["light"]);process.exit(99)}catch{process.exit(0)}')
    path=tmp_path/'design/render-environment.json';write(path,{'schema_version':1})
    run=subprocess.run([NODE,'--input-type=module','-e',js,str(path)],cwd=tmp_path,capture_output=True,text=True)
    assert run.returncode==0,run.stderr


@pytest.mark.skipif(os.environ.get('DESIGN_BROWSER_TEST')!='1',reason='explicit real browser run required')
def test_fixed_two_runs_detect_fallback_instability_and_runtime_mismatch(project,tmp_path):
    from http.server import ThreadingHTTPServer,SimpleHTTPRequestHandler
    from functools import partial
    from threading import Thread
    root,_=project;env=dict(os.environ)
    if env.get('DESIGN_TEST_NODE_MODULES'):env['NODE_PATH']=env['DESIGN_TEST_NODE_MODULES']
    config=pin(root,env);write(root/'design/render-environment.json',config)
    tokens={'color':{'$type':'color'},'font':{'$type':'fontFamily','$value':['Inter','sans-serif'],
        '$extensions':{'org.own.design':{'origin':'proposal'}}},
        'radius':{'$type':'dimension','$value':{'value':0,'unit':'px'},'$extensions':{'org.own.design':{'origin':'proposal'}}}}
    for name,components,alpha in [('primary',[0,0,0],1),('white',[1,1,1],1),('clear',[0,0,0],0)]:
        tokens['color'][name]={'$value':{'colorSpace':'srgb','components':components,'hex':'#ffffff' if name=='white' else '#000000','alpha':alpha},
            '$extensions':{'org.own.design':{'origin':'proposal'}}}
    write(root/'design/tokens/brand.tokens.json',tokens)
    normal=('<html lang="en"><head><title>Fixture</title><style>'
        '@font-face{font-family:Inter;src:url(/fixture.woff2)}'
        'body{font:16px/1.5 Inter,sans-serif;background:white;color:black}'
        '*{border-color:black;outline-color:black;text-decoration-color:black;border-radius:0}'
        '</style></head><body><main><h1>Fixture</h1><p>Stable text.</p></main></body></html>')
    class Handler(SimpleHTTPRequestHandler):
        def log_message(self,*args):pass
    server=ThreadingHTTPServer(('127.0.0.1',0),partial(Handler,directory=str(root)))
    thread=Thread(target=server.serve_forever,daemon=True);thread.start()
    script=SKILL.parent/'own-website-audit/scripts/measure_lp.mjs'
    def measure(html):
        (root/'index.html').write_text(html);l0=check(project)
        assert l0['exit_code']!=2,l0['errors']
        output=root/'design/.resolved/fixed-test.json'
        run=subprocess.run([NODE,str(script),'http://127.0.0.1:'+str(server.server_port),
            '--manifest',str(root/'design/.resolved/web.json'),'--environment',str(root/'design/render-environment.json'),
            '--viewports','375','--schemes','light','--json',str(output)],cwd=root,env=env,capture_output=True,text=True)
        return run,json.loads(output.read_text()) if output.exists() else None
    try:
        run,data=measure(normal);assert run.returncode==0,run.stderr+run.stdout
        assert len(data['runs'])==2 and data['environment']['repeat_equal']
        assert data['environment']['runtime']==config['runtime']
        assert data['runs'][0]['render_evidence']['font_glyphs']>0
        assert data['runs'][0]['render_evidence']['font_sources'][0]['sha256s']==[config['fonts'][0]['sha256']]
        assert data['runs'][0]['render_evidence']['screenshot_sha256']==data['runs'][1]['render_evidence']['screenshot_sha256']
        assert data['runs'][0]['render_evidence']['clock']==config['clock']
        run,data=measure(normal.replace('font:16px/1.5 Inter,sans-serif','font:16px/1.5 serif'))
        assert run.returncode==1 and any(f['id']=='font-fallback' for f in data['failures']),run.stdout
        changing=normal.replace('</body>','<script>document.querySelector("p").textContent=crypto.randomUUID()</script></body>')
        run,data=measure(changing)
        assert run.returncode==1 and not data['environment']['repeat_equal'],run.stdout
        assert 'render-repeat-mismatch' in data['uncovered']
        import base64
        inline=normal.replace('url(/fixture.woff2)','url(data:font/woff2;base64,'+base64.b64encode((root/'design/assets/fixture.woff2').read_bytes()).decode()+')')
        inline=inline.replace('<style>','<link rel="preload" href="/fixture.woff2" as="font" crossorigin><style>')
        run,data=measure(inline)
        assert run.returncode==1 and any(f['id']=='font-fallback' for f in data['failures']),run.stdout
        assert any('unbound custom font' in s for s in data['runs'][0]['render_evidence']['fallbacks'])
        config['runtime']['browser_version']='mismatch';write(root/'design/render-environment.json',config)
        run,data=measure(normal);assert run.returncode==2,run.stdout+run.stderr
        assert data is None
        env['PLAYWRIGHT_BROWSERS_PATH']=str(tmp_path/'missing-browser-cache')
        run,data=measure(normal);assert run.returncode==2,run.stdout+run.stderr
        assert 'browser runtime unavailable' in run.stderr and data is None
    finally:
        server.shutdown();server.server_close();thread.join()
