#!/usr/bin/env python3
"""Bind L0/L1/L2 to clean source; shipment additionally requires exact artifacts."""
import argparse
from contextlib import contextmanager
import functools
import hashlib
import json
import os
from pathlib import Path
import shutil
import signal
import socket
import subprocess
import sys
import tempfile
import time
from urllib.error import HTTPError
from urllib.request import urlopen

import layer_paths
from design_author import git_root, read_data, CATALOG
from ship_contract import validate_channels, validate_integration_gate, artifact_hashes, validate_web_result, safe_path
from web_check import validation_endpoint
# Git, refs, bases, predicted merges and the dispatcher check are shared with the IA ship gate.
from ship_base import hook_free_env, git, read_push_refs, merge_bases, predicted_commit, dispatcher_propagates

SCRIPT=Path(__file__).resolve().parent


def clean_commit(root):
    sha=git(root,'rev-parse','--verify','HEAD^{commit}')
    if git(root,'status','--porcelain','--untracked-files=all'):
        raise ValueError('ship requires a clean committed source tree')
    return sha,git(root,'rev-parse','HEAD^{tree}')


def json_command(argv,root,env):
    result=subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True)
    try:data=json.loads(result.stdout)
    except (ValueError,TypeError) as exc:raise ValueError('invalid validation JSON') from exc
    if not isinstance(data,dict) or type(data.get('exit_code')) is not int or data['exit_code']!=result.returncode or result.returncode not in (0,1,2):
        raise ValueError('validation exit status mismatch')
    return data


def command(argv,root,env,*,purpose="ship"):
    result=subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True)
    if result.returncode:raise ValueError('configured '+purpose+' command failed')


# The longest a render row may take to answer. A server that starts local emulators and seeds them before its app
# (the web pilot waits up to 120s for its emulator ports alone) does not answer within the 30s a bare process
# needs; measured 2026-11-20, the web row was still seeding when a 30s wait stopped it. A healthy row answers
# as soon as it is up, so the bound only sets how long a broken one takes to fail.
READY_SECONDS=180


def wait_ready(name,url,process,log):
    """Wait for one render row. A failure names the row, what it last answered, and where its server's output is."""
    deadline=time.monotonic()+READY_SECONDS;last='no connection'
    where=name+' ('+url+')'
    while True:
        if process is not None and process.poll() is not None:
            raise ValueError('bound render server '+where+' exited before validation (exit '+str(process.returncode)+'); log: '+str(log))
        try:
            with urlopen(url,timeout=0.5) as response:
                if response.status<400:return
                last='HTTP '+str(response.status)
        except HTTPError as exc:last='HTTP '+str(exc.code)
        except (OSError,ValueError):pass
        if time.monotonic()>=deadline:
            raise ValueError('bound render server '+where+' did not become ready in '+str(READY_SECONDS)+'s (last: '+last+'); log: '+str(log))
        time.sleep(0.1)


@contextmanager
def render_servers(root,servers,env):
    """Own each endpoint's process and ephemeral origin; ignore inherited URLs.

    Each server's output goes to a private log file, kept only when a row fails, so the failure can be read
    rather than guessed. The log is never echoed: it is the app's own output and may hold what it was given."""
    processes=[];urls={};child_env={key:value for key,value in env.items() if key not in servers}
    log_dir=Path(tempfile.mkdtemp(prefix='design-ship-render-'));failed=False
    try:
        origins={};logs={}
        # Rows that start a server first, then rows that render another path of one of them.
        for name,row in sorted(servers.items(),key=lambda item:'same_as' in item[1]):
            if 'same_as' in row:
                url=origins[row['same_as']]+row['path']
                wait_ready(name,url,None,logs[row['same_as']])
                urls[name]=url;child_env[name]=url
                continue
            with socket.socket() as listener:
                listener.bind(('127.0.0.1',0));port=listener.getsockname()[1]
            argv=[value.replace('{port}',str(port)) for value in row['command']]
            logs[name]=log_dir/(name+'.log')
            with open(os.open(logs[name],os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600),'w') as out:
                process=subprocess.Popen(argv,cwd=root,env=child_env,stdout=out,stderr=subprocess.STDOUT,start_new_session=True)
            process.log_dir=str(log_dir)
            processes.append(process)
            origins[name]='http://127.0.0.1:'+str(port)
            url=origins[name]+row['path']
            wait_ready(name,url,process,logs[name])
            urls[name]=url;child_env[name]=url
        yield child_env,urls,processes
    except BaseException:
        failed=True
        raise
    finally:
        for process in processes:
            try:os.killpg(process.pid,signal.SIGTERM)
            except ProcessLookupError:pass
        for process in processes:
            try:process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                try:os.killpg(process.pid,signal.SIGKILL)
                except ProcessLookupError:pass
                process.wait()
        if not failed:shutil.rmtree(log_dir,ignore_errors=True)


def gate(root,name,channel,catalogs,*,deploy=False,source_commit=None,predicted=False,local=False):
    root=Path(root).resolve()
    commit,tree=clean_commit(root)
    env=hook_free_env()
    if local and deploy:raise ValueError('local integration cannot deploy')
    # L0 reads only committed design/ inputs, so it runs before the (slow) preparation: a design system that
    # cannot ship never pays for installing or building the app.
    l0_command=env.get('DESIGN_AUTHOR_COMMAND') or str(SCRIPT/'design-author')
    l0=json_command([l0_command,'check','--ship','--repo',str(root),'--catalog-dir',str(catalogs),'--include-manifest'],root,env)
    if l0['exit_code']==2:return {'exit_code':2,'errors':l0.get('errors',[]) or ['ship L0 failed'],'l0':l0}
    if local:
        if 'prepare_command' in channel:command(channel['prepare_command'],root,env,purpose='prepare')
        if clean_commit(root)!=(commit,tree):raise ValueError('preparation changed the committed source')
    if l0.get('counts',{}).get('items',0)<=0:raise ValueError('zero L0 ship targets')
    inputs=l0['manifest']['web']['input_hashes']
    tracked=set(git(root,'ls-files').splitlines())
    if any(layer_paths.rel('design_dir',name) not in tracked for name in inputs):
        raise ValueError('all design inputs must be committed, including ignored inputs')
    validation=read_data(safe_path(root,layer_paths.rel('validation')))
    urls=validation.get('urls',[])
    if not isinstance(urls,list) or not urls:
        raise ValueError('render server coverage differs from validation endpoints')
    endpoints=[validation_endpoint(row)[0] for row in urls]
    if len(set(endpoints))!=len(endpoints) or set(endpoints)!=set(channel['render_servers']):
        raise ValueError('render server coverage differs from validation endpoints')
    if not local and 'build_command' in channel:command(channel['build_command'],root,env)
    before=None if local else artifact_hashes(root,channel['artifacts'])
    clean_commit(root)
    with render_servers(root,channel['render_servers'],env) as (bound_env,bound_urls,processes):
        validator=bound_env.get('DESIGN_CHECK_COMMAND') or str(SCRIPT/'design-check')
        web=json_command([validator,'--repo',str(root),'--catalog-dir',str(catalogs)],root,bound_env)
        if web['exit_code'] not in (0,1) or web.get('errors'):
            return {'exit_code':2,'errors':['L0/L1/L2 validation failed'],'validation':web}
        evidence=validate_web_result(root,web,bound_urls)
        if any(process.poll() is not None for process in processes):raise ValueError('bound render server exited during validation')
    if clean_commit(root)!=(commit,tree) or (not local and artifact_hashes(root,channel['artifacts'])!=before):
        raise ValueError('source or ship artifacts changed during validation')
    from resolved_contract import load_manifest
    if load_manifest(layer_paths.path(root,'resolved','web.json'))['input_hashes']!=inputs:
        raise ValueError('design inputs changed after ship L0')
    receipt={'schema_version':1,'exit_code':0,'channel':name,'commit':commit,'tree':tree,
        'source_commit':source_commit or commit,'predicted_merge':predicted,'artifacts':before,
        'design_input_hashes':inputs,'render_evidence':evidence,'stages':web['stages'],'report':web.get('report',{}),
        'warnings':l0.get('warnings',[])+web.get('warnings',[]),'deployment':'not-requested'}
    if clean_commit(root)!=(commit,tree) or (not local and artifact_hashes(root,channel['artifacts'])!=before):
        raise ValueError('source or artifacts changed before shipment')
    if local:
        receipt.pop('artifacts')
        receipt.update(scope='local-integration',production_protected=False,deployment='not-applicable')
    if deploy:
        env.update(DESIGN_SHIP_COMMIT=commit,DESIGN_SHIP_TREE=tree,
                   DESIGN_SHIP_ARTIFACTS_JSON=json.dumps(before,sort_keys=True))
        command(channel['ship_command'],root,env)
        if clean_commit(root)!=(commit,tree) or (not local and artifact_hashes(root,channel['artifacts'])!=before):
            raise ValueError('ship command changed the checked source or artifact set')
        receipt['deployment']='command-completed'
    output=safe_path(root,layer_paths.rel('resolved','integration.json' if local else 'ship-'+name+'.json'))
    output.parent.mkdir(parents=True,exist_ok=True)
    output.write_text(json.dumps(receipt,ensure_ascii=False,sort_keys=True,indent=2)+'\n')
    return receipt


def doctor(root,channels,integration=None):
    errors=[]
    if not channels and integration is None:errors.append('no ship:true channels or local integration gate declared')
    marker=root/'.vibe-guard/design-ship.enabled'
    if marker.is_symlink() or not marker.is_file() or marker.read_text().strip()!='1':errors.append('design ship marker missing or invalid')
    try:
        hook=Path(git(root,'rev-parse','--git-path','hooks/pre-push'))
        if not hook.is_absolute():hook=root/hook
        extension=hook.parent/'pre-push.d/60-design-ship'
        if not hook.is_file() or not os.access(hook,os.X_OK):raise ValueError('pre-push missing')
        text=hook.read_text()
        if 'pre-push.d' not in text or '"$ext"' not in text or 'scan-push' not in text:raise ValueError('scanner-preserving pre-push dispatcher missing')
        dispatcher_propagates(text)  # run it: a dispatcher that swallows a failing extension cannot stop a push
        if not extension.is_file() or not os.access(extension,os.X_OK) or extension.read_bytes()!=(SCRIPT.parent/'hooks/pre-push.d/60-design-ship').read_bytes():
            raise ValueError('current design pre-push extension missing')
    except (OSError,ValueError) as exc:errors.append(str(exc))
    return {'exit_code':2 if errors else 0,'errors':errors,'channels':sorted(channels),
        'local_integration':integration is not None,'production_protected':False,
        'scope':'configuration-diagnostic'}


def pre_push(root,channels,catalogs,refs_file,integration=None):
    updates,pushed=read_push_refs(root,refs_file)
    sources=[]
    for sha,_,_ in updates:
        if sha not in sources:sources.append(sha)
    jobs=[(name,row,False) for name,row in channels.items()]
    if integration is not None:jobs.append(('local-integration',integration,True))
    if not jobs:raise ValueError('no ship:true channels or local integration gate declared')
    results=[]
    for source in sources:
        for name,channel,local in jobs:
            # The base as it is now, and what this same push sets it to (shared with the IA gate).
            for base in merge_bases(root,channel['merge_base'],pushed):
                # No merge conflict or missing merge-base is treated as a passed check.
                # Like the IA gate, it takes the pushing user's configured email; no identity stops the push.
                synthetic=predicted_commit(root,base,source,name='Design Gate',message='Design predicted merge validation')
                with tempfile.TemporaryDirectory(prefix='design-ship-') as temporary:
                    snapshot=Path(temporary)/'checkout'
                    git(root,'worktree','add','--detach',str(snapshot),synthetic)
                    try:
                        # Config and commands come from the predicted tree itself.
                        config=read_data(layer_paths.path(snapshot,'design_config'))
                        predicted_channels=validate_channels(config)
                        predicted_integration=validate_integration_gate(config)
                        if set(predicted_channels)!=set(channels) or (predicted_integration is not None)!=(integration is not None):
                            raise ValueError('predicted tree changed gate coverage')
                        if local:
                            predicted_channel=predicted_integration
                            if predicted_channel is None:raise ValueError('predicted tree removed local integration gate')
                        else:
                            if name not in predicted_channels:raise ValueError('predicted tree removed ship channel')
                            predicted_channel=predicted_channels[name]
                        if predicted_channel['merge_base']!=channel['merge_base']:
                            raise ValueError('predicted tree changed the gate merge base')
                        result=gate(snapshot,name,predicted_channel,catalogs,source_commit=source,predicted=True,local=local)
                        results.append(result)
                        if result['exit_code']==2:return {'exit_code':2,'errors':['predicted merge ship gate failed'],'checks':results}
                    finally:git(root,'worktree','remove','--force',str(snapshot))
    return {'exit_code':0,'checks':results,'skipped':[] if sources else ['deleted refs only; no shipment']}


def probe(repo):
    """Exit 0 when the repository has a design system, 3 when it has none, 2 when its layout cannot be read.

    The pre-push hook asks this instead of looking for the config file itself, so the place stays the rule's."""
    try:
        config=layer_paths.path(git_root(repo).resolve(),'design_config')  # the guard runs first
    except ValueError as exc:  # LayerLayoutError is a ValueError
        print(str(exc),file=sys.stderr)
        return 2
    return 0 if config.exists() or config.is_symlink() else 3


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo',default='.')
    parser.add_argument('--catalog-dir',type=Path,default=CATALOG)
    modes=parser.add_mutually_exclusive_group()
    modes.add_argument('--channel')
    modes.add_argument('--pre-push',action='store_true')
    modes.add_argument('--doctor',action='store_true')
    modes.add_argument('--probe',action='store_true',help='is a design system present? exit 0 yes, 3 no, 2 unreadable layout; used by the pre-push hook')
    modes.add_argument('--integration',action='store_true',help='check the declared local integration gate; never ships')
    parser.add_argument('--refs-file',type=Path)
    args=parser.parse_args(argv)
    if args.probe:return probe(args.repo)
    try:
        root=git_root(args.repo).resolve()
        layer_paths.guard(root)  # before any read: a moved repository must stop, never skip
        config=read_data(layer_paths.path(root,'design_config'))
        channels=validate_channels(config)
        integration=validate_integration_gate(config)
        if args.doctor:result=doctor(root,channels,integration)
        elif args.pre_push:result=pre_push(root,channels,args.catalog_dir.resolve(),args.refs_file,integration)
        elif args.integration:
            if integration is None:raise ValueError('no local integration gate declared')
            result=gate(root,'local-integration',integration,args.catalog_dir.resolve(),local=True)
        elif not channels:raise ValueError('no ship:true channels declared; L0-only shipment is unsupported')
        elif args.channel:
            if args.channel not in channels:raise ValueError('unknown or nonshipping channel')
            result=gate(root,args.channel,channels[args.channel],args.catalog_dir.resolve(),deploy=True)
        else:
            checks=[gate(root,name,row,args.catalog_dir.resolve()) for name,row in channels.items()]
            result={'exit_code':2 if any(r['exit_code']==2 for r in checks) else 0,'checks':checks}
    except ValueError as exc:
        result={'exit_code':2,'errors':[str(exc)]}
    except (OSError,KeyError,TypeError,subprocess.SubprocessError):
        result={'exit_code':2,'errors':['ship input or runtime unavailable; shipment stopped']}
    print(json.dumps(result,ensure_ascii=False,sort_keys=True))
    return result['exit_code']


if __name__=='__main__':raise SystemExit(main())
