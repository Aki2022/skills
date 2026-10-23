#!/usr/bin/env python3
"""Bind L0/L1/L2 to a clean commit and exact built artifacts before shipment."""
import argparse
from contextlib import contextmanager
import hashlib
import json
import os
from pathlib import Path
import signal
import socket
import subprocess
import sys
import tempfile
import time
from urllib.request import urlopen

from design_author import git_root, read_data, CATALOG
from ship_contract import validate_channels, artifact_hashes, validate_web_result, safe_path

SCRIPT=Path(__file__).resolve().parent


def git(root,*args,input=None):
    result=subprocess.run(['git','-C',str(root),*args],input=input,capture_output=True,text=True)
    if result.returncode:raise ValueError('git '+args[0]+' failed')
    return result.stdout.strip()


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


def command(argv,root,env):
    result=subprocess.run(argv,cwd=root,env=env,capture_output=True,text=True)
    if result.returncode:raise ValueError('configured ship command failed')


@contextmanager
def render_servers(root,servers,env):
    """Own each endpoint's process and ephemeral origin; ignore inherited URLs."""
    processes=[];urls={};child_env=dict(env)
    try:
        for name,row in servers.items():
            with socket.socket() as listener:
                listener.bind(('127.0.0.1',0));port=listener.getsockname()[1]
            argv=[value.replace('{port}',str(port)) for value in row['command']]
            process=subprocess.Popen(argv,cwd=root,env=child_env,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,start_new_session=True)
            processes.append(process)
            url='http://127.0.0.1:'+str(port)+row['path']
            deadline=time.monotonic()+30
            while True:
                if process.poll() is not None:raise ValueError('bound render server exited before validation')
                try:
                    with urlopen(url,timeout=0.5) as response:
                        if response.status<400:break
                except (OSError,ValueError):pass
                if time.monotonic()>=deadline:raise ValueError('bound render server did not become ready')
                time.sleep(0.1)
            urls[name]=url;child_env[name]=url
        yield child_env,urls,processes
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


def gate(root,name,channel,catalogs,*,deploy=False,source_commit=None,predicted=False):
    root=Path(root).resolve()
    commit,tree=clean_commit(root)
    env=dict(os.environ)
    l0_command=env.get('DESIGN_AUTHOR_COMMAND') or str(SCRIPT/'design-author')
    l0=json_command([l0_command,'check','--ship','--repo',str(root),'--catalog-dir',str(catalogs),'--include-manifest'],root,env)
    if l0['exit_code']==2:return {'exit_code':2,'errors':l0.get('errors',[]) or ['ship L0 failed'],'l0':l0}
    if l0.get('counts',{}).get('items',0)<=0:raise ValueError('zero L0 ship targets')
    inputs=l0['manifest']['web']['input_hashes']
    tracked=set(git(root,'ls-files').splitlines())
    if any('design/'+name not in tracked for name in inputs):
        raise ValueError('all design inputs must be committed, including ignored inputs')
    validation=read_data(safe_path(root,'design/validation.json'))
    urls=validation.get('urls',[])
    if not urls or any(not isinstance(row,dict) or set(row)!={'env'} for row in urls) or {row['env'] for row in urls}!=set(channel['render_servers']):
        raise ValueError('render server coverage differs from validation endpoints')
    if 'build_command' in channel:command(channel['build_command'],root,env)
    before=artifact_hashes(root,channel['artifacts'])
    clean_commit(root)
    with render_servers(root,channel['render_servers'],env) as (bound_env,bound_urls,processes):
        validator=bound_env.get('DESIGN_CHECK_COMMAND') or str(SCRIPT/'design-check')
        web=json_command([validator,'--repo',str(root),'--catalog-dir',str(catalogs)],root,bound_env)
        if web['exit_code'] not in (0,1) or web.get('errors'):
            return {'exit_code':2,'errors':['L0/L1/L2 validation failed'],'validation':web}
        evidence=validate_web_result(root,web,bound_urls)
        if any(process.poll() is not None for process in processes):raise ValueError('bound render server exited during validation')
    if clean_commit(root)!=(commit,tree) or artifact_hashes(root,channel['artifacts'])!=before:
        raise ValueError('source or ship artifacts changed during validation')
    from resolved_contract import load_manifest
    if load_manifest(root/'design/.resolved/web.json')['input_hashes']!=inputs:
        raise ValueError('design inputs changed after ship L0')
    receipt={'schema_version':1,'exit_code':0,'channel':name,'commit':commit,'tree':tree,
        'source_commit':source_commit or commit,'predicted_merge':predicted,'artifacts':before,
        'design_input_hashes':inputs,'render_evidence':evidence,'stages':web['stages'],'report':web.get('report',{}),
        'warnings':l0.get('warnings',[]),'deployment':'not-requested'}
    if clean_commit(root)!=(commit,tree) or artifact_hashes(root,channel['artifacts'])!=before:
        raise ValueError('source or artifacts changed before shipment')
    if deploy:
        env.update(DESIGN_SHIP_COMMIT=commit,DESIGN_SHIP_TREE=tree,
                   DESIGN_SHIP_ARTIFACTS_JSON=json.dumps(before,sort_keys=True))
        command(channel['ship_command'],root,env)
        if clean_commit(root)!=(commit,tree) or artifact_hashes(root,channel['artifacts'])!=before:
            raise ValueError('ship command changed the checked source or artifact set')
        receipt['deployment']='command-completed'
    output=safe_path(root,'design/.resolved/ship-'+name+'.json')
    output.parent.mkdir(parents=True,exist_ok=True)
    output.write_text(json.dumps(receipt,ensure_ascii=False,sort_keys=True,indent=2)+'\n')
    return receipt


def doctor(root,channels):
    errors=[]
    if not channels:errors.append('no ship:true channels declared')
    marker=root/'.vibe-guard/design-ship.enabled'
    if marker.is_symlink() or not marker.is_file() or marker.read_text().strip()!='1':errors.append('design ship marker missing or invalid')
    try:
        hook=Path(git(root,'rev-parse','--git-path','hooks/pre-push'))
        if not hook.is_absolute():hook=root/hook
        extension=hook.parent/'pre-push.d/60-design-ship'
        if not hook.is_file() or not os.access(hook,os.X_OK):raise ValueError('pre-push missing')
        text=hook.read_text()
        if 'pre-push.d' not in text or '"$ext"' not in text or 'scan-push' not in text:raise ValueError('scanner-preserving pre-push dispatcher missing')
        if not extension.is_file() or not os.access(extension,os.X_OK) or extension.read_bytes()!=(SCRIPT.parent/'hooks/pre-push.d/60-design-ship').read_bytes():
            raise ValueError('current design pre-push extension missing')
    except (OSError,ValueError) as exc:errors.append(str(exc))
    return {'exit_code':2 if errors else 0,'errors':errors,'channels':sorted(channels)}


def pre_push(root,channels,catalogs,refs_file):
    if refs_file is None:raise ValueError('pre-push requires the preserved refs file')
    records=Path(refs_file).read_text().splitlines()
    if not records:raise ValueError('zero push refs')
    sources=[]
    for row in records:
        parts=row.split()
        if len(parts)!=4:raise ValueError('invalid push ref record')
        sha=parts[1]
        if set(sha)=={'0'}:continue
        verified=git(root,'rev-parse','--verify',sha+'^{commit}')
        if verified!=sha:raise ValueError('push source must be a commit')
        if sha not in sources:sources.append(sha)
    results=[]
    for source in sources:
        for name,channel in channels.items():
            base=git(root,'rev-parse','--verify',channel['merge_base']+'^{commit}')
            # No merge conflict or missing merge-base is treated as a passed check.
            tree=git(root,'merge-tree','--write-tree',base,source).splitlines()[0]
            synthetic=git(root,'-c','user.name=Design Gate','-c','user.email=design-gate@example.invalid',
                'commit-tree',tree,'-p',source,'-p',base,input='Design predicted merge validation\n')
            with tempfile.TemporaryDirectory(prefix='design-ship-') as temporary:
                snapshot=Path(temporary)/'checkout'
                git(root,'worktree','add','--detach',str(snapshot),synthetic)
                try:
                    # Config and commands come from the predicted tree itself.
                    predicted_channels=validate_channels(read_data(snapshot/'design/design.yaml'))
                    if name not in predicted_channels:raise ValueError('predicted tree removed ship channel')
                    result=gate(snapshot,name,predicted_channels[name],catalogs,source_commit=source,predicted=True)
                    results.append(result)
                    if result['exit_code']==2:return {'exit_code':2,'errors':['predicted merge ship gate failed'],'checks':results}
                finally:git(root,'worktree','remove','--force',str(snapshot))
    return {'exit_code':0,'checks':results,'skipped':[] if sources else ['deleted refs only; no shipment']}


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo',default='.')
    parser.add_argument('--catalog-dir',type=Path,default=CATALOG)
    modes=parser.add_mutually_exclusive_group()
    modes.add_argument('--channel')
    modes.add_argument('--pre-push',action='store_true')
    modes.add_argument('--doctor',action='store_true')
    parser.add_argument('--refs-file',type=Path)
    args=parser.parse_args(argv)
    try:
        root=git_root(args.repo).resolve()
        channels=validate_channels(read_data(root/'design/design.yaml'))
        if args.doctor:result=doctor(root,channels)
        elif not channels:raise ValueError('no ship:true channels declared; L0-only shipment is unsupported')
        elif args.pre_push:result=pre_push(root,channels,args.catalog_dir.resolve(),args.refs_file)
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
