#!/usr/bin/env python3
"""Run real L0 -> L1 -> L2. Rendering requires caller authorization before invocation."""
import argparse
import copy
import json
import os
from pathlib import Path
import subprocess
import sys
from design_author import check_design, git_root, inside, read_data, CATALOG
from reporting import empty_report,stage,skip


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo',default='.')
    parser.add_argument('--catalog-dir',type=Path,default=CATALOG)
    args=parser.parse_args()
    report={'exit_code':2,'stages':[],'errors':[],'report':empty_report()}
    try:
        root=git_root(args.repo)
        l0=check_design(root,catalog_dir=args.catalog_dir)
        report['report']=copy.deepcopy(l0['report'])
        skip(report['report'],'L1','not reached after L0 or unavailable validation inputs')
        skip(report['report'],'L2','not reached after L0/L1 or unavailable rendering inputs')
        report['stages'].append({'stage':'L0','exit_code':l0['exit_code'],'counts':l0['counts']})
        if l0['exit_code']==2:
            report['errors']=l0['errors']
            print(json.dumps(report,ensure_ascii=False));return 2
        config=read_data(inside(root/'design','validation.json'))
        if not isinstance(config,dict) or set(config)!={'schema_version','html','urls','viewports','schemes'} or config['schema_version']!=1:
            raise ValueError('invalid validation configuration')
        html=config['html'];urls=config['urls']
        if not isinstance(html,list) or not html or not isinstance(urls,list) or not urls:
            raise ValueError('zero L1/L2 targets')
        files=[inside(root,path) for path in html]
        resolved=root/'design/.resolved/web.json'
        skills=Path(__file__).resolve().parents[2]
        l1=subprocess.run([sys.executable,str(skills/'own-web-design/scripts/audit_html.py'),
                           *map(str,files),'--manifest',str(resolved)],cwd=root,capture_output=True,text=True)
        report['stages'].append({'stage':'L1','exit_code':l1.returncode,'targets':len(files)})
        stage(report['report'],'L1',l1.returncode,len(files))
        if l1.returncode:
            report['exit_code']=l1.returncode
            report['errors'].append('L1 failed; rendering was not run')
            print(json.dumps(report,ensure_ascii=False));return report['exit_code']
        if not isinstance(config['viewports'],list) or not config['viewports'] or any(type(v)!=int or v<=0 for v in config['viewports']):
            raise ValueError('zero or invalid viewport targets')
        if not isinstance(config['schemes'],list) or not config['schemes'] or any(v not in ('light','dark') for v in config['schemes']):
            raise ValueError('zero or invalid scheme targets')
        environment=inside(root/'design','render-environment.json')
        if not environment.is_file():raise ValueError('fixed rendering environment missing')
        uncovered=set(l0['manifest']['web']['coverage']['uncovered'])
        for index,row in enumerate(urls):
            if not isinstance(row,dict) or set(row)!={'env'} or not isinstance(row['env'],str):
                raise ValueError('rendering URL must name an environment variable')
            url=os.environ.get(row['env'],'')
            if not url.startswith(('http://','https://')):
                raise ValueError('missing or invalid rendering URL: '+row['env'])
            output=inside(root/'design/.resolved','l2-'+str(index)+'.json')
            output.unlink(missing_ok=True)
            l2=subprocess.run([os.environ.get('DESIGN_AUDIT_NODE','node'),
                str(skills/'own-website-audit/scripts/measure_lp.mjs'),url,'--manifest',str(resolved),
                '--environment',str(environment),
                '--viewports',','.join(map(str,config['viewports'])),'--schemes',','.join(config['schemes']),
                '--json',str(output)],cwd=root,capture_output=True,text=True)
            if not output.is_file():
                report['report']['missing'].append({'id':'L2:'+row['env'],'reason':'no fresh fixed rendering report'})
                raise ValueError('fresh fixed rendering report missing')
            measured=read_data(output)
            fixed=measured.get('environment')
            if not isinstance(fixed,dict) or type(fixed.get('repeat_equal')) is not bool:
                raise ValueError('missing fixed rendering evidence')
            if l2.returncode==0 and not fixed['repeat_equal']:raise ValueError('inconsistent rendering result')
            if not isinstance(measured.get('uncovered'),list):
                raise ValueError('missing L2 coverage report')
            uncovered.update(measured['uncovered'])
            report['uncovered']=sorted(uncovered)
            report['report']['uncovered'].extend({'id':'L2:'+row['env']+':'+ident,'reason':'rendering scope not covered'} for ident in measured['uncovered'])
            stage(report['report'],'L2',l2.returncode,index+1)
            report['stages'].append({'stage':'L2','exit_code':l2.returncode,'target_env':row['env']})
            if l2.returncode:
                report['exit_code']=l2.returncode
                report['errors'].append('L2 failed; inspect the generated report or rerun the measurement command')
                print(json.dumps(report,ensure_ascii=False));return report['exit_code']
        report['exit_code']=l0['exit_code']
        report['uncovered']=sorted(uncovered)
    except (ValueError,OSError,KeyError,TypeError):
        report['errors'].append('validation input, runtime dependency or rendering endpoint unavailable')
        report['report']['missing'].append({'id':'validation','reason':'input, runtime dependency or rendering endpoint unavailable'})
    print(json.dumps(report,ensure_ascii=False))
    return report['exit_code']


if __name__=='__main__':
    raise SystemExit(main())
