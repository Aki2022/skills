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
from page_inventory import check_pages
from afm_judge import judge, AfmUnavailable
from item_checks import L2_GATED


def validation_endpoint(row):
    if not isinstance(row,dict) or set(row) not in ({'env'},{'env','profile'}) or not isinstance(row.get('env'),str):
        raise ValueError('rendering URL must name an environment variable and an optional profile')
    profile=row.get('profile','lp')
    if profile not in ('lp','product-ui'):
        raise ValueError('invalid rendering profile')
    return row['env'],profile


def page_evidence(root,endpoints,spec):
    """Per rendered page: which components are present and what each collected checker found (all runs)."""
    pages=[]
    for index,(env,_profile) in enumerate(endpoints):
        measured=read_data(inside(root/'design/.resolved','l2-'+str(index)+'.json'))
        runs=[r for r in measured.get('runs',[]) if r.get('repeat',0)==0]
        if not runs or any(not isinstance(r.get('item_checks'),dict) for r in runs):
            raise ValueError('L2 item check evidence missing')
        presence={doc:max(int(r['item_checks']['presence'].get(doc,0)) for r in runs) for doc in spec['applicability']}
        results={}
        for checker in spec['collect']:
            rows=[r['item_checks']['results'].get(checker) for r in runs]
            if any(not isinstance(row,dict) or row.get('checked') is not True for row in rows):
                results[checker]=None
            else:
                results[checker]={'count':sum(row['count'] for row in rows),
                                  'sample':[x for row in rows for x in row.get('sample',[])][:4]}
        pages.append({'env':env,'index':index,'runs':runs,'presence':presence,'results':results})
    return pages


def verify_items(root,l0,endpoints,report):
    """Machine items are checked by code on every page where they apply; AFM gets only applicable leftovers."""
    manifest=l0['manifest']['web'];routes=manifest['coverage']['item_routes'];spec=manifest['item_checks']
    pages=page_evidence(root,endpoints,spec)
    machine,afm_pairs,failures,not_applicable,causes={},[],[],0,{}
    for ident,row in sorted(routes.items()):
        for page in pages:
            if row['document'] in spec['applicability'] and page['presence'][row['document']]==0:
                not_applicable+=1
                machine.setdefault(ident,{})[page['env']]='n/a' if row['route']=='machine' else None
                continue
            if row['route']=='afm':
                afm_pairs.append([ident,page['env']]);continue
            status='pass'
            for checker in row['checkers']:
                if checker in L2_GATED or checker.startswith('rule:'):continue
                found=page['results'].get(checker)
                if found is None:
                    cause=f"{page['env']}: {checker} was not measured"
                elif found['count']:
                    cause=f"{page['env']}: {checker} {found['count']} ({'; '.join(found['sample'])})"
                else:continue
                status='fail';failures.append(f"{ident} on {cause}");causes.setdefault(cause,[]).append(ident)
            machine.setdefault(ident,{})[page['env']]=status
    machine={k:{e:v for e,v in pages_.items() if v} for k,pages_ in machine.items()}
    record=inside(root/'design/.resolved','items.json')
    record.write_text(json.dumps({'schema_version':1,'pages':[p['env'] for p in pages],'machine':machine,
        'afm_pairs':afm_pairs,'failures':failures},ensure_ascii=False,indent=1))
    code=2 if failures else 0
    report['stages'].append({'stage':'ITEMS','exit_code':code,'machine_items':sum(1 for r in routes.values() if r['route']=='machine'),
        'afm_pairs':len(afm_pairs),'not_applicable':not_applicable,'failures':len(failures)})
    stage(report['report'],'ITEMS',code,len(pages))
    if failures:
        report['exit_code']=2
        # One line per cause, naming every adopted item it breaks; per-item rows stay in items.json.
        report['errors'].extend(f"{cause} — breaks adopted items {', '.join(ids)}" for cause,ids in causes.items())
    return pages,afm_pairs,code


def run_afm(root,l0,config,pages,afm_pairs,report):
    if not afm_pairs:
        skip(report['report'],'AFM','no applicable adopted item needs AFM judgment')
        return 0
    candidates={row['id']:row for row in l0.get('afm_candidates',[])}
    scheme=config['schemes'][0];width=max(config['viewports'])
    merged={'errors':[],'warnings':[],'uncovered':[],'screened':[],'calls':0,'cached':0}
    artifacts=[]
    for page in pages:
        wanted=[candidates[i] for i,env in afm_pairs if env==page['env']]
        if not wanted:continue
        chosen=[r for r in page['runs'] if r.get('width')==width and r.get('scheme')==scheme]
        if len(chosen)!=1:raise ValueError('L2 capture for AFM missing')
        digest=chosen[0].get('render_evidence',{}).get('screenshot_sha256','')
        path=inside(root/'design/.resolved','screens/'+digest+'.png')
        if not path.is_file():raise ValueError('L2 screenshot for AFM missing')
        artifact={'id':f"L2:{page['env']}:{width}:{scheme}",'kind':'image','path':str(path)}
        artifacts.append(artifact['id'])
        for run in page['runs']:
            if run is not chosen[0]:
                report['report']['skipped'].append({'id':f"AFM:{page['env']}:{run.get('width')}:{run.get('scheme')}",
                    'reason':'AFM judges the widest viewport in the first scheme'})
        result=judge(wanted,[artifact])
        for key in ('errors','warnings','uncovered','screened'):merged[key].extend(result[key])
        for key in ('calls','cached'):merged[key]+=result[key]
    code=2 if merged['errors'] or merged['uncovered'] else (1 if merged['warnings'] else 0)
    report['stages'].append({'stage':'AFM','exit_code':code,'pairs':len(afm_pairs),'artifacts':len(artifacts),
                             'judged':len(merged['screened']),'calls':merged['calls'],'cached':merged['cached']})
    stage(report['report'],'AFM',code,len(artifacts))
    inside(root/'design/.resolved','afm.json').write_text(json.dumps({'schema_version':1,'pairs':afm_pairs,
        'artifacts':artifacts,'verdicts':merged['screened'],'uncovered':merged['uncovered']},ensure_ascii=False,indent=1))
    report['report']['uncovered'].extend({'id':'AFM:'+r['id']+'@'+str(r['artifact']),'reason':r['reason']} for r in merged['uncovered'])
    report['report']['screened'].append({'id':'AFM','judged':len(merged['screened']),'advisory_violations':len(merged['warnings'])})
    report.setdefault('warnings',[]).extend(merged['warnings'])
    if code==2:
        report['exit_code']=2
        report['errors'].extend(merged['errors'] or [f"AFM left {len(merged['uncovered'])} item/page pairs unjudged"])
    return code


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
        if not isinstance(config,dict) or set(config)-{'pages'}!={'schema_version','html','urls','viewports','schemes'} or config['schema_version']!=1:
            raise ValueError('invalid validation configuration')
        inventory=check_pages(root,config)
        report['pages']=inventory['pages']
        if inventory['errors']:
            report['errors']=inventory['errors']
            report['report']['missing'].append({'id':'pages','reason':'page inventory incomplete or invalid'})
            print(json.dumps(report,ensure_ascii=False));return 2
        report['warnings']=list(inventory['warnings'])
        for warning in inventory['warnings']:
            report['report']['uncovered'].append({'id':'pages','reason':warning})
        html=config['html'];urls=config['urls']
        if not isinstance(html,list) or not html or not isinstance(urls,list) or not urls:
            raise ValueError('zero L1/L2 targets')
        endpoints=[validation_endpoint(row) for row in urls]
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
        for index,(target_env,profile) in enumerate(endpoints):
            row={'env':target_env}
            url=os.environ.get(row['env'],'')
            if not url.startswith(('http://','https://')):
                raise ValueError('missing or invalid rendering URL: '+row['env'])
            output=inside(root/'design/.resolved','l2-'+str(index)+'.json')
            output.unlink(missing_ok=True)
            screens=root/'design/.resolved/screens'
            l2=subprocess.run([os.environ.get('DESIGN_AUDIT_NODE','node'),
                str(skills/'own-website-audit/scripts/measure_lp.mjs'),url,'--manifest',str(resolved),
                '--environment',str(environment),'--profile',profile,
                '--viewports',','.join(map(str,config['viewports'])),'--schemes',','.join(config['schemes']),
                '--json',str(output)],cwd=root,capture_output=True,text=True,
                env={**os.environ,'DESIGN_SCREENSHOT_DIR':str(screens)})
            if not output.is_file():
                report['report']['missing'].append({'id':'L2:'+row['env'],'reason':'no fresh fixed rendering report'})
                raise ValueError('fresh fixed rendering report missing')
            measured=read_data(output)
            if measured.get('profile')!=profile:raise ValueError('rendering profile evidence mismatch')
            for skipped in measured.get('skipped',[]):
                report['report']['skipped'].append({'id':'L2:'+target_env+':'+skipped['id'],'reason':skipped['reason']})
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
            report['stages'].append({'stage':'L2','exit_code':l2.returncode,'target_env':row['env'],'profile':profile})
            if l2.returncode:
                report['exit_code']=l2.returncode
                reason=next((f for f in measured.get('failures',[]) if isinstance(f,dict) and f.get('id')=='render-execution-failed'),None)
                if reason:
                    report['errors'].append('L2 '+row['env']+' failed at '+str(reason.get('stage'))+': '
                        +str(reason.get('error'))+' ('+str(reason.get('detail',''))[:200]+')')
                report['errors'].append('L2 failed; inspect the generated report or rerun the measurement command')
                print(json.dumps(report,ensure_ascii=False));return report['exit_code']
        pages,afm_pairs,items_code=verify_items(root,l0,endpoints,report)
        if items_code==2:
            print(json.dumps(report,ensure_ascii=False));return 2
        afm_code=run_afm(root,l0,config,pages,afm_pairs,report)
        if afm_code==2:
            print(json.dumps(report,ensure_ascii=False));return 2
        report['exit_code']=max(l0['exit_code'],afm_code,1 if report.get('warnings') else 0)
        report['uncovered']=sorted(uncovered)
    except AfmUnavailable as error:
        report['exit_code']=2
        report['errors'].append(str(error)+'; AFM judgment is required and was not run')
        report['report']['missing'].append({'id':'AFM','reason':'on-device model unavailable'})
    except (ValueError,OSError,KeyError,TypeError):
        report['errors'].append('validation input, runtime dependency or rendering endpoint unavailable')
        report['report']['missing'].append({'id':'validation','reason':'input, runtime dependency or rendering endpoint unavailable'})
    print(json.dumps(report,ensure_ascii=False))
    return report['exit_code']


if __name__=='__main__':
    raise SystemExit(main())
