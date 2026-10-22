"""Offline semantic evidence, migration and the shared runner boundary."""
import copy
import json
import subprocess
import sys
import pytest
from test_core import project,check,write,SKILL,approve_rows,approve
from judgment_ledger import candidate,load_policy,evaluate,migrate_ledger,policy_digest


def example():
    return candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Brand'})


def ledger_for(candidates,scores,policy=None):
    policy=policy or load_policy()
    return {'schema_version':1,'policy':policy,'rows':[
        {'judgment_hash':candidate['judgment_hash'],'score':score} for candidate,score in zip(candidates,scores)],
        'holdout':{'source':'owner','cases':[
            {'id':'fixture-positive','positive':True,'score':0.9},
            {'id':'fixture-negative','positive':False,'score':0.1}]}}


def test_unknown_input_is_uncovered_and_gate_has_no_model_call(monkeypatch):
    row=example()
    assert evaluate([row],None)['uncovered'][0]['id']==row['id']
    def forbidden(*args,**kwargs):raise AssertionError('model/network subprocess used by ledger reader')
    monkeypatch.setattr(subprocess,'run',forbidden)
    result=evaluate([row],ledger_for([row],[0.1]))
    assert not result['errors'] and not result['uncovered']
    assert result['screened'][0]['holdout']['recall']==1.0
    assert result['screened'][0]['status']=='screened'


def test_input_context_model_prompt_and_threshold_invalidate_keys():
    row=example();policy=load_policy()
    for changed in [candidate('doc-relevance','dads-000001','Changed',{'title':'Brand'}),
                    candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Other'})]:
        assert evaluate([changed],ledger_for([row],[0.1]))['uncovered']
    for field,value in [('theta',0.6),('prompt_version','other'),('model','jev-other')]:
        changed={**policy,field:value}
        target=candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Brand'},policy=changed)
        assert target['judgment_hash']!=row['judgment_hash']
        with pytest.raises(ValueError,match='policy'):
            evaluate([target],ledger_for([row],[0.1]),policy=changed)


def test_at_threshold_is_violation_and_empty_or_invalid_holdout_is_never_green():
    row=example();policy=load_policy()
    assert evaluate([row],ledger_for([row],[policy['theta']]))['errors']
    for cases in [[],[{'id':'one','positive':False,'score':0.1}]]:
        data=ledger_for([row],[0.1]);data['holdout']['cases']=cases
        with pytest.raises(ValueError):evaluate([row],data)
    data=ledger_for([row],[float('nan')])
    with pytest.raises(ValueError):evaluate([row],data)
    data=ledger_for([row],[0.1]);data['rows']*=2
    with pytest.raises(ValueError):evaluate([row],data)


def test_migration_requires_same_labeled_holdout_and_records_recall_delta(tmp_path):
    old_policy=load_policy();new_policy={**old_policy,'theta':0.95}
    old=ledger_for([example()],[0.1],old_policy)
    row=candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Brand'},policy=new_policy)
    new=ledger_for([row],[0.1],new_policy)
    migrated=migrate_ledger(old,new,new_policy)
    assert migrated['migration']['before']['recall']==1.0
    assert migrated['migration']['after']['recall']==0.0
    assert migrated['migration']['changed']==['fixture-positive']
    assert evaluate([row],migrated,policy=new_policy)['screened'][0]['holdout']['recall']==0.0
    wrong=copy.deepcopy(new);wrong['holdout']['cases'][0]['positive']=False
    with pytest.raises(ValueError):migrate_ledger(old,wrong,new_policy)
    tampered=copy.deepcopy(migrated);tampered['migration']['after']['recall']=1.0
    with pytest.raises(ValueError):evaluate([row],tampered,policy=new_policy)


def define_doc(project):
    root,_=project
    config=json.loads((root/'design/design.yaml').read_text())
    config['items']['dads-000001']={'state':'defined','evidence':['doc:DESIGN.md#dads-000001']}
    write(root/'design/design.yaml',config)
    (root/'design/DESIGN.md').write_text('# Brand\n## dads-000001\nUse token:color.primary consistently.\n')


def test_real_l0_reports_missing_judgment_and_rejects_stale_screening(project):
    root,_=project;define_doc(project)
    result=check(project)
    assert set(result['report'])=={'executed','skipped','missing','uncovered','screened'}
    assert result['judgment_candidates'] and result['report']['uncovered']
    manifest=result['manifest']['web'];item=manifest['items']['dads-000001']
    approve_rows(project,[{'key':'ds:web:item:dads-000001','value_hash':item['value_hash']},
        {'key':'ds:web:token:color.primary','value_hash':manifest['tokens']['color.primary']['value_hash']}])
    assert check(project,ship=True)['exit_code']==2
    rows=check(project)['judgment_candidates']
    write(root/'design/judgments.json',ledger_for(rows,[0.1]*len(rows)))
    result=check(project,ship=True)
    assert result['exit_code']==0,result
    assert result['report']['screened'][0]['holdout']['positives']==1
    (root/'design/DESIGN.md').write_text('# Brand\n## dads-000001\nChanged policy.\n')
    result=check(project,ship=True)
    assert result['exit_code']==2 and result['report']['uncovered']


def test_ia_input_without_installed_module_fails_visibly(project):
    root,_=project;approve(project);assert check(project,ship=True)['exit_code']==0
    assert (root/'design/.resolved/web.json').exists()
    (root/'design/info_arch').mkdir()
    result=check(project)
    assert result['exit_code']==2
    assert any(row['id']=='ia' for row in result['report']['missing'])
    assert not (root/'design/.resolved/web.json').exists()


def test_record_and_migrate_cli_write_real_ledger_and_reject_label_changes(project,tmp_path):
    root,_=project;row=example();data=ledger_for([row],[0.1])
    replacement=tmp_path/'replacement.json';write(replacement,data)
    argv=[sys.executable,str(SKILL/'scripts/judgment_ledger.py')]
    args=['--repo',str(root),'--replacement',str(replacement)]
    preview=subprocess.run(argv+['record']+args,capture_output=True,text=True)
    assert preview.returncode==0,preview.stdout+preview.stderr
    assert not (root/'design/judgments.json').exists()
    run=subprocess.run(argv+['record']+args+['--write'],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    assert evaluate([row],json.loads((root/'design/judgments.json').read_text()))['screened']
    before=(root/'design/judgments.json').read_bytes()
    data['holdout']['cases'][0]['positive']=False;write(replacement,data)
    bad=subprocess.run(argv+['migrate']+args+['--write'],capture_output=True,text=True)
    assert bad.returncode==2
    assert (root/'design/judgments.json').read_bytes()==before
    data=ledger_for([row],[0.2]);data['holdout']['cases'][0]['score']=0.4;write(replacement,data)
    run=subprocess.run(argv+['migrate']+args+['--write'],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    assert json.loads(run.stdout)['migration']['after']['recall']==0.0


def test_real_cli_resolves_installed_ia_once_and_consumers_pin_module_and_policy(project,tmp_path):
    import shutil,os
    from test_core import approve
    root,catalogs=project;approve(project)
    installed=tmp_path/'installed';copied=installed/'own-design-author'
    shutil.copytree(SKILL,copied,ignore=shutil.ignore_patterns('__pycache__'))
    implementation=installed/'own-infoarch-author/scripts/ia_module.py'
    implementation.parent.mkdir(parents=True)
    implementation.write_text("""VERSION='fixture-v1'
def resolve(*,root,approvals,profiles):
    with (root/'.ia-resolution-count').open('a') as stream:stream.write('1')
    return {'media':{media:{'messages':['fixture-message']} for media in profiles},
        'approval_hashes':{'ia:web:item:fixture-message':'a'*64},'must_approval_keys':[],
        'errors':[],'warnings':[],
        'report':{key:[] for key in ('executed','skipped','missing','uncovered','screened')}}
""")
    module_source=implementation.read_text()
    config=json.loads((root/'design/design.yaml').read_text());config['modules']={'ia':'fixture-v1'}
    write(root/'design/design.yaml',config)
    run=subprocess.run([sys.executable,str(copied/'scripts/design_author.py'),'check','--repo',str(root),
        '--catalog-dir',str(catalogs),'--include-manifest'],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    manifest=json.loads(run.stdout)['manifest']['web']
    assert manifest['modules']['ia']['resolved']['messages']==['fixture-message']
    assert (root/'.ia-resolution-count').read_text()=='1'
    path=root/'design/.resolved/web.json'
    python=[sys.executable,str(copied/'scripts/resolved_contract.py'),str(path)]
    js='import {loadManifest} from '+json.dumps((copied/'scripts/resolved_contract.mjs').as_uri())+';loadManifest(process.argv.at(-1));'
    node=[os.environ.get('DESIGN_TEST_NODE','node'),'--input-type=module','-e',js,str(path)]
    for command in [python,node]:assert subprocess.run(command,capture_output=True).returncode==0
    policy=copied/'references/judgment-policy.json';saved=policy.read_bytes();policy.write_bytes(saved+b'\n')
    for command in [python,node]:assert subprocess.run(command,capture_output=True).returncode!=0
    policy.write_bytes(saved);implementation.write_text(implementation.read_text()+'\n# changed resolver\n')
    for command in [python,node]:assert subprocess.run(command,capture_output=True).returncode!=0
    assert (root/'.ia-resolution-count').read_text()=='1'
    # Each malformed result must fail through the same real CLI, with no stale output.
    for old,new,error in [
        ("'ia:web:item:fixture-message'","'ia:'",'invalid IA approval hash contract'),
        ("'must_approval_keys':[]","'must_approval_keys':['unknown']",'invalid IA required approval keys'),
        ("{'messages':['fixture-message']}","[]",'IA media resolution mismatch')]:
        implementation.write_text(module_source.replace(old,new))
        bad=subprocess.run([sys.executable,str(copied/'scripts/design_author.py'),'check','--repo',str(root),
            '--catalog-dir',str(catalogs)],capture_output=True,text=True)
        assert bad.returncode==2,bad.stdout+bad.stderr
        assert error in json.loads(bad.stdout)['errors']
        assert not path.exists()


def test_candidate_export_is_explicit_and_normalized_span_is_stable(project):
    root,catalogs=project;define_doc(project)
    command=[sys.executable,str(SKILL/'scripts/design_author.py'),'check',
        '--repo',str(root),'--catalog-dir',str(catalogs)]
    default=subprocess.run(command,capture_output=True,text=True)
    assert 'judgment_candidates' not in json.loads(default.stdout)
    exported=subprocess.run(command+['--include-candidates'],capture_output=True,text=True)
    rows=json.loads(exported.stdout)['judgment_candidates']
    assert rows[0]['input']['text']=='Use token:color.primary consistently.'
    a=candidate('doc-relevance','fixture','Two  words.\n',{'title':'same'})
    b=candidate('doc-relevance','fixture',' Two words. ',{'title':'same'})
    assert a['judgment_hash']==b['judgment_hash'] and a['input']['text']==b['input']['text']


def test_policy_change_during_l0_is_rejected_and_old_output_removed(project,monkeypatch,tmp_path):
    import design_author,judgment_ledger
    root,_=project;approve(project);assert check(project,ship=True)['exit_code']==0
    policy=tmp_path/'policy.json';policy.write_bytes(judgment_ledger.POLICY.read_bytes())
    monkeypatch.setattr(judgment_ledger,'POLICY',policy)
    monkeypatch.setattr(design_author,'JUDGMENT_POLICY',policy)
    original=design_author.media_hashes
    def changed(*args):
        hashes=original(*args);policy.write_bytes(policy.read_bytes()+b'\n');return hashes
    monkeypatch.setattr(design_author,'media_hashes',changed)
    result=check(project)
    assert result['exit_code']==2
    assert 'judgment policy changed during validation' in result['errors']
    assert not (root/'design/.resolved/web.json').exists()
