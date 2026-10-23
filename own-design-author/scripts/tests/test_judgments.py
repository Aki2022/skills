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


def test_at_threshold_is_advisory_and_empty_or_invalid_holdout_is_never_green():
    row=example();policy=load_policy()
    assert not evaluate([row],ledger_for([row],[policy['theta']]))['errors']
    assert evaluate([row],ledger_for([row],[policy['theta']]))['warnings']
    for cases in [[],[{'id':'one','positive':False,'score':0.1}]]:
        data=ledger_for([row],[0.1]);data['holdout']['cases']=cases
        with pytest.raises(ValueError):evaluate([row],data)
    data=ledger_for([row],[float('nan')])
    with pytest.raises(ValueError):evaluate([row],data)
    data=ledger_for([row],[0.1]);data['rows']*=2
    with pytest.raises(ValueError):evaluate([row],data)


def test_migration_requires_same_labeled_holdout_and_records_recall_delta(tmp_path):
    old_policy=load_policy();new_policy={**old_policy,'theta':0.95}
    old=current_ledger([example()],[0.1],old_policy,holdout=ledger_for([],[])['holdout'])
    row=candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Brand'},policy=new_policy)
    new=current_ledger([row],[0.1],new_policy,holdout=ledger_for([],[])['holdout'])
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
    data=ledger_for([row],[0.1]);write(replacement,data)
    run=subprocess.run(argv+['migrate']+args+['--write'],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    assert json.loads(run.stdout)['migration']['after']['recall']==1.0


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


def current_ledger(candidates,scores,policy=None,holdout=None):
    return {'schema_version':2,'policy':policy or load_policy(),'rows':[
        {'candidate':row,'score':score} for row,score in zip(candidates,scores)],
        'holdout':holdout or {'source':'unmeasured'}}


def test_unmeasured_high_scores_are_visible_advisory_without_fabricated_metrics(monkeypatch):
    row=example()
    def forbidden(*args,**kwargs):raise AssertionError('gate invoked a model')
    monkeypatch.setattr(subprocess,'run',forbidden)
    result=evaluate([row],current_ledger([row],[0.99]))
    assert not result['errors'] and not result['uncovered']
    screened=result['screened'][0]
    assert screened['flagged_advisory'] is True and screened['semantic_guarantee'] is False
    assert 'input' not in screened and screened['evaluation']=={'status':'unmeasured'}
    assert 'recall' not in json.dumps(result) and result['warnings']


def test_current_ledger_validates_candidate_identity_duplicates_and_unmeasured_shape():
    row=example();data=current_ledger([row],[0.9])
    for mutate in [lambda d:d['rows'].extend(copy.deepcopy(d['rows'])),
                   lambda d:d['rows'][0]['candidate']['input'].update(text='tampered'),
                   lambda d:d['holdout'].update(recall=1.0),
                   lambda d:d['holdout'].update(cases=[]),
                   lambda d:d['rows'][0].update(score=True)]:
        changed=copy.deepcopy(data);mutate(changed)
        with pytest.raises(ValueError):evaluate([row],changed)


def test_same_input_migration_records_actual_flags_without_fake_label_metrics():
    row=example();old_policy=load_policy();new_policy={**old_policy,'theta':0.8}
    newrow=candidate(**{'kind':row['input']['kind'],'ident':row['input']['catalog_id'],
        'text':row['input']['text'],'context':row['input']['context']},policy=new_policy)
    before=current_ledger([row],[0.7]);after=current_ledger([newrow],[0.7],new_policy)
    migrated=migrate_ledger(before,after,new_policy)
    proof=migrated['migration']
    assert proof['changed_inputs']==[{'id':row['id'],'input_hash':proof['inputs'][0]['input_hash']}]
    assert proof['inputs'][0]['before_flag'] is True and proof['inputs'][0]['after_flag'] is False
    assert proof['before']==proof['after']=={'status':'unmeasured'}
    assert 'recall' not in json.dumps(proof)
    assert evaluate([newrow],migrated,policy=new_policy)['screened']
    for mutate in [lambda d:d['migration']['inputs'][0].update(after_flag=True),
                   lambda d:d['migration']['previous_rows'].append(copy.deepcopy(d['migration']['previous_rows'][0])),
                   lambda d:d['migration'].update(changed_inputs=[]),
                   lambda d:d['migration']['previous_rows'][0]['candidate']['input'].update(text='wrong')]:
        bad=copy.deepcopy(migrated);mutate(bad)
        with pytest.raises(ValueError):evaluate([newrow],bad,policy=new_policy)
    wrong=copy.deepcopy(after);wrong['rows'][0]['candidate']=candidate('doc-relevance','dads-000002','different',{},policy=new_policy)
    with pytest.raises(ValueError,match='same input'):migrate_ledger(before,wrong,new_policy)
    with pytest.raises(ValueError,match='same input'):migrate_ledger(before,current_ledger([],[],new_policy),new_policy)


def test_legacy_owner_holdout_can_upgrade_without_losing_metrics():
    row=example();before=ledger_for([row],[0.8])
    after=current_ledger([row],[0.8],holdout=copy.deepcopy(before['holdout']))
    upgraded=migrate_ledger(before,after,load_policy())
    assert upgraded['migration']['before']['recall']==1.0
    assert upgraded['migration']['after']['recall']==1.0
    assert evaluate([row],upgraded)['screened'][0]['holdout']['recall']==1.0


def test_real_l0_high_score_does_not_block_approved_targets_but_stale_missing_block(project):
    root,_=project;define_doc(project)
    result=check(project);manifest=result['manifest']['web']
    approve_rows(project,[{'key':'ds:web:item:dads-000001','value_hash':manifest['items']['dads-000001']['value_hash']},
        {'key':'ds:web:token:color.primary','value_hash':manifest['tokens']['color.primary']['value_hash']}])
    candidates=result['judgment_candidates']
    write(root/'design/judgments.json',current_ledger(candidates,[0.99]*len(candidates)))
    result=check(project,ship=True)
    assert result['exit_code']==1 and not result['errors'],result
    assert result['warnings'] and result['report']['screened'][0]['flagged_advisory']
    assert (root/'design/.resolved/web.json').exists()
    assert 'recall' not in json.dumps(result)
    ledger=current_ledger([],[]);write(root/'design/judgments.json',ledger)
    assert check(project,ship=True)['exit_code']==2
    write(root/'design/judgments.json',current_ledger(candidates,[0.1]*len(candidates)))
    (root/'design/DESIGN.md').write_text('# Brand\n## dads-000001\nOther current wording.\n')
    result=check(project,ship=True)
    assert result['exit_code']==2 and result['report']['uncovered']
    assert not (root/'design/.resolved/web.json').exists()


def test_real_l0_ship_rejects_missing_optional_item_judgment(project):
    from test_core import item
    root,catalogs=project
    catalog=json.loads((catalogs/'test-v1.json').read_text());catalog['items'].append(item('dads-000002','nice'))
    write(catalogs/'test-v1.json',catalog)
    ledger=json.loads((catalogs/'id-ledger.json').read_text());ledger['next_id']=3
    ledger['entries'].append({'id':'dads-000002','match_key':'dads-000002','retired':False});write(catalogs/'id-ledger.json',ledger)
    config=json.loads((root/'design/design.yaml').read_text());config['items']['dads-000002']={'state':'defined','evidence':['doc:DESIGN.md#dads-000002']}
    write(root/'design/design.yaml',config)
    (root/'design/DESIGN.md').write_text('# Brand\n## dads-000002\nA specific optional policy.\n')
    approve(project)
    result=check(project,ship=True)
    assert result['exit_code']==2 and result['report']['uncovered'],result


def test_migration_keeps_existing_owner_labels_and_strictly_validates_flag_types():
    row=example();old=ledger_for([row],[0.9]);new=current_ledger([row],[0.9])
    with pytest.raises(ValueError,match='holdout'):migrate_ledger(old,new,load_policy())
    new['holdout']=copy.deepcopy(old['holdout']);migrated=migrate_ledger(old,new,load_policy())
    migrated['migration']['inputs'][0]['before_flag']=1
    with pytest.raises(ValueError,match='mismatch'):evaluate([row],migrated)


@pytest.mark.parametrize('kind',['doc-relevance','na-rationale'])
def test_both_semantic_kinds_are_advisory_and_cli_records_unmeasured(kind,project,tmp_path):
    root,_=project;row=candidate(kind,'dads-000001','A retained source statement.',{'title':'Brand'})
    data=current_ledger([row],[1.0]);replacement=tmp_path/'replacement.json';write(replacement,data)
    argv=[sys.executable,str(SKILL/'scripts/judgment_ledger.py'),'record','--repo',str(root),
        '--replacement',str(replacement),'--write']
    run=subprocess.run(argv,capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    assert json.loads(run.stdout)['evaluation']=={'status':'unmeasured'}
    assert 'recall' not in run.stdout
    evaluated=evaluate([row],json.loads((root/'design/judgments.json').read_text()))
    assert evaluated['warnings'] and not evaluated['errors']


def test_policy_change_requires_legacy_identity_upgrade_first():
    old=ledger_for([example()],[0.1]);newpolicy={**load_policy(),'theta':0.7}
    newrow=candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Brand'},policy=newpolicy)
    with pytest.raises(ValueError,match='upgrading legacy'):
        migrate_ledger(old,current_ledger([newrow],[0.1],newpolicy,holdout=old['holdout']),newpolicy)


def test_original_schema1_migration_proof_remains_readable_and_tamper_rejected():
    from judgment_ledger import holdout_metrics
    data=ledger_for([example()],[0.1]);prior_policy={**load_policy(),'theta':0.95}
    prior_holdout=copy.deepcopy(data['holdout'])
    data['migration']={'previous_policy':prior_policy,'previous_holdout':prior_holdout,
        'before':holdout_metrics(prior_holdout,prior_policy['theta']),
        'after':holdout_metrics(data['holdout'],data['policy']['theta']),
        'changed':['fixture-positive']}
    assert evaluate([example()],data)['screened'][0]['holdout']['recall']==1.0
    bad=copy.deepcopy(data);bad['migration']['after']['recall']=0.0
    with pytest.raises(ValueError):evaluate([example()],bad)
    bad=current_ledger([example()],[0.1],holdout=data['holdout']);bad['migration']=data['migration']
    with pytest.raises(ValueError):evaluate([example()],bad)


def test_refresh_preserves_scores_labels_and_verifies_added_removed_input_sets():
    from judgment_ledger import refresh_ledger
    retained=example();removed=candidate('na-rationale','dads-000002','Old rationale.',{})
    added=candidate('na-rationale','dads-000002','New rationale.',{})
    before=current_ledger([retained,removed],[0.8,0.6]);after=current_ledger([retained,added],[0.8,0.9])
    refreshed=refresh_ledger(before,after,load_policy())
    proof=refreshed['refresh']
    assert proof['added']==[added['judgment_hash']] and proof['removed']==[removed['judgment_hash']]
    assert proof['retained']==[retained['judgment_hash']]
    assert evaluate([retained,added],refreshed)['screened']
    assert 'recall' not in json.dumps(proof)
    for mutate in [lambda d:d['refresh'].update(added=[]),
                   lambda d:d['refresh']['previous_rows'][0].update(score=0.7),
                   lambda d:d['refresh']['previous_rows'].append(copy.deepcopy(d['refresh']['previous_rows'][0]))]:
        bad=copy.deepcopy(refreshed);mutate(bad)
        with pytest.raises(ValueError):evaluate([retained,added],bad)
    bad=copy.deepcopy(after);bad['rows'][0]['score']=0.1
    with pytest.raises(ValueError,match='unchanged'):refresh_ledger(before,bad,load_policy())
    policy={**load_policy(),'theta':0.7};target=candidate('doc-relevance','dads-000001','A clear policy.',{'title':'Brand'},policy=policy)
    with pytest.raises(ValueError,match='policy'):refresh_ledger(before,current_ledger([target],[0.8],policy),policy)
    labeled=current_ledger([retained,removed],[0.8,0.6],holdout=ledger_for([],[])['holdout'])
    with pytest.raises(ValueError,match='holdout'):refresh_ledger(labeled,after,load_policy())
    after['holdout']=copy.deepcopy(labeled['holdout']);after['holdout']['cases'][0]['score']=0.4
    with pytest.raises(ValueError,match='holdout'):refresh_ledger(labeled,after,load_policy())


def test_refresh_cli_records_changed_source_inputs_then_real_l0_uses_them(project,tmp_path):
    from judgment_ledger import refresh_ledger
    root,_=project;define_doc(project);oldrows=check(project)['judgment_candidates']
    before=current_ledger(oldrows,[0.8]*len(oldrows));write(root/'design/judgments.json',before)
    (root/'design/DESIGN.md').write_text('# Brand\n## dads-000001\nAn explicitly updated policy.\n')
    current=check(project);newrows=current['judgment_candidates']
    replacement=tmp_path/'refresh.json';write(replacement,current_ledger(newrows,[0.9]*len(newrows)))
    argv=[sys.executable,str(SKILL/'scripts/judgment_ledger.py'),'refresh','--repo',str(root),
        '--replacement',str(replacement)]
    preview=subprocess.run(argv,capture_output=True,text=True)
    assert preview.returncode==0,preview.stdout+preview.stderr
    assert json.loads((root/'design/judgments.json').read_text())==before
    run=subprocess.run(argv+['--write'],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    recorded=json.loads((root/'design/judgments.json').read_text())
    assert recorded['refresh']['removed']==[oldrows[0]['judgment_hash']]
    manifest=current['manifest']['web']
    approve_rows(project,[{'key':'ds:web:item:dads-000001','value_hash':manifest['items']['dads-000001']['value_hash']},
        {'key':'ds:web:token:color.primary','value_hash':manifest['tokens']['color.primary']['value_hash']}])
    result=check(project,ship=True)
    assert result['exit_code']==1 and not result['errors'],result
    assert result['report']['screened'][0]['judgment_hash']==newrows[0]['judgment_hash']
    assert result['report']['screened'][0]['flagged_advisory'] is True


def test_same_policy_migrate_cannot_bypass_refresh_score_and_holdout_retention():
    row=example();before=current_ledger([row],[0.9]);after=current_ledger([row],[0.1])
    with pytest.raises(ValueError,match='unchanged'):migrate_ledger(before,after,load_policy())
    after=current_ledger([row],[0.9],holdout=ledger_for([],[])['holdout'])
    with pytest.raises(ValueError,match='holdout'):migrate_ledger(before,after,load_policy())
    with pytest.raises(ValueError,match='holdout'):
        from judgment_ledger import refresh_ledger
        refresh_ledger(before,after,load_policy())
