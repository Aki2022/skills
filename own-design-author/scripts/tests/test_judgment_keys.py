"""Judgment identity is per kind (SPEC-information-architecture-standardization v5, requirement 19).

Adding a kind or rewording one kind's prompt must not invalidate the judgments of
the other kinds; a shared policy value (model, theta, prompt_version) still does.
"""
import copy
import subprocess
import pytest
from test_core import project  # noqa: F401 -- fixture; importing test_core also puts scripts/ on sys.path
from judgment_ledger import KINDS,candidate,evaluate,load_policy,rekey_ledger,validate_ledger

TEXT='新規事業の人材を、学術尺度で見極める。'
CASES=[(kind,'id-'+kind,TEXT,{'brief':kind}) for kind in sorted(KINDS)]


def without(policy,kind):
    reduced=copy.deepcopy(policy);reduced['prompts'].pop(kind);return reduced


def reworded(policy,kind):
    changed=copy.deepcopy(policy);changed['prompts'][kind]+=' Reworded.';return changed


def ledger(rows,scores,policy,version=3):
    return {'schema_version':version,'policy':policy,'rows':[
        {'candidate':row,'score':score} for row,score in zip(rows,scores)],'holdout':{'source':'unmeasured'}}


def test_strategy_fit_kind_is_part_of_the_shared_policy():
    assert 'ia-strategy-fit' in KINDS
    assert load_policy()['prompts']['ia-strategy-fit'].strip()


def test_adding_a_kind_keeps_every_other_key():
    policy=load_policy();before=without(policy,'ia-strategy-fit')
    for kind,ident,text,context in CASES:
        if kind=='ia-strategy-fit':continue
        assert (candidate(kind,ident,text,context,policy=before)['judgment_hash']
                ==candidate(kind,ident,text,context,policy=policy)['judgment_hash']),kind


def test_ledger_recorded_before_a_kind_was_added_still_screens():
    policy=load_policy();before=without(policy,'ia-strategy-fit')
    rows=[candidate(*case,policy=before) for case in CASES if case[0]!='ia-strategy-fit']
    current=[candidate(*case) for case in CASES if case[0]!='ia-strategy-fit']
    result=evaluate(current,ledger(rows,[0.1]*len(rows),before))
    assert not result['uncovered'] and len(result['screened'])==len(rows)


def test_rewording_one_prompt_invalidates_only_that_kind():
    policy=load_policy();changed=reworded(policy,'ia-emphasis')
    rows=[candidate(*case,policy=policy) for case in CASES]
    for kind,ident,text,context in CASES:
        moved=(candidate(kind,ident,text,context,policy=changed)['judgment_hash']
               !=candidate(kind,ident,text,context,policy=policy)['judgment_hash'])
        assert moved==(kind=='ia-emphasis'),kind
    current=[candidate(*case,policy=changed) for case in CASES]
    result=evaluate(current,ledger(rows,[0.1]*len(rows),policy),policy=changed)
    assert [r['id'] for r in result['uncovered']]==['ia-emphasis:id-ia-emphasis']
    assert len(result['screened'])==len(CASES)-1


@pytest.mark.parametrize('field,value',[('theta',0.61),('model','jev-other'),('prompt_version','other')])
def test_a_shared_policy_value_still_requires_migration(field,value):
    policy=load_policy();changed={**policy,field:value}
    rows=[candidate(*case,policy=policy) for case in CASES]
    with pytest.raises(ValueError,match='policy'):
        evaluate([candidate(*CASES[0],policy=changed)],ledger(rows,[0.1]*len(rows),policy),policy=changed)


def test_a_row_inconsistent_with_its_own_ledger_policy_is_rejected_not_skipped():
    policy=load_policy();rows=[candidate(*case,policy=policy) for case in CASES]
    forged=ledger(rows,[0.1]*len(rows),reworded(policy,'doc-relevance'))
    with pytest.raises(ValueError):validate_ledger(forged,policy)


def test_schema2_ledger_is_rekeyed_once_without_rescoring(monkeypatch):
    def forbidden(*args,**kwargs):raise AssertionError('rekey called a model')
    monkeypatch.setattr(subprocess,'run',forbidden)
    policy=load_policy();before=without(policy,'ia-strategy-fit')
    old=[candidate(*case,policy=before,scheme=2) for case in CASES if case[0]!='ia-strategy-fit']
    scores=[0.1+0.01*i for i in range(len(old))]
    previous=ledger(old,scores,before,version=2)
    validate_ledger(previous,before)
    migrated=rekey_ledger(previous,policy)
    assert migrated['schema_version']==3 and migrated['migration']['changed_inputs']==[]
    assert [r['score'] for r in migrated['rows']]==scores
    current=[candidate(*case) for case in CASES if case[0]!='ia-strategy-fit']
    result=evaluate(current,migrated)
    assert not result['uncovered'] and len(result['screened'])==len(old)


def test_rekey_refuses_when_a_used_prompt_changed():
    policy=load_policy();before=reworded(without(policy,'ia-strategy-fit'),'doc-relevance')
    old=[candidate(*case,policy=before,scheme=2) for case in CASES if case[0]=='doc-relevance']
    with pytest.raises(ValueError,match='rekey'):
        rekey_ledger(ledger(old,[0.1],before,version=2),policy)


def test_rekey_cli_moves_a_real_schema2_ledger_once(project):
    import json,sys
    from test_core import SKILL,write
    root,_=project;policy=load_policy();before=without(policy,'ia-strategy-fit')
    old=[candidate(*CASES[0],policy=before,scheme=2)]
    write(root/'design/judgments.json',ledger(old,[0.2],before,version=2))
    argv=[sys.executable,str(SKILL/'scripts/judgment_ledger.py'),'rekey','--repo',str(root)]
    with_replacement=subprocess.run(argv+['--replacement','x.json'],capture_output=True,text=True)
    assert with_replacement.returncode!=0
    preview=subprocess.run(argv,capture_output=True,text=True)
    assert preview.returncode==0,preview.stdout+preview.stderr
    assert json.loads((root/'design/judgments.json').read_text())['schema_version']==2
    run=subprocess.run(argv+['--write'],capture_output=True,text=True)
    assert run.returncode==0,run.stdout+run.stderr
    data=json.loads((root/'design/judgments.json').read_text())
    assert data['schema_version']==3 and data['rows'][0]['score']==0.2
    assert evaluate([candidate(*CASES[0])],data)['screened']
    again=subprocess.run(argv+['--write'],capture_output=True,text=True)
    assert again.returncode==2
