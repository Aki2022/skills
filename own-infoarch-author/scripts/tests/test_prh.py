import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import prh  # noqa: E402

ENTITY = {'ent-0001': {'state': 'proposed', 'provenance': 'owner', 'type': 'product',
                       'name': 'Webサーバー', 'forbidden': ['ウェブサーバー', 'Webサーバ']}}

# 表記ゆれ 6 形: 大小文字・全角英字・半角カナ・空白挿入・長音の字形・別表記
SIX_FORMS = ['webサーバー', 'ＷＥＢサーバー', 'Webｻｰﾊﾞｰ', 'Web サーバー', 'Webサ−バ−', 'ウェブサーバー']


def generated_rules(entities=ENTITY):
    import yaml
    text = 'version: 1\nrules:\n' + prh.render_generated(entities) + '\n'
    return prh.load_rules(yaml.safe_load(text))


@pytest.mark.parametrize('form', SIX_FORMS)
def test_each_of_six_variant_forms_is_flagged(form):
    hits = list(prh.check_text(generated_rules(), f'当社の{form}は速い'))
    assert hits and hits[0]['expected'] == 'Webサーバー'


@pytest.mark.parametrize('text', ['当社のWebサーバーは速い', 'Webサーバーの設定', 'Websiteを公開'])
def test_canonical_name_and_longer_words_are_not_flagged(text):
    assert list(prh.check_text(generated_rules(), text)) == []


def test_forbidden_prefix_of_name_still_hits_alone():
    hits = list(prh.check_text(generated_rules(), 'Webサーバを使う'))
    assert [h['found'] for h in hits] == ['Webサーバ']


def test_literal_regex_wordboundary_and_must_empty():
    rules = prh.load_rules({'version': 1, 'rules': [
        {'expected': 'JavaScript', 'patterns': ['javascript', '/java\\s+script/i']},
        {'expected': 'Go', 'pattern': 'golang', 'options': {'wordBoundary': True}},
        {'expected': '$1データ', 'pattern': '/(ビッグ)?デ－タ/', 'regexpMustEmpty': '$1'},
    ]})
    found = [h['found'] for h in prh.check_text(rules, 'javascript と Java Script と golangs と golang')]
    assert found == ['javascript', 'Java Script', 'golang']
    assert [h['found'] for h in prh.check_text(rules, 'デ－タ ビッグデ－タ')] == ['デ－タ']


def test_specs_failures_are_reported():
    rules = prh.load_rules({'version': 1, 'rules': [
        {'expected': 'サーバー', 'pattern': 'サーバ', 'specs': [{'from': 'サーバ', 'to': 'サーバー'}, {'from': 'サーバー', 'to': 'サーバー'}]}]})
    failures = prh.run_specs(rules)
    assert len(failures) == 1 and failures[0]['got'] == 'サーバーー'


@pytest.mark.parametrize('data', [
    {'version': 1, 'imports': ['x.yml'], 'rules': []},
    {'version': 1, 'rules': [{'expected': 'A'}]},
    {'version': 1, 'rules': [{'expected': 'A', 'pattern': '/a/y'}]},
    {'version': 1, 'rules': [{'expected': 'A', 'pattern': '/(?<x>a)/'}]},
    {'version': 1, 'rules': [{'expected': 'A', 'pattern': 'a', 'prh': 'x'}]},
    {'version': 2, 'rules': []},
    {'version': 1, 'rules': [{'expected': 'A', 'pattern': 'a', 'options': {'foo': 1}}]},
])
def test_unsupported_shapes_are_rejected_not_ignored(data):
    with pytest.raises(prh.PrhError):
        prh.load_rules(data)


def test_generated_block_roundtrip_and_drift_detection():
    block = prh.render_generated(ENTITY)
    text = f'version: 1\nrules:\n{block}\n  - expected: A\n    pattern: a\n'
    assert prh.generated_block(text) == block
    edited = text.replace('ウェブサーバー', 'ウェッブサーバー')
    assert prh.generated_block(edited) != prh.render_generated(ENTITY)
