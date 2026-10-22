import sys
from decimal import Decimal
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import numeric  # noqa: E402


def one(text):
    tokens = numeric.candidate_tokens(text)
    assert len(tokens) == 1, (text, tokens)
    return tokens[0]


@pytest.mark.parametrize('text', ['1 万円', '10,000 円', '1万円', '10000円', '一万円'])
def test_yen_forms_resolve_to_same_value(text):
    tok = one(text)
    assert (tok.value, tok.unit) == ((Decimal(10000), Decimal(10000)), '円')


@pytest.mark.parametrize('text', ['１００社', '百社', '100社', '100 社'])
def test_company_forms_resolve_to_same_value(text):
    tok = one(text)
    assert (tok.value, tok.unit) == ((Decimal(100), Decimal(100)), '社')


def test_company_and_yen_are_different_units():
    assert one('100社').unit != one('100円').unit


def test_wari_is_percent():
    tok = one('9割')
    assert (tok.value, tok.unit) == ((Decimal(90), Decimal(90)), '%')


def test_bare_number_without_unit_is_still_a_candidate():
    tok = one('お客様の 9 が満足')
    assert tok.unit == ''


@pytest.mark.parametrize('text,quals', [
    ('100社以上', {'gte'}), ('約100社', {'approx'}), ('最大50%', {'max'}),
    ('月額 1,000円(税込)', {'monthly', 'tax_in'}), ('最短3日', {'min'}),
    ('100社超', {'gt'}), ('１００社＋', {'plus'}),
])
def test_qualifiers(text, quals):
    assert one(text).qualifiers == frozenset(quals)


def test_range_with_wave_dash_variants():
    for text in ['10〜20社', '10～20社', '10~20社']:
        tok = one(text)
        assert (tok.value, tok.unit) == ((Decimal(10), Decimal(20)), '社')


def test_compound_kanji_large_units():
    assert one('1億2000万円').value == (Decimal(120000000), Decimal(120000000))


@pytest.mark.parametrize('text', [
    '2026年10月22日', '2026/10/22', '令和8年度', '10月22日', '10:30', '9時から',
    '03-1234-5678', '〒100-0001', 'v1.2.3', 'STEP 1', '第3章', '3つの理由',
    '© 2026', '※1', 'https://example.com/a/123', '2026年',
])
def test_non_claim_patterns_remove_tokens(text):
    assert numeric.candidate_tokens(text) == []


@pytest.mark.parametrize('text', ['一人ひとり', '千葉県', '百貨店', '一緒に'])
def test_kanji_words_are_not_numbers(text):
    assert numeric.candidate_tokens(text) == []


@pytest.mark.parametrize('text', ['第1位', '創業30年', '3人に1人', '24時間365日', '3年間'])
def test_claims_that_look_structural_stay_candidates(text):
    assert numeric.candidate_tokens(text)


def test_dash_variants_fold_before_patterns():
    assert numeric.candidate_tokens('03−1234−5678') == []
    assert numeric.candidate_tokens('03ー1234ー5678') == []
