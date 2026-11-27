"""The v6 answer key's labels carry over to the script's pairs (SPEC-information-architecture-standardization
requirement 6/migration: a pair carries over when the copy text and the value are the same). The judgment id changed
from `<outline>/sec-…/cpy-…@<value>` to `<script>/blk-…/cpy-…@<value>`, so the old labels name pairs the new key does not
know: measuring with them stops ("unknown pairs"). The carry-over rewrites them for the new key; it never invents a
label, never carries a conflicting one, and says what it left out."""
import copy
import json
from pathlib import Path
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import answer_key  # noqa: E402
from test_answer_key import items  # noqa: E402
from test_check import OUTLINE  # noqa: E402


def old_key_of(new_items, rename=lambda i: i.replace('/blk-', '/sec-')):
    """A v6-shaped key: the same pairs under the old ids (sec-), as `key_file` writes them."""
    rows = [{**{k: i[k] for k in ('kind', 'text', 'value')}, 'id': rename(i['id']), 'judgment_hash': 'old-' + i['candidate']['judgment_hash'][:8]}
            for i in new_items]
    return {'answer_key': 'a' * 64, 'items': rows}


def labels_for(old_key, pick):
    return {'answer_key': old_key['answer_key'], 'labels': {r['id']: pick(r) for r in old_key['items']}}


def test_labels_follow_the_pair_to_its_new_id():
    new = items()
    old = old_key_of(new)
    labels = labels_for(old, lambda r: r['kind'] == 'own')
    out = answer_key.carry_over(old, labels, new)
    assert out['labels']['answer_key'] == answer_key.key_hash(new)
    assert set(out['labels']['labels']) == {i['id'] for i in new}
    assert all(out['labels']['labels'][i['id']] is (i['kind'] == 'own') for i in new)
    assert out['carried'] == len(new) and out['dropped'] == [] and out['unmatched'] == []


def test_the_carried_labels_measure_against_the_new_key():
    new = items()
    labels = answer_key.carry_over(old_key_of(new), labels_for(old_key_of(new), lambda r: r['kind'] == 'own'), new)['labels']
    ledger = {'rows': []}
    result = answer_key.measure(new, labels, ledger, 0.5, None)  # a key it was made for: no "another answer key" error
    assert result['answer_key'] == answer_key.key_hash(new)


def test_a_pair_whose_copy_changed_is_not_carried():
    new = items()
    old = old_key_of(new)
    changed = old['items'][0]['id']
    old['items'][0] = {**old['items'][0], 'text': old['items'][0]['text'] + '（旧い文）'}
    out = answer_key.carry_over(old, labels_for(old, lambda r: True), new)
    assert out['carried'] == len(new) - 1 and out['dropped'] == [changed]
    assert len(out['unmatched']) == 1 and out['labels']['labels'][out['unmatched'][0]] is None  # not guessed: unknown


def test_a_pair_whose_value_text_changed_is_not_carried():
    new = items()
    old = old_key_of(new)
    old['items'][0] = {**old['items'][0], 'value': {**old['items'][0]['value'], 'text': '別の提供価値の文'}}
    out = answer_key.carry_over(old, labels_for(old, lambda r: True), new)
    assert out['carried'] == len(new) - 1 and len(out['dropped']) == 1


def test_an_unknown_label_stays_unknown():
    new = items()
    old = old_key_of(new)
    out = answer_key.carry_over(old, labels_for(old, lambda r: None), new)
    assert set(out['labels']['labels'].values()) == {None} and out['carried'] == len(new)


def test_conflicting_old_labels_for_the_same_pair_are_not_carried():
    new = items()
    old = old_key_of(new)
    twin = {**old['items'][0], 'id': old['items'][0]['id'] + '-twin'}  # the same text and value under another old id
    old['items'].append(twin)
    labels = labels_for(old, lambda r: True)
    labels['labels'][twin['id']] = False
    out = answer_key.carry_over(old, labels, new)
    assert old['items'][0]['id'] in out['dropped'] and twin['id'] in out['dropped']
    assert out['labels']['labels'][next(i['id'] for i in new if i['id'].replace('/blk-', '/sec-') == old['items'][0]['id'])] is None


def test_labels_made_for_another_key_are_refused():
    new = items()
    old = old_key_of(new)
    labels = labels_for(old, lambda r: True)
    labels['answer_key'] = 'b' * 64
    with pytest.raises(ValueError):
        answer_key.carry_over(old, labels, new)


def test_labels_naming_a_pair_the_old_key_never_had_are_refused():
    new = items()
    old = old_key_of(new)
    labels = labels_for(old, lambda r: True)
    labels['labels']['gone/sec-0009/cpy-9999@val-0001'] = True
    with pytest.raises(ValueError):
        answer_key.carry_over(old, labels, new)
