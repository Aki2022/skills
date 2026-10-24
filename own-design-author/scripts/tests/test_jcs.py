"""Canonicalization vectors and input rejection; no production signing keys."""
import hashlib
import math
from pathlib import Path
import sys
import pytest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from jcs import canonicalize_batch, media_hashes


def test_ecmascript_numbers_utf16_sort_and_numeric_property_names():
    values=[{'numbers':[333333333.33333329,1e30,4.50,2e-3,1e-27,-0.0]},
            {'2':'two','10':'ten','😀':'astral','דּ':'hebrew'},
            {'z':{'b':True,'a':None},'a':[False,'é','e\u0301']}]
    result=canonicalize_batch(values)
    assert result[0]==b'{"numbers":[333333333.3333333,1e+30,4.5,0.002,1e-27,0]}'
    assert result[1]=='{"10":"ten","2":"two","😀":"astral","דּ":"hebrew"}'.encode()
    assert result[2]=='{"a":[false,"é","e\u0301"],"z":{"a":null,"b":true}}'.encode()


@pytest.mark.parametrize('value',[float('nan'),float('inf'),9007199254740993,'\ud800',{'\udfff':1},{1:'bad-key'}])
def test_invalid_i_json_is_rejected(value):
    with pytest.raises(ValueError):canonicalize_batch([value])


def test_media_hash_binds_profile_and_alias_resolution():
    values={'item':{'token:alias':{'type':'number','value':1}}}
    original=media_hashes('web',{'minimums':{'text_contrast':4.5}},values)
    assert len(original['item'])==64
    assert original!=media_hashes('web',{'minimums':{'text_contrast':7}},values)
    assert original!=media_hashes('web',{'minimums':{'text_contrast':4.5}},
        {'item':{'token:alias':{'type':'number','value':2}}})
    expected=canonicalize_batch([{'media':'web','profile':{'minimums':{'text_contrast':4.5}},'resolved':values['item']}])[0]
    assert original['item']==hashlib.sha256(expected).hexdigest()
