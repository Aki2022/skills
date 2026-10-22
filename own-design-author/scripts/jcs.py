"""Batch RFC 8785 canonicalization using the existing ECMAScript runtime.

No install/network fallback. These helpers do not write approvals or change trust.
"""
import hashlib
import json
import math
import os
from pathlib import Path
import subprocess


def _validate(value):
    if value is None or isinstance(value,(str,bool)):
        return
    if isinstance(value,int):
        try:
            if int(float(value)) != value:
                raise ValueError('integer loses IEEE 754 precision; encode as string')
        except OverflowError as exc:
            raise ValueError('integer outside IEEE 754 range') from exc
        return
    if isinstance(value,float):
        if not math.isfinite(value):
            raise ValueError('non-finite JSON number')
        return
    if isinstance(value,list):
        for child in value:_validate(child)
        return
    if isinstance(value,dict) and all(isinstance(key,str) for key in value):
        for child in value.values():_validate(child)
        return
    raise ValueError('not I-JSON data')


def canonicalize_batch(values):
    if not isinstance(values,list):
        raise ValueError('canonicalization requires a batch')
    _validate(values)
    try:
        run=subprocess.run([os.environ.get('DESIGN_AUTHOR_NODE','node'),str(Path(__file__).with_suffix('.cjs'))],
            input=json.dumps(values,ensure_ascii=True,allow_nan=False),text=True,
            capture_output=True,timeout=30,check=False)
        if run.returncode:
            raise ValueError('JCS canonicalization failed')
        result=json.loads(run.stdout)
        if not isinstance(result,list) or len(result)!=len(values) or any(not isinstance(x,str) for x in result):
            raise ValueError('invalid JCS runtime result')
        return [value.encode('utf-8') for value in result]
    except (OSError,subprocess.TimeoutExpired,UnicodeError,json.JSONDecodeError) as exc:
        raise ValueError('JCS runtime unavailable or invalid input') from exc


def media_hashes(media,profile,resolved):
    if not isinstance(media,str) or not media or not isinstance(profile,dict) or not isinstance(resolved,dict):
        raise ValueError('invalid media hash context')
    keys=sorted(resolved)
    content=canonicalize_batch([{'media':media,'profile':profile,'resolved':resolved[key]} for key in keys])
    return {key:hashlib.sha256(value).hexdigest() for key,value in zip(keys,content)}
