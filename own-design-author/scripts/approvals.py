"""Validate human approval records; authenticity is an operational boundary."""
from datetime import date
import hashlib
import re
import subprocess

SCHEME = 'hash-jcs-v1'


def repository_id(root):
    run=subprocess.run(['git','-C',str(root),'remote','get-url','origin'],capture_output=True,text=True)
    if run.returncode or not run.stdout.strip():
        raise ValueError('approval request requires an origin repository identity')
    # Bind review requests to a repository without emitting its potentially private URL.
    return hashlib.sha256(run.stdout.strip().encode()).hexdigest()


def verify_rows(rows,root):
    if not isinstance(rows,list):raise ValueError('approval ledger must be a list')
    approved={}
    for row in rows:
        if not isinstance(row,dict) or set(row)!={'key','value_hash','approver','date'}:
            raise ValueError('approval fields required: key, value_hash, approver, date')
        if any(not isinstance(row[k],str) or not row[k].strip() for k in row):
            raise ValueError('invalid approval row')
        if not re.fullmatch(r'(ds:[a-z][a-z0-9-]*:(item|token):.+|ia:[a-z][a-z0-9-]*:[A-Za-z0-9_.:/-]+)',row['key']):
            raise ValueError('invalid approval key namespace')
        if not re.fullmatch('[0-9a-f]{64}',row['value_hash']):
            raise ValueError('invalid approval hash')
        if not re.fullmatch(r'[A-Za-z0-9_.@+-]+',row['approver']):
            raise ValueError('invalid owner principal')
        if date.fromisoformat(row['date']).isoformat()!=row['date']:
            raise ValueError('approval date must be ISO date')
        if row['key'] in approved:raise ValueError('duplicate approval target')
        approved[row['key']]=row['value_hash']
    return approved
