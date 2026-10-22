"""Namespaced SSH signatures. Real trust/key provisioning is an owner operation."""
from datetime import date
import hashlib
import os
from pathlib import Path
import re
import subprocess
import tempfile
from jcs import canonicalize_batch


def repository_id(root):
    run=subprocess.run(['git','-C',str(root),'remote','get-url','origin'],capture_output=True,text=True)
    if run.returncode or not run.stdout.strip():
        raise ValueError('signed approval requires an origin repository identity')
    # Only the digest is emitted; remote URLs can contain private locations.
    return hashlib.sha256(run.stdout.strip().encode()).hexdigest()


def payload(repo_id,row):
    return canonicalize_batch([{'repo_id':repo_id,'key':row['key'],
        'value_hash':row['value_hash'],'date':row['date']}])[0]


def namespace(key):
    match=re.fullmatch(r'(ds|ia):([a-z][a-z0-9-]*):(item|token):(.+)',key)
    if not match:
        raise ValueError('invalid approval key namespace')
    return match[1]+'-approval'


def _store_metadata(path,root):
    """Reject same-user, writable or replaceable trust files, including parents."""
    if not path:
        raise ValueError('owner-controlled allowed_signers is not configured')
    path=Path(path)
    if not path.is_absolute() or path.is_symlink() or not path.is_file():
        raise ValueError('allowed_signers must be an external regular file')
    resolved=path.resolve()
    if resolved.is_relative_to(Path(root).resolve()):
        raise ValueError('allowed_signers must be outside the repository')
    for node in (resolved,*resolved.parents):
        info=node.stat()
        if info.st_uid==os.geteuid() or os.access(node,os.W_OK):
            raise ValueError('allowed_signers or its parent can be changed by the agent account')
    return resolved



def protected_store(path,root):
    try:return _store_metadata(path,root)
    except OSError as exc:raise ValueError('allowed_signers metadata unavailable') from exc


def verify_signature(repo_id,row,signers):
    """Cryptographic primitive; callers must validate the trust store first."""
    with tempfile.TemporaryDirectory(prefix='design-signature-') as directory:
        signature=Path(directory)/'signature'
        signature.write_text(row['signature'])
        try:
            run=subprocess.run(['ssh-keygen','-Y','verify','-f',str(signers),'-I',row['approver'],
                '-n',namespace(row['key']),'-s',str(signature)],input=payload(repo_id,row),
                capture_output=True,timeout=30)
        except (OSError,subprocess.TimeoutExpired) as exc:
            raise ValueError('SSH signature verifier unavailable') from exc
    if run.returncode:
        raise ValueError('approval signature invalid or signer unregistered')


def verify_rows(rows,root,*,with_context=False):
    if not isinstance(rows,list):raise ValueError('approval ledger must be a list')
    if not rows:return ({},None) if with_context else {}
    for row in rows:
        if not isinstance(row,dict) or set(row)!={'key','value_hash','approver','date','signature'}:
            raise ValueError('signed approval fields required; unsigned legacy rows are rejected')
        if any(not isinstance(row[k],str) or not row[k].strip() for k in row):
            raise ValueError('invalid signed approval row')
        if not re.fullmatch('[0-9a-f]{64}',row['value_hash']) or not re.fullmatch(r'[A-Za-z0-9_.@+-]+',row['approver']):
            raise ValueError('invalid approval hash or principal')
        if date.fromisoformat(row['date']).isoformat()!=row['date']:
            raise ValueError('approval date must be ISO date')
        namespace(row['key'])
        if not row['signature'].startswith('-----BEGIN SSH SIGNATURE-----\n'):
            raise ValueError('SSH signature required')
    signers=protected_store(os.environ.get('DESIGN_ALLOWED_SIGNERS'),root)
    identity=repository_id(root)
    trust_digest=hashlib.sha256(signers.read_bytes()).hexdigest()
    approved={}
    for row in rows:
        if row['key'] in approved:raise ValueError('duplicate approval target')
        verify_signature(identity,row,signers)
        approved[row['key']]=row['value_hash']
    if hashlib.sha256(signers.read_bytes()).hexdigest()!=trust_digest:
        raise ValueError('allowed_signers changed during verification')
    context={'repo_id':identity,'trust_sha256':trust_digest}
    return (approved,context) if with_context else approved
