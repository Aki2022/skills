#!/usr/bin/env python3
"""Create/check/migrate a repository's sole design/ directory. Owner approve command is the only ledger writer."""
import argparse
import copy
from datetime import date, datetime, timezone
import hashlib
import json
import math
from pathlib import Path
import re
import shutil
import subprocess
import sys
import unicodedata

from dtcg import resolve_tokens, value_hash
from classify import ClassificationError, classify_response, load_policy
from approvals import SCHEME,verify_rows
from jcs import media_hashes
from judgment_ledger import candidate as judgment_candidate,evaluate as evaluate_judgments,read_policy as read_judgment_policy,POLICY as JUDGMENT_POLICY
from reporting import empty_report,stage,skip
from design_modules import resolve_ia

VERSION = '1.0.0'
SKILL = Path(__file__).resolve().parents[1]
CATALOG = SKILL / 'catalog'


def _unique(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f'duplicate key: {key}')
        result[key] = value
    return result


def read_data(path):
    text = path.read_text()
    try:
        return json.loads(text, object_pairs_hook=_unique, parse_constant=lambda v: (_ for _ in ()).throw(ValueError(v)))
    except json.JSONDecodeError:
        try:
            import yaml
        except ImportError as exc:
            raise ValueError('YAML requires an existing Python environment with PyYAML') from exc
        class UniqueLoader(yaml.SafeLoader):
            pass
        def construct(loader, node):
            return _unique([(loader.construct_object(k), loader.construct_object(v)) for k, v in node.value])
        UniqueLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, construct)
        try:
            return yaml.load(text, Loader=UniqueLoader)
        except yaml.YAMLError as exc:
            raise ValueError('malformed YAML') from exc


def write_data(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, sort_keys=True, indent=2, allow_nan=False) + '\n')
    temporary.replace(path)


def inside(root, relative):
    path = (root / relative).resolve()
    if not path.is_relative_to(root.resolve()):
        raise ValueError(f'path outside design directory: {relative}')
    return path


def load_catalog(version, catalog_dir):
    if not isinstance(version, str) or not re.fullmatch(r'[a-zA-Z0-9][a-zA-Z0-9.-]*', version):
        raise ValueError('unknown catalog version')
    path = Path(catalog_dir) / f'{version}.json'
    if not path.is_file():
        raise ValueError('unknown catalog version; migrate explicitly')
    catalog = read_data(path)
    if catalog.get('schema_version') != 1 or catalog.get('version') != version:
        raise ValueError('catalog schema/version mismatch')
    items = catalog.get('items')
    if not isinstance(items, list) or not items:
        raise ValueError('zero catalog targets')
    identifiers = [x['id'] for x in items]
    if len(set(identifiers)) != len(identifiers):
        raise ValueError('duplicate catalog ids')
    ledger = read_data(Path(catalog_dir) / 'id-ledger.json')
    entries = {entry['id']: entry for entry in ledger['entries']}
    if len(entries) != len(ledger['entries']):
        raise ValueError('duplicate id ledger entries')
    for item in items:
        entry = entries.get(item['id'])
        if entry is None or item['match_key'] not in [entry['match_key'], *entry.get('previous_match_keys', [])]:
            raise ValueError(f'catalog id ledger mismatch: {item["id"]}')
    return catalog


def _presets():
    return read_data(SKILL / 'references/presets.json')


def init_design(root, preset, version, *, catalog_dir=CATALOG):
    root = Path(root)
    presets = _presets()
    if preset not in presets or preset == 'schema_version':
        raise ValueError('unknown preset')
    load_catalog(version, catalog_dir)
    design = root / 'design'
    if design.exists():
        raise ValueError('design/ already exists; use migrate')
    shutil.copytree(SKILL / 'templates/design', design)
    (design / 'assets').mkdir(exist_ok=True)
    (design / 'checks').mkdir(exist_ok=True)
    write_data(design / 'design.yaml', {'schema_version': 1, 'skill_version': VERSION,
        'catalog_version': version, 'preset': preset, 'media': presets[preset]['media'], 'items': {}})
    return {'created': 'design/', 'catalog_version': version, 'preset': preset}


def migrate_design(root, version, *, catalog_dir=CATALOG):
    design = Path(root) / 'design'
    config = read_data(design / 'design.yaml')
    old = load_catalog(config['catalog_version'], catalog_dir)
    new = load_catalog(version, catalog_dir)
    old_ids = {x['id'] for x in old['items']}
    new_ids = {x['id'] for x in new['items']}
    removed = old_ids - new_ids
    # Preserve retired items outside the active overrides so L0 can reject unknown ids.
    retired = copy.deepcopy(config.get('retired_items', {}))
    overrides = copy.deepcopy(config.get('items', {}))
    for ident in removed:
        retired[ident] = {**overrides.pop(ident, {'state':'deferred'}), 'catalog_version':old['version']}
    for ident in new_ids & set(retired):
        restored = retired.pop(ident)
        restored.pop('catalog_version', None)
        overrides.setdefault(ident, restored)
    config.update(catalog_version=version, items=overrides)
    if retired:
        config['retired_items'] = retired
    else:
        config.pop('retired_items', None)
    report = {'from': old['version'], 'to': version, 'added': sorted(new_ids-old_ids),
              'removed': sorted(removed), 'retained': sorted(old_ids & new_ids)}
    write_data(design / 'migration.json', report)
    write_data(design / 'design.yaml', config)
    # No stale generated contract survives migration, even before the next L0.
    _clear_manifests(design)
    return report


def _clear_manifests(design):
    assets_index = design / 'assets.json'
    if assets_index.is_file() or assets_index.is_symlink():
        assets_index.unlink()
    resolved = design / '.resolved'
    if resolved.is_symlink():
        resolved.unlink()
        raise ValueError('.resolved symlink was unlinked; its target was not modified')
    if resolved.is_dir():
        for path in resolved.iterdir():
            if path.is_file() and path.suffix in ('.json', '.css', '.cjs'):
                path.unlink()


def _anchor_text(path, anchor):
    text = path.read_text()
    lines = text.splitlines()
    headings, fence = [], None
    for i, line in enumerate(lines):
        if fence:
            if re.fullmatch(r' {0,3}' + re.escape(fence[0]) + '{' + str(fence[1]) + r',}\s*', line):
                fence = None
            continue
        opening = re.match(r'^ {0,3}(`{3,}|~{3,})(.*)$', line)
        if opening and (opening[1][0] != '`' or '`' not in opening[2]):
            fence = (opening[1][0], len(opening[1]))
            continue
        match = re.match(r'^(#{1,6})\s+(.+?)\s*#*$', line)
        if not match:
            continue
        heading = match[2]
        slug = re.sub(r'[^\w -]', '', unicodedata.normalize('NFKC', heading).lower()).replace(' ', '-')
        explicit = re.search(r'\{#([^}]+)\}$', heading)
        headings.append((i, len(match[1]), explicit[1] if explicit else slug))
    matches = [h for h in headings if h[2] == anchor]
    if len(matches) > 1:
        raise ValueError('duplicate DESIGN.md evidence anchor')
    if not matches:
        return ''
    start, depth, _ = matches[0]
    end = next((i for i, level, _ in headings if i > start and level <= depth), len(lines))
    return '\n'.join(lines[start+1:end]).strip()


def _rules(design):
    rules = {}
    allowed = {'id', 'kind', 'targets', 'allowlist', 'minimum', 'paths', 'pattern'}
    kinds = {'contrast', 'target-size', 'focus', 'source-hash', 'palette', 'source-pattern'}
    for path in sorted((design / 'checks').glob('*.yaml')):
        inside(design, str(path.relative_to(design)))
        data = read_data(path)
        if not isinstance(data, dict) or set(data) != {'rules'} or not isinstance(data['rules'], list):
            raise ValueError('checks must contain a declarative rules list')
        for rule in data['rules']:
            if not isinstance(rule, dict) or set(rule) - allowed or rule.get('kind') not in kinds:
                raise ValueError('unknown rule kind/field; arbitrary code is forbidden')
            ident = rule.get('id')
            if not isinstance(ident, str) or not ident or ident in rules:
                raise ValueError('duplicate or invalid rule id')
            targets = rule.get('targets', rule.get('paths'))
            if not isinstance(targets, list) or not targets or any(not isinstance(t, str) or not t.strip() for t in targets):
                raise ValueError('zero rule targets')
            rules[ident] = rule
    return rules


def _asset_rule_evidence(root, rules, errors):
    measured = {}
    for ident, rule in rules.items():
        targets = rule.get('targets',rule.get('paths',[]))
        if rule['kind'] != 'source-hash' or not all(t.startswith('design/assets/') for t in targets):
            continue
        allowed = rule.get('allowlist')
        if not isinstance(allowed,list) or not allowed or any(not isinstance(v,str) or not re.fullmatch('[0-9a-f]{64}',v) for v in allowed):
            raise ValueError('asset source hash allowlist must contain SHA256 values')
        files = {}
        for target in targets:
            if '..' in Path(target).parts or Path(target).is_absolute():
                raise ValueError('invalid asset source hash path')
            matches = sorted(Path(root).glob(target))
            if not matches:
                errors.append(f'zero asset source hash targets: {ident}')
            for path in matches:
                if not path.is_file() or not path.resolve().is_relative_to((Path(root)/'design/assets').resolve()):
                    raise ValueError('asset source hash target escapes assets')
                digest = hashlib.sha256(path.read_bytes()).hexdigest()
                files[path.relative_to(root).as_posix()] = digest
                if digest not in allowed:
                    errors.append(f'asset source hash mismatch: {ident}')
        measured[ident] = {'rule':rule,'files':files}
    return measured


def _approvals(design):
    return verify_rows(read_data(design/'approvals.yaml'),design.parent)


def check_design(root, *, catalog_dir=CATALOG, ship=False):
    design = Path(root) / 'design'
    result = {'evaluated_at':datetime.now(timezone.utc).isoformat(),
        'environment':{'python':sys.version.split()[0],
            'jcs_sha256':hashlib.sha256((SKILL/'scripts/jcs.cjs').read_bytes()).hexdigest()},
        'report':empty_report(), 'judgment_candidates':[], 'approval_targets':{}, 'approval_required':[],
        'exit_code': 2, 'errors': [], 'warnings': [], 'counts': {
        'items': 0, 'defined': 0, 'na': 0, 'deferred': 0, 'unclassified': 0,
        'proposed': 0, 'uncovered': 0, 'tokens': 0, 'fulfillment': 0.0}, 'manifest': {}}
    errors, warnings, counts = result['errors'], result['warnings'], result['counts']
    report=result['report']
    skip(report,'L1','L0 command does not execute source validation')
    skip(report,'L2','L0 command does not execute rendering')
    try:
        if not design.is_dir() or design.is_symlink():
            raise ValueError('design/ missing or symlink; run own-design-author init at git toplevel')
        _clear_manifests(design)
        config = read_data(inside(design, 'design.yaml'))
        if not isinstance(config, dict) or set(config) - {'schema_version','skill_version','catalog_version','preset','media','items','retired_items','checks','modules'}:
            raise ValueError('unknown design configuration fields')
        if config.get('schema_version') != 1:
            raise ValueError('unknown design schema version')
        catalog = load_catalog(config.get('catalog_version'), catalog_dir)
        presets = _presets()
        preset = config.get('preset')
        if preset not in presets or preset == 'schema_version':
            raise ValueError('unknown preset')
        if config.get('skill_version') != VERSION:
            warnings.append('skill version mismatch; migrate/review the pinned configuration')
        items = {item['id']: item for item in catalog['items']}
        counts['items'] = len(items)
        overrides = copy.deepcopy(config.get('items', {}))
        retired = config.get('retired_items', {})
        if not isinstance(retired, dict):
            raise ValueError('retired_items must be an object')
        for ident, state in retired.items():
            if ident in items or not isinstance(state, dict) or not isinstance(state.get('catalog_version'), str):
                raise ValueError('invalid retired item identity or source version')
            previous_catalog = load_catalog(state['catalog_version'], catalog_dir)
            previous_items = {item['id']:item for item in previous_catalog['items']}
            if ident not in previous_items:
                raise ValueError('unknown retired item id')
            items[ident] = previous_items[ident]
            overrides[ident] = {k:v for k,v in state.items() if k != 'catalog_version'}
        counts['items'] = len(items)
        if not isinstance(overrides, dict):
            raise ValueError('items must be an object')
        if set(overrides) - set(items):
            raise ValueError('unknown item id')
        media = config.get('media')
        if not isinstance(media, list) or not media or len(set(media)) != len(media):
            raise ValueError('zero or duplicate media targets')
        minimums = read_data(SKILL / 'references/minimums.json')
        profiles = {}
        for medium in media:
            if medium not in minimums or medium == 'schema_version':
                raise ValueError('unsupported media profile')
            profile = read_data(inside(design, f'profiles/{medium}.json'))
            if not isinstance(profile, dict) or profile.get('media') != medium:
                raise ValueError('media profile mismatch')
            if set(profile) - {'schema_version','media','minimums','canvas','root_font_size_px','units'}:
                raise ValueError('unknown media profile field')
            floor = minimums[medium]['minimums']
            supplied = profile.get('minimums', {})
            if not isinstance(supplied, dict) or set(supplied) - set(floor):
                raise ValueError('unknown media minimum')
            for key, value in supplied.items():
                if isinstance(value, bool) or not isinstance(value, (int,float)) or not math.isfinite(value) or value < floor[key]:
                    raise ValueError(f'relaxed media minimum: {key}')
            profiles[medium] = {**profile, 'minimums': {**floor, **supplied}}
        documents = [read_data(inside(design, str(p.relative_to(design))))
                     for p in sorted((design / 'tokens').glob('*.tokens.json'))]
        try:
            tokens = resolve_tokens(documents)
        except ValueError as exc:
            if str(exc) != 'zero token targets':
                raise
            tokens = {}
            errors.append('zero token targets')
        counts['tokens'] = len(tokens)
        approvals = _approvals(design)
        notice = read_data(inside(design, 'NOTICE'))
        sources = {}
        for row in notice.get('sources', []):
            if not isinstance(row, dict) or not all(k in row for k in ('id','url','commit','license','copyright','edited')):
                raise ValueError('incomplete NOTICE source')
            if row['id'] in sources or not all(isinstance(row[k],str) and row[k].strip()
                for k in ('id','url','commit','license','copyright')) or not isinstance(row['edited'],bool):
                raise ValueError('invalid NOTICE source')
            sources[row['id']] = row
        ia=resolve_ia(Path(root),config,approvals,profiles,report)
        if ia:
            errors.extend(ia['errors']);warnings.extend(ia['warnings'])
            for key,digest in ia['approval_hashes'].items():
                if key in approvals and approvals[key]!=digest:errors.append('approval hash mismatch: '+key)
        valid_targets = {f'ds:{m}:token:{path}' for m in media for path in tokens} | {f'ds:{m}:item:{ident}' for m in media for ident in items}
        if ia:valid_targets.update(ia['approval_hashes'])
        if set(approvals) - valid_targets:
            raise ValueError('approval target absent from catalog/tokens')
        for path, token in tokens.items():
            origin = token['origin']
            if origin not in ('owner','proposal') and not (isinstance(origin,str) and origin.startswith('reference:') and origin[10:] in sources):
                errors.append(f'missing origin/NOTICE for token: {path}')
            token['state'] = 'proposed'
        rules = _rules(design)
        measured_rules = _asset_rule_evidence(root,rules,errors)
        uncovered_rules = set(rules) - set(measured_rules)
        counts['uncovered'] = len(uncovered_rules)
        item_states, item_values, must_proposed, must_uncovered = {}, {}, [], []
        judgment_must={}
        judgment_policy,judgment_policy_sha256=read_judgment_policy()
        classification_policy = load_policy()
        for ident, item in items.items():
            classification = item.get('classification')
            if not isinstance(classification, dict) or classification.get('grade') not in ('must','nice-to-have'):
                counts['unclassified'] += 1
                errors.append(f'unclassified catalog item: {ident}')
                continue
            try:
                expected = classify_response(item, classification['raw_answer'], classification['model'], classification_policy)
                if any(classification.get(key) != val for key, val in expected.items()):
                    raise ValueError('classification differs from recorded input/policy/result')
            except (ClassificationError, KeyError, TypeError, ValueError):
                counts['unclassified'] += 1
                errors.append(f'stale/invalid catalog classification: {ident}')
                continue
            item_context={key:item[key] for key in ('document','title','headings','text','source_url') if key in item}
            must = classification['grade'] == 'must' or preset in classification.get('preset_must', [])
            state = copy.deepcopy(overrides.get(ident, {'state': 'deferred'}))
            if not isinstance(state, dict) or set(state) - {'state','evidence','reason'}:
                errors.append(f'invalid item fields: {ident}')
                continue
            status = state.get('state')
            item_states[ident] = {**state, 'must': must}
            evidence_values = {}
            if status == 'deferred':
                counts['deferred'] += 1
                if must:
                    errors.append(f'must item is empty/deferred: {ident}')
                continue
            if status == 'n-a':
                counts['na'] += 1
                reason = state.get('reason')
                if not isinstance(reason,str) or not reason.strip():
                    errors.append(f'n-a needs a reason: {ident}')
                    continue
                evidence_values = {'state': 'n-a', 'reason': reason}
                row=judgment_candidate('na-rationale',ident,reason,{'item':item_context,'preset':preset},policy=judgment_policy)
                result['judgment_candidates'].append(row);judgment_must[row['judgment_hash']]=must
            elif status == 'defined':
                counts['defined'] += 1
                evidence = state.get('evidence')
                if not isinstance(evidence,list) or not evidence or any(not isinstance(e,str) for e in evidence):
                    errors.append(f'defined requires evidence: {ident}')
                    continue
                for ref in evidence:
                    kind, _, target = ref.partition(':')
                    if kind not in item.get('evidence_types',[]) or not target:
                        errors.append(f'unknown evidence kind: {ident}')
                    elif kind == 'token':
                        if target not in tokens:
                            errors.append(f'missing token evidence: {ident}')
                        else:
                            evidence_values[ref] = {'type':tokens[target]['type'], 'value':tokens[target]['value']}
                    elif kind == 'doc':
                        name, separator, anchor = target.partition('#')
                        if name != 'DESIGN.md' or not separator or not anchor:
                            errors.append(f'doc evidence must name DESIGN.md and anchor: {ident}')
                            continue
                        body = _anchor_text(inside(design,name),anchor)
                        if not body or re.fullmatch(r'(TODO|TBD|未決定|未定)[.!。\s]*',body,re.I):
                            errors.append(f'missing/empty doc evidence: {ident}')
                        evidence_values[ref] = body
                        if body:
                            row=judgment_candidate('doc-relevance',ident,body,{'item':item_context,'reference':ref,'preset':preset},policy=judgment_policy)
                            result['judgment_candidates'].append(row);judgment_must[row['judgment_hash']]=must
                    elif kind == 'rule':
                        if target not in rules:
                            errors.append(f'missing rule evidence: {ident}')
                        else:
                            evidence_values[ref] = measured_rules.get(target,rules[target])
                            if must:
                                if target in uncovered_rules:
                                    must_uncovered.append(ident)
            else:
                errors.append(f'blank/unknown item state: {ident}')
            item_values[ident] = evidence_values
        ledger=read_data(inside(design,'judgments.json')) if (design/'judgments.json').exists() else None
        judged=evaluate_judgments(result['judgment_candidates'],ledger,policy=judgment_policy)
        errors.extend(judged['errors'])
        report['uncovered'].extend(judged['uncovered']);report['screened'].extend(judged['screened'])
        for row in judged['uncovered']:
            if judgment_must[row['judgment_hash']]:must_uncovered.append(row['id'])
        report['uncovered'].extend({'id':ident,'reason':'rule requires L1/L2 evidence'} for ident in sorted(uncovered_rules))
        counts['uncovered']=len(report['uncovered'])
        if not result['judgment_candidates']:skip(report,'semantic-evidence','no doc or n-a spans were flagged')
        prose = inside(design,'DESIGN.md').read_text()
        if re.search(r'(?<!\w)#[0-9a-fA-F]{3,8}\b', prose):
            errors.append('DESIGN.md contains a literal color; reference tokens instead')
        resolved_by_media = {}
        proposed_paths, must_proposed = set(), []
        for medium,profile in profiles.items():
            media_tokens,media_items = copy.deepcopy(tokens),copy.deepcopy(item_states)
            values = {'token:'+path:{'type':t['type'],'value':t['value']} for path,t in tokens.items()}
            values.update({'item:'+ident:v for ident,v in item_values.items()})
            hashes = media_hashes(medium,profile,values)
            for key,digest in hashes.items():
                kind,target = key.split(':',1)
                obj = media_tokens[target] if kind=='token' else media_items[target]
                obj['value_hash'] = digest
                approved = approvals.get(f'ds:{medium}:{key}')
                if approved is not None and approved != digest:
                    errors.append('approval hash mismatch: '+key)
                obj['state' if kind=='token' else 'approval_state'] = 'approved' if approved==digest else 'proposed'
            for ident,obj in media_items.items():
                if obj.get('approval_state')=='approved':
                    for ref in obj.get('evidence',[]):
                        if ref.startswith('token:') and ref[6:] in media_tokens:
                            media_tokens[ref[6:]]['state']='approved'
                if obj['state']=='n-a' and obj['must'] and obj.get('approval_state')!='approved':
                    errors.append('must n-a approval hash missing/mismatch: '+ident)
                if f'ds:{medium}:item:{ident}' in approvals and ident not in item_values:
                    errors.append('approval on deferred/invalid item: '+ident)
            for ident,obj in media_items.items():
                refs=obj.get('evidence',[])
                token_only=bool(refs) and all(ref.startswith('token:') and ref[6:] in media_tokens
                    and media_tokens[ref[6:]]['state']=='approved' for ref in refs)
                if obj['must'] and obj['state']=='defined' and obj.get('approval_state')!='approved' and not token_only:
                    must_proposed.append(medium+':'+ident)
            proposed_paths.update(path for path,t in media_tokens.items() if t['state']!='approved')
            resolved_by_media[medium]=(media_tokens,media_items)
        if ia:
            must_proposed.extend(key for key in ia['must_approval_keys'] if approvals.get(key)!=ia['approval_hashes'][key])
        counts['proposed']=len(proposed_paths)
        if must_proposed:
            warnings.append(f'proposed must items: {len(set(must_proposed))}')
        if counts['proposed']:
            warnings.append(f'proposed values: {counts["proposed"]}')
        if counts['uncovered']:
            warnings.append(f'uncovered rules/spans: {counts["uncovered"]}; recorded or measured evidence required')
        if ship and (must_proposed or must_uncovered):
            errors.append(f'ship gate: proposed must {len(set(must_proposed))}, uncovered must {len(set(must_uncovered))}')
        assets = []
        for path in sorted((design/'assets').rglob('*')):
            if path.is_file():
                inside(design,str(path.relative_to(design)))
                assets.append({'path':path.relative_to(design).as_posix(), 'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
        input_hashes = {p.relative_to(design).as_posix():hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted(design.rglob('*')) if p.is_file() and '.resolved' not in p.relative_to(design).parts
            and p.name not in ('assets.json','migration.json')}
        source_hash = value_hash(input_hashes)
        asset_exports = _asset_exports(root,design)
        from tailwind_contract import build_artifacts
        tailwind_templates = read_data(inside(design,'tailwind.json')) if (design/'tailwind.json').exists() else None
        artifacts = build_artifacts(tokens,tailwind_templates)
        if hashlib.sha256(JUDGMENT_POLICY.read_bytes()).hexdigest()!=judgment_policy_sha256:
            raise ValueError('judgment policy changed during validation')
        for medium, profile in profiles.items():
            result['manifest'][medium] = {'schema_version':1, 'skill_version':VERSION,
                'approval_scheme':SCHEME,'judgment_policy_sha256':judgment_policy_sha256,
                'catalog_version':catalog['version'], 'catalog_hash':value_hash(catalog), 'source_hash':source_hash,
                'input_hashes':input_hashes, 'media':medium, 'profile':profile,
                'minimums':profile['minimums'], 'tokens':resolved_by_media[medium][0], 'items':resolved_by_media[medium][1],
                'css_variables':{t['css_name']:t['css'] for t in tokens.values()},
                'tailwind_templates':tailwind_templates, 'artifacts':artifacts, 'rules':list(rules.values()), 'coverage':{'uncovered':sorted(uncovered_rules | {row['id'] for row in judged['uncovered']})},
                'modules':({'ia':{'version':config['modules']['ia'],'implementation_sha256':ia['_implementation_sha256'],
                    'resolved':ia['media'][medium]}} if ia else {}),
                'palette':{kind: sorted({t['css'] for t in tokens.values() if t['type']==kind})
                           for kind in ('color','fontFamily','dimension')}, 'assets':assets, 'asset_exports':asset_exports}
        for medium,manifest in result['manifest'].items():
            for kind,rows in [('item',manifest['items']),('token',manifest['tokens'])]:
                for ident,row in rows.items():
                    if 'value_hash' not in row:continue
                    key=f'ds:{medium}:{kind}:{ident}'
                    result['approval_targets'][key]=row['value_hash']
                    if kind=='item' and row['must']:result['approval_required'].append(key)
        if ia:
            result['approval_targets'].update(ia['approval_hashes'])
            result['approval_required'].extend(ia['must_approval_keys'])
        result['exit_code'] = 2 if errors else (1 if warnings else 0)
        counts['fulfillment'] = (counts['defined']+counts['na'])/counts['items']
        stage(report,'L0',result['exit_code'],counts['items'])
        for manifest in result['manifest'].values():manifest['report']=copy.deepcopy(report)
        if not errors:
            write_data(design/'assets.json', {'schema_version':1,'assets':assets, 'asset_exports':asset_exports})
            for name, artifact in artifacts.items():
                generated = design/'.resolved'/artifact.get('file',name)
                generated.parent.mkdir(parents=True,exist_ok=True)
                generated.write_text(artifact['content'])
            for medium, manifest in result['manifest'].items():
                write_data(design/'.resolved'/f'{medium}.json',manifest)
                css = ':root {\n' + ''.join(f'  {key}: {value};\n' for key,value in manifest['css_variables'].items()) + '}\n'
                (design/'.resolved'/f'{medium}.css').write_text(css)
    except (ClassificationError, ValueError, OSError, KeyError, TypeError, AttributeError) as exc:
        errors.append(str(exc).replace(str(Path(root).resolve()), '<repo>'))
        if design.is_dir() and not design.is_symlink():
            try:
                _clear_manifests(design)
            except (OSError, ValueError):
                errors.append('generated output cleanup failed; do not consume generated files')
        result['exit_code'] = 2
    stage(report,'L0',result['exit_code'],counts['items'])
    for manifest in result['manifest'].values():manifest['report']=copy.deepcopy(report)
    return result


def _deployment_path(root, target):
    if not isinstance(target,str) or not target or Path(target).is_absolute() or '..' in Path(target).parts or '\\' in target:
        raise ValueError('invalid deployment target')
    path = Path(root)/target
    if any(p.is_symlink() for p in (path, *path.parents) if p.is_relative_to(Path(root))):
        raise ValueError('symlinked deployment target is unsupported')
    if not path.resolve().is_relative_to(Path(root).resolve()) or path.resolve().is_relative_to((Path(root)/'design').resolve()) or any(p.startswith('.') for p in Path(target).parts):
        raise ValueError('deployment target escapes artifact directories')
    return path


def _asset_exports(root, design):
    path = design/'asset-exports.json'
    if not path.exists():
        return []
    data = read_data(inside(design,'asset-exports.json'))
    if not isinstance(data,dict) or set(data) != {'schema_version','exports'} or data['schema_version'] != 1 or not isinstance(data['exports'],list):
        raise ValueError('invalid asset export schema')
    rows, seen = [], set()
    for row in data['exports']:
        if not isinstance(row,dict) or set(row) != {'source','targets'} or not isinstance(row['source'],str) or not row['source'].startswith('assets/'):
            raise ValueError('asset export source must be a design asset')
        source = inside(design,row['source'])
        targets = row['targets']
        if not isinstance(targets,list) or not targets:
            raise ValueError('zero asset export targets')
        for target in targets:
            output = _deployment_path(root,target)
            if target in seen or output.suffix.lower() not in ('.png','.jpg','.jpeg','.webp','.svg','.ico') or output.suffix.lower() != source.suffix.lower():
                raise ValueError('invalid or duplicate asset export target')
            seen.add(target)
        rows.append({**row,'sha256':hashlib.sha256(source.read_bytes()).hexdigest()})
    return rows


def sync_artifacts(root, manifest):
    writes = {}
    for target, artifact in manifest.get('artifacts',{}).items():
        writes[_deployment_path(root,target)] = artifact['content'].encode()
    for row in manifest.get('asset_exports',[]):
        content = inside(Path(root)/'design',row['source']).read_bytes()
        for target in row['targets']:
            output = _deployment_path(root,target)
            if output in writes:
                raise ValueError('duplicate artifact output')
            writes[output] = content
    before = {path:path.read_bytes() if path.exists() else None for path in writes}
    applied = []
    try:
        for path,content in writes.items():
            path.parent.mkdir(parents=True,exist_ok=True)
            temporary = path.with_name(path.name+'.design-tmp')
            try:
                temporary.write_bytes(content)
                temporary.replace(path)
            finally:
                if temporary.exists(): temporary.unlink()
            applied.append(path)
    except OSError:
        for path in reversed(applied):
            if before[path] is None: path.unlink()
            else: path.write_bytes(before[path])
        raise
    return [path.relative_to(root).as_posix() for path in writes]


def git_root(path):
    run = subprocess.run(['git','-C',str(path),'rev-parse','--show-toplevel'],capture_output=True,text=True)
    if run.returncode:
        raise ValueError('design commands require an existing git repository')
    return Path(run.stdout.strip())


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command',required=True)
    for command in ('init','check','migrate','sync','approve'):
        p = sub.add_parser(command)
        p.add_argument('--repo',default='.')
        p.add_argument('--catalog-dir',type=Path,default=CATALOG)
        if command == 'init':
            p.add_argument('--preset',required=True,choices=['lp','product-ui','presentation'])
        if command in ('init','migrate'):
            p.add_argument('--catalog-version',required=True)
        if command == 'check':
            p.add_argument('--ship',action='store_true')
            p.add_argument('--include-manifest',action='store_true')
            p.add_argument('--include-candidates',action='store_true')
        if command == 'approve':
            p.add_argument('--prepare',action='store_true')
            p.add_argument('--must',action='store_true')
            p.add_argument('--target',action='append',default=[])
            p.add_argument('--request',type=Path)
            p.add_argument('--approver',default='owner')
    args = parser.parse_args(argv)
    try:
        root = git_root(args.repo)
        if args.command == 'approve':
            from owner_approval import prepare,approve_request
            if args.prepare:
                if args.request:raise ValueError('prepare does not use a request')
                result=prepare(root,args.catalog_dir,args.target,args.must)
            else:
                if not args.request or args.must or args.target:
                    raise ValueError('approval recording requires only --request and --approver')
                result=approve_request(root,args.catalog_dir,read_data(args.request),args.approver)
            print(json.dumps(result,ensure_ascii=False,sort_keys=True))
            return result.get('exit_code',0)
        if args.command in ('check','sync'):
            result = check_design(root,catalog_dir=args.catalog_dir,ship=getattr(args,'ship',False))
            if args.command == 'sync' and result['exit_code'] != 2:
                result['synced'] = sync_artifacts(root,result['manifest']['web'])
            for key in ('approval_targets','approval_required'):result.pop(key,None)
            if not getattr(args,'include_candidates',False):result.pop('judgment_candidates',None)
            if not getattr(args,'include_manifest',False):
                result.pop('manifest',None)
            print(json.dumps(result,ensure_ascii=False,sort_keys=True))
            return result['exit_code']
        result = (init_design(root,args.preset,args.catalog_version,catalog_dir=args.catalog_dir)
                  if args.command == 'init' else migrate_design(root,args.catalog_version,catalog_dir=args.catalog_dir))
        print(json.dumps(result,ensure_ascii=False,sort_keys=True))
        return 0
    except (ValueError,OSError,KeyError,TypeError) as exc:
        print(json.dumps({'exit_code':2,'errors':[str(exc)]},ensure_ascii=False))
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
