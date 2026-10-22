#!/usr/bin/env python3
"""Create/check/migrate a repository's sole design/ directory. Never writes approvals."""
import argparse
import copy
from datetime import date
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


def _approvals(design):
    rows = read_data(design / 'approvals.yaml')
    if not isinstance(rows, list):
        raise ValueError('approval ledger must be a list')
    approvals = {}
    for row in rows:
        if not isinstance(row, dict) or set(row) - {'item_id', 'token_path', 'value_hash', 'approver', 'date'}:
            raise ValueError('invalid approval fields')
        keys = [key for key in ('item_id', 'token_path') if key in row]
        if len(keys) != 1 or not all(isinstance(row.get(k), str) and row[k].strip()
                                    for k in (keys + ['value_hash', 'approver', 'date'])):
            raise ValueError('invalid approval row')
        try:
            date.fromisoformat(row['date'])
        except ValueError as exc:
            raise ValueError('approval date must be ISO date') from exc
        key = (keys[0], row[keys[0]])
        if key in approvals:
            raise ValueError('duplicate approval target')
        approvals[key] = row['value_hash']
    return approvals


def check_design(root, *, catalog_dir=CATALOG, ship=False):
    design = Path(root) / 'design'
    result = {'exit_code': 2, 'errors': [], 'warnings': [], 'counts': {
        'items': 0, 'defined': 0, 'na': 0, 'deferred': 0, 'unclassified': 0,
        'proposed': 0, 'uncovered': 0, 'tokens': 0, 'fulfillment': 0.0}, 'manifest': {}}
    errors, warnings, counts = result['errors'], result['warnings'], result['counts']
    try:
        if not design.is_dir() or design.is_symlink():
            raise ValueError('design/ missing or symlink; run own-design-author init at git toplevel')
        _clear_manifests(design)
        config = read_data(inside(design, 'design.yaml'))
        if not isinstance(config, dict) or set(config) - {'schema_version','skill_version','catalog_version','preset','media','items','retired_items','checks'}:
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
        valid_targets = {('token_path', path) for path in tokens} | {('item_id', ident) for ident in items}
        if set(approvals) - valid_targets:
            raise ValueError('approval target absent from catalog/tokens')
        for path, token in tokens.items():
            origin = token['origin']
            if origin not in ('owner','proposal') and not (isinstance(origin,str) and origin.startswith('reference:') and origin[10:] in sources):
                errors.append(f'missing origin/NOTICE for token: {path}')
            approved = approvals.get(('token_path', path))
            if approved is not None and approved != token['value_hash']:
                errors.append(f'approval hash mismatch: {path}')
            token['state'] = 'approved' if approved == token['value_hash'] else 'proposed'
            if token['state'] == 'proposed':
                counts['proposed'] += 1
        rules = _rules(design)
        counts['uncovered'] = len(rules)
        item_states, must_proposed, must_uncovered = {}, [], []
        classification_policy = load_policy()
        used_item_approvals = set()
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
                approval = approvals.get(('item_id', ident))
                item_states[ident]['value_hash'] = value_hash(evidence_values)
                if (must or approval is not None) and approval != value_hash(evidence_values):
                    errors.append(f'must n-a approval hash missing/mismatch: {ident}')
                used_item_approvals.add(ident)
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
                            if must and tokens[target]['state'] != 'approved':
                                must_proposed.append(ident)
                    elif kind == 'doc':
                        name, separator, anchor = target.partition('#')
                        if name != 'DESIGN.md' or not separator or not anchor:
                            errors.append(f'doc evidence must name DESIGN.md and anchor: {ident}')
                            continue
                        body = _anchor_text(inside(design,name),anchor)
                        if not body or re.fullmatch(r'(TODO|TBD|未決定|未定)[.!。\s]*',body,re.I):
                            errors.append(f'missing/empty doc evidence: {ident}')
                        evidence_values[ref] = body
                        if must:
                            must_proposed.append(ident)
                    elif kind == 'rule':
                        if target not in rules:
                            errors.append(f'missing rule evidence: {ident}')
                        else:
                            evidence_values[ref] = rules[target]
                            if must:
                                must_uncovered.append(ident)
                if ('item_id',ident) in approvals:
                    used_item_approvals.add(ident)
                    if approvals[('item_id',ident)] != value_hash(evidence_values):
                        errors.append(f'item approval hash mismatch: {ident}')
                    else:
                        must_proposed = [i for i in must_proposed if i != ident]
                item_states[ident]['value_hash'] = value_hash(evidence_values)
            else:
                errors.append(f'blank/unknown item state: {ident}')
        for (kind,target) in approvals:
            if kind == 'item_id' and target not in used_item_approvals:
                errors.append(f'approval on deferred/invalid item: {target}')
        prose = inside(design,'DESIGN.md').read_text()
        if re.search(r'(?<!\w)#[0-9a-fA-F]{3,8}\b', prose):
            errors.append('DESIGN.md contains a literal color; reference tokens instead')
        if must_proposed:
            warnings.append(f'proposed must items: {len(set(must_proposed))}')
        if counts['proposed']:
            warnings.append(f'proposed values: {counts["proposed"]}')
        if rules:
            warnings.append(f'uncovered rules: {len(rules)}; L1/L2 evidence required')
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
        from tailwind_contract import tailwind_artifact
        artifacts = ({'tailwind.config.cjs': tailwind_artifact(tokens)}
                     if any(t['type'] == 'color' for t in tokens.values()) else {})
        tailwind_content = artifacts.get('tailwind.config.cjs', {}).get('content')
        for medium, profile in profiles.items():
            result['manifest'][medium] = {'schema_version':1, 'skill_version':VERSION,
                'catalog_version':catalog['version'], 'catalog_hash':value_hash(catalog), 'source_hash':source_hash,
                'input_hashes':input_hashes, 'media':medium, 'profile':profile,
                'minimums':profile['minimums'], 'tokens':tokens, 'items':item_states,
                'css_variables':{t['css_name']:t['css'] for t in tokens.values()},
                'artifacts':artifacts, 'rules':list(rules.values()), 'coverage':{'uncovered':sorted(rules)},
                'palette':{kind: sorted({t['css'] for t in tokens.values() if t['type']==kind})
                           for kind in ('color','fontFamily','dimension')}, 'assets':assets}
        result['exit_code'] = 2 if errors else (1 if warnings else 0)
        counts['fulfillment'] = (counts['defined']+counts['na'])/counts['items']
        if not errors:
            write_data(design/'assets.json', {'schema_version':1,'assets':assets})
            if tailwind_content is not None:
                (design/'.resolved'/'tailwind.config.cjs').parent.mkdir(parents=True, exist_ok=True)
                (design/'.resolved'/'tailwind.config.cjs').write_text(tailwind_content)
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
    return result


def git_root(path):
    run = subprocess.run(['git','-C',str(path),'rev-parse','--show-toplevel'],capture_output=True,text=True)
    if run.returncode:
        raise ValueError('design commands require an existing git repository')
    return Path(run.stdout.strip())


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command',required=True)
    for command in ('init','check','migrate'):
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
    args = parser.parse_args(argv)
    try:
        root = git_root(args.repo)
        if args.command == 'check':
            result = check_design(root,catalog_dir=args.catalog_dir,ship=args.ship)
            if not args.include_manifest:
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
