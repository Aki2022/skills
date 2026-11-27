#!/usr/bin/env python3
"""Create/check/migrate a repository's sole ux/ layers directory. Owner approve command is the only ledger writer."""
import argparse
import copy
from datetime import date, datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path, PurePosixPath
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
import item_checks
import layer_paths
import layout_overrides
import layout_patterns

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


def design_rel():
    """The design layer's own directory, repository-relative: the one that holds the design configuration."""
    return layer_paths.rel('design_dir')


def design_dir(root):
    """The design layer's own directory in `root`, after the layout guard."""
    return layer_paths.path(root, 'design_dir')


def inside_layer(root, key, *rest):
    """A layer path (guarded) that must stay under the layers root, symlinks resolved."""
    path = layer_paths.path(root, key, *rest).resolve()
    if not path.is_relative_to(layer_paths.path(root, 'layers').resolve()):
        raise ValueError(f'path outside the layers root: {layer_paths.rel(key, *rest)}')
    return path


def _leaf(key):
    """The file or directory name of a layer entry (generated outputs are recognised by it inside the design dir)."""
    return PurePosixPath(layer_paths.rel(key)).name


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
    layer_paths.guard(root)
    presets = _presets()
    if preset not in presets or preset == 'schema_version':
        raise ValueError('unknown preset')
    load_catalog(version, catalog_dir)
    design = design_dir(root)
    if design.exists():
        raise ValueError(f"{design_rel()}/ already exists; use migrate")
    shutil.copytree(SKILL / 'templates/design', design)
    ledger = layer_paths.path(root, 'ledger')  # the approval ledger sits next to the layers, not inside the design layer
    ledger.parent.mkdir(parents=True, exist_ok=True)
    if not ledger.exists():
        shutil.copyfile(SKILL / 'templates/approvals.yaml', ledger)
    layer_paths.path(root, 'assets').mkdir(parents=True, exist_ok=True)
    layer_paths.path(root, 'checks').mkdir(parents=True, exist_ok=True)
    write_data(layer_paths.path(root, 'design_config'), {'schema_version': 1, 'skill_version': VERSION,
        'catalog_version': version, 'preset': preset, 'media': presets[preset]['media'], 'items': {}})
    return {'created': design_rel() + '/', 'catalog_version': version, 'preset': preset}


def migrate_design(root, version, *, catalog_dir=CATALOG):
    root = Path(root)
    design = design_dir(root)
    config_path = layer_paths.path(root, 'design_config')
    config = read_data(config_path)
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
    write_data(config_path, config)
    # No stale generated contract survives migration, even before the next L0.
    _clear_manifests(root)
    return report


def _clear_manifests(root):
    assets_index = layer_paths.path(root, 'assets_index')
    if assets_index.is_file() or assets_index.is_symlink():
        assets_index.unlink()
    resolved = layer_paths.path(root, 'resolved')
    if resolved.is_symlink():
        resolved.unlink()
        raise ValueError('.resolved symlink was unlinked; its target was not modified')
    if resolved.is_dir():
        for path in resolved.iterdir():
            if path.is_file() and path.suffix in ('.json', '.css', '.cjs') and path.name != 'maintenance-stamp.json':
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


AFM_RULE_CHARS, AFM_EVIDENCE_CHARS = 500, 400


def afm_candidate(ident, context, evidence_values, statements=()):
    """A repository's judgment statement replaces the catalog text: AFM judges the decision the repository made."""
    heading = ' > '.join(context.get('headings') or [context.get('title', '')])
    text = ' '.join(statements) if statements else str(context.get('text', ''))
    rule = (heading + ': ' + ' '.join(text.split()))[:AFM_RULE_CHARS]
    prose = ' '.join(' '.join(v.split()) for k, v in sorted(evidence_values.items())
                     if k.startswith('doc:') and isinstance(v, str))
    return {'id': ident, 'document': item_checks.document_key(context), 'rule': rule, 'evidence': prose[:AFM_EVIDENCE_CHARS]}


PAGE_RULE_FIELDS = {'count': ({'targets'}, {'scope', 'minimum', 'maximum'}),
                    'text': ({'targets', 'pattern'}, {'minimum', 'maximum'}),
                    'style': ({'targets', 'property', 'allowlist'}, {'scope', 'state'}),
                    'judgment': ({'targets', 'statement'}, set())}


def _selectors(value, name):
    if not isinstance(value, list) or not value or any(not isinstance(s, str) or not s.strip() for s in value):
        raise ValueError(f'page rule {name} must be a non-empty selector list')


def _page_rule(rule):
    """Validate a rule the repository declares for its own decisions; it runs on every rendered page."""
    required, optional = PAGE_RULE_FIELDS[rule['kind']]
    fields = set(rule) - {'id', 'kind'}
    if required - fields or fields - required - optional:
        raise ValueError(f"page rule {rule['id']}: {rule['kind']} needs {sorted(required)} and allows {sorted(optional)}")
    _selectors(rule['targets'], 'targets')
    if 'scope' in rule:
        _selectors(rule['scope'], 'scope')
    for bound in ('minimum', 'maximum'):
        if bound in rule and (isinstance(rule[bound], bool) or not isinstance(rule[bound], int) or rule[bound] < 0):
            raise ValueError(f"page rule {rule['id']}: {bound} must be a non-negative integer")
    if rule['kind'] == 'count' and not {'minimum', 'maximum'} & fields:
        raise ValueError(f"page rule {rule['id']}: count needs minimum or maximum")
    if rule['kind'] == 'text':
        if not isinstance(rule['pattern'], str) or not rule['pattern']:
            raise ValueError(f"page rule {rule['id']}: text needs a pattern")
        try:
            re.compile(rule['pattern'])
        except re.error as exc:
            raise ValueError(f"page rule {rule['id']}: invalid pattern ({exc})") from None
    if rule['kind'] == 'style':
        if not isinstance(rule['property'], str) or not re.fullmatch(r'[a-z][a-z-]*', rule['property']):
            raise ValueError(f"page rule {rule['id']}: style property must be a CSS property name")
        if not isinstance(rule['allowlist'], list) or not rule['allowlist'] or any(not isinstance(v, str) or not v.strip() for v in rule['allowlist']):
            raise ValueError(f"page rule {rule['id']}: style needs a non-empty allowlist")
        if rule.get('state', 'focus') != 'focus':
            raise ValueError(f"page rule {rule['id']}: style state can only be focus")
    if rule['kind'] == 'judgment' and (not isinstance(rule['statement'], str) or not rule['statement'].strip()):
        raise ValueError(f"page rule {rule['id']}: judgment needs a statement")


def _rules(root):
    design = design_dir(root)
    rules = {}
    allowed = {'id', 'kind', 'targets', 'allowlist', 'minimum', 'paths', 'pattern'}
    kinds = {'contrast', 'target-size', 'focus', 'source-hash', 'palette', 'source-pattern'}
    for path in sorted(layer_paths.path(root, 'checks').glob('*.yaml')):
        inside(design, os.path.relpath(path, design))
        data = read_data(path)
        if not isinstance(data, dict) or set(data) != {'rules'} or not isinstance(data['rules'], list):
            raise ValueError('checks must contain a declarative rules list')
        for rule in data['rules']:
            if not isinstance(rule, dict) or rule.get('kind') not in kinds | item_checks.PAGE_KINDS:
                raise ValueError('unknown rule kind; arbitrary code is forbidden')
            ident = rule.get('id')
            if not isinstance(ident, str) or not ident or ident in rules:
                raise ValueError('duplicate or invalid rule id')
            if rule['kind'] in item_checks.PAGE_KINDS:
                _page_rule(rule)
                rules[ident] = rule
                continue
            if set(rule) - allowed:
                raise ValueError('unknown rule field; arbitrary code is forbidden')
            targets = rule.get('targets', rule.get('paths'))
            if not isinstance(targets, list) or not targets or any(not isinstance(t, str) or not t.strip() for t in targets):
                raise ValueError('zero rule targets')
            rules[ident] = rule
    return rules


def _resolve_style_tokens(rules, tokens, errors):
    """A style rule may allow token:<path>; the manifest carries the token's CSS value for the page to compare."""
    for ident, rule in rules.items():
        if rule['kind'] != 'style':
            continue
        values = []
        for value in rule['allowlist']:
            if value.startswith('token:'):
                if value[6:] not in tokens:
                    errors.append(f'style rule token missing: {ident} {value}')
                    continue
                values.append(tokens[value[6:]]['css'])
            else:
                values.append(value)
        rule['resolved_allowlist'] = values


def _asset_rule_evidence(root, rules, errors):
    measured = {}
    for ident, rule in rules.items():
        targets = rule.get('targets',rule.get('paths',[]))
        if rule['kind'] != 'source-hash' or not all(t.startswith(layer_paths.rel('assets') + '/') for t in targets):
            continue
        allowed = rule.get('allowlist')
        if not isinstance(allowed,list) or not allowed or any(not isinstance(v,str) or not re.fullmatch('[0-9a-f]{64}',v) for v in allowed):
            raise ValueError('asset source hash allowlist must contain SHA256 values')
        files = {}
        # The rule and its measured files are hashed into the approvals of the items that cite it. They are hashed in the
        # form they had before the layers moved (a frozen prefix), so a move does not drop those approvals.
        assets_prefix = layer_paths.rel('assets') + '/'
        def canonical(target):
            return layer_paths.identity('asset_path_prefix') + '/' + target[len(assets_prefix):]
        for target in targets:
            if '..' in Path(target).parts or Path(target).is_absolute():
                raise ValueError('invalid asset source hash path')
            matches = sorted(Path(root).glob(target))
            if not matches:
                errors.append(f'zero asset source hash targets: {ident}')
            for path in matches:
                if not path.is_file() or not path.resolve().is_relative_to(layer_paths.path(root, 'assets').resolve()):
                    raise ValueError('asset source hash target escapes assets')
                digest = hashlib.sha256(path.read_bytes()).hexdigest()
                files[canonical(path.relative_to(root).as_posix())] = digest
                if digest not in allowed:
                    errors.append(f'asset source hash mismatch: {ident}')
        canonical_rule = dict(rule)
        for field in ('targets', 'paths'):
            if field in canonical_rule:
                canonical_rule[field] = [canonical(t) for t in canonical_rule[field]]
        measured[ident] = {'rule':canonical_rule,'files':files}
    return measured


def _render_paths(config, errors):
    """The path each URL env renders, from the declared gate and channels; None when none are declared."""
    rows = [config['integration_gate']] if isinstance(config.get('integration_gate'), dict) else []
    rows += [row for row in (config.get('channels') or {}).values() if isinstance(row, dict)]
    paths = {}
    for row in rows:
        for env, server in (row.get('render_servers') or {}).items():
            path = server.get('path') if isinstance(server, dict) else None
            if env in paths and paths[env] != path:
                errors.append(f'render server {env} renders different paths in different gates')
            paths[env] = path
    return paths or None


def _page_inventory(root, design, config, approvals, counts, errors, warnings, ship):
    """Every page is checked or excluded; an exclusion is the owner's decision, bound to its reason by hash."""
    if not layer_paths.path(root, 'validation').exists():
        return {}
    from page_inventory import check_pages
    validation = read_data(inside_layer(root, 'validation'))
    pages = validation.get('pages') if isinstance(validation, dict) else None
    if pages is None:
        if ship:
            errors.append('ship gate: page inventory required; declare validation.json pages so every page is checked or excluded')
        return {}
    inventory = check_pages(root, validation, paths=_render_paths(config, errors))
    errors.extend(inventory['errors'])
    excluded = pages.get('excluded') if isinstance(pages, dict) and isinstance(pages.get('excluded'), dict) else {}
    covered = pages.get('covered') if isinstance(pages, dict) and isinstance(pages.get('covered'), dict) else {}
    keys = {f'ds:web:page:{page}': value_hash({'page': page, 'reason': reason})
            for page, reason in excluded.items() if isinstance(reason, str)}
    counts['pages'] = {'found': inventory['pages'], 'covered': len(covered), 'excluded': len(excluded)}
    unapproved = sorted(key for key, digest in keys.items() if approvals.get(key) != digest)
    if unapproved and ship:
        errors.append(f'ship gate: unapproved page exclusions {len(unapproved)}')
    elif unapproved:
        warnings.append(f'page exclusions awaiting owner approval: {len(unapproved)}')
    return keys


def _layout_overrides(root, approvals, counts, errors, warnings, ship):
    """A placement against the allocation rules is the owner's decision, row by row (SPEC-deliverable-production-layers requirement 4).
    -> ({key: hash} of the rows that can be approved, the keys a ledger row may name: every row's and every stale one's)."""
    state = layout_overrides.evaluate_repo(root, approvals)
    errors.extend(f'{code}: {message}' for code, message in state.findings)
    unapproved = state.unapproved()
    if state.rows:
        counts['overrides'] = {'rows': len(state.rows), 'approved': sum(r.status == 'approved' for r in state.rows)}
    if unapproved and ship:
        errors.append(f'ship gate: unapproved layout overrides {len(unapproved)}')
    elif unapproved:
        warnings.append(f'layout overrides awaiting owner approval: {len(unapproved)}')
    stale = layout_overrides.stale_keys(state, approvals)
    if stale:  # the row is gone from the file: its approval reads nothing, but it is not left unsaid
        warnings.append(f'approvals for layout overrides no longer in the file: {len(stale)}')
    return state.targets, state.declared_keys | set(stale)


def _prohibitions(config, rules):
    """What the design does not allow, in a form code can read: each row pairs a statement with the page rule that
    checks it (ux/design/checks). The prose in DESIGN.md stays as it is, so the approvals that bind it stay valid.
    -> (rows for the manifest, {approval key suffix: hash}); a row with no rule is uncovered (SPEC-deliverable-production-layers 27)."""
    rows = config.get('prohibitions', [])
    if not isinstance(rows, list):
        raise ValueError('prohibitions must be a list')
    seen, out, hashes = set(), [], {}
    for row in rows:
        if not isinstance(row, dict) or set(row) - {'id', 'statement', 'rule'}:
            raise ValueError('prohibition: unknown field (allowed: id, statement, rule)')
        ident, statement, rule = row.get('id'), row.get('statement'), row.get('rule')
        if not isinstance(ident, str) or not re.fullmatch(r'[a-z0-9][a-z0-9-]*', ident):
            raise ValueError('prohibition: id must be lower-case kebab-case')
        if ident in seen:
            raise ValueError(f'prohibition: duplicate id {ident}')
        seen.add(ident)
        if not isinstance(statement, str) or not statement.strip():
            raise ValueError(f'prohibition {ident}: statement is required')
        if rule is not None:
            if rule not in rules:
                raise ValueError(f'prohibition {ident}: unknown rule {rule}')
            if rules[rule]['kind'] not in item_checks.PAGE_KINDS - {'judgment'}:
                raise ValueError(f'prohibition {ident}: rule {rule} must be a page rule (count, text or style) that code measures')
        out.append({'id': ident, 'statement': statement, 'rule': rule})
        hashes[ident] = value_hash({'id': ident, 'statement': statement, 'rule': rule, 'definition': rules[rule] if rule else None})
    return out, hashes


def _comp_size(medium, defaults, supplied):
    """The size of the comp image for `medium`: the skill's default, or the profile's override (not a brand value, so it
    is popped from the profile before any approval hash sees it)."""
    default = defaults.get(medium)
    if not isinstance(default, dict):
        raise ValueError(f'DS_COMP_SIZE_INVALID: no comp size for media {medium}')
    if supplied is None:
        return {**default, 'source': 'skill'}
    if not isinstance(supplied, dict) or not supplied or set(supplied) - set(default) or any(
            isinstance(v, bool) or not isinstance(v, int) or v <= 0 for v in supplied.values()):
        raise ValueError(f'DS_COMP_SIZE_INVALID: profiles/{medium}.json の comp は {sorted(default)} の正の整数だけ')
    return {**default, **supplied, 'source': 'profile'}


def _approvals(root):
    return verify_rows(read_data(layer_paths.path(root, 'ledger')), root)


def check_design(root, *, catalog_dir=CATALOG, ship=False):
    root = Path(root)
    design = root / design_rel()
    result = {'evaluated_at':datetime.now(timezone.utc).isoformat(),
        'environment':{'python':sys.version.split()[0],
            'jcs_sha256':hashlib.sha256((SKILL/'scripts/jcs.cjs').read_bytes()).hexdigest()},
        'report':empty_report(), 'judgment_candidates':[], 'afm_candidates':[], 'approval_targets':{}, 'approval_required':[],
        'exit_code': 2, 'errors': [], 'warnings': [], 'counts': {
        'items': 0, 'defined': 0, 'na': 0, 'deferred': 0, 'unclassified': 0,
        'proposed': 0, 'uncovered': 0, 'tokens': 0, 'fulfillment': 0.0}, 'manifest': {}}
    errors, warnings, counts = result['errors'], result['warnings'], result['counts']
    report=result['report']
    skip(report,'L1','L0 command does not execute source validation')
    skip(report,'L2','L0 command does not execute rendering')
    try:
        layer_paths.guard(root)
        if not design.is_dir() or design.is_symlink():
            raise ValueError(f"{design_rel()}/ missing or symlink; run own-design-author init at git toplevel")
        _clear_manifests(root)
        config = read_data(inside_layer(root, 'design_config'))
        if not isinstance(config, dict) or set(config) - {'schema_version','skill_version','catalog_version','preset','media','items','retired_items','checks','modules','spec','maintenance','channels','integration_gate','layout_patterns','prohibitions'}:
            raise ValueError('unknown design configuration fields')
        if config.get('schema_version') != 1:
            raise ValueError('unknown design schema version')
        catalog = load_catalog(config.get('catalog_version'), catalog_dir)
        if 'channels' in config:
            from ship_contract import validate_channels
            validate_channels(config)
        if 'integration_gate' in config:
            from ship_contract import validate_integration_gate
            validate_integration_gate(config)
        if 'maintenance' in config:
            from maintenance import validate_maintenance
            validate_maintenance(config, {x['id'] for x in catalog['items']} | set(config.get('retired_items', {})))
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
        comp_sizes = read_data(SKILL / 'references/comp-sizes.json')
        profiles, comps = {}, {}
        for medium in media:
            if medium not in minimums or medium == 'schema_version':
                raise ValueError('unsupported media profile')
            profile = read_data(inside(design, f'profiles/{medium}.json'))
            if not isinstance(profile, dict) or profile.get('media') != medium:
                raise ValueError('media profile mismatch')
            if set(profile) - {'schema_version','media','minimums','canvas','root_font_size_px','units','comp'}:
                raise ValueError('unknown media profile field')
            comps[medium] = _comp_size(medium, comp_sizes, profile.pop('comp', None))
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
        approvals = _approvals(root)
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
        page_keys = _page_inventory(root, design, config, approvals, counts, errors, warnings, ship)
        valid_targets.update(page_keys)
        override_keys, override_names = _layout_overrides(root, approvals, counts, errors, warnings, ship)
        valid_targets.update(override_names)
        vocabulary = layout_patterns.resolve(config.get('layout_patterns'), approvals, media, ship=ship)
        errors.extend(vocabulary['errors'])
        warnings.extend(vocabulary['warnings'])
        valid_targets.update(vocabulary['targets'])
        prohibition_rows, prohibition_hashes = _prohibitions(config, _rules(root))
        prohibition_keys = {f'ds:web:prohibition:{ident}': digest for ident, digest in prohibition_hashes.items()}
        # the row is gone from the file: its approval reads nothing, but it is not left unsaid
        stale_prohibitions = {k for k in approvals if k.startswith('ds:web:prohibition:')} - set(prohibition_keys)
        valid_targets.update(prohibition_keys, stale_prohibitions)
        if set(approvals) - valid_targets:
            raise ValueError('approval target absent from catalog/tokens')
        uncovered_prohibitions = sorted('prohibition:' + r['id'] for r in prohibition_rows if r['rule'] is None)
        unapproved_prohibitions = sorted(k for k, digest in prohibition_keys.items() if approvals.get(k) != digest)
        if stale_prohibitions:
            warnings.append(f'approvals for prohibitions no longer in the file: {len(stale_prohibitions)}')
        if unapproved_prohibitions and ship:
            errors.append(f'ship gate: unapproved design prohibitions {len(unapproved_prohibitions)}')
        elif unapproved_prohibitions:
            warnings.append(f'design prohibitions awaiting owner approval: {len(unapproved_prohibitions)}')
        if uncovered_prohibitions and ship:
            errors.append(f'ship gate: prohibitions without a declared check {len(uncovered_prohibitions)}')
        for path, token in tokens.items():
            origin = token['origin']
            if origin not in ('owner','proposal') and not (isinstance(origin,str) and origin.startswith('reference:') and origin[10:] in sources):
                errors.append(f'missing origin/NOTICE for token: {path}')
            token['state'] = 'proposed'
        rules = _rules(root)
        _resolve_style_tokens(rules,tokens,errors)
        measured_rules = _asset_rule_evidence(root,rules,errors)
        # Page rules are measured by L2 on every rendered page and enforced per item after rendering.
        uncovered_rules = {i for i,r in rules.items() if r['kind'] not in item_checks.PAGE_KINDS} - set(measured_rules)
        counts['uncovered'] = len(uncovered_rules)
        item_states, item_values, must_proposed, must_uncovered = {}, {}, [], []
        item_routes, check_map, active = {}, item_checks.load_map(), item_checks.active_checkers(tokens)
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
                result['judgment_candidates'].append(row)
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
                            result['judgment_candidates'].append(row)
                    elif kind == 'rule':
                        if target not in rules:
                            errors.append(f'missing rule evidence: {ident}')
                        else:
                            evidence_values[ref] = measured_rules.get(target,rules[target])
                            # Adopted means verified: a non-must item with an unmeasured rule fails ship too.
                            if target in uncovered_rules:
                                must_uncovered.append(ident)
            else:
                errors.append(f'blank/unknown item state: {ident}')
            item_values[ident] = evidence_values
            if status == 'defined':
                cited = state.get('evidence',[])
                way, checkers, why = item_checks.route(item, cited, tokens, rules, check_map, active)
                judged = [ref for ref in cited if ref.startswith('rule:') and rules.get(ref[5:],{}).get('kind') == 'judgment']
                item_routes[ident] = {'route':way, 'checkers':checkers, 'document':item_checks.document_key(item),
                                      **({'reason':why} if why else {}),
                                      **({'presence':judged} if way == 'afm' and judged else {}),
                                      **({'decision':True} if item_checks.mapped_route(item,check_map) == 'decision' else {})}
                if way == 'afm':
                    statements = [rules[ref[5:]]['statement'] for ref in judged]
                    result['afm_candidates'].append(afm_candidate(ident,item_context,evidence_values,statements))
        ledger=read_data(inside_layer(root,'judgments')) if layer_paths.path(root,'judgments').exists() else None
        judged=evaluate_judgments(result['judgment_candidates'],ledger,policy=judgment_policy)
        errors.extend(judged['errors'])
        warnings.extend(f"semantic advisory: {row['id']} score {row['score']} >= {row['theta']}; source review remains; no semantic guarantee" for row in judged['warnings'])
        report['uncovered'].extend(judged['uncovered']);report['screened'].extend(judged['screened'])
        for row in judged['uncovered']:
            must_uncovered.append(row['id'])
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
        # Not a warning: L0 cannot judge artifacts, and a permanent yellow would hide real warnings.
        counts['afm'] = len(result['afm_candidates'])
        counts['machine'] = sum(1 for r in item_routes.values() if r['route']=='machine')
        counts['decision'] = sum(1 for r in item_routes.values() if r.get('decision'))
        # A design decision is checked by the rule the repository declared for it; the owner also approves the item,
        # and the approval hash covers the cited rule, so changing the rule needs a new approval.
        decision_unapproved = sorted(i for i,r in item_routes.items() if r.get('decision') and r['route']!='uncovered' and any(
            resolved_by_media[m][1][i].get('approval_state')!='approved' for m in resolved_by_media))
        route_uncovered = sorted(i for i,r in item_routes.items() if r['route']=='uncovered')
        report['uncovered'].extend({'id':i,'reason':item_routes[i]['reason']} for i in route_uncovered)
        report['uncovered'].extend({'id':i,'reason':'design decision not approved'} for i in decision_unapproved)
        counts['uncovered']=len(report['uncovered'])
        if route_uncovered:
            warnings.append(f'adopted items with no way to verify them: {len(route_uncovered)}')
        if decision_unapproved:
            warnings.append(f'design decisions awaiting owner approval: {len(decision_unapproved)}')
        if counts['uncovered']:
            warnings.append(f'uncovered rules/spans: {counts["uncovered"]}; recorded or measured evidence required')
        if ship and (must_proposed or must_uncovered):
            errors.append(f'ship gate: proposed must {len(set(must_proposed))}, uncovered required evidence {len(set(must_uncovered))}')
        if ship and (route_uncovered or decision_unapproved):
            errors.append(f'ship gate: adopted items without a verification route {len(route_uncovered)}, '
                          f'unapproved design decisions {len(decision_unapproved)}')
        assets = []
        for path in sorted(layer_paths.path(root,'assets').rglob('*')):
            if path.is_file():
                inside(design,os.path.relpath(path,design))
                assets.append({'path':path.relative_to(design).as_posix(), 'sha256':hashlib.sha256(path.read_bytes()).hexdigest()})
        # Every file of every layer, the approval ledger and the judgment record included (keys are relative to the
        # layers root): a change to any of them after L0 makes the manifest stale. Generated files are not inputs.
        layers = layer_paths.path(root,'layers')
        input_hashes = {p.relative_to(layers).as_posix():hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted(layers.rglob('*')) if p.is_file() and _leaf('resolved') not in p.relative_to(layers).parts
            and layer_paths.design_input(p.relative_to(layers).as_posix())
            and p.name not in (_leaf('assets_index'),'migration.json')}
        source_hash = value_hash(input_hashes)
        asset_exports = _asset_exports(root)
        from tailwind_contract import build_artifacts
        tailwind_templates = read_data(inside_layer(root,'tailwind')) if layer_paths.path(root,'tailwind').exists() else None
        artifacts = build_artifacts(tokens,tailwind_templates)
        if hashlib.sha256(JUDGMENT_POLICY.read_bytes()).hexdigest()!=judgment_policy_sha256:
            raise ValueError('judgment policy changed during validation')
        for medium, profile in profiles.items():
            result['manifest'][medium] = {'schema_version':1, 'skill_version':VERSION,
                'approval_scheme':SCHEME,'judgment_policy_sha256':judgment_policy_sha256,
                'catalog_version':catalog['version'], 'catalog_hash':value_hash(catalog), 'source_hash':source_hash,
                'input_hashes':input_hashes, 'media':medium, 'profile':profile,
                'minimums':profile['minimums'], 'tokens':resolved_by_media[medium][0], 'items':resolved_by_media[medium][1],
                'comp':comps[medium],
                'layout_patterns':{'catalog_version':vocabulary['catalog_version'],'patterns':vocabulary['patterns']},
                'css_variables':{t['css_name']:t['css'] for t in tokens.values()},
                'prohibitions':prohibition_rows, 'prohibition_hashes':prohibition_hashes, 'tailwind_templates':tailwind_templates, 'artifacts':artifacts, 'rules':list(rules.values()), 'coverage':{'uncovered':sorted(uncovered_rules | set(uncovered_prohibitions) | {row['id'] for row in judged['uncovered']}),
                    'items':{way:sorted(i for i,r in item_routes.items() if r['route']==way) for way in ('machine','afm','uncovered')},
                    'decisions':sorted(i for i,r in item_routes.items() if r.get('decision')),
                    'item_routes':item_routes},
                'item_checks':{'collect':sorted(({c for r in item_routes.values() for c in r['checkers']+r.get('presence',[])}
                    & (item_checks.COLLECTED | {'rule:'+i for i,r in rules.items() if r['kind'] in item_checks.PAGE_KINDS})) | {'rule:'+p['rule'] for p in prohibition_rows if p['rule']}),
                    'applicability':{d:check_map['applicability'][d] for d in sorted({r['document'] for r in item_routes.values()})
                                     if d in check_map['applicability']}},
                'modules':({'ia':{'version':config['modules']['ia'],'implementation_sha256':ia['_implementation_sha256'],
                    'resolved':ia['media'][medium]}} if ia else {}),
                'palette':{kind: sorted({t['css'] for t in tokens.values() if t['type']==kind})
                           for kind in ('color','fontFamily','dimension','fontWeight','shadow')}, 'assets':assets, 'asset_exports':asset_exports}
        for medium,manifest in result['manifest'].items():
            for kind,rows in [('item',manifest['items']),('token',manifest['tokens'])]:
                for ident,row in rows.items():
                    if 'value_hash' not in row:continue
                    key=f'ds:{medium}:{kind}:{ident}'
                    result['approval_targets'][key]=row['value_hash']
                    if kind=='item' and row['must']:result['approval_required'].append(key)
        result['approval_targets'].update(page_keys)
        result['approval_targets'].update(override_keys)
        result['approval_targets'].update(vocabulary['targets'])  # an added layout pattern is approved one at a time  # not in approval_required: an exception is approved one row at a time, never with the must items
        if ia:
            result['approval_targets'].update(ia['approval_hashes'])
            result['approval_required'].extend(ia['must_approval_keys'])
        result['exit_code'] = 2 if errors else (1 if warnings else 0)
        counts['fulfillment'] = (counts['defined']+counts['na'])/counts['items']
        stage(report,'L0',result['exit_code'],counts['items'])
        for manifest in result['manifest'].values():manifest['report']=copy.deepcopy(report)
        if not errors:
            write_data(layer_paths.path(root,'assets_index'), {'schema_version':1,'assets':assets, 'asset_exports':asset_exports})
            for name, artifact in artifacts.items():
                generated = layer_paths.path(root,'resolved',artifact.get('file',name))
                generated.parent.mkdir(parents=True,exist_ok=True)
                generated.write_text(artifact['content'])
            for medium, manifest in result['manifest'].items():
                write_data(layer_paths.path(root,'resolved',f'{medium}.json'),manifest)
                css = ':root {\n' + ''.join(f'  {key}: {value};\n' for key,value in manifest['css_variables'].items()) + '}\n'
                layer_paths.path(root,'resolved',f'{medium}.css').write_text(css)
    except layer_paths.LayerLayoutError as exc:
        # A repository this version does not read: report and touch nothing (no cleanup, no generated files).
        errors.append(str(exc))
        result['exit_code'] = 2
    except (ClassificationError, ValueError, OSError, KeyError, TypeError, AttributeError) as exc:
        errors.append(str(exc).replace(str(Path(root).resolve()), '<repo>'))
        if design.is_dir() and not design.is_symlink():
            try:
                _clear_manifests(root)
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
    if not path.resolve().is_relative_to(Path(root).resolve()) or path.resolve().is_relative_to(layer_paths.path(root, 'layers').resolve()) or any(p.startswith('.') for p in Path(target).parts):
        raise ValueError('deployment target escapes artifact directories')
    return path


def _asset_exports(root):
    design = design_dir(root)
    if not layer_paths.path(root, 'asset_exports').exists():
        return []
    data = read_data(inside_layer(root, 'asset_exports'))
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
    root = Path(root)
    layer_paths.guard(root)
    writes = {}
    for target, artifact in manifest.get('artifacts',{}).items():
        writes[_deployment_path(root,target)] = artifact['content'].encode()
    for row in manifest.get('asset_exports',[]):
        content = inside(design_dir(root),row['source']).read_bytes()
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
    return Path(run.stdout.strip()).resolve()


def _override_request(root, args):
    """Write the approval request page for the layout overrides, or read the owner's saved answers back. Never writes the ledger."""
    import override_request
    layer_paths.guard(root)
    state = layout_overrides.evaluate_repo(root)
    if state.findings:
        return {'exit_code': 2, 'error': '配置の例外の行が失格している（先に直す。直せない行は承認の依頼に載せない）',
                'findings': [f'{code}: {message}' for code, message in state.findings]}
    items = override_request.build(root, state)
    if not items:
        return {'exit_code': 2, 'error': '承認の対象（配置の例外の行）が 0 件'}
    if args.read_decisions:
        decided = override_request.ar.decisions_from_rows(items, override_request.ar.load_rows(args.read_decisions))
        decided['exit_code'] = 0 if not (decided['stale'] or decided['unknown']) else 1
        return decided
    out_dir = Path(args.out)
    files = {out_dir / 'approval-request.html': override_request.page(items, args.title),
             out_dir / 'approval-request.json': json.dumps({'rows': override_request.ar.request_rows(items)}, ensure_ascii=False, indent=2)}
    existing = [str(p) for p in files if p.exists() or p.is_symlink()]
    if existing:
        return {'exit_code': 2, 'error': f'上書きしない: {existing}'}
    out_dir.mkdir(parents=True, exist_ok=True)
    for path, text in files.items():
        path.write_text(text + ('' if text.endswith('\n') else '\n'), encoding='utf-8')
    return {'exit_code': 0, 'cards': len(items), 'approved': sum(i['approved'] for i in items), 'written': sorted(str(p) for p in files),
            'publish': 'この html を非公開の Artifact として出す（capabilities: {db: {}}）。保存先は decisions/<行の id>',
            'record': 'オーナーの言葉を受けたら、design-author approve --prepare --target <キー> で束を作り、approve --request … --owner-words … --approved-on … で台帳に記録する'}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command',required=True)
    for command in ('init','check','migrate','sync','approve','maintain','migrate-layout'):
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
        if command == 'maintain':
            p.add_argument('--force',action='store_true')
        if command == 'migrate-layout':
            p.add_argument('--apply',action='store_true',help='move the files (without it: a dry run that only says what would move)')
        if command == 'approve':
            p.add_argument('--prepare',action='store_true')
            p.add_argument('--must',action='store_true')
            p.add_argument('--target',action='append',default=[])
            p.add_argument('--request',type=Path)
            p.add_argument('--approver',default='owner')
            p.add_argument('--owner-words')
            p.add_argument('--approved-on')
            p.add_argument('--date')
    rq = sub.add_parser('override-request', help='the approval request page for the layout overrides (○ / ×), or read the answers back')
    rq.add_argument('--repo', default='.')
    rq.add_argument('--out', default='override-approval-request', help='directory to write into (files are never overwritten)')
    rq.add_argument('--title', default='配置の例外の承認依頼')
    rq.add_argument('--read-decisions', help="the owner's saved answers (the Artifact db `decisions` rows, or the page's copied JSON)")
    args = parser.parse_args(argv)
    try:
        root = git_root(args.repo)
        if args.command == 'override-request':
            result = _override_request(root, args)
            print(json.dumps(result, ensure_ascii=False, sort_keys=True))
            return result['exit_code']
        if args.command == 'migrate-layout':
            import migrate_layout
            try:
                result = migrate_layout.apply(root) if args.apply else migrate_layout.plan(root)
            except migrate_layout.MigrationError as exc:
                print(json.dumps({'exit_code':2,'errors':[str(exc)]},ensure_ascii=False))
                return 2
            print(json.dumps(result,ensure_ascii=False,sort_keys=True))
            return 0
        if args.command == 'maintain':
            from maintenance import maintain_all
            result = maintain_all(root, catalog_dir=args.catalog_dir, force=args.force)
            print(json.dumps(result, ensure_ascii=False, sort_keys=True))
            return result['exit_code']
        if args.command == 'approve':
            from owner_approval import prepare,approve_request
            if args.prepare:
                if args.request:raise ValueError('prepare does not use a request')
                if args.owner_words or args.approved_on:raise ValueError('prepare does not take the owner approval')
                result=prepare(root,args.catalog_dir,args.target,args.must,args.date)
            else:
                if not args.request or args.must or args.target or args.date:
                    raise ValueError('approval recording takes --request, --approver, --owner-words and --approved-on')
                result=approve_request(root,args.catalog_dir,read_data(args.request),args.approver,
                                       owner_words=args.owner_words,approved_on=args.approved_on)
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
