"""own-infoarch-author CLI: check (L0 / L1), gen-terms, normalize. Never writes approvals."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import l0  # noqa: E402
import l1  # noqa: E402
import numeric  # noqa: E402
import prh  # noqa: E402
from common import INFO_ARCH_DIR, DependencyMissing, load_yaml  # noqa: E402


def cmd_check(args):
    date = args.date or dt.date.today().isoformat()
    report = l0.check(args.root, date)[0] if args.level == 'l0' else l1.check(args.root, date)
    print(json.dumps(report.to_dict(), ensure_ascii=False, indent=2, default=str))
    return report.exit_code


def cmd_gen_terms(args):
    root = Path(args.root)
    data = load_yaml(root / INFO_ARCH_DIR / 'info_arch.yaml')
    block = prh.render_generated(data.get('entities') or {})
    terms = root / INFO_ARCH_DIR / 'terms.prh.yml'
    if not args.write:
        print(block)
        return 0
    text = terms.read_text(encoding='utf-8') if terms.exists() else f'version: 1\nrules:\n{prh.GEN_BEGIN}\n{prh.GEN_END}\n'
    terms.write_text(prh.replace_generated(text, block), encoding='utf-8')
    print(f'updated {terms.relative_to(root)}')
    return 0


def cmd_normalize(args):
    for t in numeric.candidate_tokens(args.text):
        print(json.dumps({'text': t.text, 'value': [str(v) for v in t.value], 'unit': t.unit,
                          'qualifiers': sorted(t.qualifiers)}, ensure_ascii=False))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog='infoarch-author')
    sub = ap.add_subparsers(dest='cmd', required=True)
    c = sub.add_parser('check', help='L0 (info_arch consistency) or L1 (claims and terms in outputs)')
    c.add_argument('--root', default='.')
    c.add_argument('--level', choices=['l0', 'l1'], default='l1', help='l1 runs L0 first')
    c.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    c.set_defaults(fn=cmd_check)
    g = sub.add_parser('gen-terms', help='regenerate the entity section of terms.prh.yml')
    g.add_argument('--root', default='.')
    g.add_argument('--write', action='store_true')
    g.set_defaults(fn=cmd_gen_terms)
    n = sub.add_parser('normalize', help='show claim-candidate tokens of a text')
    n.add_argument('text')
    n.set_defaults(fn=cmd_normalize)
    args = ap.parse_args(argv)
    try:
        return args.fn(args)
    except DependencyMissing as exc:
        print(json.dumps({'exit_code': 2, 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
