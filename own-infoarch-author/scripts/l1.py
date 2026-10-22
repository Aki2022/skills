"""L1: claims and terms in the text of each channel's declared outputs (token 0).

Every numeric token not removed by the non-claim table is a claim candidate: unless it sits inside a valid
claim anchor (existing, proposed) it fails; inside one it must equal the claim's normalized value, unit and
qualifiers, read from the whole block so words just outside the anchor count. Lexicon hits outside a valid
anchor fail. prh rules apply to the raw text. A channel, file or block set with zero text fails.
"""
from __future__ import annotations

import unicodedata
from functools import lru_cache

import extract
import l0
import numeric
import prh
from common import load_ref


@lru_cache(maxsize=None)
def _lexicon():
    entries = []
    for e in load_ref('claim-lexicon.json')['entries']:
        for term in e['terms']:
            entries.append((e['kind'], term, numeric.normalize(term).casefold()))
    return entries


def _files(root, globs):
    seen = []
    for g in globs:
        for p in sorted(root.glob(g)):
            if p.is_file() and p not in seen:
                seen.append(p)
    return seen


def check(root, date):
    report, m = l0.check(root, date)
    report.level = 'l0+l1'
    if m.data is None or not m.channels:
        return report
    root = m.root
    total_files = total_blocks = 0
    for name, (rel, ch) in sorted(m.channels.items()):
        files = _files(root, ch['outputs'])
        if not files:
            report.add('fail', 'IA_L1_NO_TARGETS', f'channels/{name} の outputs {ch["outputs"]} に一致するファイルが 0 件', rel)
            continue
        anchored = set()
        for f in files:
            frel = f.relative_to(root)
            try:
                result = extract.blocks_for(f)
            except UnicodeDecodeError as exc:
                report.add('fail', 'IA_L1_DECODE', f'{frel} が UTF-8 として読めない（検査できないものを合格にしない）: {exc}', frel)
                continue
            if result is None:
                report.uncovered.append({'path': str(frel), 'reason': f'形式 {f.suffix or "(なし)"} の抽出アダプタが無い'})
                report.add('warn', 'IA_L1_UNCOVERED_FORMAT', f'{frel} は抽出アダプタの無い形式（未カバー）', frel)
                continue
            blocks, constructs = result
            for c in constructs:
                report.uncovered.append({'path': str(frel), 'reason': c})
                report.add('warn', 'IA_L1_UNCOVERED_CONSTRUCT', f'{frel}: {c}（未カバー）', frel)
            if not blocks:
                report.add('fail', 'IA_L1_NO_TARGETS', f'{frel} から検査できるテキストが 0 件（描画後にしか文言が無いページは L1 では検査できない）', frel)
                continue
            total_files += 1
            total_blocks += len(blocks)
            for b in blocks:
                anchored |= _check_block(report, m, frel, b)
        for cid in ch.get('required_claims', []):
            if cid not in anchored:
                report.add('fail', 'IA_L1_REQUIRED_CLAIM_MISSING', f'channels/{name} の必須主張 {cid} のアンカーが outputs に無い', rel)
    report.targets = {'files': total_files, 'blocks': total_blocks}
    if total_blocks == 0:
        report.add('fail', 'IA_L1_NO_TARGETS', '検査したテキストが 0 件（対象 0 件を合格にしない）')
    return report


def _valid_claims(report, m, frel, b):
    valid = set()
    for cid in sorted(b.claim_ids()):
        claim = m.data['claims'].get(cid)
        if claim is None:
            report.add('fail', 'IA_L1_UNKNOWN_CLAIM', f'アンカー {cid} が登録簿に無い（墓標: {m.tombstone(cid)}）', frel, b.line)
        elif claim['state'] != 'proposed':
            report.add('fail', 'IA_L1_CLAIM_EMPTY', f'アンカー {cid} の主張が空欄', frel, b.line)
        else:
            valid.add(cid)
    return valid


def _check_block(report, m, frel, b):
    valid = _valid_claims(report, m, frel, b)
    for tok in numeric.candidate_tokens(b.norm, normalized=True):
        cid = b.claims[tok.start] if tok.start < len(b.claims) else None
        if cid not in valid:
            report.add('fail', 'IA_L1_UNANCHORED_NUMBER',
                       f'未アンカーの数値「{tok.text}」（非主張表に無い数値は主張候補）', frel, b.line, block=b.norm)
            continue
        claim = m.data['claims'][cid]
        if 'value' not in claim:
            report.add('fail', 'IA_L1_CLAIM_NO_VALUE', f'{cid} に正規化値が無いのに数値「{tok.text}」がある', frel, b.line)
            continue
        want_value, want_unit, want_quals = numeric.claim_key(claim['value'])
        if (tok.value, tok.unit) != (want_value, want_unit):
            report.add('fail', 'IA_L1_CLAIM_VALUE_MISMATCH',
                       f'「{tok.text}」が {cid} の値 {want_value} {want_unit} と違う', frel, b.line)
        elif tok.qualifiers != want_quals:
            report.add('fail', 'IA_L1_CLAIM_QUALIFIER_MISMATCH',
                       f'「{tok.text}」の修飾 {sorted(tok.qualifiers)} が {cid} の {sorted(want_quals)} と違う', frel, b.line)
    folded = b.norm.casefold()  # casefold keeps length for the Japanese and ASCII text the lexicon targets
    for kind, term, key in _lexicon():
        start = folded.find(key)
        while start != -1:
            if len(folded) != len(b.norm) or b.claim_over(start, start + len(key)) not in valid:
                report.add('fail', 'IA_L1_LEXICON_HIT', f'主張の語彙「{term}」（{kind}）が有効な主張アンカーの外にある',
                           frel, b.line, block=b.norm)
            start = folded.find(key, start + 1)
    for hit in prh.check_text(m.rules, b.raw):
        report.add('fail', 'IA_L1_TERM', f'「{hit["found"]}」→「{hit["expected"]}」', frel, b.line, origin=b.origin)
    return valid
