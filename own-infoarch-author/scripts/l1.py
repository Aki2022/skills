"""L1: claims and terms in the text of each channel's declared outputs (token 0).

Every numeric token not removed by the non-claim table is a claim candidate: outside a claim anchor it
fails; inside one it must equal the claim's normalized value, unit and qualifiers. Lexicon hits outside
a claim anchor fail. prh rules from terms.prh.yml apply to all extracted text. Zero targets fail.
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
            entries.append((e['kind'], term, unicodedata.normalize('NFKC', term).casefold()))
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
    claims = m.data['claims']
    total_files = total_segments = 0
    for name, (rel, ch) in sorted(m.channels.items()):
        files = _files(root, ch['outputs'])
        if not files:
            report.add('fail', 'IA_L1_NO_TARGETS', f'channels/{name} の outputs {ch["outputs"]} に一致するファイルが 0 件', rel)
            continue
        for f in files:
            frel = f.relative_to(root)
            segs = extract.segments_for(f)
            if segs is None:
                report.uncovered.append({'path': str(frel), 'reason': f'形式 {f.suffix or "(なし)"} の抽出アダプタが無い'})
                report.add('warn', 'IA_L1_UNCOVERED_FORMAT', f'{frel} は抽出アダプタの無い形式（未カバー）', frel)
                continue
            total_files += 1
            total_segments += len(segs)
            for s in segs:
                _check_segment(report, m, claims, frel, s)
    report.targets = {'files': total_files, 'segments': total_segments}
    if m.channels and total_segments == 0:
        report.add('fail', 'IA_L1_NO_TARGETS', '検査したテキストが 0 件（対象 0 件を合格にしない）')
    return report


def _check_segment(report, m, claims, frel, s):
    norm = numeric.normalize(s.text)
    for tok in numeric.candidate_tokens(norm, normalized=True):
        if s.claim is None:
            report.add('fail', 'IA_L1_UNANCHORED_NUMBER',
                       f'未アンカーの数値「{tok.text}」（非主張表に無い数値は主張候補）', frel, s.line, segment=s.text)
            continue
        claim = claims.get(s.claim)
        if claim is None:
            report.add('fail', 'IA_L1_UNKNOWN_CLAIM', f'アンカー {s.claim} が登録簿に無い（墓標: {m.tombstone(s.claim)}）', frel, s.line)
            continue
        if 'value' not in claim:
            report.add('fail', 'IA_L1_CLAIM_NO_VALUE', f'{s.claim} に正規化値が無いのに数値「{tok.text}」がある', frel, s.line)
            continue
        want_value, want_unit, want_quals = numeric.claim_key(claim['value'])
        if (tok.value, tok.unit) != (want_value, want_unit):
            report.add('fail', 'IA_L1_CLAIM_VALUE_MISMATCH',
                       f'「{tok.text}」が {s.claim} の値 {want_value} {want_unit} と違う', frel, s.line)
        elif tok.qualifiers != want_quals:
            report.add('fail', 'IA_L1_CLAIM_QUALIFIER_MISMATCH',
                       f'「{tok.text}」の修飾 {sorted(tok.qualifiers)} が {s.claim} の {sorted(want_quals)} と違う', frel, s.line)
    if s.claim is None:
        folded = norm.casefold()
        for kind, term, key in _lexicon():
            if key in folded:
                report.add('fail', 'IA_L1_LEXICON_HIT',
                           f'主張の語彙「{term}」（{kind}）がアンカー無しで出ている', frel, s.line, segment=s.text)
    for hit in prh.check_text(m.rules, s.text):
        report.add('fail', 'IA_L1_TERM', f'「{hit["found"]}」→「{hit["expected"]}」', frel, s.line, origin=s.origin)
