"""Regressions from the 2026-10-22 independent review: each input once passed (or crashed) when it should not."""
import copy
import sys
from decimal import Decimal
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import numeric  # noqa: E402
import prh  # noqa: E402
from test_checks import CHANNEL, HTML, INFO, MD, add_html, codes, run, write_repo  # noqa: E402


def l1_codes(tmp_path, html=HTML, md=MD, info=INFO, channel=CHANNEL, **kw):
    write_repo(tmp_path, info=info, channel=channel, html=html, md=md, **kw)
    return run(tmp_path)


A = '<span data-ia-claim="clm-0001">{}</span>'
B = '<span data-ia-claim="clm-0003">{}</span>'  # 100社, no qualifier
INFO_Q = copy.deepcopy(INFO)
INFO_Q['ledger']['next']['clm'] = 4
INFO_Q['claims']['clm-0003'] = copy.deepcopy(INFO['claims']['clm-0001'])
INFO_Q['claims']['clm-0003']['value']['qualifier'] = []


@pytest.mark.parametrize('snippet,code', [
    ('<p>約' + B.format('100社') + '</p>', 'IA_L1_CLAIM_QUALIFIER_MISMATCH'),            # 1 qualifier outside the anchor
    ('<p>' + B.format('100社') + '以上</p>', 'IA_L1_CLAIM_QUALIFIER_MISMATCH'),
    ('<ul><li data-ia-claim="clm-0001">100社以上<li>業界初</ul>', 'IA_L1_LEXICON_HIT'),     # 2 optional end tag
    ('<p><span data-ia-claim="clm-9999">業界初</span></p>', 'IA_L1_UNKNOWN_CLAIM'),        # 3 unknown anchor
    ('<p><span data-ia-claim="clm-9999">業界初</span></p>', 'IA_L1_LEXICON_HIT'),
    ('<p>業界\n初のサービス</p>', 'IA_L1_LEXICON_HIT'),                                    # 4 hard wrap
    ('<p>業界<br>初のサービス</p>', 'IA_L1_LEXICON_HIT'),
    ('<p>詳細はhttps://example.comから。導入企業120社</p>', 'IA_L1_UNANCHORED_NUMBER'),   # 5 url mask
    ('<p>300–1000人規模</p>', 'IA_L1_UNANCHORED_NUMBER'),                                 # 6 postal mask
    ('<p>※30日以内なら全額返金</p>', 'IA_L1_UNANCHORED_NUMBER'),                           # 7 footnote mask
    ('<p>数万人が利用</p>', 'IA_L1_UNANCHORED_NUMBER'),                                    # 9 vague
    ('<p>九割の企業</p>', 'IA_L1_UNANCHORED_NUMBER'),
    ('<p>' + B.format('最大で100社') + '</p>', 'IA_L1_CLAIM_QUALIFIER_MISMATCH'),      # 8 qualifier forms
    ('<p>' + B.format('100社を超える') + '</p>', 'IA_L1_CLAIM_QUALIFIER_MISMATCH'),
    ('<p>' + B.format('100社近く') + '</p>', 'IA_L1_CLAIM_QUALIFIER_MISMATCH'),
])
def test_html_false_greens_now_fail(tmp_path, snippet, code):
    rc, report = l1_codes(tmp_path, info=INFO_Q, html=add_html(HTML, snippet))
    assert rc == 2 and code in codes(report), report['findings']


@pytest.mark.parametrize('extra,code', [
    ('\n業界\n初のサービス\n', 'IA_L1_LEXICON_HIT'),
    ('\n約\n[100社]{data-ia-claim=clm-0003}\n', 'IA_L1_CLAIM_QUALIFIER_MISMATCH'),
    ('\n<img src="a.png" alt="業界初の画面">\n', 'IA_L1_LEXICON_HIT'),
])
def test_markdown_false_greens_now_fail(tmp_path, extra, code):
    rc, report = l1_codes(tmp_path, info=INFO_Q, md=MD + extra)
    assert rc == 2 and code in codes(report), report['findings']


def test_front_matter_block_scalar_is_read(tmp_path):
    md = '---\ntitle: 会社概要\ndescription: >\n  導入社数120社、業界初\n---\n' + MD
    rc, report = l1_codes(tmp_path, md=md)
    assert {'IA_L1_UNANCHORED_NUMBER', 'IA_L1_LEXICON_HIT'} <= set(codes(report)), report['findings']


def test_quarto_computed_values_are_uncovered(tmp_path):
    rc, report = l1_codes(tmp_path, md=MD + '\n導入は{{< var n >}}社\n')
    assert rc >= 1 and 'IA_L1_UNCOVERED_CONSTRUCT' in codes(report, 'warn') + codes(report)


def test_file_without_text_fails_even_if_other_files_have_text(tmp_path):
    write_repo(tmp_path)
    (tmp_path / 'site/app.html').write_text('<div id="root"></div><script src="x.js"></script>', encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 2 and 'IA_L1_NO_TARGETS' in codes(report)


def test_required_claim_without_anchor_fails(tmp_path):
    channel = copy.deepcopy(CHANNEL)
    channel['required_claims'] = ['clm-0001']
    rc, report = l1_codes(tmp_path, channel=channel, html=HTML.replace('data-ia-claim="clm-0001"', ''),
                          md=MD.replace('[100社以上]{data-ia-claim=clm-0001}', '導入'))
    assert rc == 2 and 'IA_L1_REQUIRED_CLAIM_MISSING' in codes(report), report['findings']


def test_empty_patterns_list_rejected():
    with pytest.raises(prh.PrhError):
        prh.load_rules({'version': 1, 'rules': [{'expected': 'A', 'patterns': []}]})


def test_word_boundary_works_next_to_japanese():
    rules = prh.load_rules({'version': 1, 'rules': [{'expected': 'jQuery', 'pattern': 'jquery', 'options': {'wordBoundary': True}}]})
    assert [h['found'] for h in prh.check_text(rules, 'jqueryを使う')] == ['jquery']
    assert list(prh.check_text(rules, 'jquerys')) == []


def test_duplicate_yaml_keys_fail(tmp_path):
    write_repo(tmp_path)
    p = tmp_path / 'design/info_arch/info_arch.yaml'
    p.write_text(p.read_text(encoding='utf-8') + 'schema_version: 1\n', encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 2 and 'IA_YAML' in codes(report)


@pytest.mark.parametrize('target', ['design/info_arch/info_arch.yaml', 'design/info_arch/terms.prh.yml'])
def test_invalid_yaml_is_exit_2_finding_not_crash(tmp_path, target):
    write_repo(tmp_path)
    (tmp_path / target).write_text('a: [unclosed\n', encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 2 and codes(report)


def test_undecodable_output_is_exit_2(tmp_path):
    write_repo(tmp_path)
    (tmp_path / 'site/sjis.html').write_bytes('<p>導入100社</p>'.encode('shift_jis'))
    rc, report = run(tmp_path)
    assert rc == 2 and 'IA_L1_DECODE' in codes(report)


def test_dotdot_inside_outputs_glob_rejected(tmp_path):
    channel = copy.deepcopy(CHANNEL)
    channel['outputs'] = ['site/../../x/*.html']
    rc, report = l1_codes(tmp_path, channel=channel)
    assert rc == 2 and 'IA_SCHEMA' in codes(report)


def test_unquoted_dates_are_accepted(tmp_path):
    write_repo(tmp_path)
    p = tmp_path / 'design/info_arch/info_arch.yaml'
    text = p.read_text(encoding='utf-8').replace("'2027-12-31'", '2027-12-31').replace("'2026-10-01'", '2026-10-01')
    assert "'2027-12-31'" not in text
    p.write_text(text, encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 0, report['findings']


def test_proposed_message_supporting_empty_claim_fails(tmp_path):
    info = copy.deepcopy(INFO)
    info['claims']['clm-0002'] = {'state': 'empty', 'provenance': 'proposal'}
    rc, report = l1_codes(tmp_path, info=info, html=HTML.replace(' data-ia-claim="clm-0002"', ''))
    assert 'IA_EMPTY_REFERENCED' in codes(report), report['findings']


def test_trm_tombstone_target_is_not_broken(tmp_path):
    info = copy.deepcopy(INFO)
    info['ledger']['next']['trm'] = 3
    info['ledger']['tombstones'] = {'trm-0001': 'trm-0002'}
    rc, report = l1_codes(tmp_path, info=info)
    assert rc == 0, report['findings']


@pytest.mark.parametrize('text', ['B2B', '5G回線', 'Windows11'])
def test_numbers_glued_to_ascii_letters_are_not_candidates(text):
    assert numeric.candidate_tokens(text) == []


def test_keiyaku_is_not_approx():
    assert numeric.candidate_tokens('契約100社')[0].qualifiers == frozenset()


def test_mixed_arabic_kanji_compound():
    toks = numeric.candidate_tokens('1万2千円')
    assert len(toks) == 1 and toks[0].value == (Decimal(12000), Decimal(12000)) and toks[0].unit == '円'


@pytest.mark.parametrize('text,quals', [('最大約100社', {'max', 'approx'}), ('税込み1,100円', {'tax_in'})])
def test_stacked_and_okurigana_qualifiers(text, quals):
    assert numeric.candidate_tokens(text)[0].qualifiers == frozenset(quals)


def test_ordinary_problem_statement_is_not_social_proof(tmp_path):
    rc, report = l1_codes(tmp_path, html=add_html(HTML, '<p>多くの企業が人手不足に悩んでいます</p>'))
    assert rc == 0, report['findings']


def test_only_uncovered_outputs_is_zero_targets_not_a_warning(tmp_path):
    channel = copy.deepcopy(CHANNEL)
    channel['outputs'] = ['site/*.tsx']
    write_repo(tmp_path, channel=channel)
    (tmp_path / 'site/App.tsx').write_text('export const A = () => <p>9割</p>;', encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 2 and 'IA_L1_NO_TARGETS' in codes(report), report['findings']


def test_crash_path_exits_2_not_1(tmp_path):
    import json
    import subprocess
    write_repo(tmp_path)
    scripts = Path(__file__).resolve().parents[1]
    proc = subprocess.run([sys.executable, str(scripts / 'infoarch_author.py'), 'check', '--root', str(tmp_path),
                           '--date', '2026-13-45'], capture_output=True, text=True, check=False)
    assert proc.returncode == 2 and json.loads(proc.stdout)['exit_code'] == 2


# ---- 2026-10-22 external review (critical): references to empty atoms passed with exit 0 ----

def test_claim_scope_referencing_empty_audience_fails(tmp_path):
    info = copy.deepcopy(INFO)
    info['ledger']['next']['aud'] = 4
    info['audiences']['aud-0003'] = {'state': 'empty', 'provenance': 'proposal'}
    info['claims']['clm-0001']['scope']['audiences'] = ['aud-0001', 'aud-0003']
    rc, report = l1_codes(tmp_path, info=info)
    assert rc == 2 and 'IA_EMPTY_REFERENCED' in codes(report), report['findings']


def test_rendition_placeholder_to_empty_claim_fails(tmp_path):
    info = copy.deepcopy(INFO)
    info['ledger']['next']['clm'] = 4
    info['claims']['clm-0003'] = {'state': 'empty', 'provenance': 'proposal'}
    info['messages']['msg-0001']['renditions']['web'] = '{{clm-0001}}の企業が導入、{{clm-0003}}'
    rc, report = l1_codes(tmp_path, info=info)
    assert rc == 2 and 'IA_EMPTY_REFERENCED' in codes(report), report['findings']


def test_broken_reference_from_empty_atom_fails(tmp_path):
    info = copy.deepcopy(INFO)
    info['ledger']['next']['msg'] = 6
    info['messages']['msg-0005'] = {'state': 'empty', 'provenance': 'proposal', 'supports': ['clm-0099']}
    rc, report = l1_codes(tmp_path, info=info)
    assert rc == 2 and 'IA_REF_BROKEN' in codes(report), report['findings']


def test_feature_count_is_a_claim_candidate():
    assert numeric.candidate_tokens('5つの機能を搭載')
    assert numeric.candidate_tokens('3つの理由') == []


def test_single_quoted_and_unquoted_raw_html_attrs_in_markdown(tmp_path):
    for attr in ("alt='業界初の画面'", 'alt=業界初'):
        rc, report = l1_codes(tmp_path, md=MD + f'\n<img src="a.png" {attr}>\n')
        assert rc == 2 and 'IA_L1_LEXICON_HIT' in codes(report), (attr, report['findings'])


def test_uppercase_submit_and_site_name_are_extracted(tmp_path):
    rc, report = l1_codes(tmp_path, html=add_html(HTML, '<input type="SUBMIT" value="業界初の申込">'))
    assert rc == 2 and 'IA_L1_LEXICON_HIT' in codes(report), report['findings']
    rc, report = l1_codes(tmp_path, html=HTML.replace('</head>', '<meta property="og:site_name" content="業界初サイト"></head>'))
    assert rc == 2 and 'IA_L1_LEXICON_HIT' in codes(report), report['findings']


# ---- 2026-10-22 external re-review (critical): anchors without text were never validated ----

INFO_E = copy.deepcopy(INFO)
INFO_E['ledger']['next']['clm'] = 4
INFO_E['claims']['clm-0003'] = {'state': 'empty', 'provenance': 'proposal'}


@pytest.mark.parametrize('snippet,code', [
    ('<p data-ia-claim="clm-0099"><img src="a.png"></p>', 'IA_L1_UNKNOWN_CLAIM'),
    ('<p data-ia-claim="clm-0003"><img src="a.png"></p>', 'IA_L1_CLAIM_EMPTY'),
    ('<p><span data-ia-claim="clm-0003"></span>導入</p>', 'IA_L1_CLAIM_EMPTY'),
])
def test_textless_html_anchors_are_validated(tmp_path, snippet, code):
    rc, report = l1_codes(tmp_path, info=INFO_E, html=add_html(HTML, snippet))
    assert rc == 2 and code in codes(report), report['findings']


@pytest.mark.parametrize('extra,code', [
    ('\n[]{data-ia-claim=clm-0003}\n', 'IA_L1_CLAIM_EMPTY'),
    ('\n::: {data-ia-claim=clm-0099}\n:::\n', 'IA_L1_UNKNOWN_CLAIM'),
])
def test_textless_markdown_anchors_are_validated(tmp_path, extra, code):
    rc, report = l1_codes(tmp_path, info=INFO_E, md=MD + extra)
    assert rc == 2 and code in codes(report), report['findings']


def test_raw_html_anchor_in_markdown_is_uncovered_not_silent(tmp_path):
    rc, report = l1_codes(tmp_path, md=MD + '\n<span data-ia-claim="clm-0099">注記</span>\n')
    assert rc >= 1 and 'IA_L1_UNCOVERED_CONSTRUCT' in codes(report, 'warn'), report['findings']


@pytest.mark.parametrize('rendition', ['{{clm0099}}の企業', '{{clm-0099a}}の企業', '{{ msg-0001 }}の企業'])
def test_malformed_placeholder_fails(tmp_path, rendition):
    info = copy.deepcopy(INFO)
    info['messages']['msg-0001']['renditions']['web'] = rendition
    rc, report = l1_codes(tmp_path, info=info)
    assert rc == 2 and 'IA_RENDITION_BAD_PLACEHOLDER' in codes(report), report['findings']


@pytest.mark.parametrize('name', ['ship.yml', 'sub/web2.yaml', 'README.md'])
def test_unexpected_files_in_channels_dir_fail(tmp_path, name):
    write_repo(tmp_path)
    f = tmp_path / 'design/info_arch/channels' / name
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text('media: ship\n', encoding='utf-8')
    rc, report = run(tmp_path)
    assert rc == 2 and 'IA_CHANNEL_FILE' in codes(report), report['findings']


# ---- 2026-10-22 third review (critical): anchor syntax the parser did not fully read passed with exit 0 ----

@pytest.mark.parametrize('line,code', [
    ('[導入]{data-ia-claim=clm-0001_old}', 'IA_L1_ANCHOR_MALFORMED'),
    ('[導入]{data-ia-claim="clm-0001 clm-0099"}', 'IA_L1_ANCHOR_MALFORMED'),
    ('[導入\n件]{data-ia-claim=clm-0099}', 'IA_L1_ANCHOR_UNPARSED'),
    ('[導入](https://example.com){data-ia-claim=clm-0099}', 'IA_L1_ANCHOR_UNPARSED'),
    ('<span DATA-IA-CLAIM="clm-0099">導入</span>', 'IA_L1_ANCHOR_UNPARSED'),
])
def test_markdown_anchor_syntax_is_read_fully_or_fails(tmp_path, line, code):
    rc, report = l1_codes(tmp_path, md=MD + '\n' + line + '\n')
    assert rc == 2 and code in codes(report), report['findings']


def test_uppercase_alt_in_markdown_raw_html_is_read(tmp_path):
    rc, report = l1_codes(tmp_path, md=MD + '\n本文。\n\n<img src="a.png" ALT="100社の画面">\n')
    assert rc == 2 and 'IA_L1_UNANCHORED_NUMBER' in codes(report), report['findings']


def test_unquoted_numeric_front_matter_is_read(tmp_path):
    rc, report = l1_codes(tmp_path, md='---\ntitle: 2026\ndescription: 100\n---\n' + MD)
    assert rc == 2 and 'IA_L1_UNANCHORED_NUMBER' in codes(report), report['findings']


@pytest.mark.parametrize('snippet,code', [
    ('<template data-ia-claim="clm-0099"><p>導入</p></template>', 'IA_L1_UNKNOWN_CLAIM'),
    ('<p data-ia-claim="clm-0099" data-ia-claim="clm-0001">導入</p>', 'IA_L1_ANCHOR_MALFORMED'),
    ('<p data-ia-claim="">導入</p>', 'IA_L1_ANCHOR_MALFORMED'),
    ('<p><span data-ia-claim="clm-0001 clm-0099">導入</span></p>', 'IA_L1_ANCHOR_MALFORMED'),
    ('<p data-ia-clam="clm-0001">導入</p>', 'IA_L1_ANCHOR_MALFORMED'),
])
def test_html_anchor_attributes_are_strict(tmp_path, snippet, code):
    rc, report = l1_codes(tmp_path, html=add_html(HTML, snippet))
    assert rc == 2 and code in codes(report), report['findings']
