"""The comp request (SPEC-deliverable-production-layers requirements 14, 22, 23).

comp_request.py --repo <repo> --deliverable <id> writes own-image-gen's field file and the output path for one wide
image of the whole page, under the git-ignored comp dir, sized from the resolved profile. It never generates an image
and refuses (exit 2) when the plan check fails."""
import copy
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from plan_testkit import PLAN, SCRIPT, SKILLS, Repo, brief, codes, lp, plan, run, write  # noqa: E402

sys.path.insert(0, str(SKILLS / 'own-image-gen/scripts'))
import image_gen  # noqa: E402  (the real field reader of the image window: the request must be accepted as-is)


@pytest.fixture
def repo(tmp_path):
    r = Repo(tmp_path)
    r.ignore_comps()
    r.l0()
    return r


def request(repo, deliverable='lp-sample', *args):
    rc, out, raw = run('comp_request.py', '--repo', repo.root, '--deliverable', deliverable, *args)
    assert out is not None, raw.stdout + raw.stderr
    assert out['exit_code'] == rc
    return rc, out


def fields_of(repo, out, part=0):
    return json.loads((repo.root / out['fields'][part]).read_text())


def test_writes_the_field_file_the_image_window_reads_and_an_output_under_the_comp_dir(repo):
    rc, out = request(repo)
    assert rc == 0, out
    comp = repo.comp_dir()
    assert len(out['fields']) == 1  # one image by default (the comp settings' parts is 1)
    fields_path = repo.root / out['fields'][0]
    assert fields_path.parent == comp and fields_path.is_file()
    fields = image_gen.load_fields(fields_path)  # exactly the 10 labels, all text
    assert list(fields) == image_gen.FIELD_ORDER
    assert [repo.root / p for p in out['out']] == [comp / 'comp.png']
    assert not (comp / 'comp.png').exists()  # no image is generated here
    saved = json.loads((comp / 'request.json').read_text())
    assert saved['width'] == 1280 and saved['out'] == out['out'] and saved['fields'] == out['fields']
    assert len(saved['commands']) == 1 and saved['composition'] == 'plan'
    assert '--fields' in saved['commands'][0] and '--out' in saved['commands'][0]
    assert all(not Path(a).is_absolute() for a in saved['commands'][0])  # no local paths in what is written


def test_the_composition_follows_the_script_order_with_function_pattern_and_details(repo):
    rc, out = request(repo)
    fields = fields_of(repo, out)
    composition = fields['Composition']
    positions = [composition.index(b) for b in ('blk-0001', 'blk-0002', 'blk-0003', 'blk-0004', 'blk-0005')]
    assert positions == sorted(positions)  # the script's order field, not the order the file is written in
    assert 'split-2' in composition and '主張' in composition
    assert PLAN['blocks']['blk-0002']['details'] in composition
    assert 'stack-1' not in composition  # narrow shapes are not drawn (one wide image); the plan carries them


def test_the_copy_is_placeholder_text_whose_wording_is_not_used(repo):
    rc, out = request(repo)
    fields = fields_of(repo, out)
    content = fields['Content']
    assert '毎日の手間を減らす' in content and '準備の時間を短くします' in content
    assert content.index('毎日の手間を減らす') < content.index('準備の時間を短くします')  # copy order within a block
    assert '使わない' in content  # the comp may show the copy, but its wording is not used


def test_the_palette_and_size_come_from_the_resolved_design(repo):
    rc, out = request(repo)
    fields = fields_of(repo, out)
    assert '#003366' in fields['Color palette'] and 'color.primary' in fields['Color palette']
    assert '1280' in fields['Asset type'] and '1280' in fields['Constraints']


def test_a_profile_override_changes_the_size(repo):
    write(lp.path(repo.root, 'design_dir', 'profiles/web.json'),
          {'schema_version': 1, 'media': 'web', 'minimums': {}, 'comp': {'width': 1440}})
    repo.l0()
    rc, out = request(repo)
    assert rc == 0 and out['width'] == 1440


def test_refuses_when_the_plan_check_fails_and_writes_nothing(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    del blocks['blk-0004']
    repo.plan(plan(blocks=blocks))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_PLAN_FAILED' in codes(out, 'fail')
    assert 'PLAN_BLOCK_MISSING' in json.dumps(out['plan_check'], ensure_ascii=False)
    assert not repo.comp_dir().exists()


def test_a_plan_warning_does_not_stop_the_request(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    blocks['blk-0005']['pattern'] = 'quotes'
    repo.plan(plan(blocks=blocks))
    rc, out = request(repo)
    assert rc == 1 and 'PLAN_PATTERN_UNSUITED' in json.dumps(out['plan_check'], ensure_ascii=False)
    assert (repo.comp_dir() / 'fields.json').is_file()


def test_refuses_a_comp_dir_git_would_commit(tmp_path):
    repo = Repo(tmp_path)  # no ux/build/.gitignore
    repo.l0()
    rc, out = request(repo)
    assert rc == 2 and 'COMP_DIR_NOT_IGNORED' in codes(out, 'fail')
    assert not repo.comp_dir().exists()


def test_outside_git_the_request_is_written_and_says_the_ignore_was_not_checked(tmp_path):
    repo = Repo(tmp_path, git=False)
    repo.l0()
    rc, out = request(repo)
    assert rc == 1 and 'COMP_IGNORE_NOT_CHECKED' in codes(out, 'warn')
    assert (repo.comp_dir() / 'fields.json').is_file()


def _edit_manifest(repo, change):
    path = lp.path(repo.root, 'resolved', 'web.json')
    data = json.loads(path.read_text())
    change(data)
    path.write_text(json.dumps(data))  # the manifest's inputs are unchanged, so it still reads as fresh


def test_a_manifest_without_a_comp_size_is_refused(repo):
    _edit_manifest(repo, lambda d: d.pop('comp'))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_SIZE_MISSING' in codes(out, 'fail')
    assert not (repo.comp_dir() / 'fields.json').exists()


def test_an_unexpected_error_is_a_refusal_never_a_warning(repo):
    """An uncaught exception would exit 1, which reads as 'written with warnings': it must be 2."""
    _edit_manifest(repo, lambda d: d['layout_patterns']['patterns']['band'].pop('description'))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_REQUEST_FAILED' in codes(out, 'fail')


# --- A. the build brief ---------------------------------------------------------------------------------------------

def test_a_missing_brief_is_refused_and_nothing_is_written(repo):
    repo.brief_path().unlink()
    rc, out = request(repo)
    assert rc == 2 and 'COMP_BRIEF_MISSING' in codes(out, 'fail')
    assert not repo.comp_dir().exists()


@pytest.mark.parametrize('change', [{'tokens': {'exclude': 'color.*'}}, {'extra': 1}, {'deliverable': 'other'},
                                    {'style_references': ['/abs/ref.png']}, {'schema_version': 2}])
def test_a_brief_out_of_shape_is_refused(repo, change):
    repo.brief(brief(**change), deliverable='lp-sample')
    rc, out = request(repo)
    assert rc == 2 and 'COMP_BRIEF_INVALID' in codes(out, 'fail')
    assert not repo.comp_dir().exists()


def _themes(repo):
    """Three brand themes side by side in one design (color.a.*, color.b.*, color.c.*), the way a multi-brand repo has."""
    path = lp.path(repo.root, 'design_dir', 'tokens/themes.tokens.json')
    theme = lambda hexa, comps: {'$value': {'colorSpace': 'srgb', 'components': comps, 'hex': hexa},  # noqa: E731
                                 '$extensions': {'org.own.design': {'origin': 'proposal'}}}
    write(path, {'color': {'$type': 'color',
                           'a': {'accent': theme('#112233', [0x11 / 255, 0x22 / 255, 0x33 / 255])},
                           'b': {'accent': theme('#aa0000', [0xaa / 255, 0, 0])},
                           'c': {'accent': theme('#00aa00', [0, 0xaa / 255, 0])}}})
    repo.l0()


def test_the_brief_token_filter_keeps_one_theme_of_a_multi_theme_design(repo):
    _themes(repo)
    rc, out = request(repo)
    every = fields_of(repo, out)['Color palette']
    assert '#aa0000' in every and '#00aa00' in every  # without a filter every theme is passed (the defect)
    repo.brief(brief(tokens={'exclude': ['*.b.*', '*.c.*']}))
    rc, out = request(repo)
    assert rc == 0, out
    palette = fields_of(repo, out)['Color palette']
    assert '#112233' in palette and '#003366' in palette
    assert '#aa0000' not in palette and '#00aa00' not in palette and 'color.b.' not in palette


def test_an_include_filter_keeps_only_the_matching_tokens(repo):
    _themes(repo)
    repo.brief(brief(tokens={'include': ['color.a.*', 'font.*']}))
    rc, out = request(repo)
    fields = fields_of(repo, out)
    assert fields['Color palette'] == 'color.a.accent #112233'
    assert 'Sample Sans' in fields['Style']


def test_a_filter_leaving_no_color_is_refused(repo):
    repo.brief(brief(tokens={'include': ['font.*']}))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_BRIEF_NO_COLOR' in codes(out, 'fail')
    assert not repo.comp_dir().exists()


def test_a_glob_that_matches_no_token_warns(repo):
    repo.brief(brief(tokens={'exclude': ['colour.*']}))  # a typo must not read as a filter that worked
    rc, out = request(repo)
    assert rc == 1 and 'COMP_BRIEF_GLOB_UNUSED' in codes(out, 'warn')


def test_style_references_are_passed_as_anchors(repo):
    ref = repo.root / 'references/comp-style.png'  # outside design: an anchor is not a design input
    ref.parent.mkdir(parents=True, exist_ok=True)
    ref.write_bytes(b'\x89PNG fixture')
    rel = ref.relative_to(repo.root).as_posix()
    repo.brief(brief(style_references=[rel]))
    rc, out = request(repo)
    assert rc == 0, out
    fields = fields_of(repo, out)
    assert rel in fields['Style references']
    command = json.loads((repo.comp_dir() / 'request.json').read_text())['commands'][0]
    assert command[command.index('--anchor') + 1] == rel
    # image_gen checks an anchor from the repo root and codex reads it from the workdir: both must be the repo root
    assert command[command.index('--workdir') + 1] == '.'
    assert rel in image_gen.build_prompt(image_gen.load_fields(repo.root / out['fields'][0]), [rel])


def test_a_missing_style_reference_is_refused(repo):
    repo.brief(brief(style_references=['references/none.png']))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_BRIEF_REFERENCE_MISSING' in codes(out, 'fail')


def test_no_style_reference_is_said(repo):
    rc, out = request(repo)
    assert rc == 0 and 'COMP_NO_STYLE_REFERENCE' in codes(out, 'info')


def test_emphasis_goes_into_the_line_of_its_block(repo):
    repo.brief(brief(emphasis={'blk-0003': '数字を大きく見せる'}))
    rc, out = request(repo)
    lines = fields_of(repo, out)['Composition'].splitlines()
    line = next(x for x in lines if 'blk-0003' in x)
    assert '数字を大きく見せる' in line
    assert not any('数字を大きく見せる' in x for x in lines if 'blk-0003' not in x)


def test_emphasis_on_an_unknown_block_is_refused(repo):
    repo.brief(brief(emphasis={'blk-0009': '大きく'}))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_BRIEF_BLOCK_UNKNOWN' in codes(out, 'fail')


def test_a_brief_for_another_script_than_the_plan_is_refused(repo):
    repo.brief(brief(script='other-script'))
    rc, out = request(repo)
    assert rc == 2 and 'COMP_BRIEF_SCRIPT_MISMATCH' in codes(out, 'fail')


# --- B. the style comes from the script -----------------------------------------------------------------------------

def test_the_style_is_the_scripts_tone_target_and_objective(repo):
    rc, out = request(repo)
    style = fields_of(repo, out)['Style']
    assert SCRIPT['tone'] in style and SCRIPT['target']['condition'] in style and SCRIPT['objective']['action'] in style
    assert '落ち着いた完成見本' not in style  # no mood words of the skill's own
    assert '色と書体は下の値だけ' in style


# --- C. dictionary values -------------------------------------------------------------------------------------------

def _with_copy(repo, text, cid='cpy-0003'):
    script = copy.deepcopy(SCRIPT)
    script['copies'][cid]['text'] = text
    repo.script(script)
    repo.l0()


def test_dictionary_values_are_filled_in(repo):
    _with_copy(repo, '導入は {{value-customer-count}}')
    rc, out = request(repo)
    assert rc == 0, out
    content = fields_of(repo, out)['Content']
    assert '導入は 100社以上' in content and '{{' not in content


def test_an_unresolved_value_is_kept_and_warned(repo, monkeypatch):
    sys.path.insert(0, str(SKILLS / 'own-web-build/scripts'))
    import comp_request
    _with_copy(repo, '導入は {{value-customer-count}}')
    monkeypatch.setattr(comp_request, 'dictionary_values', lambda root: ({}, []))
    out = comp_request.comp_request(repo.root, 'lp-sample')
    assert out['exit_code'] == 1 and 'COMP_VALUE_UNRESOLVED' in codes(out, 'warn')
    content = json.loads((repo.root / out['fields'][0]).read_text())['Content']
    assert '{{value-customer-count}}' in content


def test_fill_values_reports_what_it_could_not_fill():
    sys.path.insert(0, str(SKILLS / 'own-web-build/scripts'))
    import comp_request
    text, missing = comp_request.fill_values('{{value-a}} と {{value-b}}', {'value-a': '一'})
    assert text == '一 と {{value-b}}' and missing == ['value-b']


# --- D. long copy is drawn as gray lines ----------------------------------------------------------------------------

def test_copy_longer_than_text_max_chars_is_marked_as_gray_lines(repo):
    write(lp.path(repo.root, 'design_dir', 'profiles/web.json'),
          {'schema_version': 1, 'media': 'web', 'minimums': {}, 'comp': {'text_max_chars': 8}})
    repo.l0()
    rc, out = request(repo)
    assert rc == 0, out
    lines = fields_of(repo, out)['Content'].splitlines()
    long_line = next(x for x in lines if '準備の時間を短くします' in x)
    assert '灰色の線' in long_line  # 11 characters > 8: listed, but drawn as placeholder lines
    short = next(x for x in lines if 'サンプル事業' in x)
    assert '灰色の線' not in short
    saved = json.loads((repo.comp_dir() / 'request.json').read_text())
    assert saved['text_max_chars'] == 8


def test_the_default_text_max_chars_comes_from_the_skill(repo):
    rc, out = request(repo)
    defaults = json.loads((SKILLS / 'own-design-author/references/comp-sizes.json').read_text())['web']
    saved = json.loads((repo.comp_dir() / 'request.json').read_text())
    assert saved['text_max_chars'] == defaults['text_max_chars'] and saved['parts'] == defaults['parts'] == 1
    assert '［本文' not in fields_of(repo, out)['Content']  # the kit's copies are all short


# --- E. parts -------------------------------------------------------------------------------------------------------

def test_parts_split_the_blocks_in_order_into_balanced_contiguous_groups(repo):
    rc, out = request(repo, 'lp-sample', '--parts', '2')
    assert rc == 0, out
    assert [Path(p).name for p in out['fields']] == ['fields-1.json', 'fields-2.json']
    assert [Path(p).name for p in out['out']] == ['comp-1.png', 'comp-2.png']
    first, second = fields_of(repo, out, 0), fields_of(repo, out, 1)
    assert all(b in first['Composition'] for b in ('blk-0001', 'blk-0002', 'blk-0003'))
    assert not any(b in first['Composition'] for b in ('blk-0004', 'blk-0005'))
    assert all(b in second['Composition'] for b in ('blk-0004', 'blk-0005'))
    assert '2 枚のうち 1 枚目' in first['Asset type'] and 'blk-0001' in first['Asset type']
    assert '2 枚のうち 2 枚目' in second['Asset type'] and 'blk-0005' in second['Asset type']
    saved = json.loads((repo.comp_dir() / 'request.json').read_text())
    assert saved['parts'] == 2 and len(saved['commands']) == 2
    assert [p['blocks'] for p in saved['part_list']] == [['blk-0001', 'blk-0002', 'blk-0003'], ['blk-0004', 'blk-0005']]
    for part, command in enumerate(saved['commands']):
        assert command[command.index('--fields') + 1] == saved['fields'][part]
        assert command[command.index('--out') + 1] == saved['out'][part]
        image_gen.load_fields(repo.root / saved['fields'][part])


def test_a_new_request_removes_the_field_files_of_an_earlier_split(repo):
    request(repo, 'lp-sample', '--parts', '3')
    rc, out = request(repo)
    assert sorted(p.name for p in repo.comp_dir().glob('fields*.json')) == ['fields.json']


def test_the_profile_sets_the_default_parts(repo):
    write(lp.path(repo.root, 'design_dir', 'profiles/web.json'),
          {'schema_version': 1, 'media': 'web', 'minimums': {}, 'comp': {'parts': 3}})
    repo.l0()
    rc, out = request(repo)
    assert rc == 0 and len(out['fields']) == 3


@pytest.mark.parametrize('parts', ['0', '6'])
def test_more_parts_than_blocks_or_none_is_refused(repo, parts):
    rc, out = request(repo, 'lp-sample', '--parts', parts)
    assert rc == 2 and 'COMP_PARTS_INVALID' in codes(out, 'fail')


# --- F. composition mode --------------------------------------------------------------------------------------------

def test_free_composition_needs_no_plan_and_lets_the_model_propose(repo):
    lp.path(repo.root, 'build', 'lp-sample', 'plan.yaml').unlink()
    repo.brief(brief(emphasis={'blk-0002': '絵を大きく裁ち落とす'}))
    rc, out = request(repo, 'lp-sample', '--composition', 'free')
    assert rc == 0, out
    fields = fields_of(repo, out)
    composition = fields['Composition']
    for name in ('split-2', 'header-bar', 'stats-row', 'source-footer'):
        assert name not in composition
    assert '提案' in composition and '主張' in composition and '絵を大きく裁ち落とす' in composition
    positions = [composition.index(b) for b in ('blk-0001', 'blk-0002', 'blk-0003', 'blk-0004', 'blk-0005')]
    assert positions == sorted(positions)
    assert '型' not in fields['Constraints']
    saved = json.loads((repo.comp_dir() / 'request.json').read_text())
    assert saved['composition'] == 'free' and out['composition'] == 'free'


def test_plan_composition_still_keeps_each_pattern(repo):
    rc, out = request(repo, 'lp-sample', '--composition', 'plan')
    fields = fields_of(repo, out)
    assert '型' in fields['Constraints'] and 'split-2' in fields['Composition']


def test_plan_composition_without_a_plan_is_refused(repo):
    lp.path(repo.root, 'build', 'lp-sample', 'plan.yaml').unlink()
    rc, out = request(repo)
    assert rc == 2 and 'COMP_PLAN_FAILED' in codes(out, 'fail')
