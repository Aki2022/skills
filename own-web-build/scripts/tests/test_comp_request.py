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
from plan_testkit import PLAN, SKILLS, Repo, codes, lp, plan, run, write  # noqa: E402

sys.path.insert(0, str(SKILLS / 'own-image-gen/scripts'))
import image_gen  # noqa: E402  (the real field reader of the image window: the request must be accepted as-is)


@pytest.fixture
def repo(tmp_path):
    r = Repo(tmp_path)
    r.ignore_comps()
    r.l0()
    return r


def request(repo, deliverable='lp-sample'):
    rc, out, raw = run('comp_request.py', '--repo', repo.root, '--deliverable', deliverable)
    assert out is not None, raw.stdout + raw.stderr
    assert out['exit_code'] == rc
    return rc, out


def test_writes_the_field_file_the_image_window_reads_and_an_output_under_the_comp_dir(repo):
    rc, out = request(repo)
    assert rc == 0, out
    comp = repo.comp_dir()
    fields_path = repo.root / out['fields']
    assert fields_path.parent == comp and fields_path.is_file()
    fields = image_gen.load_fields(fields_path)  # exactly the 10 labels, all text
    assert list(fields) == image_gen.FIELD_ORDER
    assert [repo.root / p for p in out['out']] == [comp / 'comp.png']
    assert not (comp / 'comp.png').exists()  # no image is generated here
    saved = json.loads((comp / 'request.json').read_text())
    assert saved['width'] == 1280 and saved['out'] == out['out'] and saved['fields'] == out['fields']
    assert '--fields' in saved['command'] and '--out' in saved['command']
    assert all(not Path(a).is_absolute() for a in saved['command'])  # no local paths in what is written


def test_the_composition_follows_the_script_order_with_function_pattern_and_details(repo):
    rc, out = request(repo)
    fields = json.loads((repo.root / out['fields']).read_text())
    composition = fields['Composition']
    positions = [composition.index(b) for b in ('blk-0001', 'blk-0002', 'blk-0003', 'blk-0004', 'blk-0005')]
    assert positions == sorted(positions)  # the script's order field, not the order the file is written in
    assert 'split-2' in composition and '主張' in composition
    assert PLAN['blocks']['blk-0002']['details'] in composition
    assert 'stack-1' not in composition  # narrow shapes are not drawn (one wide image); the plan carries them


def test_the_copy_is_placeholder_text_whose_wording_is_not_used(repo):
    rc, out = request(repo)
    fields = json.loads((repo.root / out['fields']).read_text())
    content = fields['Content']
    assert '毎日の手間を減らす' in content and '準備の時間を短くします' in content
    assert content.index('毎日の手間を減らす') < content.index('準備の時間を短くします')  # copy order within a block
    assert '使わない' in content  # the comp may show the copy, but its wording is not used


def test_the_palette_and_size_come_from_the_resolved_design(repo):
    rc, out = request(repo)
    fields = json.loads((repo.root / out['fields']).read_text())
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
