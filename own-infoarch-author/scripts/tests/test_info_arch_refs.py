"""INFO_ARCH.md's ids (SPEC-information-architecture-standardization requirement 1: the prose refers to atoms by id and
holds no value) and the vertical-order check being a hand-run of a type change, not a per-check tool
(SPEC-product-ui-infoarch Quality, ADR-20261123-ia-skill-completion-judgments-and-measurement)."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from blueprint_fixtures import write_blueprint  # noqa: E402
from ia_fixtures import SSOT  # noqa: E402
from test_check import DATE, write_repo  # noqa: E402
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent


def run(command, root):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), command, '--root', str(root), '--date', DATE],
                          capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


def prose(root, text):
    path = layer_paths.path(root, 'infoarch', 'INFO_ARCH.md')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding='utf-8')


def finds(report, level=None):
    return [(f['level'], f['code'], f['detail'].get('id')) for f in report['findings']
            if f['code'].startswith('IA_INFO_ARCH') and (level is None or f['level'] == level)]


def tree(root):
    """An outline repository with a product UI as well: every kind of id the prose may cite exists."""
    write_repo(root)
    write_blueprint(root)


GREEN = ('# 情報設計\n\n提供価値 val-0001 は `script:lp-municipal` の一番のメッセージが伝える（v6 の書き方 `outline:deck-smb` も読む）。\n'
         '画面 scr-overview はストーリー sto-0001 に応える。原則 pri-0001・セグメント seg-0001。\n')


def test_ids_that_exist_pass(tmp_path):
    tree(tmp_path)
    prose(tmp_path, GREEN)
    for command in ('check', 'blueprint'):
        rc, report = run(command, tmp_path)
        assert rc == 0 and finds(report) == [], (command, report['findings'])


@pytest.mark.parametrize('mention, kind', [
    ('val-0099', 'a positioning atom'), ('seg-0042', 'a segment'), ('sto-0099', 'a product story'),
    ('pri-0099', 'a product principle'), ('scr-nowhere', 'a screen'), ('outline:no-such-outline', 'an outline'),
    ('script:no-such-script', 'a script'),
])
def test_an_id_that_does_not_exist_fails(tmp_path, mention, kind):
    tree(tmp_path)
    prose(tmp_path, GREEN + f'\nあと {mention} も大事。\n')
    for command in ('check', 'blueprint'):
        rc, report = run(command, tmp_path)
        wanted = mention.split(':')[-1]
        assert rc == 2 and finds(report, 'fail') == [('fail', 'IA_INFO_ARCH_REF_BROKEN', wanted)], (kind, command, report['findings'])
        [f] = [f for f in report['findings'] if f['code'] == 'IA_INFO_ARCH_REF_BROKEN']
        assert f['path'].endswith('INFO_ARCH.md') and f['line'] == 6, f  # the line to fix is named


def test_a_retired_id_is_a_warning_not_a_failure(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['ledger']['next']['val'] = 5
    ssot['ledger']['tombstones']['val-0004'] = None  # retired: never reused, and nothing to point at any more
    write_repo(tmp_path, ssot)
    prose(tmp_path, '# 情報設計\n\nval-0004 は取り下げた。\n')
    rc, report = run('check', tmp_path)
    assert rc == 1 and finds(report) == [('warn', 'IA_INFO_ARCH_REF_TOMBSTONED', 'val-0004')], report['findings']


def test_legacy_ids_and_plain_words_are_not_references(tmp_path):
    """v5 ids (msg-…) are the migration's business (prose_mentions); words that merely look like an id are not ids."""
    tree(tmp_path)
    prose(tmp_path, GREEN + '\n旧い msg-0001 と aud-0002。dev-valve-0001 は別の物。value-0001x も違う。\n')
    rc, report = run('check', tmp_path)
    assert rc == 0 and finds(report) == [], report['findings']


def test_every_mention_is_reported_once_per_id_with_its_first_line(tmp_path):
    tree(tmp_path)
    prose(tmp_path, '# 情報設計\n\nval-0099 と val-0099 をくりかえす。\n\nval-0098 も。\n')
    rc, report = run('check', tmp_path)
    assert rc == 2
    broken = [(f['detail']['id'], f['line']) for f in report['findings'] if f['code'] == 'IA_INFO_ARCH_REF_BROKEN']
    assert broken == [('val-0099', 3), ('val-0098', 5)]


def test_no_prose_file_is_not_a_finding(tmp_path):
    tree(tmp_path)
    assert not layer_paths.path(tmp_path, 'infoarch', 'INFO_ARCH.md').exists()
    rc, report = run('check', tmp_path)
    assert rc == 0 and finds(report) == []


def test_the_template_prose_passes_a_new_repository(tmp_path):
    """The prose a new repository starts from names no id, so it never fails on its first check."""
    from test_check import install_templates
    install_templates(tmp_path)
    rc, report = run('check', tmp_path)
    assert finds(report) == [], report['findings']


# ---- the vertical order is a hand-run of a type change (SPEC-product-ui-infoarch Quality) ----------------------

def test_skill_and_references_say_the_vertical_order_is_tried_by_hand_when_the_type_changes():
    skill = (SKILL / 'SKILL.md').read_text(encoding='utf-8')
    contract = (SKILL / 'references/contract.md').read_text(encoding='utf-8')
    for name, text in (('SKILL.md', skill), ('references/contract.md', contract)):
        assert '縦の順' in text, name
        assert 'blueprint の型' in text and '手で一度' in text, name  # when, and how: a type change, once, by hand
        assert '毎回の検査にはしない' in text, name
        assert '割り付け' in text, name  # what checks the rendered order every time: design's allocation check
    # the two never describe a tool or a per-check step for it
    for text in (skill, contract):
        assert 'vertical-order' not in text and 'infoarch-author vertical' not in text
