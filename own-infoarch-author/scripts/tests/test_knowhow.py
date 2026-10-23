"""ISSUE-02: the six steps name their know-how files, every know-how file states its sources, the worked example checks green."""
import re
import subprocess
import sys
from pathlib import Path

import pytest
import yaml

SKILL = Path(__file__).resolve().parents[2]
SCRIPTS = SKILL / 'scripts'
STEPS = ['01-brief', '02-audience', '03-select-and-cut', '04-structure', '05-copy', '06-review']
MEDIA = ['lp', 'site', 'flyer-poster', 'deck', 'app-copy']
KNOWHOW = [SKILL / f'references/steps/{s}.md' for s in STEPS] + [SKILL / f'references/media/{m}.md' for m in MEDIA] + \
          [SKILL / 'references/light-process.md', SKILL / 'references/bootstrap.md']
NUMBER_WITH_UNIT = re.compile(r'[0-9０-９]+\s*(?:〜[0-9]+\s*)?(?:秒|文字|字|語|mm|pt|行|枚|個|m)(?![a-zA-Z])')


def test_skill_md_names_every_step_and_media_file_and_they_exist():
    text = (SKILL / 'SKILL.md').read_text(encoding='utf-8')
    for f in KNOWHOW:
        rel = str(f.relative_to(SKILL))
        assert rel in text, f'SKILL.md does not name {rel}'
        assert f.is_file(), rel


@pytest.mark.parametrize('path', KNOWHOW, ids=lambda p: p.name)
def test_every_knowhow_file_states_its_sources(path):
    text = path.read_text(encoding='utf-8')
    assert '## 出典' in text, f'{path.name}: no 出典 section'
    sources = text.split('## 出典', 1)[1]
    assert 'https://' in sources or '出典なし・目安' in sources, f'{path.name}: 出典 names neither a URL nor 出典なし・目安'


@pytest.mark.parametrize('path', KNOWHOW, ids=lambda p: p.name)
def test_numbers_with_units_are_marked_as_sourced_or_rough(path):
    body = path.read_text(encoding='utf-8').split('## 出典', 1)[0]
    for line in body.splitlines():
        if NUMBER_WITH_UNIT.search(line) and not line.lstrip().startswith(('#', '|')):
            assert '目安' in line or 'https://' in line or '（出典' in line, f'{path.name}: number without 目安 or source: {line}'


def _example(name):
    return yaml.safe_load((SKILL / f'references/examples/design/info_arch/outlines/{name}.yaml').read_text(encoding='utf-8'))


def test_worked_example_checks_green():
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'check', '--root',
                           str(SKILL / 'references/examples'), '--date', '2026-10-23'], capture_output=True, text=True, check=False)
    assert proc.returncode == 0, proc.stdout


def test_worked_example_two_audiences_share_content_and_differ_in_order():
    lp, deck = _example('lp-municipal'), _example('deck-smb')
    assert lp['audience'] != deck['audience']
    def msgs(o):
        return [s['message'] for p in o['pages'] for s in p['sections'] if 'message' in s]
    def text(o):
        return ' '.join(s['copy'] for p in o['pages'] for s in p['sections'])
    assert set(msgs(deck)) <= set(msgs(lp)) and msgs(deck)[0] != msgs(lp)[0]
    assert '{{fct-0001}}' in text(lp) and '{{fct-0001}}' in text(deck)
    ssot = yaml.safe_load((SKILL / 'references/examples/design/info_arch/info_arch.yaml').read_text(encoding='utf-8'))
    for o in (lp, deck):
        first = ssot['audiences'][o['audience']]['priorities'][0]
        assert first in msgs(o), f'{o["id"]}: the audience top message {first} is missing'
