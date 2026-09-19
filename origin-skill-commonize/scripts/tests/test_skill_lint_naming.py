"""skill_lint.sh の S9（origin-* skill の命名規則）を陽性対照つきで確かめるテスト。

規則は ADR-20260915-unify-skill-naming-to-object-action（biz_ops）で確定:
許容型は `origin-<対象>-<動作>` の1つだけで、末尾は必ず動詞。

**実装前にこのテストを走らせて赤を見ること。** S9 が無い状態では
「FAIL S9 が出ない」ので赤になる。緑のまま通ったら、検査が対象を見ていない。

偽 skill には SKILL.md（name: がディレクトリ名と一致）を必ず置く —
そうしないと S1〜S3 で先に FAIL し、S9 の検証にならない。
"""
from __future__ import annotations

import subprocess
import textwrap
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
SKILL_LINT = SCRIPTS_DIR / "skill_lint.sh"
VERBS = SCRIPTS_DIR.parent / "references" / "naming-verbs.txt"
EXCEPTIONS = SCRIPTS_DIR.parent / "references" / "naming-exceptions.txt"


def _write_skill(root: Path, name: str) -> None:
    p = root / name / "SKILL.md"
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(
        textwrap.dedent(f"""\
            ---
            name: {name}
            description: fixture for the S9 naming test
            ---
            body
            """)
    )


def _lint(root: Path) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["bash", str(SKILL_LINT), str(root)],
        capture_output=True,
        text=True,
        env={"PATH": "/usr/bin:/bin:/usr/local/bin", "HOME": str(root), "SKILL_LINT_SKIP_PYTEST": "1"},
    )


def _lint_home(root: Path, home: Path) -> subprocess.CompletedProcess:
    """HOME を明示して lint を呼ぶ。S9 は root が $HOME/.agents/skills かどうかで
    グローバル/リポジトリ固有を判定するため、テストでは HOME を temp に固定する。"""
    return subprocess.run(
        ["bash", str(SKILL_LINT), str(root)],
        capture_output=True, text=True, stdin=subprocess.DEVNULL,
        env={"PATH": "/usr/bin:/bin:/usr/local/bin", "HOME": str(home), "SKILL_LINT_SKIP_PYTEST": "1"},
    )


class NamingRuleTests(unittest.TestCase):
    """S9 第1段: 末尾が動詞かどうか。

    2026-09-19 に接頭辞が own- に変わり、語数が所属を表すようになったので、
    このクラスはグローバル正典（3語）の形で動詞規則だけを見る。
    """

    def _global(self, home: Path) -> Path:
        r = home / ".agents" / "skills"
        r.mkdir(parents=True, exist_ok=True)
        return r

    def test_verb_ending_passes(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            root = self._global(home)
            _write_skill(root, "own-thing-update")
            r = _lint_home(root, home)
            self.assertEqual(r.returncode, 0, f"動詞末尾は通るべき\n{r.stdout}{r.stderr}")

    def test_noun_ending_fails(self):
        """陽性対照。名詞末尾（D型）は必ず赤にする。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            root = self._global(home)
            _write_skill(root, "own-thing-policy")
            r = _lint_home(root, home)
            self.assertNotEqual(r.returncode, 0, f"名詞末尾は落ちるべき\n{r.stdout}{r.stderr}")
            self.assertIn("S9", r.stdout + r.stderr)
            self.assertIn("own-thing-policy", r.stdout + r.stderr)
            self.assertIn("動詞リスト", r.stdout + r.stderr)

    def test_missing_action_fails(self):
        """対象だけ（C型）も落とす。own-<対象> の2語は不可。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            root = self._global(home)
            _write_skill(root, "own-thing")
            r = _lint_home(root, home)
            self.assertNotEqual(r.returncode, 0, f"2語は落ちるべき\n{r.stdout}{r.stderr}")
            self.assertIn("S9", r.stdout + r.stderr)

    def test_reversed_order_fails(self):
        """B型（動作-対象）も末尾が名詞になるので落ちる。

        猶予リストに載っている実在名を使うと抑止されて検査にならない。
        リストに無い名前で試すこと。
        """
        with TemporaryDirectory() as d:
            home = Path(d)
            root = self._global(home)
            _write_skill(root, "own-create-report")
            r = _lint_home(root, home)
            self.assertNotEqual(r.returncode, 0, f"逆順は落ちるべき\n{r.stdout}{r.stderr}")
            self.assertIn("S9", r.stdout + r.stderr)

    def test_third_party_skill_is_out_of_scope(self):
        """陰性対照。第三者ミラーは自前の規則の対象外で、名詞末尾でも通す。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            root = self._global(home)
            _write_skill(root, "cloudflare-one-migrations")
            r = _lint_home(root, home)
            self.assertEqual(r.returncode, 0, f"第三者 skill は対象外\n{r.stdout}{r.stderr}")

    def test_known_exception_is_suppressed(self):
        """既存の逸脱は猶予リストで抑止する（常に赤い検査にしないため）。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            root = self._global(home)
            _write_skill(root, "origin-pptx")
            r = _lint_home(root, home)
            self.assertEqual(r.returncode, 0, f"猶予リストの skill は通すべき\n{r.stdout}{r.stderr}")


class NamingDataTests(unittest.TestCase):
    def test_verb_list_exists_and_is_nonempty(self):
        self.assertTrue(VERBS.is_file(), f"{VERBS.name} が無い")
        verbs = [l.strip() for l in VERBS.read_text().splitlines() if l.strip() and not l.startswith("#")]
        self.assertGreater(len(verbs), 0)

    def test_exception_list_exists(self):
        self.assertTrue(EXCEPTIONS.is_file(), f"{EXCEPTIONS.name} が無い")

    def test_exception_entries_are_all_pre_migration_names(self):
        """猶予リストは移行前の名前だけを載せる。

        own- で始まる名前が残っていたら、その skill は既に移行済みなのに
        リストから消し忘れている＝検査が抑止されたままになる。
        （末尾が動詞かどうかでは判定できない。origin-doc-update は動詞末尾だが
        接頭辞が違うので依然として移行対象。）
        """
        if not EXCEPTIONS.is_file():
            self.skipTest("リスト未作成（実装前）")
        migrated, malformed = [], []
        for line in EXCEPTIONS.read_text().splitlines():
            name = line.strip()
            if not name or name.startswith("#"):
                continue
            if name.startswith("own-"):
                migrated.append(name)
            elif not name.startswith("origin-"):
                malformed.append(name)
        self.assertEqual(migrated, [], f"移行済みなのに猶予リストに残っている: {migrated}")
        self.assertEqual(malformed, [], f"自前 skill の名前ではない行: {malformed}")
