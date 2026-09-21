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
        (r / "mirrors.yaml").write_text("mirrors: []\n")  # 正典の内在的な印
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
            _write_skill(root, "own-pptx-build")
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


class CanonDetectionTests(unittest.TestCase):
    """所属の判定が「正典そのもののパス」に依存しないことを確かめる。

    欠陥（2026-09-21 実測）: 判定が $HOME/.agents/skills との**パス一致**だったため、
    正典を別チェックアウト（worktree / CI / レビュアーの作業コピー）で lint すると
    全 skill が「リポジトリ固有＝4語」と誤判定され **25件 FAIL** した。

    ws-loop は独立レビュアーに `git worktree add --detach` で品質ゲートを走らせろと
    定めているので、この欠陥はレビュー工程を直撃する（レビュアーが必ず赤になる）。

    正典は `mirrors.yaml` を持つ（第三者 skill のミラー台帳）。リポジトリ固有の
    .agents/skills は5リポジトリのいずれも持たない。この内在的な印で判定する。
    """

    def test_canon_detected_outside_home_path(self):
        """陽性対照。HOME の外にある正典でも3語が通ること。"""
        with TemporaryDirectory() as d:
            home = Path(d) / "elsewhere"
            root = home / "checkout" / "skills"   # $HOME/.agents/skills ではない
            root.mkdir(parents=True)
            (root / "mirrors.yaml").write_text("mirrors: []\n")   # 正典の印
            _write_skill(root, "own-thing-update")
            r = _lint_home(root, home)
            self.assertEqual(
                r.returncode, 0,
                f"正典の印を持つ root は、パスが違っても3語で通るべき\n{r.stdout}{r.stderr}")

    def test_repo_local_still_requires_four_words(self):
        """陰性対照。印が無ければ従来どおりリポジトリ固有として4語を要求する。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            root = home / "repo" / ".agents" / "skills"
            root.mkdir(parents=True)
            _write_skill(root, "own-thing-update")   # 3語
            r = _lint_home(root, home)
            self.assertNotEqual(r.returncode, 0, f"印が無ければ3語は落ちるべき\n{r.stdout}{r.stderr}")
            self.assertIn("S9", r.stdout + r.stderr)
