"""配線監査が、台帳を持たずに実際に検査していることを固定する。

この監査を作った動機は、台帳（alias-roots.txt）が**人が書くものだから書き落とす**こと。
実測で台帳は 6 件だったが、正典を指す配線先は 13 件あった。落ちていたのは
~/.claude-private/skills・~/.claude-seat2/skills・~/.gemini/antigravity-cli/skills の
3経路（いずれも生きていた）と古いバックアップ4件。台帳は最初から半分しか見ていなかった。

したがって固定すべきは:
  1. 台帳に載っていない配線先を、正典から辿って**自分で見つける**
  2. 個数・不足・余分・宙を指す、の4つで落ちる
  3. 退役マーカーのついた配線先は対象にしない（常に赤い検査を作らない）
  4. 検査した配線先の数を**常に出す**
"""
import importlib.util
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

_spec = importlib.util.spec_from_file_location(
    "audit_skill_wiring", Path(__file__).resolve().parents[1] / "audit_skill_wiring.py")
mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(mod)


def _canon(tmp: Path, names: list[str]) -> Path:
    c = tmp / "canon"
    c.mkdir(parents=True)
    for n in names:
        d = c / n
        d.mkdir()
        (d / "SKILL.md").write_text(f"---\nname: {n}\ndescription: fixture.\n---\n")
    return c


def _wire(root: Path, canon: Path, names: list[str]) -> Path:
    root.mkdir(parents=True, exist_ok=True)
    for n in names:
        (root / n).symlink_to(canon / n)
    return root


def _run(canon: Path, search: Path):
    return mod.main(["--canonical", str(canon), "--search", str(search)])


class DiscoveryTests(unittest.TestCase):
    def test_finds_a_root_nobody_declared(self):
        """要点。どこにも宣言していない配線先を、正典から辿って見つける。"""
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run", "own-b-run"])
            search = tmp / "search"
            _wire(search / "some-tool" / "skills", canon, ["own-a-run", "own-b-run"])
            self.assertEqual(_run(canon, search), 0,
                             "宣言なしで見つけて、一致していれば通るべき")

    def test_retired_backup_is_not_audited(self):
        """退役マーカーのついた配線先は対象にしない（常に赤い検査を作らない）。"""
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run", "own-b-run"])
            search = tmp / "search"
            _wire(search / "tool" / "skills", canon, ["own-a-run", "own-b-run"])
            # 古い中身のバックアップ。これを見ると永久に赤くなる
            _wire(search / "tool" / "skills_backup_20260611", canon, ["own-a-run"])
            self.assertEqual(_run(canon, search), 0,
                             "退役マーカーつきは対象外であるべき")


class ParityTests(unittest.TestCase):
    def _one_root(self, tmp: Path, wired: list[str], canon_names: list[str]):
        canon = _canon(tmp, canon_names)
        search = tmp / "search"
        _wire(search / "tool" / "skills", canon, wired)
        return canon, search

    def test_missing_link_fails(self):
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon, search = self._one_root(tmp, ["own-a-run"], ["own-a-run", "own-b-run"])
            self.assertEqual(_run(canon, search), 1, "配線が足りなければ落ちるべき")

    def test_extra_name_fails(self):
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run"])
            search = tmp / "search"
            root = _wire(search / "tool" / "skills", canon, ["own-a-run"])
            (root / "own-ghost-run").symlink_to(canon / "own-a-run")  # 正典に無い名前
            self.assertEqual(_run(canon, search), 1, "正典に無い名前があれば落ちるべき")

    def test_broken_link_fails(self):
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run"])
            search = tmp / "search"
            root = _wire(search / "tool" / "skills", canon, ["own-a-run"])
            (root / "own-gone-run").symlink_to(canon / "does-not-exist")
            self.assertEqual(_run(canon, search), 1, "宙を指していれば落ちるべき")


class VacuityTests(unittest.TestCase):
    """対象0件を黙って通さない。"""

    def test_empty_canon_is_an_error(self):
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = tmp / "canon"; canon.mkdir()
            search = tmp / "search"; search.mkdir()
            self.assertEqual(_run(canon, search), 2, "正典が空なら error であるべき")

    def test_unwired_canon_is_skip_not_failure(self):
        """配線0件でも、探索起点が実在するなら「まだ配線していない」。

        新しい環境や、$HOME を差し替えて走る fixture がこれに当たる。
        ここを error にすると、それらの環境で常に赤い検査になる。
        """
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run"])
            search = tmp / "search"; search.mkdir()   # 実在するが配線が1本も無い
            self.assertEqual(_run(canon, search), 0,
                             "起点が実在して配線が無いだけなら skip であるべき")

    def test_missing_search_root_is_an_error(self):
        """探索起点そのものが実在しなければ、指定の誤り。空振りを通さない。"""
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run"])
            self.assertEqual(_run(canon, tmp / "does-not-exist"), 2,
                             "探索範囲の誤りは error であるべき")


if __name__ == "__main__":
    unittest.main()


class LintIntegrationTests(unittest.TestCase):
    """lint から呼ばれることを固定する。

    孤立したスクリプトは誰も走らせない。実際に
    `check_global_topology.py` は実運用で値を埋めて呼ぶ場所が 0 件のまま
    存在していた（呼び出しはテストと docs の説明文のみ）。同じ形にしない。

    ただし **実正典を lint したときだけ**走る。worktree を指す symlink は
    存在しないので、そこで走らせると配線先0件で常に赤くなる（実測 rc=2）。
    """

    LINT = Path(__file__).resolve().parents[1] / "skill_lint.sh"

    def test_lint_calls_the_audit(self):
        import re
        body = self.LINT.read_text()
        self.assertIn("audit_skill_wiring.py", body,
                      "lint が監査を呼んでいない。孤立したスクリプトは走らない")
        self.assertRegex(body, r"S10.*skip",
                         "検査しない場合にそれを言わないと、緑が何を意味するか読めない")

    def test_lint_skips_when_not_the_live_canon(self):
        """worktree を lint しても S10 で落ちない。"""
        import subprocess
        with TemporaryDirectory() as d:
            tmp = Path(d)
            root = tmp / "skills"
            root.mkdir()
            (root / "mirrors.yaml").write_text("mirrors: []\n")
            s = root / "own-thing-update"
            s.mkdir()
            (s / "SKILL.md").write_text(
                "---\nname: own-thing-update\ndescription: fixture.\n---\n")
            env = dict(os.environ, SKILL_LINT_SKIP_PYTEST="1")
            r = subprocess.run(["bash", str(self.LINT), str(root)],
                               capture_output=True, text=True, env=env)
            out = r.stdout + r.stderr
            self.assertIn("S10: skip", out, f"実正典でなければ skip すべき\n{out}")
            self.assertEqual(r.returncode, 0, f"skip で落ちてはいけない\n{out}")
