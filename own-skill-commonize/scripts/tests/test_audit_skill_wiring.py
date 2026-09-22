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

    def test_no_wiring_found_is_an_error(self):
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = _canon(tmp, ["own-a-run"])
            search = tmp / "search"; search.mkdir()   # 配線が1本も無い
            self.assertEqual(_run(canon, search), 2,
                             "配線先が0件なら error であるべき（探索範囲の誤りを黙って通さない）")


if __name__ == "__main__":
    unittest.main()
