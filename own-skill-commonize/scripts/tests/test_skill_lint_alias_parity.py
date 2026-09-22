"""S10（別名 root の一致）の検査が、実際に検査していることを固定する。

この検査を足した動機は「検査が無かった」ことではない。
check_global_topology.py は既にあり、全 root を渡せば正しく落ちた。
足りなかったのは **実運用で値を埋めて呼ぶ場所が 0 件**だったことと、
root 一覧が人の記憶の中にしかなかったこと（~/.kiro/skills が7ヶ月見逃された）。

したがってここで固定すべきは3つ:
  1. 板が足りなければ落ちる
  2. 壊れた板があれば落ちる
  3. 一覧から root を落とすと、**検査した root の数が減って見える**
     （数を出さなければ、渡し忘れと正常の区別がつかない）
"""
import os
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "skill_lint.sh"
REFS = Path(__file__).resolve().parents[2] / "references"


def _skill(root: Path, name: str) -> None:
    d = root / name
    d.mkdir(parents=True, exist_ok=True)
    (d / "SKILL.md").write_text(
        f"---\nname: {name}\ndescription: fixture.\n---\n\n# {name}\n"
    )


def _roots_file(tmp: Path, canonical: Path, aliases: dict[str, Path]) -> Path:
    lines = [f"canonical\t{canonical}"]
    for role, path in aliases.items():
        lines.append(f"{role}\t{path}")
    f = tmp / "alias-roots.txt"
    f.write_text("\n".join(lines) + "\n")
    return f


class AliasParityTests(unittest.TestCase):
    """canonical と1つの別名 root を作り、S10 の振る舞いを確かめる。"""

    def _run(self, canonical: Path, roots_file: Path):
        # references/alias-roots.txt を差し替えて lint を走らせる
        real = REFS / "alias-roots.txt"
        backup = real.read_text() if real.exists() else None
        real.write_text(roots_file.read_text())
        try:
            env = dict(os.environ, SKILL_LINT_SKIP_PYTEST="1")
            return subprocess.run(
                ["bash", str(SCRIPT), str(canonical)],
                capture_output=True, text=True, env=env,
            )
        finally:
            if backup is None:
                real.unlink(missing_ok=True)
            else:
                real.write_text(backup)

    def test_missing_link_fails(self):
        """陽性対照。別名に板が無ければ落ちる。"""
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = tmp / "canon"; canon.mkdir()
            (canon / "mirrors.yaml").write_text("mirrors: []\n")
            _skill(canon, "own-thing-update")
            alias = tmp / "alias"; alias.mkdir()   # 板を1枚も張らない
            r = self._run(canon, _roots_file(tmp, canon, {"gemini": alias}))
            self.assertNotEqual(r.returncode, 0,
                                f"板が無ければ落ちるべき\n{r.stdout}{r.stderr}")
            self.assertIn("S10", r.stdout + r.stderr)

    def test_broken_link_fails(self):
        """陽性対照。宙を指す板があれば落ちる。"""
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = tmp / "canon"; canon.mkdir()
            (canon / "mirrors.yaml").write_text("mirrors: []\n")
            _skill(canon, "own-thing-update")
            alias = tmp / "alias"; alias.mkdir()
            (alias / "own-thing-update").symlink_to(canon / "own-thing-update")
            (alias / "origin-gone").symlink_to(canon / "does-not-exist")
            r = self._run(canon, _roots_file(tmp, canon, {"gemini": alias}))
            self.assertNotEqual(r.returncode, 0,
                                f"壊れた板があれば落ちるべき\n{r.stdout}{r.stderr}")

    def test_parity_passes_and_reports_count(self):
        """陰性。一致していれば通り、**検査した root 数を必ず出す**。"""
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = tmp / "canon"; canon.mkdir()
            (canon / "mirrors.yaml").write_text("mirrors: []\n")
            _skill(canon, "own-thing-update")
            alias = tmp / "alias"; alias.mkdir()
            (alias / "own-thing-update").symlink_to(canon / "own-thing-update")
            r = self._run(canon, _roots_file(tmp, canon, {"gemini": alias}))
            out = r.stdout + r.stderr
            self.assertEqual(r.returncode, 0, f"一致していれば通るべき\n{out}")
            self.assertIn("別名 root 1 件を検査", out,
                          f"検査した root 数を出さなければ渡し忘れに気づけない\n{out}")

    def test_dropping_a_root_lowers_the_reported_count(self):
        """要点。一覧から root を落としても緑のままだが、**数が減って見える**。

        渡し忘れた root について check_global_topology.py は何も言わない
        （実測: 渡さなければ RESULT: OK）。数がその唯一の手がかりになる。
        """
        with TemporaryDirectory() as d:
            tmp = Path(d)
            canon = tmp / "canon"; canon.mkdir()
            (canon / "mirrors.yaml").write_text("mirrors: []\n")
            _skill(canon, "own-thing-update")
            a1 = tmp / "a1"; a1.mkdir()
            a2 = tmp / "a2"; a2.mkdir()
            for a in (a1, a2):
                (a / "own-thing-update").symlink_to(canon / "own-thing-update")

            both = self._run(canon, _roots_file(tmp, canon, {"gemini": a1, "claude": a2}))
            self.assertIn("別名 root 2 件を検査", both.stdout + both.stderr)

            one = self._run(canon, _roots_file(tmp, canon, {"gemini": a1}))
            self.assertEqual(one.returncode, 0)
            self.assertIn("別名 root 1 件を検査", one.stdout + one.stderr,
                          "落とした root は黙って消える。数だけが手がかり")


if __name__ == "__main__":
    unittest.main()
