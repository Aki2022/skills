"""SessionStart hook に置く高速検査を固定する。

これは audit_skill_wiring.py（1.1 秒）と skill_lint（88 秒）の代わりではなく、
**hook に置ける唯一の粒度**として切り出したもの。見る条件は1つだけ:
per-skill 配線先に symlink でない実体があるか。

なぜその1条件か: Codex の `$skill-installer` は `$CODEX_HOME/skills/<名前>` へ
実体を書き込む。2026-09-23 に実測したところ、正典にある名前は
`InstallError: Destination already exists` で拒否され、**正典に無い名前だけが実体になる**。
実体が出来た skill は 1 席だけで使え、Claude / Gemini から見えず、git にも
mirrors.yaml にも残らない（ADR-20260906 違反）。

固定すべきは:
  1. 逸脱が無ければ**無言で rc=0**（常に赤い／常に喋る検査を作らない）
  2. 実体があれば rc=1 で、どのパスかを出す
  3. 製品所有（`.` 始まり。Codex の .system）は対象外
  4. stdin に hook の JSON を渡してもハングしない
"""
import os
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "check_wiring_fast.sh"


def _run(home: Path, stdin: str = ""):
    env = dict(os.environ, HOME=str(home))
    return subprocess.run(["bash", str(SCRIPT)], env=env, input=stdin,
                          capture_output=True, text=True, timeout=30)


def _seat(home: Path, name: str = ".codex") -> Path:
    root = home / name / "skills"
    root.mkdir(parents=True)
    return root


class FastWiringCheckTests(unittest.TestCase):
    def test_clean_wiring_is_silent(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            canon = home / ".agents" / "skills" / "own-a-run"
            canon.mkdir(parents=True)
            root = _seat(home)
            (root / "own-a-run").symlink_to(canon)
            r = _run(home)
            self.assertEqual(r.returncode, 0)
            self.assertEqual(r.stdout + r.stderr, "", "逸脱が無ければ何も言わない")

    def test_real_directory_is_reported(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            root = _seat(home)
            inst = root / "gh-fix-ci"
            inst.mkdir()
            (inst / "SKILL.md").write_text("---\nname: gh-fix-ci\n---\n")
            r = _run(home)
            self.assertEqual(r.returncode, 1)
            self.assertIn("gh-fix-ci", r.stderr)

    def test_product_owned_dot_dir_is_ignored(self):
        """Codex が自分で書き込む .system を逸脱として報告しない。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            root = _seat(home)
            (root / ".system" / "skill-installer").mkdir(parents=True)
            r = _run(home)
            self.assertEqual(r.returncode, 0, ".system は製品所有なので対象外")
            self.assertEqual(r.stdout + r.stderr, "")

    def test_hook_json_on_stdin_does_not_hang(self):
        """hook は stdin で JSON を渡す。読まない実装でも詰まらないこと。"""
        with TemporaryDirectory() as d:
            home = Path(d)
            _seat(home)
            r = _run(home, stdin='{"hook_event_name":"SessionStart"}')
            self.assertEqual(r.returncode, 0)


if __name__ == "__main__":
    unittest.main()
