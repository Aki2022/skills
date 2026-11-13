"""SessionStart hook に置く「正典のチェックアウトが main か」の検査を固定する。

正典 `~/.agents/skills` は全エージェントの読み込み元で、同時に git の作業ツリーでもある。
2026-11-12 から 11-13 にかけて、別の作業がそこで `feat/own-marketing-route` に切り替えたまま
残り、main に入った 13 コミット（IA の評価の仕組み一式）が実際の作業で読まれていなかった。
誰も気づかなかったのは、黙って古い skill が読まれるだけで、何も失敗しないから。

固定すべきは:
  1. main で origin/main と同じなら無言で rc=0（常に喋る検査を作らない）
  2. main 以外のブランチ・detached HEAD は rc=1 で、ブランチ名を出す
  3. main が（最後に fetch した）origin/main より遅れていれば rc=1 で、件数を出す
  4. 未コミットの変更だけでは警告しない（正典で別の作業の下書きが残るのは許す）
  5. 正典が git でなければ無言（配置の検査は別の script が持つ）
  6. ネットワークに出ない（fetch しない）
"""
import os
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "check_canon_branch.sh"
# git はメール欄に任意の文字列を受け付ける。アドレスの形にしない（個人情報の検出に掛からないように）
GIT_ENV = {"GIT_AUTHOR_NAME": "t", "GIT_AUTHOR_EMAIL": "fixture-owner",
           "GIT_COMMITTER_NAME": "t", "GIT_COMMITTER_EMAIL": "fixture-owner",
           "GIT_CONFIG_GLOBAL": os.devnull, "GIT_CONFIG_NOSYSTEM": "1"}


def _git(cwd: Path, *args: str) -> str:
    return subprocess.run(["git", *args], cwd=cwd, env=dict(os.environ, **GIT_ENV),
                          check=True, capture_output=True, text=True).stdout.strip()


def _run(home: Path):
    env = dict(os.environ, HOME=str(home), **GIT_ENV)
    return subprocess.run(["bash", str(SCRIPT)], env=env, input="{}",
                          capture_output=True, text=True, timeout=30)


def _canon(home: Path) -> Path:
    """origin（bare）と、それを clone した正典を作る。正典は main で origin/main と同じ。"""
    origin = home / "origin.git"
    _git(home, "init", "-q", "--bare", "-b", "main", str(origin))
    seed = home / "seed"
    _git(home, "clone", "-q", str(origin), str(seed))
    (seed / "README.md").write_text("x\n")
    _git(seed, "add", "-A")
    _git(seed, "commit", "-qm", "init")
    _git(seed, "push", "-q", "origin", "HEAD:main")
    canon = home / ".agents" / "skills"
    canon.parent.mkdir(parents=True)
    _git(home, "clone", "-q", str(origin), str(canon))
    return canon


class CanonBranchCheckTests(unittest.TestCase):
    def test_main_in_sync_is_silent(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            _canon(home)
            r = _run(home)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(r.stdout + r.stderr, "", "逸脱が無ければ何も言わない")

    def test_other_branch_is_reported(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            canon = _canon(home)
            _git(canon, "switch", "-q", "-c", "feat/x")
            r = _run(home)
            self.assertEqual(r.returncode, 1)
            self.assertIn("feat/x", r.stderr)

    def test_detached_head_is_reported(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            canon = _canon(home)
            _git(canon, "switch", "-q", "--detach")
            r = _run(home)
            self.assertEqual(r.returncode, 1)
            self.assertIn("detached", r.stderr)

    def test_main_behind_origin_is_reported(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            canon = _canon(home)
            seed = home / "seed"
            for i in range(2):
                (seed / f"f{i}").write_text("y\n")
                _git(seed, "add", "-A")
                _git(seed, "commit", "-qm", f"c{i}")
            _git(seed, "push", "-q", "origin", "HEAD:main")
            _git(canon, "fetch", "-q", "origin")  # 検査自身は fetch しない。最後の fetch の結果で比べる
            r = _run(home)
            self.assertEqual(r.returncode, 1)
            self.assertIn("2", r.stderr)

    def test_uncommitted_changes_alone_are_silent(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            canon = _canon(home)
            (canon / "README.md").write_text("draft\n")
            (canon / "new.md").write_text("draft\n")
            r = _run(home)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(r.stdout + r.stderr, "")

    def test_not_a_git_repository_is_silent(self):
        with TemporaryDirectory() as d:
            home = Path(d)
            (home / ".agents" / "skills").mkdir(parents=True)
            r = _run(home)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertEqual(r.stdout + r.stderr, "")

    def test_session_entry_runs_this_check(self):
        entry = SCRIPT.parent / "check_session_fast.sh"
        self.assertIn("check_canon_branch.sh", entry.read_text(),
                      "SessionStart の入口から呼ばれなければ毎セッションの検出にならない")


if __name__ == "__main__":
    unittest.main()
