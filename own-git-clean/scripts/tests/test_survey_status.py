import subprocess
import tempfile
import unittest
from pathlib import Path


SURVEY = Path(__file__).resolve().parents[1] / "survey.sh"


def run(command, *, cwd=None):
    return subprocess.run(
        command,
        cwd=cwd,
        capture_output=True,
        text=True,
        check=False,
        timeout=30,
    )


class SurveyStatusTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.repo = Path(self.tmp.name) / "repo"
        self.repo.mkdir()
        run(["git", "init", "-b", "main"], cwd=self.repo)
        run(["git", "config", "user.name", "Survey Test"], cwd=self.repo)
        run(["git", "config", "user.email", "survey-test@example.invalid"], cwd=self.repo)
        run(["git", "config", "status.showUntrackedFiles", "all"], cwd=self.repo)
        (self.repo / "tracked.txt").write_text("base\n", encoding="utf-8")
        run(["git", "add", "tracked.txt"], cwd=self.repo)
        run(["git", "commit", "-m", "initial"], cwd=self.repo)

    def tearDown(self):
        self.tmp.cleanup()

    def survey(self):
        result = run(["bash", str(SURVEY), "--repo", str(self.repo)])
        self.assertEqual(result.returncode, 0, result.stderr)
        return result.stdout

    def test_clean_worktree_is_reported_clean(self):
        output = self.survey()
        working_tree = output.split("=== WORKING TREE (uncommitted / untracked) ===", 1)[1]
        working_tree = working_tree.split("=== ROOT WORKTREE COMPLETION CHECK ===", 1)[0]
        self.assertIn("(clean)", working_tree)
        self.assertIn("root_ready: yes", output)

    def test_large_dirty_worktree_is_not_reported_clean(self):
        (self.repo / "tracked.txt").write_text("changed\n", encoding="utf-8")
        untracked = self.repo / "many-untracked"
        untracked.mkdir()
        for index in range(1500):
            (untracked / f"file-{index:04d}.txt").touch()

        output = self.survey()
        working_tree = output.split("=== WORKING TREE (uncommitted / untracked) ===", 1)[1]
        working_tree = working_tree.split("=== ROOT WORKTREE COMPLETION CHECK ===", 1)[0]
        self.assertNotIn("(clean)", working_tree)
        self.assertIn(" M tracked.txt", working_tree)
        self.assertIn("?? many-untracked/file-0000.txt", working_tree)
        self.assertIn("root_ready: no", output)
        self.assertIn("dirty: yes", output)
        self.assertIn("untracked_count: 1500", output)

    def test_missing_local_main_compares_against_origin_main(self):
        run(["git", "update-ref", "refs/remotes/origin/main", "HEAD"], cwd=self.repo)
        run(["git", "branch", "-m", "own-runtime-identifiers-20260923"], cwd=self.repo)
        (self.repo / "tracked.txt").write_text("changed\n", encoding="utf-8")

        output = self.survey()
        self.assertIn("commits ahead of origin/main: 0", output)
        self.assertIn("NO COMMITS YET, working tree dirty", output)
        self.assertNotIn("PATCH-EQUIVALENT", output)

    def test_unavailable_integration_ref_is_not_patch_equivalent(self):
        run(["git", "branch", "-m", "own-runtime-identifiers-20260923"], cwd=self.repo)
        (self.repo / "tracked.txt").write_text("changed\n", encoding="utf-8")

        output = self.survey()
        self.assertIn("integration comparison ref: unavailable", output)
        self.assertIn("merge_state: UNKNOWN", output)
        self.assertNotIn("PATCH-EQUIVALENT", output)


if __name__ == "__main__":
    unittest.main()
