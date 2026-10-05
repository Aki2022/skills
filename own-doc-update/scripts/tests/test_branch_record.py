"""branch 記録まわりの 2 つの欠陥の回帰テスト。

ISSUE-20261001-fix-branch-record-mechanisms-that-fabricate-missing:

  (1) create_issue.py が `branch: <issue id>` というプレースホルダを書き、存在しない
      枝名が記録される。消費側の check_active_issue_branches.py はそれを MISSING と
      数えるので、起票 1 件が MISSING 1 件になる。
  (2) front matter パーサが `#` コメントを剥がさない。`branch: "" # memo` が
      非空の `# memo` になり、これも MISSING に数えられる。
"""
import importlib.util
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parents[1]
SKILLS = SCRIPTS.parents[1]


def load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


VALIDATE = load("validate_repo_docs", SCRIPTS / "validate_repo_docs.py")
CHECK_BRANCHES = load(
    "check_active_issue_branches",
    SKILLS / "own-git-clean/scripts/check_active_issue_branches.py",
)


def write_front_matter(line: str) -> Path:
    path = Path(tempfile.mkdtemp()) / "x.md"
    path.write_text(f"---\n{line}\n---\n\nbody\n", encoding="utf-8")
    return path


# (入力行, 期待値)。両方のパーサが同じ答えを返さなければならない。
COMMENT_CASES = [
    ('branch: ""', ""),
    ('branch: "" # memo', ""),
    ("branch:   # memo", ""),
    ("branch: feat/x # memo", "feat/x"),
    ('branch: "feat/x" # memo', "feat/x"),
    # クォートの内側の # はコメントではない
    ('branch: "has # inside"', "has # inside"),
    ("branch: 'has # inside' # tail", "has # inside"),
    # YAML では `#` の前に空白が無ければコメントではない
    ("branch: feat/x#frag", "feat/x#frag"),
]


class FrontMatterCommentTest(unittest.TestCase):
    def test_validate_parser_strips_trailing_comments(self):
        for line, want in COMMENT_CASES:
            with self.subTest(line=line):
                got = VALIDATE.parse_front_matter(write_front_matter(line))["branch"]
                value_is_bare = line.split(":", 1)[1].split("#", 1)[0].strip() == ""
                if value_is_bare:
                    # `branch:` のように値が空だと、validate 側は「次行からのリスト」と
                    # みなして [] を返す（コメントと無関係の既存仕様）。コメントの有無で
                    # 結果が変わらないことだけを縛る。
                    bare = VALIDATE.parse_front_matter(write_front_matter("branch:"))["branch"]
                    self.assertEqual(got, bare)
                else:
                    self.assertEqual(got, want)

    def test_check_branches_parser_strips_trailing_comments(self):
        for line, want in COMMENT_CASES:
            with self.subTest(line=line):
                got = CHECK_BRANCHES.parse_front_matter(str(write_front_matter(line)))["branch"]
                self.assertEqual(got, want)

    def test_inline_list_with_trailing_comment(self):
        got = VALIDATE.parse_front_matter(write_front_matter("related_specs: [a, b] # note"))
        self.assertEqual(got["related_specs"], ["a", "b"])

    def test_both_parsers_agree(self):
        """片方だけ直すと、検査同士が食い違う（本件の元の症状）。"""
        for line, _ in COMMENT_CASES:
            with self.subTest(line=line):
                path = write_front_matter(line)
                a = VALIDATE.parse_front_matter(path)["branch"] or ""
                b = CHECK_BRANCHES.parse_front_matter(str(path))["branch"] or ""
                self.assertEqual(a, b)


class StripInlineCommentMirrorTest(unittest.TestCase):
    """strip_inline_comment は 2 つの独立した skill に同じ実装を置いている。ズレたら赤。"""

    @staticmethod
    def source_of(path: Path) -> str:
        text = path.read_text(encoding="utf-8")
        start = text.index("def strip_inline_comment(")
        end = text.index("\ndef ", start + 1)
        return text[start:end].strip()

    def test_the_two_copies_are_identical(self):
        a = self.source_of(SCRIPTS / "validate_repo_docs.py")
        b = self.source_of(SKILLS / "own-git-clean/scripts/check_active_issue_branches.py")
        self.assertEqual(a, b, "片方だけ直すと、2 つの検査が食い違う（本件の元の症状）")

    def test_both_functions_behave_the_same(self):
        for line, _ in COMMENT_CASES:
            raw = line.split(":", 1)[1].strip()
            with self.subTest(raw=raw):
                self.assertEqual(
                    VALIDATE.strip_inline_comment(raw),
                    CHECK_BRANCHES.strip_inline_comment(raw),
                )


class CreateIssueBranchTest(unittest.TestCase):
    """create_issue.py は実在しない枝名を記録してはならない。"""

    def make_repo(self) -> Path:
        root = Path(tempfile.mkdtemp())
        for sub in ("docs/adrs", "docs/specs", "docs/issues/archive",
                    "docs/workstreams/archive", "docs/guides"):
            (root / sub).mkdir(parents=True, exist_ok=True)
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-10-02\ncurrent_focus: []\n---\n\n# 00 Index\n\n"
            "## Active Issues\n\n"
            "<!-- own-doc-update:generated active-issues begin -->\n"
            "<!-- own-doc-update:generated active-issues end -->\n",
            encoding="utf-8",
        )
        return root

    def create(self, root: Path) -> Path:
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "create_issue.py"), "sample-slug",
             "--repo", str(root), "--standalone", "--priority", "low", "--due", "none",
             "--no-guide-reason", "test", "--verify-machine", "true",
             "--next-action", "run the test"],
            capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        created = list((root / "docs/issues").glob("ISSUE-*-sample-slug.md"))
        self.assertEqual(len(created), 1, result.stdout)
        return created[0]

    def test_new_issue_records_empty_branch(self):
        issue = self.create(self.make_repo())
        branch = VALIDATE.parse_front_matter(issue)["branch"]
        self.assertEqual(branch, "", f"プレースホルダ {branch!r} は存在しない枝を指す")

    def test_new_issue_is_not_reported_missing(self):
        """生成直後の issue が MISSING に数えられない（消費側の実害そのもの）。"""
        root = self.make_repo()
        issue = self.create(root)
        branch = CHECK_BRANCHES.parse_front_matter(str(issue)).get("branch", "").strip()
        self.assertFalse(branch, f"branch={branch!r} は MISSING として報告される")

    def test_template_does_not_teach_the_placeholder(self):
        template = (SKILLS / "own-doc-update/references/issue.template.md").read_text(encoding="utf-8")
        branch_lines = [l for l in template.splitlines() if l.startswith("branch:")]
        self.assertEqual(len(branch_lines), 1)
        value = branch_lines[0].split(":", 1)[1].split("#", 1)[0].strip().strip('"').strip("'")
        self.assertEqual(value, "", f"テンプレートが {branch_lines[0]!r} と教えている")


    def test_workstream_template_does_not_teach_the_placeholder(self):
        """workstream も同じ罠を持つ（WS-20260926 自身が branch_note で踏んでいた）。"""
        template = (SKILLS / "own-doc-update/references/workstream.template.md").read_text(encoding="utf-8")
        branch_lines = [l for l in template.splitlines() if l.startswith("branch:")]
        self.assertEqual(len(branch_lines), 1)
        value = branch_lines[0].split(":", 1)[1].split("#", 1)[0].strip().strip('"').strip("'")
        self.assertEqual(value, "", f"テンプレートが {branch_lines[0]!r} と教えている")


class MissingBranchWarningTest(unittest.TestCase):
    """`missing branch` は「作業中の issue で再開先の枝が分からない」ことへの警告。

    起票直後（active）や未着手（pending）の issue は、まだ枝を切っていないのが正常なので
    警告しない。警告するのは in_progress だけ。
    ISSUE-20260903-improve-loop-branch-record-checks-disagree: 空にすると警告、存在しない
    枝名を書くと MISSING、で未着手の issue に両方が黙る値が無かった。
    """

    def make_issue(self, status: str, branch: str) -> list[str]:
        root = CreateIssueBranchTest().make_repo()
        result = subprocess.run(
            [sys.executable, str(SCRIPTS / "create_issue.py"), "sample-slug",
             "--repo", str(root), "--standalone", "--priority", "low", "--due", "none",
             "--no-guide-reason", "test", "--verify-machine", "true",
             "--next-action", "run the test"],
            capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        issue = next((root / "docs/issues").glob("ISSUE-*-sample-slug.md"))
        text = issue.read_text(encoding="utf-8")
        text = text.replace("status: active", f"status: {status}", 1)
        text = text.replace('branch: ""', f"branch: {branch}" if branch else 'branch: ""', 1)
        issue.write_text(text, encoding="utf-8")
        _errors, warnings = VALIDATE.validate_repo(root)
        return [w for w in warnings if "missing branch" in w]

    def test_in_progress_without_a_branch_is_warned(self):
        self.assertEqual(len(self.make_issue("in_progress", "")), 1)

    def test_in_progress_with_a_branch_is_not_warned(self):
        self.assertEqual(self.make_issue("in_progress", "feat/x"), [])

    def test_a_fresh_active_issue_is_not_warned(self):
        """起票直後に枝が無いのは正常。ここが警告だと、起票のたびに警告が 1 件増える。"""
        self.assertEqual(self.make_issue("active", ""), [])

    def test_a_pending_issue_is_not_warned(self):
        self.assertEqual(self.make_issue("pending", ""), [])

    def test_a_blocked_issue_is_not_warned(self):
        self.assertEqual(self.make_issue("blocked", ""), [])


class MissingScopeTest(unittest.TestCase):
    """MISSING は「作業中のはずの枝が消えた」異常だけを指す（status: in_progress のみ）。

    枝は merge 後に削除する運用なので、枝を記録した issue は完了・保留を含めて全件が
    MISSING になり、母集団の 100% を指す検査は何も指していなかった
    （yorisoi_kaigo 2026-10-02: OK 0 件 / MISSING 95 件）。
    """

    CHECK = SKILLS / "own-git-clean/scripts/check_active_issue_branches.py"

    def make_repo(self, issues: dict, local_branches=()) -> Path:
        root = Path(tempfile.mkdtemp())
        git = lambda *a: subprocess.run(["git", "-C", str(root), *a], capture_output=True, text=True, timeout=30)
        git("init", "-b", "main")
        git("config", "user.name", "t")
        git("config", "user.email", "t@example.invalid")
        (root / "docs/issues").mkdir(parents=True)
        for iid, (status, branch) in issues.items():
            (root / f"docs/issues/{iid}.md").write_text(
                f"---\nschema_version: 2\nid: {iid}\nstatus: {status}\nbranch: {branch}\n---\n\nbody\n",
                encoding="utf-8")
        (root / "README").write_text("x\n")
        git("add", "-A")
        git("commit", "-m", "base")
        for b in local_branches:
            git("branch", b)
        return root

    def run_check(self, root: Path) -> str:
        r = subprocess.run([sys.executable, str(self.CHECK), str(root)], capture_output=True, text=True, timeout=30)
        self.assertEqual(r.returncode, 0, r.stderr)
        return r.stdout

    def test_only_in_progress_is_reported_missing(self):
        out = self.run_check(self.make_repo({
            "ISSUE-1-wip": ("in_progress", "feat/gone"),
            "ISSUE-2-active": ("active", "feat/gone2"),
            "ISSUE-3-pending": ("pending", "feat/gone3"),
        }))
        self.assertIn("MISSING: ISSUE-1-wip", out)
        self.assertNotIn("MISSING: ISSUE-2-active", out)
        self.assertNotIn("MISSING: ISSUE-3-pending", out)

    def test_a_live_branch_is_still_ok_for_in_progress(self):
        out = self.run_check(self.make_repo({"ISSUE-1-wip": ("in_progress", "feat/live")}, ["feat/live"]))
        self.assertIn("OK: ISSUE-1-wip -> feat/live", out)
        self.assertNotIn("MISSING", out)

    def test_a_live_branch_recorded_by_an_active_issue_is_not_an_orphan(self):
        """MISSING だけ絞って、枝の集合まで絞ってはいけない。絞ると生きた枝が孤児候補になる。"""
        out = self.run_check(self.make_repo({"ISSUE-2-active": ("active", "ISSUE-2-active")}, ["ISSUE-2-active"]))
        orphan_section = out.split("=== ORPHAN CANDIDATE BRANCHES ===")[1]
        self.assertNotIn("ISSUE-2-active (looks issue-shaped", orphan_section)

    def test_duplicate_branch_references_still_cover_every_status(self):
        out = self.run_check(self.make_repo({
            "ISSUE-1-a": ("active", "feat/shared"),
            "ISSUE-2-b": ("pending", "feat/shared"),
        }, ["feat/shared"]))
        dup = out.split("=== DUPLICATE BRANCH REFERENCES ===")[1].split("===")[0]
        self.assertIn("feat/shared", dup)


if __name__ == "__main__":
    unittest.main()
