"""check_mirrors.py の検査能力を陽性対照で確かめるテスト（赤を先に見る）。"""
import importlib.util
import io
import os
import textwrap
import unittest
from contextlib import redirect_stdout
from datetime import date, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "check_mirrors.py"
SPEC = importlib.util.spec_from_file_location("check_mirrors", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


def _write(p: Path, text: str) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text)


class CheckMirrorsTest(unittest.TestCase):
    def _setup(self, tmp: str, *, mutate_copy=False, fetched_days_ago=1, with_local_copy=True):
        root = Path(tmp) / "canon"
        copy = Path(tmp) / "upstream_copy"
        _write(root / "foo" / "SKILL.md", "---\nname: foo\n---\nbody\n")
        _write(root / "foo" / "LICENSE.txt", "MIT\n")
        _write(copy / "SKILL.md", "---\nname: foo\n---\nbody\n")
        _write(copy / "LICENSE.txt", "MIT\n")
        if mutate_copy:
            (root / "foo" / "SKILL.md").write_text("---\nname: foo\n---\nbody\n\n")  # 空行1つ
        fetched = (date.today() - timedelta(days=fetched_days_ago)).isoformat()
        local = f"    local_copy: {copy}\n" if with_local_copy else ""
        _write(root / "mirrors.yaml", textwrap.dedent(f"""\
            mirrors:
              - dir: foo
                upstream: example/skills
                upstream_path: skills/foo
                upstream_version: deadbeef
                fetched_at: {fetched}
            """) + local)
        return root

    def _run(self, root: Path, max_age_days=90):
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = MODULE.main([str(root), "--max-age-days", str(max_age_days)])
        return rc, buf.getvalue()

    def test_parity_match_and_fresh_is_green(self):
        with TemporaryDirectory() as tmp:
            rc, out = self._run(self._setup(tmp))
            self.assertEqual(rc, 0, out)
            self.assertIn("OK", out)

    def test_positive_control_parity_mismatch_is_red(self):
        """陽性対照: 正典側に空行1つ → FAIL。"""
        with TemporaryDirectory() as tmp:
            rc, out = self._run(self._setup(tmp, mutate_copy=True))
            self.assertEqual(rc, 1, out)
            self.assertIn("FAIL parity foo", out)

    def test_positive_control_stale_is_warn_not_fail(self):
        """陽性対照: fetched_at が N+1 日前 → WARN（exit 0・常時赤にしない）。"""
        with TemporaryDirectory() as tmp:
            rc, out = self._run(self._setup(tmp, fetched_days_ago=91), max_age_days=90)
            self.assertEqual(rc, 0, out)
            self.assertIn("WARN stale foo", out)

    def test_missing_local_copy_skips_parity_only(self):
        with TemporaryDirectory() as tmp:
            rc, out = self._run(self._setup(tmp, with_local_copy=False))
            self.assertEqual(rc, 0, out)
            self.assertIn("SKIP parity foo", out)

    def test_missing_manifest_is_error(self):
        with TemporaryDirectory() as tmp:
            rc, out = self._run(Path(tmp))
            self.assertEqual(rc, 2, out)

    def test_output_states_scope(self):
        """検査の範囲（何を見て何を見ないか）を出力に明記する。"""
        with TemporaryDirectory() as tmp:
            _, out = self._run(self._setup(tmp))
            self.assertIn("scope:", out)


if __name__ == "__main__":
    unittest.main()
