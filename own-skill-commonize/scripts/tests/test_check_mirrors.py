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

    def test_result_line_counts_skipped_parity(self):
        """緑でも、比べなかった件数を RESULT 行に出す（空振りを隠さない）。"""
        with TemporaryDirectory() as tmp:
            rc, out = self._run(self._setup(tmp, with_local_copy=False))
            self.assertEqual(rc, 0, out)
            self.assertIn("parity compared 0, skipped 1", out)

    def _plugin_setup(self, tmp: str, *, plugin_body: str):
        root = self._setup(tmp, with_local_copy=False)
        with open(root / "mirrors.yaml", "a") as fh:
            fh.write('    plugin: "foo@market"\n    plugin_path: "skills/foo"\n')
        install = Path(tmp) / "cache" / "market" / "foo" / "1.2.3"
        _write(install / "skills" / "foo" / "SKILL.md", plugin_body)
        _write(install / "skills" / "foo" / "LICENSE.txt", "MIT\n")
        plugins_file = Path(tmp) / "installed_plugins.json"
        plugins_file.write_text(
            '{"version": 2, "plugins": {"foo@market": [{"scope": "user", '
            f'"installPath": "{install}", "version": "1.2.3", '
            '"gitCommitSha": "0123456789abcdef0123456789abcdef01234567"}]}}'
        )
        return root, plugins_file

    def _run_plugins(self, root: Path, plugins_file: Path):
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = MODULE.main([str(root), "--plugins-file", str(plugins_file)])
        return rc, buf.getvalue()

    def test_positive_control_plugin_drift_is_reported_as_warn(self):
        """陽性対照: plugin の導入版が正典と違う → DRIFT を知らせる（exit 0）。"""
        with TemporaryDirectory() as tmp:
            root, plugins = self._plugin_setup(tmp, plugin_body="---\nname: foo\n---\nnew\n")
            rc, out = self._run_plugins(root, plugins)
            self.assertEqual(rc, 0, out)
            self.assertIn("WARN drift foo", out)
            self.assertIn("0123456789abcdef0123456789abcdef01234567", out)
            self.assertIn("plugin drift 1", out)

    def test_plugin_matching_canon_is_ok(self):
        with TemporaryDirectory() as tmp:
            root, plugins = self._plugin_setup(tmp, plugin_body="---\nname: foo\n---\nbody\n")
            rc, out = self._run_plugins(root, plugins)
            self.assertEqual(rc, 0, out)
            self.assertIn("OK plugin foo", out)
            self.assertIn("plugin drift 0", out)

    def test_parser_decodes_quoted_metadata_with_spaces(self):
        entries = MODULE.parse_manifest(
            'mirrors:\n  - dir: foo\n    local_copy: "~/plugin cache/skill"\n'
        )
        self.assertEqual(entries[0]["local_copy"], "~/plugin cache/skill")


if __name__ == "__main__":
    unittest.main()
