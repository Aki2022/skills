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

    def _run_only_drift(self, root: Path, plugins_file: Path):
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = MODULE.main([str(root), "--plugins-file", str(plugins_file), "--only-drift"])
        return rc, buf.getvalue()

    def test_only_drift_prints_just_the_drift_line(self):
        """起動時 hook 用: ずれがある時だけ 1 行。scope 行や OK 行は出さない。"""
        with TemporaryDirectory() as tmp:
            root, plugins = self._plugin_setup(tmp, plugin_body="---\nname: foo\n---\nnew\n")
            rc, out = self._run_only_drift(root, plugins)
            self.assertEqual(rc, 0, out)
            self.assertEqual([l for l in out.splitlines() if l], [l for l in out.splitlines() if l.startswith("WARN drift foo")])
            self.assertEqual(len(out.splitlines()), 1, out)

    def test_only_drift_is_silent_without_drift(self):
        with TemporaryDirectory() as tmp:
            root, plugins = self._plugin_setup(tmp, plugin_body="---\nname: foo\n---\nbody\n")
            rc, out = self._run_only_drift(root, plugins)
            self.assertEqual(rc, 0, out)
            self.assertEqual(out, "")

    def _marketplace_setup(self, tmp: str, *, clone_body: str):
        root = self._setup(tmp, with_local_copy=False)
        with open(root / "mirrors.yaml", "a") as fh:
            fh.write('    marketplace: "mk"\n    marketplace_path: "plugins/foo-plugin/skills/foo"\n')
        clone = Path(tmp) / "marketplaces" / "mk"
        _write(clone / "plugins" / "foo-plugin" / "skills" / "foo" / "SKILL.md", clone_body)
        _write(clone / "plugins" / "foo-plugin" / "skills" / "foo" / "LICENSE.txt", "MIT\n")
        known = Path(tmp) / "known_marketplaces.json"
        known.write_text('{"mk": {"source": {"source": "github", "repo": "x/y"}, '
                         f'"installLocation": "{clone}", "lastUpdated": "2026-11-08T00:00:00Z"}}}}')
        return root, known

    def _run_marketplaces(self, root: Path, known: Path, *extra: str):
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = MODULE.main([str(root), "--marketplaces-file", str(known), *extra])
        return rc, buf.getvalue()

    def test_positive_control_marketplace_drift_is_reported_as_warn(self):
        """陽性対照: marketplace の写しが正典と違う → WARN drift（exit 0）。"""
        with TemporaryDirectory() as tmp:
            root, known = self._marketplace_setup(tmp, clone_body="---\nname: foo\n---\nnew\n")
            rc, out = self._run_marketplaces(root, known)
            self.assertEqual(rc, 0, out)
            self.assertIn("WARN drift foo: marketplace mk", out)
            self.assertIn("plugin drift 1", out)

    def test_marketplace_matching_canon_is_ok_and_only_drift_is_silent(self):
        with TemporaryDirectory() as tmp:
            root, known = self._marketplace_setup(tmp, clone_body="---\nname: foo\n---\nbody\n")
            rc, out = self._run_marketplaces(root, known)
            self.assertEqual(rc, 0, out)
            self.assertIn("OK marketplace foo", out)
            rc, out = self._run_marketplaces(root, known, "--only-drift")
            self.assertEqual(out, "")

    def test_marketplace_snapshot_without_git_uses_gcs_sha_as_pin(self):
        """公式 marketplace の写しは .git を持たず .gcs-sha に 40 桁 SHA を置く（2026-11-08 実測）。"""
        with TemporaryDirectory() as tmp:
            root, known = self._marketplace_setup(tmp, clone_body="---\nname: foo\n---\nnew\n")
            sha = "0412be76f50ab4f4f8abff3753efa9269fbbca5b"
            (Path(tmp) / "marketplaces" / "mk" / ".gcs-sha").write_text(sha)
            _, out = self._run_marketplaces(root, known)
            self.assertIn(f"marketplace mk {sha} differs", out)
            self.assertIn(f"--version {sha}", out)

    def test_marketplace_drift_shows_in_only_drift(self):
        with TemporaryDirectory() as tmp:
            root, known = self._marketplace_setup(tmp, clone_body="---\nname: foo\n---\nnew\n")
            _, out = self._run_marketplaces(root, known, "--only-drift")
            self.assertEqual(len(out.splitlines()), 1, out)
            self.assertTrue(out.startswith("WARN drift foo"), out)

    def test_parser_decodes_quoted_metadata_with_spaces(self):
        entries = MODULE.parse_manifest(
            'mirrors:\n  - dir: foo\n    local_copy: "~/plugin cache/skill"\n'
        )
        self.assertEqual(entries[0]["local_copy"], "~/plugin cache/skill")


if __name__ == "__main__":
    unittest.main()
