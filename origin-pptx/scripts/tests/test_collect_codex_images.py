import importlib.util
import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPT = Path(__file__).resolve().parents[1] / "collect_codex_images.py"
SPEC = importlib.util.spec_from_file_location("collect_codex_images", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class SelectFilesTest(unittest.TestCase):
    """edit-mode では入力画像のコピーも session dir に残るため、生成枚数 > 出力名の数になる。

    既定は今までどおり件数不一致で止める（黙って間引かない）。救済は明示フラグでのみ効く。
    """

    def setUp(self):
        self.tmp = TemporaryDirectory()
        self.session = Path(self.tmp.name)
        # mtime 昇順が生成順。edit-mode では [0] が入力コピー、[1] が生成物。
        self.files = []
        for i, name in enumerate(["base.png", "generated.png", "generated2.png"]):
            p = self.session / name
            p.write_bytes(b"x")
            os.utime(p, (1000 + i, 1000 + i))
            self.files.append(str(p))

    def tearDown(self):
        self.tmp.cleanup()

    def test_exact_count_returns_all_in_mtime_order(self):
        got = MODULE.select_files(self.files, 3, take=None)
        self.assertEqual(got, self.files)

    def test_count_mismatch_without_flag_raises(self):
        with self.assertRaises(SystemExit) as cm:
            MODULE.select_files(self.files, 2, take=None)
        self.assertIn("count mismatch", str(cm.exception))
        self.assertIn("--take-latest", str(cm.exception))

    def test_take_latest_keeps_the_newest_n_in_mtime_order(self):
        got = MODULE.select_files(self.files, 2, take="latest")
        self.assertEqual(got, self.files[1:])

    def test_take_first_keeps_the_oldest_n_in_mtime_order(self):
        got = MODULE.select_files(self.files, 2, take="first")
        self.assertEqual(got, self.files[:2])

    def test_take_flag_never_invents_missing_images(self):
        """足りない側（生成 < 出力名）は救済しない。無音で欠落を埋めないこと。"""
        with self.assertRaises(SystemExit) as cm:
            MODULE.select_files(self.files[:1], 2, take="latest")
        self.assertIn("count mismatch", str(cm.exception))


if __name__ == "__main__":
    unittest.main()
