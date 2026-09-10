"""vibe-guard scan-files — 前置パスの誤検出を塞ぐ。陽性対照を先に書き、実装前に赤を見る。

なぜホーム配下に一時ディレクトリを作るか: local-info パターンは `/Users/<name>/...` の形に
当たる。`/var/folders/...` の既定の一時ディレクトリではバグが**再現しない**ので、
検査が空振りになる。再現条件を満たす場所に作る。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

VIBE_GUARD = Path.home() / ".config" / "vibe-guard" / "bin" / "vibe-guard"
PATTERNS = Path.home() / ".config" / "vibe-guard" / "config" / "personal-identifiers.txt"


def _sample_matching_first_pattern() -> str:
    """personal-identifiers.txt の1件目に**実際にマッチする実文**を機械的に作る。

    パターンは拡張正規表現なので、パターン文字列そのものを貼っても自分自身にマッチしない
    （2026-09-09 に陽性対照が空振りした実測がある）。生成してから re.search で確かめる。
    """
    import re

    pats = [l.rstrip("\n") for l in PATTERNS.read_text().splitlines()
            if l.strip() and not l.strip().startswith("#")]
    if not pats:
        raise unittest.SkipTest("personal-identifiers.txt が空")
    p = pats[0]
    out, i = [], 0
    while i < len(p):
        c = p[i]
        if c == "[":
            j = p.index("]", i)
            cls = p[i + 1:j]
            i = j + 1
            if i < len(p) and p[i] == "?":
                i += 1
                continue
            out.append(cls[0])
        elif c == "\\":
            out.append(p[i + 1]); i += 2
        elif c == "+":
            out.append(out[-1] if out else "x"); i += 1
        else:
            out.append(c); i += 1
    s = "".join(out)
    assert re.search(p, s), f"生成した実文がパターンにマッチしない: {s!r}"
    return s


@unittest.skipUnless(VIBE_GUARD.is_file(), f"{VIBE_GUARD} が無い（この端末に vibe-guard 未配置）")
class ScanFilesTest(unittest.TestCase):
    """scan-files をディレクトリに掛けたとき、指摘が**ファイル内容**由来だけであること。"""

    def _run(self, *args):
        r = subprocess.run([str(VIBE_GUARD), *args], capture_output=True, text=True)
        return r.returncode, r.stdout + r.stderr

    def _homedir_tmp(self):
        """ホーム配下の一時ディレクトリ。パスが local-info パターンに当たる状態を作る。"""
        return TemporaryDirectory(dir=str(Path.home()), prefix=".vgtest-")

    def test_target_exists(self):
        """空振り防止: 被検査物が実在することを独立に確かめる。"""
        self.assertTrue(VIBE_GUARD.is_file())

    def test_clean_directory_under_home_is_green(self):
        """陰性対照（**これが今 赤い**）: 内容が清潔なら、パスがホーム配下でも exit 0。"""
        with self._homedir_tmp() as t:
            d = Path(t)
            (d / "a.md").write_text("# hello\nnothing sensitive here\n")
            (d / "b.py").write_text("def f():\n    return 1\n")
            rc, out = self._run("scan-files", str(d))
            hits = [l for l in out.splitlines() if l.startswith("[local-info]")]
            self.assertEqual(hits, [], f"内容は清潔なのに指摘が出た（前置パス由来）:\n" + "\n".join(hits[:5]))
            self.assertEqual(rc, 0, out)

    def test_content_hit_is_detected_and_names_the_file(self):
        """陽性対照: **ファイル内容**に local-info があれば exit 1 で、どのファイルか分かる。"""
        with self._homedir_tmp() as t:
            d = Path(t)
            (d / "clean.md").write_text("nothing here\n")
            (d / "dirty.md").write_text("line1\ncontact = " + _sample_matching_first_pattern() + "\nline3\n")
            rc, out = self._run("scan-files", str(d))
            self.assertEqual(rc, 1, out)
            hits = [l for l in out.splitlines() if l.startswith("[local-info]")]
            self.assertTrue(hits, "内容に入れたのに検出されない")
            self.assertTrue(any("dirty.md" in l for l in hits),
                            f"どのファイルか分からない指摘:\n" + "\n".join(hits[:5]))
            self.assertFalse(any("clean.md" in l for l in hits), "清潔なファイルを指摘している")

    def test_single_file_target_still_works(self):
        """退行防止: ファイルを直接渡す経路（もともとバグが無い側）が壊れていない。"""
        with self._homedir_tmp() as t:
            d = Path(t)
            f = d / "x.md"
            f.write_text("plain text\n")
            rc, out = self._run("scan-files", str(f))
            self.assertEqual(rc, 0, out)

    def test_scan_text_not_regressed(self):
        """退行防止: scan-text（push の判定に使う経路）が清潔/汚染で 0/1 を返す。"""
        with self._homedir_tmp() as t:
            d = Path(t)
            clean = d / "clean.txt"; clean.write_text("just words\n")
            rc, out = self._run("scan-text", str(clean))
            self.assertEqual(rc, 0, out)
            dirty = d / "dirty.txt"
            dirty.write_text("contact = " + _sample_matching_first_pattern() + "\n")
            rc, out = self._run("scan-text", str(dirty))
            self.assertEqual(rc, 1, out)


if __name__ == "__main__":
    unittest.main()
