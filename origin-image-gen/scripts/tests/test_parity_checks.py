"""check_shell_parity.py / check_script_parity.py のテスト。

段階移行の期間、origin-image-gen と origin-pptx には「同じもの」が二重に存在する:

  (a) プロンプトの器（10フィールドのラベルと順序） — prompt-shell.md と
      imagegen-prompt-convention.md
  (b) collect_codex_images.py の実体コピー

このズレが黙って起きないことを機械で見張るのがこの2本。TDD として、まず陽性対照
（わざと壊して赤になることを実測できるケース）を含めてテストを先に書く。

一時ファイルは必ず TemporaryDirectory を使い、リポジトリ内の正典ファイルは一切書き換えない。
"""
from __future__ import annotations

import importlib.util
import io
import re
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = SCRIPTS_DIR.parents[1]  # .../origin-image-gen/scripts -> parents[1] = リポジトリルート

SHELL_SCRIPT = SCRIPTS_DIR / "check_shell_parity.py"
SCRIPT_PARITY_SCRIPT = SCRIPTS_DIR / "check_script_parity.py"

CANON_SHELL = REPO_ROOT / "origin-image-gen" / "references" / "prompt-shell.md"
CANON_CONVENTION = REPO_ROOT / "origin-pptx" / "style-guide" / "imagegen-prompt-convention.md"

CANON_COPY_A = REPO_ROOT / "origin-image-gen" / "scripts" / "collect_codex_images.py"
CANON_COPY_B = REPO_ROOT / "origin-pptx" / "scripts" / "collect_codex_images.py"

SHELL_SCOPE_LINE = (
    "scope: この検査は器の骨（ラベルと順序）のみを比較する。各フィールドの中身のズレは検出しない。"
)
SCRIPT_SCOPE_LINE = (
    "scope: この検査は2つのコピーがバイト一致かのみを見る。どちらが正しいかは判定しない。"
)

ROW_RE = re.compile(r"^\|\s*\d+\s*\|\s*`([^`]+)`\s*\|")


def _load(path: Path):
    spec = importlib.util.spec_from_file_location(path.stem, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def _run(module, argv):
    buf = io.StringIO()
    with redirect_stdout(buf):
        code = module.main(argv)
    return code, buf.getvalue()


def _swap_two_rows(text: str, label_a: str, label_b: str) -> str:
    """表の2行（番号セルごと）を丸ごと入れ替える。物理的な出現順が変わるので、
    番号順ではなく出現順でラベルを拾う抽出であれば不一致として検出できる。"""
    lines = text.splitlines(keepends=True)
    idx_a = idx_b = None
    for i, line in enumerate(lines):
        m = ROW_RE.match(line)
        if m:
            if m.group(1) == label_a:
                idx_a = i
            elif m.group(1) == label_b:
                idx_b = i
    assert idx_a is not None, f"行が見つからない: {label_a}"
    assert idx_b is not None, f"行が見つからない: {label_b}"
    lines[idx_a], lines[idx_b] = lines[idx_b], lines[idx_a]
    return "".join(lines)


def _drop_rows(text: str, keep_labels: set[str]) -> str:
    """表の行を keep_labels 以外すべて削り、抽出件数を10件未満に落とす。"""
    out = []
    for line in text.splitlines(keepends=True):
        m = ROW_RE.match(line)
        if m and m.group(1) not in keep_labels:
            continue
        out.append(line)
    return "".join(out)


class CheckShellParityTest(unittest.TestCase):
    def test_a1_canonical_files_match(self):
        module = _load(SHELL_SCRIPT)
        code, out = _run(module, ["--shell", str(CANON_SHELL), "--convention", str(CANON_CONVENTION)])
        self.assertEqual(code, 0, out)

    def test_a2_swapped_rows_are_detected_as_mismatch(self):
        """陽性対照: 表の2行を入れ替えると exit 1 になることを実測する。"""
        module = _load(SHELL_SCRIPT)
        original = CANON_SHELL.read_text(encoding="utf-8")
        swapped = _swap_two_rows(original, "Style", "House structure")
        self.assertNotEqual(original, swapped)
        with TemporaryDirectory() as tmp:
            shell_path = Path(tmp) / "prompt-shell.md"
            shell_path.write_text(swapped, encoding="utf-8")
            code, out = _run(module, ["--shell", str(shell_path), "--convention", str(CANON_CONVENTION)])
        self.assertEqual(code, 1, out)

    def test_a3_truncated_table_is_extraction_failure_not_match(self):
        """陽性対照: 抽出が10件に満たない場合は exit 2。両方0件でも「一致」にしてはいけない。"""
        module = _load(SHELL_SCRIPT)
        original = CANON_SHELL.read_text(encoding="utf-8")
        truncated = _drop_rows(original, keep_labels={"Use case", "Asset type", "Style"})
        with TemporaryDirectory() as tmp:
            shell_path = Path(tmp) / "prompt-shell.md"
            shell_path.write_text(truncated, encoding="utf-8")
            code, out = _run(module, ["--shell", str(shell_path), "--convention", str(CANON_CONVENTION)])
        self.assertEqual(code, 2, out)

    def test_a3b_both_sides_empty_is_still_extraction_failure(self):
        """両側とも0件のとき「両方0件だから一致」と判定してはいけない（exit 2 のはず）。"""
        module = _load(SHELL_SCRIPT)
        with TemporaryDirectory() as tmp:
            empty_shell = Path(tmp) / "empty-shell.md"
            empty_conv = Path(tmp) / "empty-convention.md"
            empty_shell.write_text("# 空\n", encoding="utf-8")
            empty_conv.write_text("# 空\n", encoding="utf-8")
            code, out = _run(module, ["--shell", str(empty_shell), "--convention", str(empty_conv)])
        self.assertEqual(code, 2, out)

    def test_a4_first_line_is_scope(self):
        module = _load(SHELL_SCRIPT)
        _, out = _run(module, ["--shell", str(CANON_SHELL), "--convention", str(CANON_CONVENTION)])
        self.assertEqual(out.splitlines()[0], SHELL_SCOPE_LINE)

    def test_default_paths_point_at_canonical_files(self):
        """引数省略時は正典2ファイルを既定値として使う。"""
        module = _load(SHELL_SCRIPT)
        code, out = _run(module, [])
        self.assertEqual(code, 0, out)


class CheckScriptParityTest(unittest.TestCase):
    @unittest.skipUnless(
        CANON_COPY_A.is_file() and CANON_COPY_B.is_file(),
        "collect_codex_images.py の二重コピーがまだ両方揃っていない"
        "（親セッションが origin-image-gen 側を並行して配置中のため）。"
        "両方揃った時点で親が再実行する。",
    )
    def test_b1_canonical_copies_match(self):
        module = _load(SCRIPT_PARITY_SCRIPT)
        code, out = _run(module, ["--image-gen", str(CANON_COPY_A), "--pptx", str(CANON_COPY_B)])
        self.assertEqual(code, 0, out)

    def test_b2_one_extra_blank_line_is_detected_as_mismatch(self):
        """陽性対照: 一時コピーの片方に空行を1行足すと exit 1 になることを実測する。"""
        module = _load(SCRIPT_PARITY_SCRIPT)
        content = "# dummy collect_codex_images.py\nprint('x')\n"
        with TemporaryDirectory() as tmp:
            a = Path(tmp) / "a.py"
            b = Path(tmp) / "b.py"
            a.write_text(content, encoding="utf-8")
            b.write_text(content + "\n", encoding="utf-8")
            code, out = _run(module, ["--image-gen", str(a), "--pptx", str(b)])
        self.assertEqual(code, 1, out)

    def test_b2b_identical_temp_copies_match(self):
        """陰性対照: 一時コピーが完全一致していれば exit 0。"""
        module = _load(SCRIPT_PARITY_SCRIPT)
        content = "# dummy collect_codex_images.py\nprint('x')\n"
        with TemporaryDirectory() as tmp:
            a = Path(tmp) / "a.py"
            b = Path(tmp) / "b.py"
            a.write_text(content, encoding="utf-8")
            b.write_text(content, encoding="utf-8")
            code, out = _run(module, ["--image-gen", str(a), "--pptx", str(b)])
        self.assertEqual(code, 0, out)

    def test_b3_missing_file_is_exit_2(self):
        module = _load(SCRIPT_PARITY_SCRIPT)
        with TemporaryDirectory() as tmp:
            a = Path(tmp) / "a.py"
            a.write_text("x\n", encoding="utf-8")
            missing = Path(tmp) / "does-not-exist.py"
            code, out = _run(module, ["--image-gen", str(a), "--pptx", str(missing)])
        self.assertEqual(code, 2, out)

    def test_b4_first_line_is_scope(self):
        module = _load(SCRIPT_PARITY_SCRIPT)
        content = "x\n"
        with TemporaryDirectory() as tmp:
            a = Path(tmp) / "a.py"
            b = Path(tmp) / "b.py"
            a.write_text(content, encoding="utf-8")
            b.write_text(content, encoding="utf-8")
            _, out = _run(module, ["--image-gen", str(a), "--pptx", str(b)])
        self.assertEqual(out.splitlines()[0], SCRIPT_SCOPE_LINE)


if __name__ == "__main__":
    unittest.main()
