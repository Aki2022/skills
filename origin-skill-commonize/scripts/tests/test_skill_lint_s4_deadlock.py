"""S4 の参照ループが自己デッドロックしないことの回帰テスト。

欠陥（2026-09-19 に 1eca216 で修正）: S4 が `done <<< "$refs_output"` を使っていた。
here-string では bash が内容をパイプへ書き、その読み手が**同じプロセスのこのループ**になる。
参照が多い skill で内容が**パイプバッファ（16KB）を超える**と、ループが読み始める前に
write(2) が埋まって自己デッドロックする。実測で skill_lint が1日以上ぶら下がっていた。

修正は `done < <(printf '%s\\n' "$refs_output")`（process substitution・別プロセスが書く）。

**このテストは 16KB を超える参照量でなければ空振りする。** 参照60件（約1.4KB）では
欠陥がある版でも通ってしまう — 一度それで「再現しない」と誤って結論した。
"""
from __future__ import annotations

import subprocess
import textwrap
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
SKILL_LINT = SCRIPTS_DIR / "skill_lint.sh"

# パーサ出力は 1 参照あたり約 25 バイト。16KB のパイプバッファを確実に超える量にする
# （1200 件で実測 28,893 バイト）。
REFERENCE_COUNT = 1200
TIMEOUT_SECONDS = 60


def _bash5() -> str | None:
    """here-doc をパイプで渡す最適化を持つ bash を探す。

    **bash 3.2 では再現しない** — here-doc に常に一時ファイルを使うため
    自己デッドロックが起きない。macOS の /bin/bash は 3.2 なので、
    PATH 既定で走らせると欠陥がある版でもテストが通ってしまう（一度それで
    「再現しない」と誤って結論した）。5.x を明示的に探すこと。
    """
    import shutil
    candidates = [shutil.which("bash"), "/opt/homebrew/bin/bash",
                  "/usr/local/bin/bash", "/run/current-system/sw/bin/bash"]
    for c in candidates:
        if not c or not Path(c).exists():
            continue
        try:
            out = subprocess.run([c, "--version"], capture_output=True, text=True, timeout=10).stdout
        except Exception:
            continue
        head = out.splitlines()[0] if out else ""
        ver = head.split("version ")[-1].split("(")[0] if "version " in head else ""
        if ver and ver.split(".")[0].isdigit() and int(ver.split(".")[0]) >= 5:
            return c
    return None


def _build(root: Path, name: str, n: int) -> None:
    d = root / name
    (d / "references").mkdir(parents=True, exist_ok=True)
    body = ["---", f"name: {name}", "description: many references", "---"]
    for i in range(n):
        (d / "references" / f"f{i}.md").write_text("x\n")
        body.append(f"See `references/f{i}.md`.")
    (d / "SKILL.md").write_text("\n".join(body) + "\n")


class S4DeadlockRegressionTests(unittest.TestCase):
    def test_lint_completes_when_reference_output_exceeds_pipe_buffer(self):
        with TemporaryDirectory() as tmp:
            home = Path(tmp)
            root = home / ".agents" / "skills"
            _build(root, "own-many-run", REFERENCE_COUNT)
            bash5 = _bash5()
            if bash5 is None:
                self.skipTest("bash 5 以上が見つからない — 3.2 では欠陥が再現しないので検査にならない")
            try:
                r = subprocess.run(
                    [bash5, str(SKILL_LINT), str(root)],
                    capture_output=True, text=True, timeout=TIMEOUT_SECONDS,
                    stdin=subprocess.DEVNULL,
                    env={"PATH": "/usr/bin:/bin:/usr/local/bin", "HOME": str(home),
                         "SKILL_LINT_SKIP_PYTEST": "1"},
                )
            except subprocess.TimeoutExpired:
                self.fail(
                    f"参照 {REFERENCE_COUNT} 件（>16KB）で skill_lint.sh が "
                    f"{TIMEOUT_SECONDS} 秒以内に終わらない — S4 の自己デッドロックが再発している"
                )
            self.assertEqual(r.returncode, 0, f"{r.stdout}{r.stderr}")
            self.assertIn("OK", r.stdout)

    def test_missing_reference_still_detected_at_that_scale(self):
        """陰性対照。デッドロックを避けても S4 の検出力が落ちていないこと。"""
        with TemporaryDirectory() as tmp:
            home = Path(tmp)
            root = home / ".agents" / "skills"
            _build(root, "own-many-run", REFERENCE_COUNT)
            md = root / "own-many-run" / "SKILL.md"
            md.write_text(md.read_text() + "See `references/absent.md`.\n")
            bash5 = _bash5() or "bash"
            r = subprocess.run(
                [bash5, str(SKILL_LINT), str(root)],
                capture_output=True, text=True, timeout=TIMEOUT_SECONDS,
                stdin=subprocess.DEVNULL,
                env={"PATH": "/usr/bin:/bin:/usr/local/bin", "HOME": str(home),
                     "SKILL_LINT_SKIP_PYTEST": "1"},
            )
            self.assertNotEqual(r.returncode, 0)
            self.assertIn("absent.md", r.stdout + r.stderr)


if __name__ == "__main__":
    unittest.main()
