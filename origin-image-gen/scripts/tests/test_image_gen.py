"""image_gen.py — 完結した窓口。陽性対照を先に書き、実装前に赤を見る。

codex は呼ばない。IMAGE_GEN_CODEX_BIN でダミーに差し替えて、窓口のロジックだけを検証する。
"""
from __future__ import annotations

import json
import os
import subprocess
import textwrap
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS = Path(__file__).resolve().parents[1]
IMAGE_GEN = SCRIPTS / "image_gen.py"

FIELDS = {
    "Use case": "infographic-diagram",
    "Asset type": "illustration (1:1)",
    "Style references": "",
    "Style": "flat, no gradients",
    "House structure": "single subject centered",
    "Composition": "one object, generous margin",
    "Content": "a paper airplane",
    "Color palette": "#0F62FE (blue), #F4F4F4 (light grey)",
    "Constraints": "no text",
    "Avoid": "photorealism, drop shadows",
}

# 生成枚数を環境変数で決めるダミー codex。session dir を作り png を置き、
# ログに `session id: <uuid>` を出す（本物と同じ形）。
DUMMY_CODEX = textwrap.dedent(
    """\
    #!/usr/bin/env python3
    import os, sys, pathlib, uuid
    n = int(os.environ.get("DUMMY_N_IMAGES", "0"))
    fail = os.environ.get("DUMMY_FAIL", "0") == "1"
    sid = os.environ.get("DUMMY_SID") or str(uuid.uuid4())
    home = pathlib.Path(os.environ["DUMMY_CODEX_HOME"])
    d = home / "generated_images" / sid
    d.mkdir(parents=True, exist_ok=True)
    for i in range(n):
        (d / f"ig_{i:03d}.png").write_bytes(b"\\x89PNG" + bytes([i]))
    print(f"session id: {sid}")
    print("done")
    sys.exit(1 if fail else 0)
    """
)


class ImageGenTest(unittest.TestCase):
    def _setup(self, tmp: Path, n_images: int, fail: bool = False):
        """ダミー codex と fields.json を用意し、環境変数を返す。"""
        dummy = tmp / "dummy_codex.py"
        dummy.write_text(DUMMY_CODEX)
        fields = tmp / "fields.json"
        fields.write_text(json.dumps(FIELDS, ensure_ascii=False))
        env = dict(
            os.environ,
            IMAGE_GEN_CODEX_BIN=f"python3 {dummy}",
            DUMMY_N_IMAGES=str(n_images),
            DUMMY_FAIL="1" if fail else "0",
            DUMMY_CODEX_HOME=str(tmp / "codexhome"),
            CODEX_HOME=str(tmp / "codexhome"),
        )
        return fields, env

    # 窓口は必ず1行目に scope 行を出す。これを検査に使うことで、
    # 「スクリプトが無いので python3 が rc=2 を返し、rc!=0 を見るテストが勝手に通る」
    # という空振りを塞ぐ（実装前にこの形で3件が緑になったのを実測して直した）。
    MARKER = "scope: image_gen"

    def _run(self, args, env, cwd=None, expect_ran=True):
        r = subprocess.run(
            ["python3", str(IMAGE_GEN), *args],
            capture_output=True, text=True, env=env, cwd=cwd,
        )
        out = r.stdout + r.stderr
        if expect_ran:
            assert self.MARKER in out, (
                f"窓口が動いた証拠（{self.MARKER!r}）が出力に無い。"
                f"スクリプト不在・起動失敗を『不合格』と読み違えないための検査。\n{out}"
            )
        return r.returncode, out

    def test_script_exists(self):
        """空振り防止: 被検査物が実在することを独立に確かめる。"""
        self.assertTrue(IMAGE_GEN.is_file(), f"{IMAGE_GEN} が無い")

    # --- AC(1): 出力名3件・生成2枚 → exit≠0 かつ3件目が作られない ---
    def test_shortfall_stops_and_does_not_fabricate(self):
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=2)
            outs = [str(tmp / "out" / f"a{i}.png") for i in range(3)]
            rc, log = self._run(["--fields", str(fields), "--out", *outs,
                                 "--workdir", str(tmp), "--skip-git-check"], env)
            self.assertNotEqual(rc, 0, log)
            self.assertFalse(Path(outs[2]).exists(), "3件目を作ってはいけない")
            # 埋めないので、そもそも1件も成果物として置かない
            self.assertFalse(Path(outs[0]).exists(), "枚数不足なら部分成果物も置かない")

    # --- AC(2): .stale 退避後に回収失敗 → 旧ファイルが成果物パスとして返らない ---
    def test_stale_not_returned_as_artifact(self):
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=0, fail=True)
            out = tmp / "out" / "a0.png"
            out.parent.mkdir(parents=True)
            out.write_bytes(b"OLD-ARTIFACT")
            rc, log = self._run(["--fields", str(fields), "--out", str(out),
                                 "--workdir", str(tmp), "--skip-git-check"], env)
            self.assertNotEqual(rc, 0, log)
            self.assertTrue((tmp / "out" / "a0.png.stale").exists(), ".stale へ退避すること")
            self.assertFalse(out.exists(), "旧ファイルを成果物パスに残してはいけない")
            self.assertNotIn(str(out), log.splitlines(), "旧ファイルを成果物として報告してはいけない")

    # --- AC(3): 禁止フラグが組み立てられたコマンドに現れない ---
    def test_bypass_flag_never_in_assembled_command(self):
        BAD = "--dangerously-bypass-approvals-and-sandbox"
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=1)
            out = str(tmp / "out" / "a0.png")
            rc, log = self._run(["--fields", str(fields), "--out", out,
                                 "--workdir", str(tmp), "--skip-git-check", "--dry-run"], env)
            self.assertEqual(rc, 0, log)
            self.assertNotIn(BAD, log, "dry-run が出すコマンドに禁止フラグが現れてはいけない")
            self.assertIn("--sandbox workspace-write", log)
            self.assertIn("sandbox_workspace_write.network_access=true", log)

    def test_bypass_flag_smuggled_through_field_value_is_rejected(self):
        """陽性対照: フィールド値に紛れ込ませても組み立てに載らない（exit 2 で拒否する）。"""
        BAD = "--dangerously-bypass-approvals-and-sandbox"
        with TemporaryDirectory() as t:
            tmp = Path(t)
            _, env = self._setup(tmp, n_images=1)
            f = dict(FIELDS, Content=f"a paper airplane {BAD}")
            fields = tmp / "smuggled.json"   # _setup が書く fields.json と別名にする
            fields.write_text(json.dumps(f, ensure_ascii=False))
            out = str(tmp / "out" / "a0.png")
            rc, log = self._run(["--fields", str(fields), "--out", out,
                                 "--workdir", str(tmp), "--skip-git-check", "--dry-run"], env)
            self.assertEqual(rc, 2, log)
            # 拒否理由の文はフラグ名を引用してよい（それが説明）。見たいのは
            # 「コマンドが1つも組み立てられていない」こと。
            self.assertNotIn("command:", log, "拒否したのにコマンドを組み立ててはいけない")
            self.assertIn("禁止フラグ", log)

    # --- 正常系と入力不備 ---
    def test_happy_path_returns_artifacts_and_leaves_prompt_and_log(self):
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=2)
            outs = [str(tmp / "out" / f"a{i}.png") for i in range(2)]
            rc, log = self._run(["--fields", str(fields), "--out", *outs,
                                 "--workdir", str(tmp), "--skip-git-check"], env)
            self.assertEqual(rc, 0, log)
            for o in outs:
                self.assertTrue(Path(o).exists(), f"{o} が無い")
            self.assertTrue((tmp / "out" / "prompt.txt").exists(), "prompt.txt を残すこと")
            self.assertTrue((tmp / "out" / "run.log").exists(), "run.log を残すこと")

    def test_prompt_has_ten_fields_in_fixed_order(self):
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=1)
            out = str(tmp / "out" / "a0.png")
            self._run(["--fields", str(fields), "--out", out,
                       "--workdir", str(tmp), "--skip-git-check", "--dry-run"], env)
            prompt = (tmp / "out" / "prompt.txt").read_text()
            labels = ["Use case", "Asset type", "Style references", "Style", "House structure",
                      "Composition", "Content", "Color palette", "Constraints", "Avoid"]
            positions = [prompt.index(f"{l}:") for l in labels]
            self.assertEqual(positions, sorted(positions), "10フィールドの順序が固定されていない")

    def test_missing_field_key_is_exit2(self):
        with TemporaryDirectory() as t:
            tmp = Path(t)
            _, env = self._setup(tmp, n_images=1)
            f = dict(FIELDS); del f["Avoid"]
            fields = tmp / "incomplete.json"  # _setup が書く fields.json と別名にする
            fields.write_text(json.dumps(f, ensure_ascii=False))
            rc, log = self._run(["--fields", str(fields), "--out", str(tmp / "out" / "a.png"),
                                 "--workdir", str(tmp), "--skip-git-check"], env)
            self.assertEqual(rc, 2, log)
            self.assertIn("Avoid", log)

    def test_skip_git_repo_check_is_added_automatically_outside_repo(self):
        """窓口が呼び出し側の規律に依存しないこと: --skip-git-check を渡さなくても、
        workdir が git リポジトリ外なら --skip-git-repo-check が自動で載る。"""
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=1)
            out = str(tmp / "out" / "a0.png")
            rc, log = self._run(["--fields", str(fields), "--out", out,
                                 "--workdir", str(tmp), "--dry-run"], env)
            self.assertEqual(rc, 0, log)
            self.assertIn("--skip-git-repo-check", log)

    def test_skip_git_repo_check_absent_inside_repo(self):
        """陰性対照: git リポジトリ内なら足さない（過剰に付けない）。"""
        with TemporaryDirectory() as t:
            tmp = Path(t)
            subprocess.run(["git", "init", "-q", str(tmp)], check=True)
            fields, env = self._setup(tmp, n_images=1)
            out = str(tmp / "out" / "a0.png")
            rc, log = self._run(["--fields", str(fields), "--out", out,
                                 "--workdir", str(tmp), "--dry-run"], env)
            self.assertEqual(rc, 0, log)
            self.assertNotIn("--skip-git-repo-check", log)

    def test_no_out_is_exit2(self):
        with TemporaryDirectory() as t:
            tmp = Path(t)
            fields, env = self._setup(tmp, n_images=1)
            rc, log = self._run(["--fields", str(fields), "--out",
                                 "--workdir", str(tmp)], env)
            self.assertEqual(rc, 2, log)


if __name__ == "__main__":
    unittest.main()
