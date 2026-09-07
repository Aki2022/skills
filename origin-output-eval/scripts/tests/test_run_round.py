"""run_round.sh — 1巡ぶんの「判定→record→収束判定」の機械化。実装前に赤を見る。

judge_round.py の実装は別実装者が同時進行中のため、ここでは JUDGE_ROUND_BIN 環境変数で
契約どおりの出力を返すダミーに差し替えて検証する（本番既定値は同じディレクトリの judge_round.py）。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import textwrap
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
RUN_ROUND = SCRIPTS_DIR / "run_round.sh"

DUMMY_JUDGE_ROUND_SRC = textwrap.dedent(
    """\
    #!/usr/bin/env python3
    # run_round.sh のテスト専用ダミー。judge_round.py の契約どおりの出力だけを返す。
    import argparse, json, os, sys

    def main():
        ap = argparse.ArgumentParser()
        ap.add_argument("round_dir")
        ap.add_argument("--thresholds", required=True)
        ap.add_argument("--round", type=int, default=None)
        ap.add_argument("--json-out", default=None)
        ap.add_argument("--expect-judges", type=int, default=None)
        ap.add_argument("--expect-readers", type=int, default=None)
        a = ap.parse_args()

        exit_code = int(os.environ.get("DUMMY_JUDGE_EXIT", "0"))
        findings = int(os.environ.get("DUMMY_JUDGE_FINDINGS", "0"))
        round_no = a.round if a.round is not None else int(os.environ.get("DUMMY_JUDGE_ROUND", "1"))
        passed = exit_code == 0

        verdict = {
            "pass": passed,
            "round": round_no,
            "methods": {},
            "reasons": [] if passed else ["dummy: 不合格"],
            "findings_count": findings,
            "must_fix": [],
            "rejected": [],
            "provenance": {},
        }
        print("scope: run_round.sh テスト用ダミー（judge_round.py の代役）")
        print(json.dumps(verdict, ensure_ascii=False))
        if a.json_out:
            with open(a.json_out, "w", encoding="utf-8") as f:
                json.dump(verdict, f, ensure_ascii=False)
        return exit_code

    if __name__ == "__main__":
        sys.exit(main())
    """
)


class RunRoundTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = TemporaryDirectory()
        dummy_path = Path(cls._tmp.name) / "dummy_judge_round.py"
        dummy_path.write_text(DUMMY_JUDGE_ROUND_SRC, encoding="utf-8")
        dummy_path.chmod(0o755)
        cls.dummy_bin = f"python3 {dummy_path}"

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def _env(self, **overrides):
        env = dict(os.environ)
        env["JUDGE_ROUND_BIN"] = self.dummy_bin
        env.update({k: str(v) for k, v in overrides.items()})
        return env

    def _run(self, eval_dir: Path, *args, env=None):
        cmd = ["bash", str(RUN_ROUND), str(eval_dir), "--thresholds", str(eval_dir / "thresholds.json"), *args]
        proc = subprocess.run(cmd, capture_output=True, text=True, env=env or self._env())
        return proc.returncode, proc.stdout, proc.stderr

    def _make_thresholds(self, eval_dir: Path):
        (eval_dir / "thresholds.json").write_text("{}", encoding="utf-8")

    def test_pass_case_exit0_and_state(self):
        with TemporaryDirectory() as t:
            eval_dir = Path(t)
            self._make_thresholds(eval_dir)
            (eval_dir / "round_1").mkdir()
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=0, DUMMY_JUDGE_FINDINGS=0))
            self.assertEqual(rc, 0, f"stdout={out}\nstderr={err}")
            verdict_path = eval_dir / "round_1" / "verdict.json"
            self.assertTrue(verdict_path.is_file(), "round_1/verdict.json が生成されていない")
            state = json.loads((eval_dir / "state.json").read_text())
            self.assertEqual(state["round"], 1)

    def test_missing_round_dir_exit2(self):
        with TemporaryDirectory() as t:
            eval_dir = Path(t)
            self._make_thresholds(eval_dir)
            # round_1/ を用意しない
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=0, DUMMY_JUDGE_FINDINGS=0))
            self.assertEqual(rc, 2, f"stdout={out}\nstderr={err}")

    def test_fail_then_improving_exit1(self):
        with TemporaryDirectory() as t:
            eval_dir = Path(t)
            self._make_thresholds(eval_dir)
            (eval_dir / "round_1").mkdir()
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=1, DUMMY_JUDGE_FINDINGS=5))
            self.assertEqual(rc, 1, f"stdout={out}\nstderr={err}")

            (eval_dir / "round_2").mkdir()
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=1, DUMMY_JUDGE_FINDINGS=2))
            self.assertEqual(rc, 1, f"stdout={out}\nstderr={err}")  # 不合格だが続行可
            state = json.loads((eval_dir / "state.json").read_text())
            self.assertEqual(state["round"], 2)

    def test_same_findings_two_rounds_exit3(self):
        with TemporaryDirectory() as t:
            eval_dir = Path(t)
            self._make_thresholds(eval_dir)
            (eval_dir / "round_1").mkdir()
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=1, DUMMY_JUDGE_FINDINGS=4))
            self.assertEqual(rc, 1, f"stdout={out}\nstderr={err}")

            (eval_dir / "round_2").mkdir()
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=1, DUMMY_JUDGE_FINDINGS=4))
            self.assertEqual(rc, 3, f"stdout={out}\nstderr={err}")

    def test_max_rounds_exit4(self):
        with TemporaryDirectory() as t:
            eval_dir = Path(t)
            self._make_thresholds(eval_dir)
            (eval_dir / "round_1").mkdir()
            rc, out, err = self._run(eval_dir, "--max-rounds", "1",
                                      env=self._env(DUMMY_JUDGE_EXIT=1, DUMMY_JUDGE_FINDINGS=3))
            self.assertEqual(rc, 1, f"stdout={out}\nstderr={err}")  # 1巡目は不合格・続行可

            rc, out, err = self._run(eval_dir, "--max-rounds", "1",
                                      env=self._env(DUMMY_JUDGE_EXIT=1, DUMMY_JUDGE_FINDINGS=3))
            self.assertEqual(rc, 4, f"stdout={out}\nstderr={err}")  # 上限到達

    def test_provenance_violation_exit2_propagates(self):
        with TemporaryDirectory() as t:
            eval_dir = Path(t)
            self._make_thresholds(eval_dir)
            (eval_dir / "round_1").mkdir()
            rc, out, err = self._run(eval_dir, env=self._env(DUMMY_JUDGE_EXIT=2, DUMMY_JUDGE_FINDINGS=0))
            self.assertEqual(rc, 2, f"stdout={out}\nstderr={err}")


class RunRoundIntegratorHardeningTest(unittest.TestCase):
    """統合者が run_round.sh に追加した2件の陽性対照。

    どちらも「判定していないのに判定したように見える」経路を塞ぐ。
    """

    def _dummy_judge(self, tmp: Path, body: str) -> str:
        d = tmp / "dummy_judge.sh"
        d.write_text("#!/usr/bin/env bash\n" + body + "\n")
        d.chmod(0o755)
        return str(d)

    def _run(self, eval_dir: Path, judge_bin: str):
        env = dict(os.environ, JUDGE_ROUND_BIN=judge_bin)
        return subprocess.run(
            ["bash", str(RUN_ROUND), str(eval_dir),
             "--thresholds", str(eval_dir / "thresholds.json")],
            capture_output=True, text=True, env=env,
        )

    def test_positive_control_broken_state_is_exit2_not_fail(self):
        """陽性対照: state.json が壊れている（round キー無し）→ exit 2。

        eval_state next がそこで落ちるが、その exit 1 をそのまま中継すると run_round の契約では
        「不合格・続行可」に化ける。壊れた state が品質問題に見えるのを防ぐ。
        """
        with TemporaryDirectory() as t:
            tmp = Path(t)
            e = tmp / "eval"
            (e / "round_1").mkdir(parents=True)
            (e / "thresholds.json").write_text("{}")
            (e / "state.json").write_text(json.dumps({"max_rounds": 3, "history": []}))
            r = self._run(e, self._dummy_judge(tmp, "exit 0"))
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)

    def test_positive_control_judge_crash_is_exit2_not_fail(self):
        """陽性対照: judge_round が verdict を書かずに終了 → exit 2。不合格(1)と混同しない。"""
        with TemporaryDirectory() as t:
            tmp = Path(t)
            e = tmp / "eval"
            (e / "round_1").mkdir(parents=True)
            (e / "thresholds.json").write_text("{}")
            r = self._run(e, self._dummy_judge(tmp, 'echo "boom" >&2; exit 1'))
            self.assertEqual(r.returncode, 2, r.stdout + r.stderr)


if __name__ == "__main__":
    unittest.main()
