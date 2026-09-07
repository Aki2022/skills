"""judge_round.py — 1巡の合否を機械判定する。陽性対照を先に書き、実装前に赤を見る。"""
import importlib.util, io, json, unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "judge_round.py"
SPEC = importlib.util.spec_from_file_location("judge_round", SCRIPT)
M = importlib.util.module_from_spec(SPEC); assert SPEC.loader; SPEC.loader.exec_module(M)

THRESH = {"judgeMinEach": 75, "judgeAvgMin": 80, "criticalMustFixMax": 0,
          "freshReaderUnclearMax": 0, "reviewFindingsMax": 0, "reviewMaxSeverity": "minor"}

def judge(total, must_fix=()):
    return {"persona": "J", "round": 1,
            "scores": [{"item": "a", "score": total, "max": 100, "reason": "r"}],
            "total_score": total, "total_max": 100, "verdict": "pass" if total >= 75 else "revise",
            "strengths": [], "improvements": [], "must_fix": list(must_fix), "advice": []}

def reader(unclear=0):
    per = [{"n": i + 1, "verdict": "分からない" if i < unclear else "分かる", "note": "x"} for i in range(3)]
    return {"persona": "初見読者", "round": 1, "per_slide": per, "through_line": "t",
            "naive_questions": [], "top_fixes": []}

class JudgeRoundTest(unittest.TestCase):
    def _round(self, tmp, judges=None, readers=None, review=None, tests=None):
        d = Path(tmp) / "round1"; (d / "judges").mkdir(parents=True); (d / "readers").mkdir()
        for i, j in enumerate(judges or []): (d / "judges" / f"j{i}.json").write_text(json.dumps(j))
        for i, r in enumerate(readers or []): (d / "readers" / f"r{i}.json").write_text(json.dumps(r))
        if review is not None: (d / "review.json").write_text(json.dumps(review))
        if tests is not None: (d / "tests.json").write_text(json.dumps(tests))
        (Path(tmp) / "thresholds.json").write_text(json.dumps(THRESH))
        return d

    def _run(self, d):
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = M.main([str(d), "--thresholds", str(d.parent / "thresholds.json")])
        return rc, buf.getvalue()

    def test_all_pass_is_green(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, [judge(85), judge(90)], [reader(0)], {"findings": []}, {"exit_code": 0}))
            self.assertEqual(rc, 0, out); self.assertIn('"pass": true', out)

    def test_positive_control_threshold_miss_is_red(self):
        """陽性対照: 審査員1人が 74 → FAIL（各75以上を割る）。"""
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, [judge(74), judge(95)], [reader(0)]))
            self.assertEqual(rc, 1, out); self.assertIn("judgeMinEach", out)

    def test_positive_control_schema_violation_is_error(self):
        """陽性対照: verdict が語彙外 → exit 2（評価者の出力が壊れている）。"""
        with TemporaryDirectory() as t:
            bad = judge(90); bad["verdict"] = "maybe"
            rc, out = self._run(self._round(t, [bad], [reader(0)]))
            self.assertEqual(rc, 2, out); self.assertIn("schema", out)

    def test_must_fix_blocks(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._run_ok_shape(t, [judge(90, must_fix=["x"])]))
            self.assertEqual(rc, 1, out); self.assertIn("criticalMustFixMax", out)

    def _run_ok_shape(self, t, judges): return self._round(t, judges, [reader(0)])

    def test_fresh_reader_unclear_blocks(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, [judge(90)], [reader(1)]))
            self.assertEqual(rc, 1, out); self.assertIn("freshReaderUnclearMax", out)

    def test_method_b_findings_block(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, review={"findings": [{"severity": "minor", "msg": "x"}]}))
            self.assertEqual(rc, 1, out); self.assertIn("reviewFindingsMax", out)

    def test_method_a_exit_code_blocks(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, tests={"exit_code": 1}))
            self.assertEqual(rc, 1, out); self.assertIn("tests", out)

    def test_empty_round_is_error(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t))
            self.assertEqual(rc, 2, out)

    def test_output_states_scope(self):
        with TemporaryDirectory() as t:
            _, out = self._run(self._round(t, [judge(90)], [reader(0)]))
            self.assertIn("scope:", out)

if __name__ == "__main__": unittest.main()
