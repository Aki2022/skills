"""eval_state.py — 巡をまたぐ状態と収束判定。実装前に赤を見る。"""
import importlib.util, io, json, unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "eval_state.py"
SPEC = importlib.util.spec_from_file_location("eval_state", SCRIPT)
M = importlib.util.module_from_spec(SPEC); assert SPEC.loader; SPEC.loader.exec_module(M)

def verdict(n_findings, must_fix, passed=False):
    return {"pass": passed, "findings_count": n_findings, "must_fix": must_fix, "rejected": []}

class EvalStateTest(unittest.TestCase):
    def _run(self, *args):
        buf = io.StringIO()
        with redirect_stdout(buf): rc = M.main(list(args))
        return rc, buf.getvalue()

    def test_two_rounds_carry_forward(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"; v1 = Path(t) / "v1.json"; v1.write_text(json.dumps(verdict(5, ["fix A", "fix B"])))
            rc, _ = self._run("init", str(st), "--max-rounds", "3"); self.assertEqual(rc, 0)
            rc, _ = self._run("record", str(st), "--verdict", str(v1)); self.assertEqual(rc, 0)
            v2 = Path(t) / "v2.json"; v2.write_text(json.dumps(verdict(2, ["fix C"])))
            rc, _ = self._run("record", str(st), "--verdict", str(v2)); self.assertEqual(rc, 0)
            s = json.loads(st.read_text())
            self.assertEqual(s["round"], 2)
            rc, out = self._run("carry", str(st))
            self.assertEqual(rc, 0)
            carry = json.loads(out)
            self.assertIn("fix C", carry["verify_must_fix_from_prev_round"])  # 前巡の must-fix だけ
            self.assertNotIn("scores", carry)  # 過去評価の本文は渡さない（フレッシュ評価）

    def test_positive_control_not_converging(self):
        """陽性対照: 前巡と同数の指摘 → not-converging（exit 3）。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            for n in (4, 4):
                v = Path(t) / f"v{n}.json"; v.write_text(json.dumps(verdict(n, ["x"]))); self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("converged", str(st))
            self.assertEqual(rc, 3, out); self.assertIn("not-converging", out)

    def test_converging_when_findings_drop(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            for n in (4, 2):
                v = Path(t) / f"v{n}.json"; v.write_text(json.dumps(verdict(n, ["x"]))); self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("converged", str(st)); self.assertEqual(rc, 0, out)

    def test_max_rounds_stops(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "2")
            for n in (5, 4):
                v = Path(t) / f"v{n}.json"; v.write_text(json.dumps(verdict(n, ["x"]))); self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("next", str(st))
            self.assertEqual(rc, 4, out); self.assertIn("max-rounds", out)  # 上限到達＝終了して報告

if __name__ == "__main__": unittest.main()
