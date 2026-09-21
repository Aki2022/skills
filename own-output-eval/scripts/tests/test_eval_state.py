"""eval_state.py — 巡をまたぐ状態と収束判定。実装前に赤を見る。"""
import importlib.util, io, json, unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "eval_state.py"
SPEC = importlib.util.spec_from_file_location("eval_state", SCRIPT)
M = importlib.util.module_from_spec(SPEC); assert SPEC.loader; SPEC.loader.exec_module(M)

def verdict(n_findings, must_fix, passed=False, round=None, include_findings_count=True,
            blocking=None, include_blocking_count=True, methods=("C_judges",)):
    v = {"pass": passed, "must_fix": must_fix, "rejected": [],
         "methods": {m: {} for m in methods}}
    if include_findings_count:
        v["findings_count"] = n_findings
    if include_blocking_count:
        v["blocking_count"] = n_findings if blocking is None else blocking
    if round is not None:
        v["round"] = round
    return v

class EvalStateTest(unittest.TestCase):
    def _run(self, *args):
        buf = io.StringIO()
        with redirect_stdout(buf): rc = M.main(list(args))
        return rc, buf.getvalue()

    def test_two_rounds_carry_forward(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"; v1 = Path(t) / "v1.json"; v1.write_text(json.dumps(verdict(5, ["fix A", "fix B"], round=1)))
            rc, _ = self._run("init", str(st), "--max-rounds", "3"); self.assertEqual(rc, 0)
            rc, _ = self._run("record", str(st), "--verdict", str(v1)); self.assertEqual(rc, 0)
            v2 = Path(t) / "v2.json"; v2.write_text(json.dumps(verdict(2, ["fix C"], round=2)))
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
            for i, n in enumerate((4, 4), start=1):
                v = Path(t) / f"v{i}.json"; v.write_text(json.dumps(verdict(n, ["x"], round=i))); self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("converged", str(st))
            self.assertEqual(rc, 3, out); self.assertIn("not-converging", out)

    def test_converging_when_findings_drop(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            for i, n in enumerate((4, 2), start=1):
                v = Path(t) / f"v{i}.json"; v.write_text(json.dumps(verdict(n, ["x"], round=i))); self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("converged", str(st)); self.assertEqual(rc, 0, out)

    def test_max_rounds_stops(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "2")
            for i, n in enumerate((5, 4), start=1):
                v = Path(t) / f"v{i}.json"; v.write_text(json.dumps(verdict(n, ["x"], round=i))); self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("next", str(st))
            self.assertEqual(rc, 4, out); self.assertIn("max-rounds", out)  # 上限到達＝終了して報告

    # --- ここから硬化スライス追加分 ---

    def test_missing_findings_count_exit2(self):
        """①の陽性対照: findings_count の無い旧形式 verdict は record を拒否する。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v = Path(t) / "v.json"
            v.write_text(json.dumps(verdict(4, ["x"], round=1, include_findings_count=False)))
            rc, out = self._run("record", str(st), "--verdict", str(v))
            self.assertEqual(rc, 2, out)

    def test_round_mismatch_exit2(self):
        """verdict の round が state の round+1 と食い違えば拒否する（巡の取り違え検出）。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v = Path(t) / "v.json"
            v.write_text(json.dumps(verdict(4, ["x"], round=2)))  # 期待値は 1
            rc, out = self._run("record", str(st), "--verdict", str(v))
            self.assertEqual(rc, 2, out)

    def test_rejected_flows_to_history_and_carry(self):
        """--rejected は history に積まれ、carry の already_rejected_proposals に出る。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v = Path(t) / "v.json"; v.write_text(json.dumps(verdict(3, ["x"], round=1)))
            rc, _ = self._run("record", str(st), "--verdict", str(v), "--rejected", "proposal X")
            self.assertEqual(rc, 0)
            rc, out = self._run("carry", str(st))
            self.assertEqual(rc, 0)
            carry = json.loads(out)
            self.assertIn("proposal X", carry["already_rejected_proposals"])

    def test_carry_has_note_for_readers(self):
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            rc, out = self._run("carry", str(st))
            self.assertEqual(rc, 0)
            carry = json.loads(out)
            self.assertIn("note_for_readers", carry)
            self.assertTrue(carry["note_for_readers"])

    # --- ここから ISSUE-07（skill eval で実測した設計欠陥）追加分 ---

    def test_missing_blocking_count_exit2(self):
        """①の陽性対照: blocking_count の無い verdict は record を拒否する。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v = Path(t) / "v.json"
            v.write_text(json.dumps(verdict(4, ["x"], round=1, include_blocking_count=False)))
            rc, out = self._run("record", str(st), "--verdict", str(v))
            self.assertEqual(rc, 2, out)

    def test_converged_uses_blocking_count_not_findings_count(self):
        """eval-1 の実測再生: 文書を再構成して評価単位が増え findings 6→7 になったが、
        読者の「分からない」は 4→1 に減っている。収束は blocking_count で見るので exit 0。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v1 = Path(t) / "v1.json"
            v1.write_text(json.dumps(verdict(6, ["x"], round=1, blocking=4)))
            self._run("record", str(st), "--verdict", str(v1))
            v2 = Path(t) / "v2.json"
            v2.write_text(json.dumps(verdict(7, ["x"], round=2, blocking=1)))
            self._run("record", str(st), "--verdict", str(v2))
            rc, out = self._run("converged", str(st))
            self.assertEqual(rc, 0, out)
            self.assertIn("blocking", out)

    def test_converged_still_red_when_blocking_not_decreasing(self):
        """陽性対照: 方式集合が同じで blocking が減っていなければ従来どおり exit 3。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            for i, b in enumerate((4, 4), start=1):
                v = Path(t) / f"v{i}.json"
                v.write_text(json.dumps(verdict(9, ["x"], round=i, blocking=b)))
                self._run("record", str(st), "--verdict", str(v))
            rc, out = self._run("converged", str(st))
            self.assertEqual(rc, 3, out)
            self.assertIn("not-converging", out)

    def test_converged_method_set_change_is_not_comparable(self):
        """eval-0 の実測再生: 方式A が緑になり方式B を追加したため件数の分母が変わった。
        件数比較は不成立なので停止させない（exit 0）。理由を明示する。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v1 = Path(t) / "v1.json"
            v1.write_text(json.dumps(verdict(1, ["x"], round=1, blocking=1, methods=("A_tests",))))
            self._run("record", str(st), "--verdict", str(v1))
            v2 = Path(t) / "v2.json"
            v2.write_text(json.dumps(verdict(6, ["x"], round=2, blocking=6,
                                             methods=("A_tests", "B_review"))))
            self._run("record", str(st), "--verdict", str(v2))
            rc, out = self._run("converged", str(st))
            self.assertEqual(rc, 0, out)
            self.assertIn("方式", out)

    def test_record_rejects_verdict_without_methods(self):
        """方式集合を比較に使うので、methods の無い verdict は受けない。"""
        with TemporaryDirectory() as t:
            st = Path(t) / "state.json"
            self._run("init", str(st), "--max-rounds", "3")
            v = json.loads(json.dumps(verdict(4, ["x"], round=1)))
            del v["methods"]
            vp = Path(t) / "v.json"; vp.write_text(json.dumps(v))
            rc, out = self._run("record", str(st), "--verdict", str(vp))
            self.assertEqual(rc, 2, out)

if __name__ == "__main__": unittest.main()
