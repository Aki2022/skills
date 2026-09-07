"""judge_round.py — 1巡の合否を機械判定する。陽性対照を先に書き、実装前に赤を見る。"""
import importlib.util, io, json, os, unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPT = Path(__file__).resolve().parents[1] / "judge_round.py"
SPEC = importlib.util.spec_from_file_location("judge_round", SCRIPT)
M = importlib.util.module_from_spec(SPEC); assert SPEC.loader; SPEC.loader.exec_module(M)

THRESH = {"judgeMinEach": 75, "judgeAvgMin": 80, "criticalMustFixMax": 0,
          "freshReaderUnclearMax": 0, "reviewFindingsMax": 0, "reviewMaxSeverity": "minor"}


def judge(total, must_fix=(), comments=(), round_=1):
    return {"persona": "J", "round": round_,
            "scores": [{"item": "a", "score": total, "max": 100, "reason": "r"}],
            "total_score": total, "total_max": 100, "verdict": "pass" if total >= 75 else "revise",
            "strengths": [], "improvements": [], "must_fix": list(must_fix),
            "comments": list(comments), "advice": []}


def reader(unclear=0, hikkakaru=0, total=3, round_=1):
    n = max(total, unclear + hikkakaru)
    verdicts = (["分からない"] * unclear) + (["引っかかる"] * hikkakaru)
    verdicts += ["分かる"] * (n - len(verdicts))
    per = [{"n": i + 1, "verdict": v, "note": "x"} for i, v in enumerate(verdicts)]
    return {"persona": "初見読者", "round": round_, "per_slide": per, "through_line": "t",
            "naive_questions": [], "top_fixes": []}


def make_provenance(round_=1, judge_ids=(), reader_ids=(), cache_cleared=True, fresh=True,
                     judge_inputs=("artifact", "rubric"), reader_inputs=("artifact",),
                     gates=("dummy gate",)):
    evaluators = [{"role": "judge", "id": jid, "fresh": fresh, "inputs": list(judge_inputs)}
                  for jid in judge_ids]
    evaluators += [{"role": "reader", "id": rid, "fresh": fresh, "inputs": list(reader_inputs)}
                   for rid in reader_ids]
    return {"round": round_, "evaluators": evaluators, "gates_passed": list(gates),
            "cache_cleared": cache_cleared}


class JudgeRoundTest(unittest.TestCase):
    def _round(self, tmp, judges=None, readers=None, review=None, tests=None, round_=1, prov=None):
        d = Path(tmp) / "round1"
        (d / "judges").mkdir(parents=True)
        (d / "readers").mkdir()
        judge_ids, reader_ids = [], []
        for i, j in enumerate(judges or []):
            jid = f"j{i}"
            (d / "judges" / f"{jid}.json").write_text(json.dumps(j))
            judge_ids.append(jid)
        for i, r in enumerate(readers or []):
            rid = f"r{i}"
            (d / "readers" / f"{rid}.json").write_text(json.dumps(r))
            reader_ids.append(rid)
        if review is not None:
            (d / "review.json").write_text(json.dumps(review))
        if tests is not None:
            (d / "tests.json").write_text(json.dumps(tests))
        p = prov if prov is not None else make_provenance(round_=round_, judge_ids=judge_ids, reader_ids=reader_ids)
        (d / "provenance.json").write_text(json.dumps(p))
        (Path(tmp) / "thresholds.json").write_text(json.dumps(THRESH))
        return d

    def _run(self, d, extra_args=()):
        buf = io.StringIO()
        with redirect_stdout(buf):
            rc = M.main([str(d), "--thresholds", str(d.parent / "thresholds.json"), *extra_args])
        return rc, buf.getvalue()

    @staticmethod
    def _json_part(out):
        return json.loads(out[out.index("{"):])

    # --- 既存契約（閾値判定そのもの。provenance.json 必須化に合わせて更新） ---

    def test_all_pass_is_green(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, [judge(85), judge(90)], [reader(0)],
                                             {"findings": []}, {"exit_code": 0}))
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
            rc, out = self._run(self._round(t, [judge(90, must_fix=["x"])], [reader(0)]))
            self.assertEqual(rc, 1, out); self.assertIn("criticalMustFixMax", out)

    def test_fresh_reader_unclear_blocks(self):
        with TemporaryDirectory() as t:
            rc, out = self._run(self._round(t, [judge(90)], [reader(unclear=1)]))
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

    # --- 硬化契約: provenance.json による起動条件の機械検証（陽性対照 1-8, 10） ---

    def test_pc1_missing_provenance_is_error(self):
        """陽性対照1: provenance.json 無し → exit 2。"""
        with TemporaryDirectory() as t:
            d = self._round(t, [judge(90)], [reader(0)])
            os.remove(d / "provenance.json")
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("provenance", out)

    def test_pc2_provenance_overclaims_judges_is_error(self):
        """陽性対照2: 計画3人分を provenance が申告するが実ファイルは1件 → exit 2（②）。"""
        with TemporaryDirectory() as t:
            d = self._round(t, [judge(90)], [reader(0)])
            prov = json.loads((d / "provenance.json").read_text())
            prov["evaluators"] = [
                {"role": "judge", "id": "j0", "fresh": True, "inputs": ["artifact", "rubric"]},
                {"role": "judge", "id": "j1", "fresh": True, "inputs": ["artifact", "rubric"]},
                {"role": "judge", "id": "j2", "fresh": True, "inputs": ["artifact", "rubric"]},
                {"role": "reader", "id": "r0", "fresh": True, "inputs": ["artifact"]},
            ]
            (d / "provenance.json").write_text(json.dumps(prov))
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out)

    def test_pc2b_expect_judges_flag_catches_shortfall(self):
        """--expect-judges を明示すると、provenance が実態に合わせ正直に縮小申告していても検出する。"""
        with TemporaryDirectory() as t:
            d = self._round(t, [judge(90)], [reader(0)])  # provenance は j0 だけを正直に申告
            rc, out = self._run(d, extra_args=["--expect-judges", "3"])
            self.assertEqual(rc, 2, out); self.assertIn("expect-judges", out)

    def test_pc3_undeclared_evaluator_file_is_error(self):
        """陽性対照3: judges/j9.json が実在するが provenance に申告が無い → exit 2（未申告）。"""
        with TemporaryDirectory() as t:
            d = self._round(t, [judge(90)], [reader(0)])
            (d / "judges" / "j9.json").write_text(json.dumps(judge(85)))
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("未申告", out)

    def test_pc4_not_fresh_is_error(self):
        """陽性対照4: fresh: false の entry → exit 2。"""
        with TemporaryDirectory() as t:
            prov = make_provenance(judge_ids=["j0"], reader_ids=["r0"], fresh=False)
            d = self._round(t, [judge(90)], [reader(0)], prov=prov)
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("fresh", out)

    def test_pc5_reader_side_info_is_error(self):
        """陽性対照5: reader の inputs に rubric が入っている → exit 2。"""
        with TemporaryDirectory() as t:
            prov = make_provenance(judge_ids=["j0"], reader_ids=["r0"], reader_inputs=["artifact", "rubric"])
            d = self._round(t, [judge(90)], [reader(0)], prov=prov)
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("inputs", out)

    def test_pc6_cache_not_cleared_is_error(self):
        """陽性対照6: cache_cleared: false → exit 2。"""
        with TemporaryDirectory() as t:
            prov = make_provenance(judge_ids=["j0"], reader_ids=["r0"], cache_cleared=False)
            d = self._round(t, [judge(90)], [reader(0)], prov=prov)
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("cache_cleared", out)

    def test_pc7_stale_round_evaluator_is_error(self):
        """陽性対照7: round 2 のディレクトリに round 1 の審査員 JSON が混ざる → exit 2（④）。"""
        with TemporaryDirectory() as t:
            d = self._round(t, [judge(90, round_=1)], [reader(0, round_=2)], round_=2)
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("round", out)

    def test_pc8_severity_out_of_vocab_is_error(self):
        """陽性対照8: severity: "blocker" → exit 2（③）。"""
        with TemporaryDirectory() as t:
            d = self._round(t, review={"findings": [{"severity": "blocker", "msg": "x"}]})
            rc, out = self._run(d)
            self.assertEqual(rc, 2, out); self.assertIn("severity", out)

    def test_pc9_findings_count_formula(self):
        """陽性対照9: must_fix 2 + comments 3 + 読者の非「分かる」1 = findings_count 6。"""
        with TemporaryDirectory() as t:
            d = self._round(
                t,
                [judge(90, must_fix=["a", "b"],
                       comments=[{"loc": "p1", "comment": "c1"}, {"loc": "p2", "comment": "c2"},
                                 {"loc": "p3", "comment": "c3"}])],
                [reader(unclear=1)],
            )
            _, out = self._run(d)
            self.assertEqual(self._json_part(out)["findings_count"], 6, out)

    def test_pc10_json_out_is_pure_json(self):
        """陽性対照10: --json-out は scope 行を含まない純 JSON を書く。"""
        with TemporaryDirectory() as t:
            d = self._round(t, [judge(90)], [reader(0)])
            out_path = Path(t) / "verdict.json"
            rc, out = self._run(d, extra_args=["--json-out", str(out_path)])
            self.assertEqual(rc, 0, out)
            content = out_path.read_text()
            self.assertFalse(content.lstrip().startswith("scope"), content)
            parsed = json.loads(content)
            self.assertIn("findings_count", parsed)
            self.assertEqual(parsed["round"], 1)


if __name__ == "__main__": unittest.main()
