#!/usr/bin/env python3
"""judge_round.py — 評価ループ1巡の合否を機械判定する（決定論・LLM 不使用）。

使い方: python3 judge_round.py <round_dir> --thresholds <thresholds.json>

round_dir の中身（存在するものだけ判定する。1つも無ければ exit 2）:
  tests.json        方式A: {"exit_code": 0}
  review.json       方式B: {"findings": [{"severity": "critical|major|minor", ...}, ...]}
  judges/*.json     方式C 審査員（content-eval 固定スキーマ）
  readers/*.json    方式C 初見読者（同）
thresholds.json: judgeMinEach / judgeAvgMin / criticalMustFixMax / freshReaderUnclearMax /
                 reviewFindingsMax / reviewMaxSeverity
scope: 本スクリプトは「渡された出力が閾値を満たすか」だけを見る。評価者が独立か・フレッシュか・
       何を渡されたかは見ない（それは最終報告の起動条件欄で事後検証する）。
exit: 0 pass / 1 fail / 2 スキーマ違反・入力なし
出力: 1行目に scope、以降 JSON（pass, methods, reasons）。
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path

SEV = {"minor": 1, "major": 2, "critical": 3}
JUDGE_REQ = {"persona", "round", "scores", "total_score", "total_max", "verdict", "must_fix"}
READER_REQ = {"persona", "round", "per_slide"}
READER_VERDICTS = {"分かる", "引っかかる", "分からない"}


class SchemaError(Exception):
    pass


def load_json(p: Path):
    try:
        return json.loads(p.read_text())
    except Exception as e:  # noqa: BLE001
        raise SchemaError(f"{p.name}: JSON として読めない ({e})")


def check_judge(j: dict, name: str):
    miss = JUDGE_REQ - set(j)
    if miss:
        raise SchemaError(f"{name}: 必須キー欠落 {sorted(miss)}")
    if j["verdict"] not in ("pass", "revise"):
        raise SchemaError(f"{name}: verdict が語彙外 '{j['verdict']}'")
    if not isinstance(j["scores"], list) or not j["scores"]:
        raise SchemaError(f"{name}: scores が空")
    for s in j["scores"]:
        if not {"item", "score", "max"} <= set(s) or s["score"] > s["max"]:
            raise SchemaError(f"{name}: scores の要素が不正 {s}")
    if not j["total_max"]:
        raise SchemaError(f"{name}: total_max が 0")


def check_reader(r: dict, name: str):
    miss = READER_REQ - set(r)
    if miss:
        raise SchemaError(f"{name}: 必須キー欠落 {sorted(miss)}")
    for ps in r["per_slide"]:
        if ps.get("verdict") not in READER_VERDICTS:
            raise SchemaError(f"{name}: per_slide.verdict が語彙外 {ps.get('verdict')!r}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("round_dir")
    ap.add_argument("--thresholds", required=True)
    a = ap.parse_args(argv)
    d = Path(a.round_dir)
    print("scope: 渡された評価出力が閾値を満たすか（A: exit code / B: 指摘件数と重大度 / C: 審査員の各下限・平均・"
          "must-fix・初見読者の「分からない」）。評価者の独立性・フレッシュ性・入力内容は見ない。")
    try:
        th = load_json(Path(a.thresholds))
        methods: dict = {}
        reasons: list[str] = []
        seen = False
        # A
        tp = d / "tests.json"
        if tp.is_file():
            seen = True
            t = load_json(tp)
            ok = t.get("exit_code") == 0
            methods["A_tests"] = {"exit_code": t.get("exit_code"), "pass": ok}
            if not ok:
                reasons.append(f"tests: exit_code={t.get('exit_code')} (0 が必要)")
        # B
        rp = d / "review.json"
        if rp.is_file():
            seen = True
            rv = load_json(rp)
            f = rv.get("findings")
            if not isinstance(f, list):
                raise SchemaError("review.json: findings が list でない")
            maxsev = th.get("reviewMaxSeverity", "minor")
            counted = [x for x in f if SEV.get(x.get("severity", "minor"), 1) >= SEV.get(maxsev, 1)]
            ok = len(counted) <= th.get("reviewFindingsMax", 0)
            methods["B_review"] = {"findings": len(f), "counted": len(counted), "pass": ok}
            if not ok:
                reasons.append(f"reviewFindingsMax: {len(counted)} 件 (> {th.get('reviewFindingsMax', 0)})")
        # C
        judges = sorted((d / "judges").glob("*.json")) if (d / "judges").is_dir() else []
        readers = sorted((d / "readers").glob("*.json")) if (d / "readers").is_dir() else []
        if judges or readers:
            seen = True
            pcts, mf = [], 0
            for jp in judges:
                j = load_json(jp)
                check_judge(j, jp.name)
                pcts.append(100.0 * j["total_score"] / j["total_max"])
                mf += len(j["must_fix"])
            unclear = 0
            for rp2 in readers:
                r = load_json(rp2)
                check_reader(r, rp2.name)
                unclear += sum(1 for ps in r["per_slide"] if ps["verdict"] == "分からない")
            c_ok = True
            if judges:
                if min(pcts) < th.get("judgeMinEach", 75):
                    c_ok = False
                    reasons.append(f"judgeMinEach: min={min(pcts):.1f} (< {th.get('judgeMinEach', 75)})")
                avg = sum(pcts) / len(pcts)
                if avg < th.get("judgeAvgMin", 80):
                    c_ok = False
                    reasons.append(f"judgeAvgMin: avg={avg:.1f} (< {th.get('judgeAvgMin', 80)})")
                if mf > th.get("criticalMustFixMax", 0):
                    c_ok = False
                    reasons.append(f"criticalMustFixMax: must_fix={mf} (> {th.get('criticalMustFixMax', 0)})")
            if readers and unclear > th.get("freshReaderUnclearMax", 0):
                c_ok = False
                reasons.append(f"freshReaderUnclearMax: 分からない={unclear} (> {th.get('freshReaderUnclearMax', 0)})")
            methods["C_judges"] = {"judges": len(judges), "readers": len(readers),
                                   "min_pct": round(min(pcts), 1) if pcts else None,
                                   "avg_pct": round(sum(pcts) / len(pcts), 1) if pcts else None,
                                   "must_fix": mf, "unclear": unclear, "pass": c_ok}
        if not seen:
            raise SchemaError(f"{d}: 判定対象の出力が1つも無い（tests.json / review.json / judges/ / readers/）")
        passed = not reasons
        print(json.dumps({"pass": passed, "methods": methods, "reasons": reasons}, ensure_ascii=False, indent=1))
        return 0 if passed else 1
    except SchemaError as e:
        print(json.dumps({"pass": False, "schema_error": str(e)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    sys.exit(main())
