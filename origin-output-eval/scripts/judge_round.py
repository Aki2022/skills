#!/usr/bin/env python3
"""judge_round.py — 評価ループ1巡の合否を機械判定する（決定論・LLM 不使用）。

使い方:
  judge_round.py <round_dir> --thresholds <thresholds.json>
                 [--expect-judges N] [--expect-readers M] [--round K] [--json-out <path>]

round_dir の中身:
  provenance.json   必須。評価者の起動条件の申告（無ければ exit 2）
  tests.json        任意（方式A）: {"exit_code": 0}
  review.json       任意（方式B）: {"findings": [{"severity": "critical|major|minor", ...}, ...]}
  judges/*.json     任意（方式C 審査員、content-eval 固定スキーマ）
  readers/*.json    任意（方式C 初見読者、同）
tests.json / review.json / judges/*.json / readers/*.json が1つも無ければ exit 2。

thresholds.json: judgeMinEach / judgeAvgMin / criticalMustFixMax / freshReaderUnclearMax /
                 reviewFindingsMax / reviewMaxSeverity

provenance.json は起動条件を機械で検証する（黙って合格させない）:
  - round が --round（省略時は provenance 自身の round）と一致する
  - evaluators の申告が実ファイル（judges/<id>.json・readers/<id>.json）と双方向に一致する
    （届いていない評価者・未申告の評価者を両方とも許さない）
  - fresh が全 entry で true（同一エージェントの再評価を許さない）
  - reader の inputs は artifact のみ／judge の inputs は artifact・rubric・carry のみ
    （読者評価にサイド情報を渡さない）
  - cache_cleared が true（採点者が旧版を見た疑いを機械で否定できないなら通さない）
  - --expect-judges / --expect-readers を明示した場合、実ファイル数と食い違えば exit 2
    （provenance 自体が実態に合わせて縮小申告していても、計画側の数値で欠落を検出する）
評価者 JSON 個別の検証:
  - round が provenance の round と不一致（前巡の残骸が紛れ込んでいる）→ exit 2
  - review.json の severity が critical/major/minor 以外 → exit 2（既定値で吸収しない）

findings_count は機械が一意に数える:
  A(tests.json の exit_code != 0 なら1、それ以外0)
  + B(review.json の counted 件数)
  + C(全審査員の must_fix 件数の合計)
  + C(全審査員の comments 件数の合計)
  + C(全初見読者の per_slide のうち verdict != "分かる" の件数)

scope: 本スクリプトは「渡された評価出力が閾値を満たすか」（A: exit code / B: 指摘件数と重大度 /
C: 審査員の各下限・平均・must-fix・初見読者の「分からない」）に加えて、provenance.json が申告する
起動条件が実ファイルと機械的に整合するか（評価者の申告漏れ・未申告・フレッシュ性・読者への入力・
cache_cleared・round 一致）を見る。申告そのものの真偽（本当にフレッシュか・本当にキャッシュを
外したか）そのものは見ない（自己申告を信じる。虚偽申告の検出はできない）。

exit: 0 pass / 1 fail / 2 入力不備（スキーマ違反・provenance 不備・評価者の欠落/未申告・round 不一致・
      severity 語彙外）
出力: 標準出力は1行目 scope ＋ JSON（pass, round, methods, reasons, findings_count, must_fix,
      rejected, provenance）。--json-out 指定時はそのパスへ scope 行を含まない純 JSON を書く
      （eval_state.py へ機械的に渡す口）。
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path

SEV = {"minor": 1, "major": 2, "critical": 3}
JUDGE_REQ = {"persona", "round", "scores", "total_score", "total_max", "verdict", "must_fix"}
READER_REQ = {"persona", "round", "per_slide"}
READER_VERDICTS = {"分かる", "引っかかる", "分からない"}
PROV_REQ = {"round", "evaluators", "gates_passed", "cache_cleared"}
EVALUATOR_REQ = {"role", "id", "fresh", "inputs"}
ROLE_ALLOWED_INPUTS = {"judge": {"artifact", "rubric", "carry"}, "reader": {"artifact"}}


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


def load_provenance(d: Path, round_arg):
    """provenance.json を検証する。違反は全て SchemaError（呼び出し側で exit 2 にする）。
    戻り値: (provenance の中身, round 値, {id: Path} な judges, {id: Path} な readers)"""
    pp = d / "provenance.json"
    if not pp.is_file():
        raise SchemaError("provenance.json: ファイルが無い（起動条件を申告する必須ファイル）")
    prov = load_json(pp)
    if not isinstance(prov, dict):
        raise SchemaError("provenance.json: トップレベルが object でない")
    miss = PROV_REQ - set(prov)
    if miss:
        raise SchemaError(f"provenance.json: 必須キー欠落 {sorted(miss)}")

    prov_round = prov.get("round")
    if not isinstance(prov_round, int) or isinstance(prov_round, bool):
        raise SchemaError("provenance.json: round が整数でない")
    if round_arg is not None and prov_round != round_arg:
        raise SchemaError(f"provenance.json: round={prov_round} が --round {round_arg} と不一致")
    round_value = prov_round

    if prov.get("cache_cleared") is not True:
        raise SchemaError("provenance.json: cache_cleared が true でない"
                           "（採点者が修正前の版を見た疑いを機械で否定できない）")

    evaluators = prov.get("evaluators")
    if not isinstance(evaluators, list):
        raise SchemaError("provenance.json: evaluators が list でない")

    declared: dict[str, dict[str, dict]] = {"judge": {}, "reader": {}}
    for i, e in enumerate(evaluators):
        if not isinstance(e, dict):
            raise SchemaError(f"provenance.json: evaluators[{i}] が object でない")
        emiss = EVALUATOR_REQ - set(e)
        if emiss:
            raise SchemaError(f"provenance.json: evaluators[{i}] 必須キー欠落 {sorted(emiss)}")
        role = e["role"]
        if role not in ("judge", "reader"):
            raise SchemaError(f"provenance.json: evaluators[{i}].role が語彙外 '{role}'")
        if e.get("fresh") is not True:
            raise SchemaError(f"provenance.json: evaluators[{i}]（id={e.get('id')!r}）の"
                               " fresh が true でない（独立性違反。毎巡フレッシュが原則）")
        inputs = e.get("inputs")
        if not isinstance(inputs, list):
            raise SchemaError(f"provenance.json: evaluators[{i}].inputs が list でない")
        allowed = ROLE_ALLOWED_INPUTS[role]
        bad = [x for x in inputs if x not in allowed]
        if bad:
            reason = "reader にはサイド情報を渡さない" if role == "reader" else "judge の許容は artifact/rubric/carry のみ"
            raise SchemaError(f"provenance.json: evaluators[{i}]（role={role}）の inputs に許容外 {bad}（{reason}）")
        eid = e["id"]
        if eid in declared[role]:
            raise SchemaError(f"provenance.json: id {eid!r}（role={role}）が重複申告されている")
        declared[role][eid] = e

    judge_dir, reader_dir = d / "judges", d / "readers"
    judge_files = {p.stem: p for p in judge_dir.glob("*.json")} if judge_dir.is_dir() else {}
    reader_files = {p.stem: p for p in reader_dir.glob("*.json")} if reader_dir.is_dir() else {}

    for eid in declared["judge"]:
        if eid not in judge_files:
            raise SchemaError(f"provenance.json: judge {eid!r} を申告しているが"
                               f" judges/{eid}.json が無い（届いていない評価者を検出）")
    for eid in judge_files:
        if eid not in declared["judge"]:
            raise SchemaError(f"judges/{eid}.json が実在するが provenance.json に申告が無い（未申告の評価者）")
    for eid in declared["reader"]:
        if eid not in reader_files:
            raise SchemaError(f"provenance.json: reader {eid!r} を申告しているが"
                               f" readers/{eid}.json が無い（届いていない評価者を検出）")
    for eid in reader_files:
        if eid not in declared["reader"]:
            raise SchemaError(f"readers/{eid}.json が実在するが provenance.json に申告が無い（未申告の評価者）")

    return prov, round_value, judge_files, reader_files


def check_expect(explicit, declared_count: int, actual_count: int, label: str):
    """--expect-judges/--expect-readers: 省略時は provenance の申告数をそのまま使う（自明に一致）。
    明示された場合はそれを優先し、実ファイル数と食い違えば違反（評価者の欠落を計画側の数値で検出）。"""
    expected = explicit if explicit is not None else declared_count
    if expected != actual_count:
        raise SchemaError(f"{label}: 期待 {expected} 件に対し実ファイル {actual_count} 件（評価者の欠落）")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("round_dir")
    ap.add_argument("--thresholds", required=True)
    ap.add_argument("--expect-judges", type=int, default=None)
    ap.add_argument("--expect-readers", type=int, default=None)
    ap.add_argument("--round", type=int, default=None)
    ap.add_argument("--json-out")
    a = ap.parse_args(argv)
    d = Path(a.round_dir)
    print("scope: 渡された評価出力が閾値を満たすか（A: exit code / B: 指摘件数と重大度 / C: 審査員の各下限・平均・"
          "must-fix・初見読者の「分からない」）に加え、provenance.json が申告する起動条件が実ファイルと機械的に"
          "整合するか（評価者の申告漏れ・未申告・フレッシュ性・読者への入力・cache_cleared・round 一致）を見る。"
          "申告そのものの真偽（本当にフレッシュか・本当にキャッシュを外したか）は見ない。")
    try:
        th = load_json(Path(a.thresholds))
        prov, round_value, judge_files, reader_files = load_provenance(d, a.round)

        declared_judges = sum(1 for e in prov["evaluators"] if e["role"] == "judge")
        declared_readers = sum(1 for e in prov["evaluators"] if e["role"] == "reader")
        check_expect(a.expect_judges, declared_judges, len(judge_files), "--expect-judges")
        check_expect(a.expect_readers, declared_readers, len(reader_files), "--expect-readers")

        methods: dict = {}
        reasons: list[str] = []
        seen = False
        a_findings = 0
        # A
        tp = d / "tests.json"
        if tp.is_file():
            seen = True
            t = load_json(tp)
            ok = t.get("exit_code") == 0
            methods["A_tests"] = {"exit_code": t.get("exit_code"), "pass": ok}
            if not ok:
                reasons.append(f"tests: exit_code={t.get('exit_code')} (0 が必要)")
                a_findings = 1
        # B
        b_findings = 0
        rp = d / "review.json"
        if rp.is_file():
            seen = True
            rv = load_json(rp)
            f = rv.get("findings")
            if not isinstance(f, list):
                raise SchemaError("review.json: findings が list でない")
            for x in f:
                sev = x.get("severity")
                if sev not in SEV:
                    raise SchemaError(f"review.json: severity が語彙外 {sev!r}（critical/major/minor のみ。既定値で吸収しない）")
            maxsev = th.get("reviewMaxSeverity", "minor")
            counted = [x for x in f if SEV[x["severity"]] >= SEV.get(maxsev, 1)]
            ok = len(counted) <= th.get("reviewFindingsMax", 0)
            methods["B_review"] = {"findings": len(f), "counted": len(counted), "pass": ok}
            b_findings = len(counted)
            if not ok:
                reasons.append(f"reviewFindingsMax: {len(counted)} 件 (> {th.get('reviewFindingsMax', 0)})")
        # C
        judges = sorted(judge_files.values())
        readers = sorted(reader_files.values())
        c_must_fix: list = []
        c_comments = 0
        c_nonwakaru = 0
        if judges or readers:
            seen = True
            pcts, mf = [], 0
            for jp in judges:
                j = load_json(jp)
                if j.get("round") != round_value:
                    raise SchemaError(f"{jp.name}: round={j.get('round')!r} が今巡（{round_value}）と不一致"
                                       "（前巡の残骸が紛れ込んでいる）")
                check_judge(j, jp.name)
                pcts.append(100.0 * j["total_score"] / j["total_max"])
                mf += len(j["must_fix"])
                c_must_fix.extend(j["must_fix"])
                c_comments += len(j.get("comments", []))
            unclear = 0
            for rp2 in readers:
                r = load_json(rp2)
                if r.get("round") != round_value:
                    raise SchemaError(f"{rp2.name}: round={r.get('round')!r} が今巡（{round_value}）と不一致"
                                       "（前巡の残骸が紛れ込んでいる）")
                check_reader(r, rp2.name)
                unclear += sum(1 for ps in r["per_slide"] if ps["verdict"] == "分からない")
                c_nonwakaru += sum(1 for ps in r["per_slide"] if ps["verdict"] != "分かる")
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
        findings_count = a_findings + b_findings + len(c_must_fix) + c_comments + c_nonwakaru
        result = {
            "pass": passed,
            "round": round_value,
            "methods": methods,
            "reasons": reasons,
            "findings_count": findings_count,
            "must_fix": c_must_fix,
            "rejected": [],
            "provenance": prov,
        }
        print(json.dumps(result, ensure_ascii=False, indent=1))
        if a.json_out:
            Path(a.json_out).write_text(json.dumps(result, ensure_ascii=False, indent=1))
        return 0 if passed else 1
    except SchemaError as e:
        print(json.dumps({"pass": False, "schema_error": str(e)}, ensure_ascii=False))
        return 2


if __name__ == "__main__":
    sys.exit(main())
