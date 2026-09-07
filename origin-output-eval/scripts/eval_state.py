#!/usr/bin/env python3
"""eval_state.py — 評価ループの巡をまたぐ状態（決定論・LLM 不使用）。

使い方:
  eval_state.py init      <state.json> --max-rounds N
  eval_state.py record    <state.json> --verdict <verdict.json> [--rejected <text> ...]
                          # verdict は judge_round.py の出力（--json-out の中身）をそのまま渡す。
                          # findings_count / must_fix / pass / round は verdict からそのまま読み、
                          # 呼び出し側は数え直さない。--rejected は「今巡で却下した提案」を人が
                          # 渡す欄（機械化できない）。複数回指定できる。
  eval_state.py carry     <state.json>   # 次巡のフレッシュ評価者へ渡してよいもの**だけ**を JSON で出す
  eval_state.py converged <state.json>   # exit 0 収束中 / 3 not-converging（前巡より指摘が減っていない）
  eval_state.py next      <state.json>   # exit 0 続行可 / 4 max-rounds 到達（終了して報告）/ 5 前巡 pass 済み
設計: 過去評価の本文（scores・comments）は次巡へ渡さない（フレッシュ評価）。渡すのは
「前巡 must-fix の反映確認」と「却下済み提案の一覧」だけ。人間ゲートはこのループの中に無い。

record の入力契約（穴①への対処）: verdict に findings_count が無ければ exit 2 で拒否する
（旧形式の手書き verdict を黙って受けない＝呼び出し側が数え直せない形にする）。verdict の
round が state の round+1 と一致しなければ exit 2（巡の取り違え検出）。
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path


def load(p: Path) -> dict:
    return json.loads(p.read_text())


def save(p: Path, s: dict) -> None:
    p.write_text(json.dumps(s, ensure_ascii=False, indent=1))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["init", "record", "carry", "converged", "next"])
    ap.add_argument("state")
    ap.add_argument("--max-rounds", type=int, default=3)
    ap.add_argument("--verdict")
    ap.add_argument("--rejected", action="append", default=None,
                     help="今巡で却下した提案（複数回指定可）。history に積み、次巡の carry へ渡す。")
    a = ap.parse_args(argv)
    sp = Path(a.state)
    if a.cmd == "init":
        save(sp, {"round": 0, "max_rounds": a.max_rounds, "history": []})
        print(f"init: max_rounds={a.max_rounds}")
        return 0
    s = load(sp)
    if a.cmd == "record":
        v = load(Path(a.verdict))
        if "findings_count" not in v:
            print("record: verdict に findings_count が無い（旧形式は拒否。数え直しではなく "
                  "judge_round の出力をそのまま渡すこと）")
            return 2
        expected_round = s["round"] + 1
        vround = v.get("round")
        if vround != expected_round:
            print(f"record: verdict の round={vround!r} が期待値 {expected_round} と不一致（巡の取り違え）")
            return 2
        rejected_from_verdict = list(v.get("rejected", []))
        rejected_from_cli = list(a.rejected) if a.rejected else []
        s["round"] += 1
        s["history"].append({"round": s["round"], "pass": bool(v.get("pass")),
                             "findings_count": int(v["findings_count"]),
                             "must_fix": list(v.get("must_fix", [])),
                             "rejected": rejected_from_verdict + rejected_from_cli})
        save(sp, s)
        print(f"record: round={s['round']} findings={s['history'][-1]['findings_count']}")
        return 0
    h = s["history"]
    if a.cmd == "carry":
        prev = h[-1] if h else {"must_fix": [], "rejected": []}
        rejected = [r for e in h for r in e.get("rejected", [])]
        print(json.dumps({"round": s["round"] + 1,
                          "verify_must_fix_from_prev_round": prev["must_fix"],
                          "already_rejected_proposals": rejected,
                          "note_for_readers": "初見読者にはこの carry を渡さない（サイド情報になる）"},
                         ensure_ascii=False))
        return 0
    if a.cmd == "converged":
        if len(h) < 2:
            print("converging: 比較対象なし（1巡目）")
            return 0
        if h[-1]["findings_count"] >= h[-2]["findings_count"] and not h[-1]["pass"]:
            print(f"not-converging: findings {h[-2]['findings_count']} -> {h[-1]['findings_count']}（減っていない）")
            return 3
        print(f"converging: findings {h[-2]['findings_count']} -> {h[-1]['findings_count']}")
        return 0
    if a.cmd == "next":
        if h and h[-1]["pass"]:
            print("done: 前巡で pass")
            return 5
        if s["round"] >= s["max_rounds"]:
            print(f"stop: max-rounds {s['max_rounds']} 到達。終了して最終報告へ（人間ゲートはループの外）")
            return 4
        print(f"continue: round {s['round'] + 1}/{s['max_rounds']}")
        return 0
    return 2


if __name__ == "__main__":
    sys.exit(main())
