#!/usr/bin/env python3
"""eval_state.py — 評価ループの巡をまたぐ状態（決定論・LLM 不使用）。

使い方:
  eval_state.py init      <state.json> --max-rounds N
  eval_state.py record    <state.json> --verdict <verdict.json>   # judge_round の出力に findings_count/must_fix/rejected を添えたもの
  eval_state.py carry     <state.json>   # 次巡のフレッシュ評価者へ渡してよいもの**だけ**を JSON で出す
  eval_state.py converged <state.json>   # exit 0 収束中 / 3 not-converging（前巡より指摘が減っていない）
  eval_state.py next      <state.json>   # exit 0 続行可 / 4 max-rounds 到達（終了して報告）/ 5 前巡 pass 済み
設計: 過去評価の本文（scores・comments）は次巡へ渡さない（フレッシュ評価）。渡すのは
「前巡 must-fix の反映確認」と「却下済み提案の一覧」だけ。人間ゲートはこのループの中に無い。
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
    a = ap.parse_args(argv)
    sp = Path(a.state)
    if a.cmd == "init":
        save(sp, {"round": 0, "max_rounds": a.max_rounds, "history": []})
        print(f"init: max_rounds={a.max_rounds}")
        return 0
    s = load(sp)
    if a.cmd == "record":
        v = load(Path(a.verdict))
        s["round"] += 1
        s["history"].append({"round": s["round"], "pass": bool(v.get("pass")),
                             "findings_count": int(v.get("findings_count", 0)),
                             "must_fix": list(v.get("must_fix", [])),
                             "rejected": list(v.get("rejected", []))})
        save(sp, s)
        print(f"record: round={s['round']} findings={s['history'][-1]['findings_count']}")
        return 0
    h = s["history"]
    if a.cmd == "carry":
        prev = h[-1] if h else {"must_fix": [], "rejected": []}
        rejected = [r for e in h for r in e.get("rejected", [])]
        print(json.dumps({"round": s["round"] + 1,
                          "verify_must_fix_from_prev_round": prev["must_fix"],
                          "already_rejected_proposals": rejected}, ensure_ascii=False))
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
