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
  eval_state.py converged <state.json>   # exit 0 収束中 / 3 not-converging（前巡より blocking が減っていない）
  eval_state.py next      <state.json>   # exit 0 続行可 / 4 max-rounds 到達（終了して報告）/ 5 前巡 pass 済み
設計: 過去評価の本文（scores・comments）は次巡へ渡さない（フレッシュ評価）。渡すのは
「前巡 must-fix の反映確認」と「却下済み提案の一覧」だけ。人間ゲートはこのループの中に無い。

record の入力契約（穴①への対処）: verdict に findings_count / blocking_count / methods が
無ければ exit 2 で拒否する（旧形式の手書き verdict を黙って受けない＝呼び出し側が数え直せない
形にする）。verdict の round が state の round+1 と一致しなければ exit 2（巡の取り違え検出）。

収束判定は blocking_count で行う（ISSUE-07）。findings_count は評価単位の総数なので、成果物を
再構成したり方式を追加したりすると分母が変わり、改善した巡を「未収束」で止める（eval-1・eval-0 で実測）。
blocking_count は「直さないと合格にならないもの」だけ（A の失敗 + B の counted + Σ must_fix +
読者の「分からない」）を数える。それでも方式集合が変われば分母は変わるので、その巡は比較しない。
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
        for key in ("findings_count", "blocking_count", "methods"):
            if key not in v:
                print(f"record: verdict に {key} が無い（旧形式・手書きは拒否。数え直しではなく "
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
                             "blocking_count": int(v["blocking_count"]),
                             "methods": sorted(v["methods"]),
                             "must_fix": list(v.get("must_fix", [])),
                             "rejected": rejected_from_verdict + rejected_from_cli})
        save(sp, s)
        print(f"record: round={s['round']} blocking={s['history'][-1]['blocking_count']}"
              f" findings={s['history'][-1]['findings_count']}")
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
        prev, cur = h[-2], h[-1]
        if prev.get("methods") != cur.get("methods"):
            # 方式集合が変われば件数の分母が変わる。比較は成立しないので停止させない。
            # 分母が違う数を大小比較すると、改善した巡を「未収束」で止める（eval-0 で実測）。
            print(f"converging: 方式集合が変わったため件数比較は不成立"
                  f"（{'+'.join(prev.get('methods') or ['-'])} -> {'+'.join(cur.get('methods') or ['-'])}）。"
                  f"blocking {prev['blocking_count']} -> {cur['blocking_count']}")
            return 0
        if cur["blocking_count"] >= prev["blocking_count"] and not cur["pass"]:
            print(f"not-converging: blocking {prev['blocking_count']} -> {cur['blocking_count']}（減っていない）"
                  f" / findings {prev['findings_count']} -> {cur['findings_count']}")
            return 3
        print(f"converging: blocking {prev['blocking_count']} -> {cur['blocking_count']}"
              f" / findings {prev['findings_count']} -> {cur['findings_count']}")
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
