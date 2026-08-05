#!/usr/bin/env python3
"""③開始前のプリフライト: process/feedback.json の verdict が全て ok か確認する。

usage: python3 check_feedback.py <デッキdir> [--slides 6,9,11]

②（デザインモックアップの人間承認）を終えずに③（ネイティブビルド）へ進むと、
未承認の構図のまま実装が積み上がり、後から構図レベルの差し戻しで実装を捨てることになる
（2026-08-05 実失敗: アンカー4枚だけ承認して残り13枚を未承認のままビルドした）。
pipeline.md ② の「全件 ok になったら②完了 → ③へ進む」を機械的に担保する。

exit 0 = 全ok（③へ進んでよい） / exit 1 = 未承認あり / exit 2 = 使い方の誤り
"""
import json
import os
import sys


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        print(__doc__)
        return 2
    deck = args[0].rstrip("/")

    only = None
    for a in sys.argv[1:]:
        if a.startswith("--slides"):
            val = a.split("=", 1)[1] if "=" in a else sys.argv[sys.argv.index(a) + 1]
            only = {int(x) for x in val.replace(" ", "").split(",") if x}

    path = os.path.join(deck, "process", "feedback.json")
    if not os.path.exists(path):
        print(f"NG: {path} が無い。②のレビューを実施し verdict を記録すること")
        print("    （②を回していないなら、③のビルダーを起動してはならない）")
        return 1

    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    slides = data.get("slides", [])
    if not slides:
        print(f"NG: {path} に slides が無い")
        return 1

    pending, revise = [], []
    for s in slides:
        n = s.get("slide")
        if only is not None and n not in only:
            continue
        v = s.get("verdict")
        if v == "ok":
            continue
        (revise if v == "revise" else pending).append((n, s.get("comment", "")))

    if pending or revise:
        print(f"NG: ②未完了 — ③のビルドに進んではならない（{path}）")
        for n, _ in sorted(pending):
            print(f"    S{n}: 未回答（人間のレビュー待ち）")
        for n, c in sorted(revise):
            print(f"    S{n}: revise — {c}")
        print("  → モックアップを再生成し、全件 ok を得てから③へ進むこと")
        return 1

    checked = len(slides) if only is None else len(only)
    print(f"OK: 対象{checked}枚すべて verdict=ok。③へ進んでよい")
    return 0


if __name__ == "__main__":
    sys.exit(main())
