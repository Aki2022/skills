---
title: 評価ループの共通プロトコル
---

# 共通プロトコル

## 1. 前提条件（呼び出し側の宣言）

対象に決定論ゲート（テスト・lint・実測）が存在するなら、**緑であること**を前提条件とする。
呼び出し側は「どのゲートを通したか」を宣言して本 skill を呼ぶ（例: `measure_lp.mjs` 失格ゼロ／
pytest 全緑／`audit_html.py` exit 0）。ゲートが無い対象（文章・企画書）は「該当ゲートなし」と宣言する。

理由（lp-review 2026-08-11 実測）: 表示が壊れたページを3人が採点し、全員が同じ表示不良を最重要欠陥として
報告。10項目中2項目が使い物にならず1人は再実行になった。

## 2. 方式の選択（skill 自身が選ぶ）

| 方式 | 使う条件 | 評価者 | 停止条件（機械判定） |
| --- | --- | --- | --- |
| **A. テスト** | 対象がコード・スクリプト・データ変換で、テストが書ける | テスト（既存 or 新設） | `tests.json` の exit_code == 0 |
| **B. レビュー skill** | 対象に既存のレビュー skill がある（コード→`code-review`、UI→`web-design-guidelines` `baseline-ui`、a11y→`fixing-accessibility`） | 既存 skill | `review.json` の findings が閾値以下（既定 0 件） |
| **C. 審査員** | 主観品質（納品物として通るか・読者に効くか・美しいか）が問われる | 独立した審査員＋初見読者（別 subagent） | 各審査員 ≥ judgeMinEach かつ平均 ≥ judgeAvgMin かつ must_fix ≤ 上限 かつ 初見読者「分からない」≤ 上限 |

規則: **A で判定できるものは A**。A にできない残りで **B が使えるものは B**。それでも残る主観品質だけ **C**。
複数該当なら **A → B → C の順**に、前が緑になってから次へ。選択と非選択の理由を最終報告に書く。

## 3. 評価計画（ループの外・事前）

評価者数・回数上限・トークン上限・対象・方式は **ws の Authorization Envelope に事前に書く**
（`eval-plan.template.md`）。単発利用では、この skill を呼んだ指示自体が計画にあたり、既定値
（`judge_round.py` の閾値既定・回数上限3）を使う。

## 4. ループ（無人・有界で完走する）

```text
# 準備（1回）
mkdir -p <eval_dir>; cp <rubric から写した> <eval_dir>/thresholds.json

# 各巡
1. eval_state.py carry <eval_dir>/state.json   # 2巡目以降。前巡 must-fix の反映確認と却下済みだけ
2. 評価者を新規 subagent として立てる（方式ごと・並列可）
   渡すのは 対象 + ルーブリック + carry のみ。**初見読者に carry を渡さない**（サイド情報になる）
3. 出力を <eval_dir>/round_K/ に置く（judges/*.json readers/*.json review.json tests.json）
4. <eval_dir>/round_K/provenance.json を書く（references/provenance.md）
5. run_round.sh <eval_dir> --thresholds <eval_dir>/thresholds.json [--max-rounds N]
      exit 0 → 合格。6 へ
      exit 1 → 不合格。指摘を A/B/C に分け、A/B を呼び出し側が直して次巡へ
      exit 2 → 入力の不備（provenance 不備・評価者の欠落や未申告・round 不一致・severity 語彙外・
               スキーマ違反）。**直してから再実行**。判定はまだ行われていない
      exit 3 → 未収束（前巡より blocking_count が減っていない）。上限前でも終了して 6 へ
               方式集合が変わった巡は分母が変わるので比較しない（停止しない）
      exit 4 → 回数上限に到達。終了して 6 へ
      exit 5 → 前巡で既に合格。6 へ
6. 最終報告（final-report.template.md）
```

このループの中に、人へ問い合わせる手順は存在しない。上限到達・未収束・C 分類は**終了して報告**する。
判定源は常に `judge_round.py` と `eval_state.py` の exit code と出力であり、LLM が数えて宣言しない。

### 収束は `blocking_count` で見る（`findings_count` ではない）

`findings_count` は評価単位の総数（A の失敗 + B の counted + Σ must_fix + Σ comments +
読者の非「分かる」）で、**分母が変わると巡間で比較できない**。実測: 文書を3段落から7見出しへ
再構成した巡で「分からない」は 4→1 に減ったのに `findings_count` は 6→7 になり、改善した巡が
`not-converging` で止まった。

そこで収束は `blocking_count`（A の失敗 + B の counted + Σ must_fix + 読者の「分からない」。
**comments と「引っかかる」を除く** = 直さないと合格にならないもの）で判定する。
さらに**方式集合が変わった巡**（例: 方式A が緑になり方式B を足した）は blocking_count でも
分母が変わるため、`eval_state.py converged` は比較を「不成立」として停止させない。
`findings_count` は人が推移を読む値として残す。

## 5. トリアージ（指摘をそのまま反映しない）

| 分類 | 内容 | 処理 |
| --- | --- | --- |
| A | 成果物のミス（仕様と食い違う・壊れている） | ループ内で直す |
| B | 表記・スタイリング（成果物内で閉じる） | ループ内で直す |
| C | ロジック・構成など**正典（仕様・outline）の変更に及ぶ** | 直さない。保留リストに積み、最終報告で提示 |

却下した提案は `rejected` として記録し、次巡の評価者に「既に却下済み」として渡す（蒸し返し防止）。

## 6. 独立性（スクリプトで守れない規律）

- **作った本人に採点させない。** 単発利用でも審査員は別 subagent。
- **毎巡フレッシュ。** 同一エージェントの再評価は文脈汚染で甘くなる（content-eval 実測）。
- **サイド情報を渡さない。** 読者評価には「その段階で読み手が受け取る成果物・それ単独」。画像段階は画像のみ。
  設計意図・仕様書・過去版を渡すと「そう書いてあるから読める」に流れる（lp-review）。
- **キャッシュを外す。** 採点者が修正前の版を見て報告した実測あり（lp-review 2026-08-11）。

これらは**`provenance.json` で機械が検証する**（`references/provenance.md`）。申告と実ファイルの不一致・
`fresh: false`・読者への余分な入力・キャッシュ未クリアは `exit 2` で止まり、判定に進めない。
規律を文書に書くだけでは破られた実績があるので、構造で拒否する形にしてある。
申告の内容はそのまま最終報告の「評価者の起動条件」欄に載る（`verdict.json` の `provenance`）。

## 7. 非代理

本 skill の合格は人間確定の代わりにならない（content-eval: AI 審査6回合格の直後に人間の初見レビューで
本編全面差し戻し・モックアップ20枚廃棄）。最終報告の後に人間ゲートが1つある — それはこの skill の外。
