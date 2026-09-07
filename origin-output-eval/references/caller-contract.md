---
title: 呼び出し規約と単発利用
---

# 呼び出し規約（工場から）と単発利用（人間から）

## A. 工場から呼ぶ（例: `origin-web-design` 工程4 — 会社サイトのイラスト5枚と LP）

渡すもの:

| 渡す | 中身 | 出所 |
| --- | --- | --- |
| 成果物 | レンダ済みページのスクショ＋HTML/CSS（読み手が受け取る形・単独） | 工場 |
| ルーブリック | `rubric.template.md` を埋めたもの。Style 観点の値は `origin-brand` の様式、禁則は web の媒体適合 | 工場（倉庫から取った値を上乗せ） |
| 前提条件の宣言 | 「`audit_html.py` exit 0・`measure_lp.mjs` 失格ゼロを通過」 | 工場 |
| 方式の希望（任意） | 例: B は `web-design-guidelines`＋`baseline-ui`、C は AD ペルソナ | 工場（最終判断は本 skill） |

渡さないもの: 設計意図・仕様書・過去版・pptx 固有の後処理・修正の実装。

流れ: 本 skill が方式を選ぶ → 評価者を立てる（**先に `provenance.json` を書く**）→
`run_round.sh <eval_dir> --thresholds <T>` → 不合格（exit 1）なら A/B 指摘を工場へ返す →
**工場が直す** → 工場が再度呼ぶ（同じ `eval_dir`）→ … → 最終報告 → 工場が `origin-goal` へ渡す → 人間ゲート。

`eval_dir` は成果物と同じ作業ディレクトリ（例: `process/eval/`）に置く。中身は
`thresholds.json` / `state.json`（`run_round.sh` が作る）/ `round_1/ round_2/ …`。

**exit 2 が返ったら判定はまだ行われていない。** 入力の不備（`provenance.json` が無い・申告と実ファイルの
食い違い・`fresh: false`・読者に余分な入力・キャッシュ未クリア・`round` 不一致・`severity` 語彙外）を
直してから再実行する。**「評価者が1人しか起動できなかったので1人で判定した」は起きない** —
申告と実ファイルの数が合わなければ止まる。

## B. 単発で呼ぶ（例: 「この PR の変更をレビューして合格まで回して」）

1. 人間が対象と合格の感覚を一言で渡す（「些事の指摘が無くなるまで」「プロが納品できる水準」）。
2. 本 skill が方式を選ぶ — コード変更なら **A（既存テスト or 新設）→ B（`code-review`）**。C は主観品質が
   問われるときだけ。
3. **審査員（C）は必ず別 subagent**。単発では修正も自分がやるため、作った本人が採点する形になりやすく、
   ここが独立性の抜け道になる。B のレビュー skill も自分の会話の中で「自分で読む」のではなく skill として起動する。
   立てた評価者は `provenance.json` に申告する（`fresh: true`・渡した `inputs`）。申告と実ファイルが
   合わなければ判定が `exit 2` で止まるので、抜け道は構造的に閉じている。
4. 既定値で回す（回数3・閾値既定）。`run_round.sh` を使い、手で組まない。上限到達（exit 4）・
   未収束（exit 3）は終了して最終報告。
5. 最終報告を人間に見せる。**pass でも人間が最後に見る。**

### 単発での最小手順（そのまま実行できる形）

```bash
E=process/eval; mkdir -p $E/round_1
cp <ルーブリックから写した閾値> $E/thresholds.json
# 審査員・レビューを別 subagent として起動し、出力を $E/round_1/ に置く
#   judges/j1.json  readers/r1.json  review.json  tests.json のうち使うもの
cat > $E/round_1/provenance.json <<'JSON'
{"round":1,
 "evaluators":[{"role":"judge","id":"j1","fresh":true,"inputs":["artifact","rubric"]}],
 "gates_passed":["pytest 全緑"],
 "cache_cleared":true}
JSON
bash <skill>/scripts/run_round.sh $E --thresholds $E/thresholds.json --max-rounds 3
```

## 受付（`origin-design-check-routing`）との関係

受付は「どの検査 skill を並べるか」を決める。本 skill は「選んだ検査を合格まで回す」。
接続の向き（受付が本 skill を1レビューアとして呼ぶか、本 skill が受付を方式 B の実体として使うか）は
spec の Deferred。それまでは相互参照のみ。
