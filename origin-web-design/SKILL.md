---
name: origin-web-design
description: >-
  web の成果物（HTML 画面・LP・コンポーネント・ダッシュボード）を作る工場。媒体規約（media/）と
  倉庫から受け取った意匠を合成し、生成規律（HTML MUST）を守って出力し、必要なら実測で点検する。
  必ず使うこと: 「UI を作って」「HTML を作って」「LP を作って」「ダッシュボードを作って」
  「web の見た目を整えて」「ブランドに沿ったサイトを作って」と言われた時。
  使わない場面: スライド・PPTX・提案資料の生成（origin-pptx が持つ。本 skill は web 専用で、
  pptx の媒体規約を持たない）／ブランド意匠の選択・値そのもの（origin-brand が持つ）／
  画像アセットの生成（origin-image-gen）／検査の振り分け（origin-design-check-routing）／
  合格まで直して再評価するループ（origin-output-eval）。
---

# origin-web-design — web 成果物の工場

このSkillは**デザインルールそのものを持たない**。媒体（`media/`）と、倉庫
（`origin-brand`）から受け取った意匠を合成し、コンテキストを組み立て、成果物を生成し、
Validatorで点検する。

**web 専用である。** `media/` は `html` と `dashboard` だけを持つ。pptx の媒体規約は持たない —
過去に「web 向けの依頼で pptx 用の実装が動き、媒体規約が誤適用された」実失敗があり、
その土台（呼ばれない pptx 媒体ルール）を退役させたのが今の形。
スライド・資料は `origin-pptx` が持つ。

## 重要ルール（常に守ること）

- **Skillはデザインを持たない**。配色・余白・フォント等の実体は倉庫 `origin-brand` の
  DESIGN.md / tokens にある。**本 skill は倉庫に問い合わせるだけで、意匠の選択規則を自前で持たない**
  （同じ規則が2実装になった時点でズレが始まる。実測済み）。
- **ライセンスゲートは倉庫が持つ**（`origin-brand`）。門を工場側に置くと、工場は門を通らずに
  第三者資産の実値を取得できてしまう。**本 skill はゲートの結果（可否）を受け取るだけで、
  自分で判定しない。**
- **第三者DESIGN.mdの参照元の限定も倉庫の所掌。** 本 skill が自律的に外部サイトを
  探索・スクレイピングしない。
- **このSkillはスタイリング層に徹する**。情報設計・コンテンツ構造・メッセージ設計といった「設計判断」は
  スコープ外（DESIGN.md フォーマット自体がスタイリング記述であり、上流の Google design.md /
  awesome-design-md-jp も同様）。重要なのは、選んだ配色・タイポ・余白・a11y制約が**成果物に確実に反映される**こと。
- **生成時のMUST（authoring）は毎回・無料で守る**が、**validate（実測点検）の実行はコストがかかるので自動発火しない**。
  サイト（HTML/web）の場合は生成後に**人間に「validateするか」を必ず聞いてから**実行する（後述「Validateは人間に聞く」）。
- **適用範囲（Tier）**: HTML / dashboard（web）＝**完全対応**（生成＋実測validate可）。PPTX・その他＝**参考Tier**
  （DESIGN.md/tokens を参考に生成方針は出すが、機械validateはしない。実生成は `pptx` Skill 等に委譲してよい）。
- 出力には使用したPlugin・主要制約・validate実施の有無/結果を必ず記録する。

## 参照ファイルと役割

| パス                                         | 役割                                        | いつ読むか                      |
| -------------------------------------------- | ------------------------------------------- | ------------------------------- |
| `references/architecture.md`                 | Runtime/Media/Plugin/Validator の境界       | 全体像を確認する時              |
| `references/media-spec.md`                   | `media.yaml` の正式スキーマ                 | Media を追加・検証する時        |
| `references/context-composer.md`             | コンフリクト解決アルゴリズム                | 手順8（コンテキスト合成）で必ず |
| `references/validator.md`                    | Validator 仕様（[静的]/[実測]/[自己点検]）  | 手順10（検証）で必ず            |
| `references/site-validation.md`              | playwright実測の手順書                      | サイト実測validateを実施する時  |
| `media/<media>/media.yaml` + `rules.md`      | 媒体固有の制約・推奨                        | 手順6                           |
| `origin-brand/plugins` の各箱                 | 意匠の実体（倉庫が解決して渡す）            | 手順5・7                        |
| `validators/common/*`                        | 媒体横断の点検項目                          | 手順10                          |

## 処理手順

```text
1. User Taskを読み、成果物の媒体（media）を判定する。→ Media Resolver（下記）
   **web 以外（スライド等）と判定したら、ここで origin-pptx へ渡して止まる。**
2. **意匠の与件があるかを分岐する。**
   - **(a) ブランド与件あり**（「自社ブランドで」「デジタル庁風に」等 / リポジトリローカルの
     DESIGN.md がある）→ `origin-brand` に文脈と媒体を渡し、解決済みの意匠を受け取る。
   - **(b) 与件なし・発案から**（「かっこいいLPを作って」等）→ 意匠を**この工場が発案してよい**が、
     発案したことを報告に明記する。倉庫の箱を勝手に選んで「ブランド準拠」と名乗らない。
3. 倉庫が返すもの: 箱の id（継承元含む）・値・**出所**・**ライセンス可否**。
   可否が false の箱の実値は使わない（原則プローズのみ参照可）。
4. Project DESIGN.md（リポジトリローカル）があれば最優先で読む。
5. 倉庫が解決した意匠（継承元があれば親も。1階層まで）を読む。
6. Media Rule（media/<media>/rules.md, media.yaml）を読む。
7. 必要なAsset / Template / Exampleだけを読む（全部は読まない）。
8. 生成用コンテキストを組み立てる。競合は references/context-composer.md の優先度順で解決する。
9. 成果物を生成する。HTMLを出力する場合は、後述の「HTML出力前 必須チェックリスト（MUST）」を
   雛形どおり満たしてから出力する（＝生成規律。ここは毎回・無料で守る）。
10. **サイト（HTML/web）なら、ユーザーに「validateするか」を聞く**（後述「Validateは人間に聞く」）。
    - 実施する場合のみ references/validator.md に従い点検する（静的 → 必要ならレンダリング実測）。
    - PPTX・その他（参考Tier）は機械validateしない。生成方針とチェックリストの提示に留める。
11. validate実施時に不合格があれば修正する。修正は最大2回まで。2回で不合格なら打ち切り、未解決項目を明示して報告する。
12. 出力時に次を必ず記録する — 使用した箱と継承元／主要制約／validate実施の有無と結果／
    **通さなかった工程とその理由**（例: 「工程2は与件が無く発案した」「実測validateは
    ユーザーが skip を選んだ」）。**省略を黙って行わない**のが本 skill の報告契約。
```

## HTML出力前 必須チェックリスト（MUST — 毎回・例外なし）

HTMLを生成するときは、以下を**必ず**満たす。これらは `media/html` の `type: constraint` に対応し、
省略・後回しにしない（過去に lang / viewport / landmark / SVGラベル / focus の欠落が頻発したため必須化）。

- [ ] `<html lang="ja">`（日本語UIなら必ず lang を付ける）
- [ ] `<meta name="viewport" content="width=device-width, initial-scale=1">`
- [ ] ページ本体を `<main>` で包み、`<header>` / `<nav>` / `<footer>` を適切に使う
- [ ] 意味を持つ SVG/図に `role="img"` + `aria-label`（装飾なら `aria-hidden="true"`）
- [ ] キーボードフォーカスを消さず `:focus-visible` を明示
- [ ] 本文コントラスト 4.5:1 以上（リンク・薄いグレー文字・淡色背景上テキストに特に注意）

コピペ可能な骨格:

```html
<!doctype html>
<html lang="ja">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>…</title>
    <style>
      :focus-visible {
        outline: 3px solid var(--color-primary, #0017c1);
        outline-offset: 2px;
      }
    </style>
  </head>
  <body>
    <header>…</header>
    <nav aria-label="主要ナビゲーション">…</nav>
    <main>
      <!-- 意味を持つ図は必ずラベルを付ける -->
      <svg role="img" aria-label="2020–2025年の総人口推移（減少傾向）">…</svg>
    </main>
    <footer>…</footer>
  </body>
</html>
```

このチェックリストは**生成時に守る規律**であって、実測validate（後述）とは別。validateを実施しない場合でも
MUSTは常に満たす。

## Validateは人間に聞く（サイトのみ・コスト配慮）

サイト（HTML/web）を生成したら、実測validateを**勝手に走らせず、まずユーザーに聞く**。理由: レンダリング
実測（ブラウザ起動）はコストがかかり、毎回自動でやると煩わしいため。次の3択を提示する:

1. **静的チェック（軽量・ほぼ無コスト）** — `python3 scripts/audit_html.py <file>`（＋リンクCSSも追跡）。
   lang / viewport / landmark / SVGラベル / focus-visible / リテラル色のコントラストを静的に判定。
2. **レンダリング実測（コストあり）** — `playwright-cli` Skill でブラウザ起動し、
   実際のcomputedスタイルでコントラスト・フォーカス可視・ランドマーク・（テーマ切替があれば）各テーマを実測。
   ブラウザ操作は `browser-automation-policy` Skill に従う。詳細手順は `references/site-validation.md`。
3. **skip** — 今回はvalidateしない（生成時MUSTは満たしている前提）。

デフォルトの薦めは「まず1の静的、必要なら2のレンダリング実測」。ユーザーが明示的に「毎回自動でvalidateして」
と言った場合に限り、都度確認を省いてよい。

## 適用範囲とTier

| 媒体                    | 対応         | 生成 | validate                   |
| ----------------------- | ------------ | ---- | -------------------------- |
| html / dashboard（web） | **完全対応** | ○    | ○（人間に聞いて静的/実測） |
| pptx / スライド         | **対象外**   | ×    | ×                          |

**pptx の「参考Tier」は廃止した**（2026-09-09）。方針だけ出せる状態が、
「web 向けの依頼で pptx 用の実装が動く」経路を残していた。スライドは `origin-pptx` が持ち、
本 skill は媒体判定で web 以外と分かった時点で渡して止まる。

## Media Resolver（手順1の判断基準）

| 入力の手がかり                                  | 選ぶ media         |
| ----------------------------------------------- | ------------------ |
| Webページ、画面、LP、コンポーネント、React/HTML | `html`                       |
| ダッシュボード、KPI、可視化、指標画面           | `dashboard`                  |
| スライド、プレゼン、提案資料、PPTX、パワポ      | **本 skill の対象外** → `origin-pptx` へ渡す |
| 明示がなく判断が割れる                          | ユーザーに確認する           |

`media/*/media.yaml` をスキャンして利用可能な媒体一覧を得る（ディレクトリ追加だけで増える）。
現在あるのは `html` と `dashboard` の2つだけ。

## 意匠の選択は倉庫が行う（本 skill は持たない）

Plugin Resolver・`extends` の1階層解決・ライセンスゲートは **`origin-brand`** が持つ。
規則は `origin-brand/references/resolver.md`。

**ここに規則を再掲しない。** 再掲した時点で2実装になり、`extends` の解釈がズレる余地ができる
（同一ブランドの値が2箇所で別物だった実測がこの決定の根拠）。

## 起動条件

UIを作る / HTMLを作る / LPを作る / ダッシュボードを作る / web の見た目を整える /
ブランドに沿ったサイトを作る — これらの依頼で起動する。

**起動しない**: スライド・PPTX・提案資料（`origin-pptx`）。意匠の値そのものの問い合わせ
（`origin-brand`）。

## 失敗時のフォールバック

- 倉庫が候補を1つも返さない → 既定方針（`references/architecture.md`）で生成し、
  **意匠なしで生成した旨を明示する**（工程2(b) の発案として報告する）。
- 倉庫の問い合わせ自体が失敗 → 参照なしで続行し、参照不可だった旨を報告する（生成全体は止めない）。
- 検証ループ2回で不合格 → 手順11に従い打ち切り、未解決項目を添えて成果物とともに報告する。
