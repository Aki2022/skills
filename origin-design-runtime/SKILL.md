---
name: origin-design-runtime
description: >
  デザイン生成のRuntime/Orchestrator。UI・HTML・PPTX/スライド・ダッシュボード・レポート・
  デザイン仕様書を作る時に必ず使うこと。DESIGN.md・媒体ルール（Media）・デザインプラグイン
  （tokens/templates/assets）・品質Validatorを機械的に選択・合成・点検して成果物を生成する。
  「UI作って」「HTML作って」「スライド/PPTX作って」「ダッシュボード作って」「デザインを整えて」
  「ブランドに沿って資料を作って」「Digital Agency風に」「DESIGN.mdを使って」「アセットを切り替えて」
  などの依頼で起動する。このSkill自身はデザインを持たず、plugins/ と media/ を参照する。
---

# origin-design-runtime — AI Design Runtime + Plugin Ecosystem

このSkillは**デザインルールそのものを持たない**。媒体（`media/`）とデザインプラグイン（`plugins/`）を
選択・合成し、DESIGN.md + Media Rule + アセットからコンテキストを組み立て、成果物を生成し、Validatorで
点検するRuntime（Orchestrator）である。

## 重要ルール（常に守ること）

- **Skillはデザインを持たない**。配色・余白・フォント等の実体は `plugins/` の DESIGN.md / tokens に置く。
- **ライセンスゲート**: 第三者アセット（デジタル庁等）を成果物へ反映する前に、その plugin の `references.md`
  にライセンス確認と人間承認の記録があることを確認する。未承認なら資産値（色トークン実値・フォント・ロゴ）を
  出力に使わない。
- **第三者DESIGN.mdの参照元は `awesome-design-md-jp` に限定**。個別サイトはユーザー指示制。Skillが自律的に
  外部サイトを探索・スクレイピングしない。
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
| `references/plugin-spec.md`                  | `plugin.yaml` / `media.yaml` の正式スキーマ | Plugin/Media を追加・検証する時 |
| `references/context-composer.md`             | コンフリクト解決アルゴリズム                | 手順8（コンテキスト合成）で必ず |
| `references/validator.md`                    | Validator 仕様（[静的]/[実測]/[自己点検]）  | 手順10（検証）で必ず            |
| `references/site-validation.md`              | playwright実測の手順書                      | サイト実測validateを実施する時  |
| `media/<media>/media.yaml` + `rules.md`      | 媒体固有の制約・推奨                        | 手順6                           |
| `plugins/<plugin>/DESIGN.md` + `plugin.yaml` | デザイン方針・資産の入口                    | 手順5・7                        |
| `validators/common/*`                        | 媒体横断の点検項目                          | 手順10                          |

## 処理手順

```text
1. User Taskを読み、成果物の媒体（media）を判定する。→ Media Resolver（下記）
2. 明示指定されたDesign Pluginがあれば採用する。
3. 明示指定がなければ、目的・業界・媒体から候補Pluginを選ぶ。→ Plugin Resolver（下記）
4. Project DESIGN.md（リポジトリローカル）があれば最優先で読む。
5. 選択したPluginのDESIGN.md（継承元があれば親も。1階層まで）を読む。
6. Media Rule（media/<media>/rules.md, media.yaml）を読む。
7. 必要なAsset / Template / Exampleだけを読む（全部は読まない）。
8. 生成用コンテキストを組み立てる。競合は references/context-composer.md の優先度順で解決する。
9. 成果物を生成する。HTMLを出力する場合は、後述の「HTML出力前 必須チェックリスト（MUST）」を
   雛形どおり満たしてから出力する（＝生成規律。ここは毎回・無料で守る）。
10. **サイト（HTML/web）なら、ユーザーに「validateするか」を聞く**（後述「Validateは人間に聞く」）。
    - 実施する場合のみ references/validator.md に従い点検する（静的 → 必要ならレンダリング実測）。
    - PPTX・その他（参考Tier）は機械validateしない。生成方針とチェックリストの提示に留める。
11. validate実施時に不合格があれば修正する。修正は最大2回まで。2回で不合格なら打ち切り、未解決項目を明示して報告する。
12. 出力時に、使用したPlugin・主要制約・validate実施の有無と結果（未解決項目があれば）を記録する。
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

| 媒体                    | Tier         | 生成                                                             | validate                   |
| ----------------------- | ------------ | ---------------------------------------------------------------- | -------------------------- |
| html / dashboard（web） | **完全対応** | ○                                                                | ○（人間に聞いて静的/実測） |
| pptx / その他           | **参考Tier** | 方針・構成・チェックリストのみ（実生成は `pptx` Skill 等に委譲） | ×（機械validateしない）    |

参考Tierでは DESIGN.md / tokens を「参考情報」として使い、配色・タイポ・余白・トーンの方針を提示する。
DESIGN.md はサイト以外（資料・スライド）でもスタイリング指針として有効なので、参考用途として残す。

## Media Resolver（手順1の判断基準）

| 入力の手がかり                                  | 選ぶ media         |
| ----------------------------------------------- | ------------------ |
| Webページ、画面、LP、コンポーネント、React/HTML | `html`             |
| スライド、プレゼン、提案資料、PPTX、パワポ      | `pptx`             |
| ダッシュボード、KPI、可視化、指標画面           | `dashboard`        |
| 明示がなく判断が割れる                          | ユーザーに確認する |

`media/*/media.yaml` をスキャンして利用可能な媒体一覧を得る（ディレクトリ追加だけで増える）。

## Plugin Resolver（手順3の判断基準）

1. `plugins/*/plugin.yaml` をスキャンし、`supported_media` に対象媒体を含む Plugin を候補にする。
2. `priority.contexts` のうちタスク文脈（public-sector / dashboard / corporate-presentation 等）に
   一致するスコアが高い Plugin を選ぶ。文脈一致がなければ `priority.default` を使う。
3. **同点のタイブレーク**: (a) プロジェクトローカルに近い Plugin、(b) それも同点なら plugin id の
   辞書順で先。決めきれない・候補が拮抗する場合はユーザーに確認する。
4. `extends` があれば親 Plugin を1階層だけ辿って合成対象に含める（多段継承は不可）。

## 起動条件

UIを作る / HTMLを作る / PPTX・スライドを作る / ダッシュボードを作る / デザインを整える /
ブランドに沿って資料を作る / 「Digital Agency風」等のデザイン指定 / DESIGN.mdを使う /
アセットを切り替える — これらの依頼で起動する。

## 失敗時のフォールバック

- 候補Pluginが1つも無い → Runtimeデフォルト（`references/architecture.md` のデフォルト方針）で生成し、
  Pluginなしで生成した旨を明示する。
- 外部DESIGN.md取得に失敗 → 参照なしで続行し、参照不可だった旨を報告する（生成全体は止めない）。
- 検証ループ2回で不合格 → 手順11に従い打ち切り、未解決項目を添えて成果物とともに報告する。
