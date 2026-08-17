---
name: origin-quarto-routing
description: Quarto 作業の入口。`.qmd`・`_quarto.yml`・Quarto プロジェクト・RevealJS スライド・Quarto からの PowerPoint 出力に関わる依頼で必ず最初に使う。「Quarto で作って」「qmd を直して」「スライドを Quarto で」「Quarto をレンダリングして」「callouts / cross-reference / citation の書き方」「R Markdown（bookdown / blogdown / xaringan / distill / Jupyter）から Quarto へ移行」で発火する。依頼の形を判定して `origin-quarto`（組織テンプレートからの成果物作成・レンダリング）と `quarto-authoring`（Quarto の機能・構文リファレンスと他フォーマットからの移行）へ振り分け、使わなかった側を理由つきで報告する。Quarto 系 skill が2つあってどちらか迷う場面では、個別 skill を直接選ばずまずこれを使う。使わない場面: Quarto を伴わない Markdown/PDF/PPTX 作業（`origin-pptx` 等）、Quarto と無関係な R コード作業（`origin-r-coding-style`）。
---

# Quarto Routing

Quarto 作業の**ルーター**。Quarto の知識もテンプレートも持たない。持つのは
**依頼の形の判定・振り分け表・実行順序・使わなかった側の申告**だけ。

## なぜルーターが要るか

Quarto に関わる skill は 2 つあり、**実体は重複ではなく補完**（2026-08-17 実測）:

|          | `origin-quarto`                                                                                                              | `quarto-authoring`                                                                                                                                   |
| -------- | ---------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| 由来     | 自前                                                                                                                         | 外部（MIT・上流更新で上書きされる）                                                                                                                  |
| SKILL.md | 35 行                                                                                                                        | 317 行                                                                                                                                               |
| 同梱     | 組織テンプレート一式（qmd テンプレート・拡張: SCSS・CSL・PowerPoint reference document・RevealJS plugin・marp テンプレート） | Quarto 機能ごとのリファレンス 21 ファイル（callouts / citations / cross-references / tables / layout / diagrams / shortcodes / engines / 移行 5 種） |
| 役割     | 組織テンプレートからの成果物作成・レンダリング・検証                                                                         | Quarto の機能と構文のリファレンス、他フォーマットからの移行                                                                                          |

競合しているのは**発火トリガだけ**で、どちらも「Quarto / `.qmd` / `_quarto.yml`」で
発火する。文脈発火（description の文字列マッチ）は**発火しなかった側が沈黙する**ため、
組織テンプレートを使うべき成果物作成が `quarto-authoring` に流れても誰にも見えない。

このルーターは逆の契約を結ぶ: **振り分けなかった側を、理由つきで必ず報告する。**

## 手順

### 1. 依頼の形を判定して振り分ける

| 依頼の形                                                                                                                                                    | 担当                                                                                 |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| 新規の成果物を作る（文書・スライド・PowerPoint 出力）                                                                                                       | `origin-quarto`（組織テンプレートの複製が先）                                        |
| 既存 Quarto プロジェクトの体裁・出力構成を変える                                                                                                            | `origin-quarto`（既存設定を先に確認し、既存デザインを勝手に置換しない）              |
| Quarto の機能・構文の書き方（callouts・cross-references・citations・tables・layout・diagrams・shortcodes・conditional content・engines・YAML front matter） | `quarto-authoring`                                                                   |
| 他フォーマットからの移行（R Markdown・bookdown・blogdown・xaringan・distill・Jupyter）                                                                      | `quarto-authoring`                                                                   |
| レンダリング失敗の切り分け                                                                                                                                  | 構文・YAML はまず `quarto-authoring`、拡張やテンプレートのパス解決は `origin-quarto` |

迷ったら聞かずに「成果物作成として扱う」と宣言して `origin-quarto` から始める
（組織テンプレートを外して作ってしまう方が、後から直す費用が高いため）。

### 2. 実行順序

成果物作成では **`origin-quarto` が先**。テンプレートを destination へ置いてから中身を書く。
途中で機能・構文の疑問が出たら `quarto-authoring` の該当 `references/` を読み、
`origin-quarto` の Validation（拡張・テーマ・CSL・reference document のパス解決、
レンダリング、スライドの目視確認）へ戻って閉じる。

### 3. 使わなかった側を申告する

報告に 1 行入れる。例: 「`quarto-authoring` は参照しなかった（既存テンプレートの
YAML 変更のみで新しい Quarto 機能を使っていない）」。
振り分けを誤ったときに、どこで誤ったかが読めるようにするためで、
判断の正しさを主張するためではない。

## 既知の限界

- **このルーターを経由しない直接発火は防げない。** 2 つの skill の description は
  どちらも Quarto で発火するままで（`origin-quarto` は `any Quarto document` と広い）、
  退役も description の書き換えもしていない。境界はこの skill の散文にしかなく、
  `skill_lint.sh` は 1 skill 単位の整合しか見ないので**黙って破れる**。
- 検出をどこに置くか（`skill_lint.sh` に発火競合の warn を足すか、トリアージの
  定常観点にするか）は biz_ops の
  `ISSUE-20260817-improve-skill-duplicate-detection` の決めること 3 で未決。
  lint に載せる案は `skill_lint.sh` が常時 exit 1 の状態
  （`ISSUE-20260817-improve-loop-skill-lint-red-on-main`）が片付くまで合否を確認できない。
- `quarto-authoring` は外部由来のため、**そちらに境界を書いても上流更新で上書きされる**。
  境界を持てるのは自前側（この skill と `origin-quarto`）だけ。
