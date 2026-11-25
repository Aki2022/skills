# 組み方の計画とカンプ

台本 → 組み方の計画 → カンプ → HTML の流れ（SPEC-deliverable-production-layers の Requirements 14・15・18・19・21〜23）の、build 側の詳細。手順の並びは SKILL.md の「処理手順」。

## 置き場所

層の置き場所の規則（own-design-author の `references/layer-paths.json`）の `build` が 1 か所で決める。コードに文字列で書かない。

| もの | 置き場所 | git |
| --- | --- | --- |
| 組み方の計画 | `ux/build/<成果物 id>/plan.yaml` | 入れる |
| カンプの依頼と画像 | `ux/build/<成果物 id>/comp/`（`fields.json`・`request.json`・`comp.png`） | 入れない |

カンプを git に入れないための ignore の規則は `templates/build/.gitignore`。最初の計画を書くときに `ux/build/.gitignore` に写す。`ux/build/` は design の入力ではない（L0 のマニフェストは計画とカンプをハッシュしない）。計画を直してもカンプを描いても、マニフェストは古くならない。

## 計画の形

形は `references/schema/plan.v1.schema.json`。

```yaml
schema_version: 1
deliverable: lp-sample        # 置き場所のディレクトリ名と同じ
script: lp-sample             # ux/infoarch/scripts/<id>.yaml
blocks:                       # 台本のまとまりをちょうど 1 回ずつ
  blk-0001: {pattern: header-bar}
  blk-0002:
    pattern: split-2
    details: 絵は右の列から上の帯へ少しはみ出す。見出しの下に細い線
  blk-0003: {pattern: stats-row, narrow: stack-1}
comp: {reviewed: false}       # カンプの門で見て、計画を直したら true
```

- `pattern` と `narrow` は design の組み方の型の名前（own-design-author の `catalog/layout-patterns.json` を repo の `design.yaml` の `layout_patterns` で差し引きしたもの。L0 が解決済みマニフェストの `layout_patterns` に入れる）。語彙に無い名前は失格。
- `narrow` は、型の既定の崩し方と違うときだけ書く。書かなければ型の既定で崩し、検査の報告がどの型になるかを出す。
- `details` は、言葉で決めきれずカンプを見て決めた細部（絵と文字の重なり・はみ出し・接地・区切り方）。カンプの画像は残らないので、ここに写さなかった細部は失われる。
- 台本の文言は書かない。
- 変種は `variant_of`（基本の計画の成果物 id）と `overrides`（`blocks/<まとまりの id>`）を宣言する。宣言していない違いは警告（交絡あり）、宣言したのに違わない上書きも警告。`infoarch-author diff` は両方の成果物に計画が 1 つずつあれば計画も比べ、`differs_in` を `script`・`plan`・`both` で出す。

## 計画の検査（`scripts/plan_check.py`）

`python3 <このskill>/scripts/plan_check.py --repo <repo> --deliverable <id>`。終了コードは 0 合格・1 警告・2 失格（確かめる対象が無いときも 2）。報告は 5 区分（実行・スキップ・欠落・未カバー・screened）と指摘（fail・warn・info）。

| code | 段 | 意味 |
| --- | --- | --- |
| PLAN_MISSING | fail | 計画が無い |
| PLAN_UNREADABLE | fail | YAML として読めない（重複する鍵を含む） |
| PLAN_SCHEMA | fail | 形が schema に合わない |
| PLAN_DELIVERABLE_MISMATCH | fail | `deliverable` がディレクトリ名と違う |
| PLAN_SCRIPT_MISSING | fail | 台本が無いか読めない |
| PLAN_NOTHING_CHECKED | fail | 台本のまとまりが 0 個 |
| PLAN_MANIFEST_UNAVAILABLE | fail | 解決済みマニフェストが無いか古い（L0 を先に流す） |
| PLAN_BLOCK_MISSING / PLAN_BLOCK_UNKNOWN | fail | 台本のまとまりが計画に無い／計画のまとまりが台本に無い |
| PLAN_PATTERN_UNKNOWN / PLAN_NARROW_UNKNOWN | fail | 型の名前が語彙に無い |
| PLAN_OVERRIDES_WITHOUT_BASE / PLAN_VARIANT_BASE_MISSING / PLAN_OVERRIDE_UNKNOWN | fail | 変種の宣言が壊れている |
| PLAN_CHECK_FAILED | fail | 想定外の誤り（警告に見せないため 2） |
| PLAN_PATTERN_UNSUITED | warn | 型がまとまりの働きに向く型として書かれていない |
| PLAN_COMP_NOT_IGNORED | warn | カンプのディレクトリが git に入る |
| PLAN_COMP_IGNORE_NOT_CHECKED | warn | git の作業ツリーでないので ignore を確かめていない（未カバーにも出す） |
| PLAN_UNDECLARED_DIFFERENCE / PLAN_OVERRIDE_UNCHANGED | warn | 変種の比べ方（交絡あり／変わらない上書き） |
| PLAN_NARROW_DEFAULT / PLAN_COMP_NOT_REVIEWED | info | 既定で崩すまとまり／カンプの門をまだ通っていない |

台本そのものの検査（引き先・提案の印など）は `infoarch-author check` が持ち、ここではスキップと報告する。

## カンプの依頼（`scripts/comp_request.py`）

`python3 <このskill>/scripts/comp_request.py --repo <repo> --deliverable <id>`。計画の検査が失格なら何も書かずに 2（COMP_PLAN_FAILED）。カンプのディレクトリが git に入るなら 2（COMP_DIR_NOT_IGNORED）。マニフェストにカンプの大きさが無ければ 2（COMP_SIZE_MISSING）。想定外の誤りも 2（COMP_REQUEST_FAILED）。git の外では書いて 1（COMP_IGNORE_NOT_CHECKED）。

書くのは own-image-gen の器の 10 フィールド（`fields.json`）と、出力先・実行するコマンドを持つ `request.json`。画像は生成しない（従量課金。オーナーの門の前に人が流す）。

- カンプはページ全体の広い画面の 1 枚。まとまりを台本の `order` の順に、働き・型（列・帯・絵の置き方・区切り）・計画の細部で描く。
- 台本の文案は仮置きの文として渡す。カンプは見せてよいが、カンプの文言は使わない。カンプが新しく描いた絵も使わない。
- 色と書体は解決済みマニフェストのトークン。大きさはマニフェストの `comp`（skill の既定 `own-design-author/references/comp-sizes.json`、repo の `profiles/<media>.json` の `comp` だけが上書きできる）。
- 狭い画面は描かない。狭い画面の形は計画が持つ。

## カンプの門のあと

オーナーがカンプを見て「進めて」と言ったら（台帳には残さない）、カンプが足した細部をすべて `details` に写し、型を変えたまとまりは `pattern`・`narrow` を直し、`comp.reviewed: true` にして `plan_check.py` を流し直す。HTML は計画に沿って組む。狭い画面の崩し方は、列の規則と同じ強さ（同じ詳細度のセレクタ、同じ層）で書く。広い画面の列の指定が狭い画面の崩し方に勝つ書き方（例：列の指定だけ詳細度が高い）にしない。
