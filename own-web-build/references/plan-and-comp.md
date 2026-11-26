# 組み方の計画とカンプ

台本 → 組み方の計画 → カンプ → HTML の流れ（SPEC-deliverable-production-layers の Requirements 14・15・18・19・21〜23）の、build 側の詳細。手順の並びは SKILL.md の「処理手順」。

## 置き場所

層の置き場所の規則（own-design-author の `references/layer-paths.json`）の `build` が 1 か所で決める。コードに文字列で書かない。

| もの | 置き場所 | git |
| --- | --- | --- |
| 組み方の計画 | `ux/build/<成果物 id>/plan.yaml` | 入れる |
| 制作の指示 | `ux/build/<成果物 id>/brief.yaml` | 入れる |
| カンプの依頼と画像 | `ux/build/<成果物 id>/comp/`（`fields.json`・`request.json`・`comp.png`。分けたときは `fields-<k>.json`・`comp-<k>.png`） | 入れない |

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
  blk-0003:
    pattern: stats-row
    narrow: stack-1
    emphasis:                 # 大きく描く文字（対象の id と、その中の文字）
      - {target: value-customer-count, text: '100'}
comp: {reviewed: false}       # カンプの門で見て、計画を直したら true
```

- `pattern` と `narrow` は design の組み方の型の名前（own-design-author の `catalog/layout-patterns.json` を repo の `design.yaml` の `layout_patterns` で差し引きしたもの。L0 が解決済みマニフェストの `layout_patterns` に入れる）。語彙に無い名前は失格。
- `narrow` は、型の既定の崩し方と違うときだけ書く。書かなければ型の既定で崩し、検査の報告がどの型になるかを出す。
- `details` は、言葉で決めきれずカンプを見て決めた細部（絵と文字の重なり・はみ出し・接地・区切り方）。カンプの画像は残らないので、ここに写さなかった細部は失われる。
- `emphasis` は、まとまりの中で残りより大きく描く文字。`target` はこのまとまりの文案の id か、このまとまりの文案が示す辞書の値の id、`text` はその中の文字（全部ではなく、残りがあること）。`details` の文は描いた画面と比べないが、`emphasis` は `plan_check.py --html` が測る。カンプで値段や数字を大きく描いたなら、ここに書く。
- 台本の文言は書かない（`emphasis` の `text` は、対象の文字の一部を指すだけ）。
- 変種は `variant_of`（基本の計画の成果物 id）と `overrides`（`blocks/<まとまりの id>`）を宣言する。宣言していない違いは警告（交絡あり）、宣言したのに違わない上書きも警告。`infoarch-author diff` は両方の成果物に計画が 1 つずつあれば計画も比べ、`differs_in` を `script`・`plan`・`both` で出す。

## 計画の検査（`scripts/plan_check.py`）

`python3 <このskill>/scripts/plan_check.py --repo <repo> --deliverable <id> [--html <描いたページ>]`。`--html` を渡すと、計画の `emphasis` を 2 つの幅（375・1280）で描いて測る（`scripts/measure_emphasis.mjs`。Playwright は design の固定環境で版を決める）。文字は空白を除き幅をそろえて読み、対象のまとまりの中の印だけを見る。終了コードは 0 合格・1 警告・2 失格（確かめる対象が無いときも 2）。報告は 5 区分（実行・スキップ・欠落・未カバー・screened）と指摘（fail・warn・info）。

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
| PLAN_EMPHASIS_TARGET_UNKNOWN | fail | 強調の対象が、このまとまりの文案でも、その文案が示す辞書の値でもない |
| PLAN_EMPHASIS_TEXT_ABSENT | fail | 強調の文字が対象に無いか、対象の全部（何より大きくするかが無い） |
| PLAN_EMPHASIS_TARGET_NOT_DRAWN | fail | `--html`：対象がそのまとまりの中に描かれていない |
| PLAN_EMPHASIS_NOT_RENDERED | fail | `--html`：強調の文字が、対象の残りの文字より大きく描かれていない（文字の大きさで比べる） |
| PLAN_EMPHASIS_NOT_MEASURED | fail | `--html`：描けなかった（ブラウザが無いなど）。合格にしない |
| PLAN_CHECK_FAILED | fail | 想定外の誤り（警告に見せないため 2） |
| PLAN_PATTERN_UNSUITED | warn | 型がまとまりの働きに向く型として書かれていない |
| PLAN_COMP_NOT_IGNORED | warn | カンプのディレクトリが git に入る |
| PLAN_COMP_IGNORE_NOT_CHECKED | warn | git の作業ツリーでないので ignore を確かめていない（未カバーにも出す） |
| PLAN_UNDECLARED_DIFFERENCE / PLAN_OVERRIDE_UNCHANGED | warn | 変種の比べ方（交絡あり／変わらない上書き） |
| PLAN_EMPHASIS_NONE | info | 計画に強調が無い（描いた画面で確かめる文字が無い） |
| PLAN_NARROW_DEFAULT / PLAN_COMP_NOT_REVIEWED | info | 既定で崩すまとまり／カンプの門をまだ通っていない |

台本そのものの検査（引き先・提案の印など）は `infoarch-author check` が持ち、ここではスキップと報告する。

## 制作の指示（`brief.yaml`）

カンプの依頼が必ず読む、成果物ごとの指示。形は `references/schema/brief.v1.schema.json`。

```yaml
schema_version: 1
deliverable: lp-sample        # 置き場所のディレクトリ名と同じ
script: lp-sample             # 台本の id（計画があるときは計画の script と同じ）
tokens:                       # マニフェストのトークンの道を fnmatch の glob で選ぶ。{} はすべて
  exclude: ["*.b.*", "*.c.*"] # 例：テーマ a だけを使う（include で残すものを書いてもよい）
style_references:             # 見本の画像（repo の root からの相対）。空でもよい（info で報告）
  - references/comp-style.png
emphasis:                     # 任意：まとまりごとの見た目の要望（組み方のそのまとまりの行に入る）
  blk-0003: 数字を大きく見せる
```

- `tokens`：ブランドのテーマが複数ある repo では、この成果物のテーマ以外を外す（外さないとすべてのテーマの色が混ざる）。`include` で残してから `exclude` で外す。
- `style_references`：人が承認した実物から採る。own-image-gen に `--anchor` で渡し（edit-mode）、`--take-latest` を付ける。見本があるときの `--workdir` は repo の root（image_gen は見本を今の作業ディレクトリから確かめ、codex は workdir から読むため）。
- 台本の文言・組み方の型は書かない。

## カンプの依頼（`scripts/comp_request.py`）

`python3 <このskill>/scripts/comp_request.py --repo <repo> --deliverable <id> [--parts N] [--composition plan|free]`。

書くのは own-image-gen の器の 10 フィールド（1 枚ごとに 1 つ）と、出力先・実行するコマンドを並べた `request.json`（`fields`・`out`・`commands`・`part_list` はどれも配列、`composition`・`parts`・`text_max_chars`・`brief` も持つ）。画像は生成しない（従量課金。オーナーの門の前に人が流す）。

- 組み方（`--composition`）：既定は `plan`。計画を検査して読み、まとまりごとに働き・型（列・帯・絵の置き方・区切り）・計画の細部で描かせ、「各まとまりの型を守る」と縛る。`free` は**試し**で、計画を読まない（無くてよい）。まとまりの順と働きは守らせたまま、格子と列・帯・絵の大きさと置き場所・重なり・はみ出し・何を目立たせるかをモデルに提案させる。流れの順（計画が先かカンプが先か。SPEC-deliverable-production-layers 要件 14）はオーナーの決定待ちなので、既定は `plan` のまま。`free` で描いた組み方は、計画に書き起こしてから HTML を組む。
- 分け方（`--parts N`）：既定はカンプの設定の `parts`（skill の既定 1＝ページ全体を 1 枚）。まとまりを台本の順のまま N 個の続いた組に、数がそろうよう（差は 1 まで、多い組が先）に分ける。1 から まとまりの数までの外は失格。各枚の Asset type に「N 枚のうち k 枚目（まとまり …）」と書く。前の依頼の `fields*.json` は消してから書く（画像は消さない）。
- 文案：台本の文案を、辞書（`docs/dictionary/`）の値（`{{value-…}}`）を埋めて渡す。読むのは own-infoarch-author の `check` と own-doc-update の辞書の読み手（同じ読み方を 2 つ持たない）。埋められない値はそのまま残して警告。カンプの設定の `text_max_chars` より長い文案は「灰色の線で描く」と印を付けて並べ、読める文字にさせない。短い文言（見出し・数字・ボタン）だけを文字で描かせる。カンプの文言とカンプが新しく描いた絵は使わない。
- 見た目（Style）：台本の `tone`・`target.condition`・`objective.action` から書く。skill の決めた雰囲気の語は足さない。色と書体は、解決済みマニフェストのトークンを制作の指示の `tokens` で絞ったものだけ。
- 大きさ・文字数・枚数はマニフェストの `comp`（`width`・`text_max_chars`・`parts`。skill の既定 `own-design-author/references/comp-sizes.json`、repo の `profiles/<media>.json` の `comp` だけが上書きできる）。
- 狭い画面は描かない。狭い画面の形は計画が持つ。

| code | 段 | 意味 |
| --- | --- | --- |
| COMP_BRIEF_MISSING | fail | 制作の指示が無い |
| COMP_BRIEF_INVALID | fail | 制作の指示が読めない・形が schema に合わない・`deliverable` がディレクトリ名と違う |
| COMP_BRIEF_REFERENCE_MISSING | fail | 見本の画像が無い |
| COMP_BRIEF_BLOCK_UNKNOWN | fail | `emphasis` のまとまりが台本に無い |
| COMP_BRIEF_SCRIPT_MISMATCH | fail | 制作の指示と計画の `script` が違う（plan のとき） |
| COMP_BRIEF_NO_COLOR | fail | `tokens` で絞ると色が 0 個 |
| COMP_PLAN_FAILED | fail | 計画の検査が失格（plan のとき。計画が無いときも） |
| COMP_SCRIPT_MISSING | fail | 台本が無いか読めないか、まとまりが 0 個 |
| COMP_DIR_NOT_IGNORED | fail | カンプのディレクトリが git に入る |
| COMP_MANIFEST_UNAVAILABLE | fail | 解決済みマニフェストが無いか古い |
| COMP_SIZE_MISSING | fail | マニフェストにカンプの設定（`width`・`text_max_chars`・`parts`）が無い（L0 を流し直す） |
| COMP_PARTS_INVALID | fail | 枚数が 1 から まとまりの数 までの外 |
| COMP_REQUEST_FAILED | fail | 想定外の誤り（警告に見せないため 2） |
| COMP_BRIEF_GLOB_UNUSED | warn | `tokens` の glob がどのトークンにも一致しない（書き間違いを効いた絞り込みに見せない） |
| COMP_VALUE_UNRESOLVED | warn | 辞書の値が見つからず `{{…}}` のまま渡した |
| COMP_IGNORE_NOT_CHECKED | warn | git の作業ツリーでないので ignore を確かめていない |
| COMP_NO_STYLE_REFERENCE | info | 見本の画像が無い |

終了コードは 0 書いた・1 警告つきで書いた（計画の検査の警告も含む）・2 断った（何も書かない）。

## カンプの門のあと

オーナーがカンプを見て「進めて」と言ったら（台帳には残さない）、カンプが足した細部をすべて `details` に写し、型を変えたまとまりは `pattern`・`narrow` を直し、`comp.reviewed: true` にして `plan_check.py` を流し直す。HTML は計画に沿って組む。狭い画面の崩し方は、列の規則と同じ強さ（同じ詳細度のセレクタ、同じ層）で書く。広い画面の列の指定が狭い画面の崩し方に勝つ書き方（例：列の指定だけ詳細度が高い）にしない。
