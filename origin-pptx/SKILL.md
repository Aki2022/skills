---
name: origin-pptx
description: >
  PowerPoint・スライド・プレゼン資料作成/編集で必ず使うこと。「PPTX作成」「スライド作って」
  「プレゼン資料」「プレゼンテーション」「パワポ」「パワーポイント」「資料作成」「deck」
  「presentation」「slides」「.pptx」など、これらの語が少しでも含まれる依頼で起動する。
  ネイティブ（PptxGenJS）レイアウト・箱・矢印・日本語テキスト ＋ Codex image_gen による
  文字なしアイコン/イラストのハイブリッド方式で、完全編集可能かつ視覚的にリッチな PPTX を
  生成する、実証済み(PoC完了)のパイプライン。旧来の template + python-pptx 方式
  を置き換える現行スキル（旧方式は廃止済み）。PPTXが入力・出力どちらであっても
  （新規作成・既存ファイルの編集・読み込みも含め）このスキルを使う。
---

# origin-pptx — ハイブリッド PPTX 生成スキル

## 核心原則

**ネイティブ層（PptxGenJS）＝ レイアウト・箱・矢印・日本語テキスト・チャート・表**（完全編集可能）
＋ **image_gen 層（Codex built-in gpt-image-2）＝ モックアップ（デザイン仕様）と文字なしアセット**（創造性）

- 日本語テキストは常にネイティブ。画像内に焼き込まない（小さい日本語はimage_genで誤字化リスクがあり、部分修正もできないため）。
- **実データのチャート・表は必ずネイティブ**（エクセルベースのOOXMLチャート）。imagegenで実データを描かせない（数値捏造が起きる）。→ `style-guide/chart-rules.md`
- モックアップ（②）は**仕様**であり、そのまま転写する**ソースではない**。テキスト・数値は outline.md が唯一の真実源。
- デザインの意思決定は②で完結させ、③以降は「忠実な施工」に徹する。
- **創造性の原則**: レイアウト文法は「デフォルトの語彙」であって「檻」ではない。印象重視のスライドは自由形（Free）でimage_genに構図を委ねる（Style DNAのみ遵守）。

ブランド（v3・2026-07-13確定、test3.pptx実測）: **グレーグラデーション原則**。
本文 #7F7F7F・見出し/強調 #404040・オブジェクト背景 #F2F2F2（**背景を塗ったら枠線なし**）。
キーメッセージ tone: neutral=#404040 / positive=#44546A / negative=#C00000。
フォントは**全デッキ一律 BIZ UDPGothic**（2026-07-28決定。2プロファイル制は
Hiragino Sans W4 がPowerPoint実機で解決されず廃止。Windows 10+標準搭載・LO検証も同一フォントで通る）。
キャンバスは **PowerPoint標準ワイド画面 33.87×19.05cm（960×540pt）**。
全トークンは `style-guide/tokens.json` が正典。

**スライドの意味論（v3・全スライド共通）**:

- **タイトル = tracker**（14pt灰・左上）: 簡易な章タイトルの**名詞句 or 疑問形**。目立たせない
- **キーメッセージ**（28pt・tone色）: **必ず1行**。タイトルが疑問形ならその回答。
  **全スライドのキーメッセージを並べて読むとストーリーになる**こと（①で必須チェック）
- **1スライド1メッセージ。takeaway box（下段まとめ）・バンパーステートメントは禁止**。
  説明詳細はキーメッセージ以下の本文で示すか、長い場合は右1/3の箇条書きセクション（ContentText型）
- 数字と単位は**50%ルール**（20pt数字→10pt単位、`deck_helpers.numUnit`）
- 改行は文節境界に手動で入れる。valign/align（middle・centering）を意識的に指定する

## パイプライン ⓪〜⑤（概要）

| Step                                     | 内容                                                                                                                                                                                            | 詳細                         |
| ---------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------- |
| ⓪ スタイルガイド                         | `style-guide/` のトークン・レイアウト文法・プロンプト規約・チャート規則を読む                                                                                                                   | 下記参照                     |
| ① outline.md 共同作成                    | テキスト・数値の Single Source of Truth を人間と確定                                                                                                                                            | `references/pipeline.md`     |
| ② デザインモックアップ生成               | image_gen フルスライドで構図決定。**1枚目=スタイルアンカーとして人間承認**、以降はedit-modeでアンカー参照。**全枚の構図承認（feedback.json 全ok）が③の前提**——アンカー承認は②の完了条件ではない | `references/image_gen.md`    |
| ②.5 ペルソナレビュー（任意）             | **人間が指示した場合のみ**。想定読者のペルソナ（サブエージェント）に合成デッキを採点させ、ビルド前に内容・構成の欠陥を落とす                                                                    | `references/persona-eval.md` |
| ③ ネイティブビルド                       | PptxGenJS でレイアウト＋テキスト＋チャート/表を構築、文字なしアセットを埋め込み                                                                                                                 | 下記「ネイティブビルド」     |
| ④ 検証ループ（≤5回）                     | render→決定的チェック→VLM比較（並列・変更分のみ再検証）                                                                                                                                         | `references/pipeline.md`     |
| ④.5 人間による画像確認ゲート（**必須**） | ④収束時点の preview PNG を人間に提示し、**明示的な OK を得るまで⑤に進まない**。修正指示が出たら 修正→再ビルド→再レンダ→再提示 を OK まで繰り返す                                                | `references/pipeline.md`     |
| ⑤ 最終pptx化 → 人間最終レビュー          | ④.5 の OK 後にテンプレ注入・最終検証・成果物pptx生成 → 人間が確認 → **承認後に `cleanup_deck.py` で中間生成物を掃除**（承認前は消さない）                                                       | `references/pipeline.md`     |

**⓪〜⑤はゼロベース生成の手順。既存デッキの部分修正・スライド追加の依頼では、パイプラインを
最初から回さず `references/edit-mode.md` を読む**——最初に「正がスクリプトか、人間が手を入れた
pptxか」を判定し、差分ビルド（ルートA）か外科的編集（ルートB）に分岐する。特に**人間が
PowerPoint上で直接編集したpptx（スクショ手貼り・文言修正等）に対して再ビルド上書きすると
手作業が全て消える**——このスキル最大の破壊的失敗モードで、edit-mode.md のガードが必須。

**②.5 ペルソナレビュー（任意工程・既定はスキップ）**: ②の人間承認後〜③の前に、
**人間が指示した場合のみ**実行する。大規模改訂時・重要提出物で推奨。評価対象は
「現行デッキの final render ＋ 改訂分の mockup を正典の順序で並べた合成デッキ
（`process/eval_deck/`）」で、mockup 単体では評価しない。**審査項目・ペルソナ割当・
モデル・コスト概算を人間に提案し、承認を得てから**並列サブエージェントで実行する
（結果は `process/eval_rubric.md`・`process/eval_results.json` に永続化）。
指摘は正典（outline.md/base.md）と照合し、**正典自体の変更に及ぶものは gate＝人間確認を
経てから正典を更新する**（スライド内で閉じる修正だけをループ内で処理）。
④のVLM比較は「承認モックアップ通りに施工されたか」しか見ないため、構図・ロジック級の
差し戻しは②.5 がなければ⑤まで残り、③の実装ごと捨てることになる。
詳細は **`references/persona-eval.md`**。

**④.5 人間による画像確認ゲート（必須工程・スキップ不可）**: ④が収束したら、
**テンプレ注入・最終pptx化・納品の前に**、その時点のレンダリング画像（preview PNG）を
人間に提示し、**明示的な OK が出るまで⑤に進まない**（初回ビルドは全枚。修正巡は変更
スライド＋前後関係が分かる範囲でよい）。修正指示が出たら 修正→再ビルド→再レンダ→
**再び画像で確認** を人間が OK するまで繰り返す（回数制限なし）。理由: pptx は
パッケージング成果物であり、**レビューの単位は常に画像**。画像段階で確定させることで、
pptx 生成後の差し戻し（注入やり直し・成果物の版管理混乱）を無くす。
**②.5（AIによる品質評価・任意）とは別物**で、②.5 を実施しても ④.5 は省略できない。

## 作業ディレクトリ規約（必須）

デッキ1本ごとに、呼び出し元プロジェクトの `presentation/` 配下に次の構造を作る。
同じプロジェクトで複数のPPTXを継続的に作成しても混ざらないための規約。

```
<プロジェクトルート>/presentation/
└── yyyymmdd_内容/                 # 例: 20260708_roa7カ国比較
    ├── input/                     # 素材（元資料md・事前指定画像=スクショ/書影等）。人間が置くか、AIが素材をコピーする
    ├── process/                   # 中間生成物すべて: outline.md・mockup_task.md・mockup_*.png・
    │                              #   生成アイコン・build_deck.js・insert_charts.py・preview-*.png・pdf中間物
    └── yyyymmdd_内容.pptx         # 最終成果物（この階層に直接置く）
```

- **デッキディレクトリ作成時に `yyyymmdd_内容/.gitignore` を必ず自動設置する**（毎リポジトリの
  .gitignore 手入れを不要にするため）。雛形は次の方針: 再生成可能な重量物とローカル残骸
  （`process/node_modules/`・`process/*.png`・`process/*.pdf`・`process/output.pptx`・
  `process/*.log`・`process/assets/`・codexが勝手に作る `generated/` や `icons_task_*/`・
  PowerPointロックファイル `~$*.pptx`）を除外し、*_再構築ソース（outline.md・slides/・
  build_deck.js・feedback.json・package_.json）と input/・最終pptx は追跡**する。
  mockup/preview 等のPNGは git に入れない——mockupの役割は「②承認〜⑤最終承認までの検証用仕様書」で、
  ⑤以降の修正の正はスクリプト側に移るため中長期価値がない（デザインやり直し時は新mockupを作る）
- 元資料が別リポジトリ等にある場合はパス参照ではなく **input/ にコピー**する（再現性の担保）
- 最終pptx以外を `yyyymmdd_内容/` 直下に置かない（成果物が一目で分かる状態を保つ）
- 反復で成果物を作り直す場合も同じファイル名に上書き（版はgitに任せる）。ただし**上書き前に必ず
  `cmp -s process/output.pptx <成果物>.pptx` を実行**——不一致なら人間が成果物を直接編集しており、
  上書きすると手作業が消える。その場合は `references/edit-mode.md` のルートBへ
- **`<プロジェクトルート>` は「今の作業ディレクトリ」基準で解決する（git worktree に注意）**。
  書き込み前に `pwd` と `git rev-parse --show-toplevel` を確認し、worktree で作業しているなら
  その worktree のパス配下に置く。main チェックアウト側に書くと、worktree にいる人からは
  「ディレクトリが存在しない」ことになる（2026-07-10 に実際に取り違え、rsync移送が必要になった）。

## style-guide/ の使い方

- `style-guide/tokens.json` — 色（グレー階調・message.neutral/positive/negative）・タイポグラフィ（BIZ UDPGothic一律・整数ptスケール）・クローム座標（chrome.\*）・表・チャート・図形のデザイントークン。**新規の色相を勝手に追加しない**。
- `style-guide/layout-grammar.md` — **基本型（表紙/目次/セクション/単一/2列/3列/比較/ハイブリッド/自由形）× 修飾子（mode: handout/preso、tone: neutral/positive/negative、content）** の体系。タイトル=tracker/キーメッセージの意味論と禁止事項（takeaway box等）もここ。imagegen用の定性バンド記述とネイティブ用pt座標表の両方を持つ。
- `style-guide/skill-config.json` — 実行設定（サブエージェントのモデル割当・検証ループ上限・並列度）。⓪で読む。
- `style-guide/template_v3.pptx` — **注入用テンプレ**（会社テンプレ124レイアウトから11枚に間引き・v3化した正典。生成元は `scripts/build_template_v3.py`）。④合格後に `inject_template.py` で最終成果物の土台にする。
- `style-guide/imagegen-prompt-convention.md`（v4） — OpenAI公式スキーマ準拠のプロンプト組立規約（ラベル付きフィールド固定順序・Style DNA・**anchors/参照画像の必須使用**・図解デバイス語彙・検収チェックリスト）。
- `style-guide/anchors/` — **実デッキから採った様式見本PNG（最重要資産）**。v3様式の正は `anchor_v3_*.png`（旧navy様式は `_deprecated_v2/`、新規デッキでは使わない）。フルスライドモックアップ生成時は必ずedit-modeの参照画像として渡す。文章のStyle記述だけでは実様式は再現されない（2026-07-07 A/Bテスト実証）。人間承認された新スライドから様式見本を随時追加する。
- `style-guide/chart-rules.md` — チャート・表の生成規則（ネイティブ必須、crtx翻訳スタイル、imagegen装飾の限定許可）。

## ネイティブビルド（③）

**★③を開始する前に `python3 <skill>/scripts/check_feedback.py <デッキdir>` を実行し、
対象スライドの verdict が全て ok であることを確認する**（exit 0 でなければビルダーを起動しない）。
様式アンカーの承認と、スライドごとの構図の承認は別物——2026-08-05 に取り違えて未承認13枚を
ビルドし、後から構図レベルの差し戻しで実装を捨てた。

**承認済みモックアップが構図の仕様書。スライドの実装は、そのスライドの `mockup_NN.png` を
実際に読んだ者だけが書ける。** メインループは画像を読まない規律（pipeline.md）があるため、
実務は「**スライドごとにビルダーサブエージェント（Sonnet）へ mockup 1枚＋outline該当節を渡して
実装させる**」形になる。outline.md / mockup_task.md のテキスト記述だけから組んではいけない——
テキストは構図の近似でしかなく、2026-07-10 に24枚中18枚が承認モックアップと別物になった
（アイコン・バッジ・吹き出し等のデバイス欠落、入れ子構図の平坦化）。
モックアップに**アイコン・イラストが写っているなら、下記のアセット生成（image_gen）は省略できない
必須工程**（各ビルダーが「このスライドに必要な文字なしアセット」を列挙→まとめて1バッチ生成）。

`scripts/build_slide.example.js` が1枚もの参照実装（PoC実証済み・**v2時代のため座標/色は旧仕様、
構成の参考のみ**）。複数枚デッキでは
**`scripts/deck_helpers.js` v3（共通ヘルパ: chrome/card/arrow/imgFit/badge/numUnit 等＋
`defineSlideMaster` によるクローム層＝日付・ページ番号・縦書きコピーライトの自動化）を
`require` して土台にし、`scripts/split_deck.py` で `slides/sNN.js` に分割してビルダーを並列化**する
（1枚岩だと並列ビルダーが同一ファイルで衝突する）。各ビルダーへは `references/builder_brief.template.md`
を渡す。**書く前に `references/pptxgenjs-gotchas.md` を必ず一読**（ShapeTypeはインスタンス側・chevron/
homePlateのテキスト切れ・フッター下端clip・ST.line矢じりが薄い→塗り三角・画像アスペクト保持・
addChartの限界・アイコン正規化）。
**アイコンは埋め込み前に `scripts/normalize_icons.py` で正規化**（白背景透過＋ライン画を #404040 単色化。
ポジ/ネガ意味のあるアイコンのみ `--color 44546A` / `--color C00000`）——
生のimage_gen産は白ボックス露出＋線が淡く視認不良になる（2026-07-10に2周した）。

- タイトル(tracker)・キーメッセージ・カード・ラベル・本文はネイティブテキスト/シェイプ。
  **できるだけPowerPoint標準プリセット図形を使い、単体で無理な形は標準図形の組み合わせで作る。
  枠線・区切り・表を画像で描かない**
- `pptx.ShapeType.roundRect` でカード（panel塗り・枠線なしが既定）、`pptx.ShapeType.rightArrow` 等の
  **塗り矢印シェイプ**でコネクタ
  （工程ステップに `chevron` を使うと左の切れ込みが左寄せ文字を食う→ `homePlate` を使う。gotchas §2）
- **チャート: chart-rules.md の「経路の選び方」に従う**。実データは **`slide.addChart`（Excel編集可能な
  ネイティブOOXML）が必須**で、チャート機能を超える装飾（末尾ラベル・国旗・コールアウト等）は
  **ネイティブのオーバーレイで重ねて再現**する（chart-rules §2.5）。チャートの画像化（SVG→PNG）は
  ユーザーがExcel編集可能性を明示的に放棄した場合の最終fallbackのみ。**imagegenで実データを描かせない**
- **例示チャート（実データが存在しない図）は「空枠」ではなく「数値なしイメージ図」で描く**（gotchas §12）。
  空のグレー帯は人間に未完成と映り差し戻される（2026-07-11実証）。棒・凡例・カテゴリラベルは
  ネイティブ図形で描き、数値・%・目盛りは一切描かず「イメージ」と明記する
- 表: **必ず** `slide.addTable`（tokens.json `table.*`）。テーブルの枠線・セルを図形や画像で模造しない
- **各スライドに `slide.addNotes("...")` でスピーカーノート（outline.mdの「語り」を1〜3文）を付与する**（詳細は `references/builder_brief.template.md`）
- **箇条書きは1つの `addText()` に breakLine ランを並べて自動フローさせる**（項目ごとに y を
  手計算して積み上げない）。**1項目=1ラン・ラン内に `\n` を入れない・bullet 指定は全ランで揃える**
  ——複数ランに割ると■が消える/増える/インデントがズレる（gotchas §20-21）
- 出所を**URLリンク付き**で出す場合は `sourceLinks(s, [{label, url}, ...])`（`source()` の
  リンク版・同一座標）
- `slide.addImage()` で文字なしアセットPNG（image_gen製）や実データチャート図を埋め込み
- **事前指定画像（スクショ・書影等の実物）はデッキの `input/` に置き、②ではedit-mode入力として構図に組み込み、③では原本を `addImage`**（prompt-convention §5.5）。歪み防止にアスペクト比保持でフィット（gotchas §4）
- キャンバス: `pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 })`（960×540pt = 33.87×19.05cm）
- フォント: `mk({ font: "BIZ UDPGothic" })`（全デッキ一律・未指定はエラー）。ビルド直後と
  inject後の最終ファイルに `scripts/set_fonts.py <pptx>` を必ず実行し、テーマ・テンプレ由来の
  レイアウト/マスター含む全XMLを BIZ UDPGothic に統一する（冪等）
- 色・サイズ・クローム座標は必ず `tokens.json`（v3）から引く（ハードコード禁止）
- 品質基準（v3・test3実測）: tracker14pt / キーメッセージ28pt(1行) / カード見出し25pt bold /
  本文14pt / 数字20pt+単位10pt。本文・脚注の最下端 y≤500、フッター帯 y=512.8 は master 任せ。
  吹き出しは`speechBubble()`ヘルパ＝標準wedgeRoundRectCallout単一オブジェクト（gotchas §10・sanitize必須）
- 人物・UIモック等の「密度が品質に直結するアセット」は**quality=high**＋人物はバストアップ指定で生成（gotchas §13）

```bash
npm install pptxgenjs
node build_slide.js                                  # → output.pptx
python3 <skill>/scripts/sanitize_pptx.py output.pptx # 必須: PowerPoint修復エラー要因の矯正（gotchas §16-17）
python3 <skill>/scripts/set_fonts.py output.pptx        # 必須: BIZ UDPGothic に統一
# ④.5（人間の画像OK）後・⑤の最終成果物化（テンプレ注入 Phase 1）:
python3 <skill>/scripts/inject_template.py output.pptx <skill>/style-guide/template_v3.pptx 最終成果物.pptx
python3 <skill>/scripts/set_fonts.py 最終成果物.pptx     # 必須: テンプレ由来部品もBIZ UDPGothicに統一
```

**`sanitize_pptx.py` はビルドの必須最終工程**（冪等）。pptxgenjs 4.0.1 は折れ線/散布図/レーダーの
チャートXMLにスキーマ違反（`invertIfNegative` 混入・`marker` の位置違反）を出力し、上向き/左向きの
線は負extentになり、いずれも PowerPoint が「修復」を要求する。LibreOffice では検出できないため、
レンダリング検証とは別にこのサニタイザを必ず通す。

**`inject_template.py` は④.5（人間の画像確認OK）後の最終成果物化**: `style-guide/template_v3.pptx`（会社テンプレ由来・
10レイアウト・v3テーマ）を土台に生成スライドを移植する。これにより人間が PowerPoint の
「新しいスライド」で会社レイアウト（11種）を使え、テーマの色（44546A等）・フォント（Noto）が正しく並ぶ。
注入後は 1回レンダして注入前と同一であること＋ `sanitize_pptx.py --check` を確認する。

## image_gen（Codex built-in, gpt-image-2）— 概要

②のモックアップ、③のアセット生成の両方で使う。組立規約は `style-guide/imagegen-prompt-convention.md`。

```bash
codex features list | grep image_generation   # → stable true を確認
codex exec --sandbox workspace-write -c sandbox_workspace_write.network_access=true --cd "$PWD" "<プロンプト>"
```

- gpt-image-2、ChatGPTサブスクOAuth。**OPENAI_API_KEY 不要・従量課金なし**
- アセットのプロンプトは必ず **"NO text, no labels"** で終える
- **一貫性はスタイルアンカー方式**（シード値は存在しない）: 承認済み1枚目をedit-modeの参照画像に

### ⚠️ SAVE-PATH GOTCHA（最重要・必読）

image_gen は生成物を**プロンプトで指定したパスには保存しない**。既定で
`$CODEX_HOME/generated_images/<session-id>/*.png` に保存される。生成後は必ず最新ファイルを探して
ワークスペースへコピーする:

```bash
find "${CODEX_HOME:-$HOME/.codex}/generated_images" -type f -name '*.png' -print0 \
  | xargs -0 ls -t | head -1
```

### ⚠️ AUTHORIZATION（実行前に必読）

- 上記の**安全モード**（`--sandbox workspace-write -c sandbox_workspace_write.network_access=true`）は
  Claude Code の通常権限フローで **AI が直接実行できる**（2026-07-26 hachinohe_sea_2026 で実証。
  実運用例: 同リポジトリ `presentation/20260726_八戸魚市場提案/process/run_mockups.sh`）。
- 残存リスク: `network_access=true` により codex が外部ネットワークへ出られる＝プロンプトへの
  インジェクション経由でワークスペース情報を持ち出す余地が理論上残る。**codex へ渡すプロンプトは
  AI 自身が組み立てた画像生成指示のみ**とし、外部由来テキストを埋め込まない。秘密情報を含む
  ディレクトリを `--cd` 対象にしない。
- `--dangerously-bypass-approvals-and-sandbox` は**使わない**。Claude Code の分類器がハードブロック
  するため、許可ルールがあっても AI からは実行できない（2026-07-26 実証）。安全モードで不足がある
  場合の**最終手段**としてのみ、**人間が自分の手で実行する**。AIは自己承認や権限回避を試みない。
- バックグラウンド起動時は **`< /dev/null` を必ず付ける**（stdin待ちハング。image_gen.md参照）。
  生成物の回収は `scripts/collect_codex_images.py`（並列競合の防止）。

詳細は **`references/image_gen.md`** を参照。

## レンダリング検証（④）

```bash
soffice --headless --convert-to pdf --outdir . output.pptx
pdftoppm -png -r 150 output.pdf preview
```

決定的チェック（テキストオーバーフロー・フォント置換・要素個数・欠落アセット列挙・全スライドの
スピーカーノート有無（`slide.notes_slide`）に加え、v3では
**python-pptxによるXML機械検査**: タイトルが名詞句/疑問形か・キーメッセージ1行か・フッター/コピーライト
存在・数字/単位50%比率・takeaway box不在）を先に、その後VLMで
②承認モックアップと構図比較 — **全スライドのフル解像度ペア精査を最低1巡必須**（モンタージュは着手順決めのみ、
合格判定に使わない）。比較プロンプトは敵対的（あら探し・差分列挙型）に書く。検収基準は
imagegen-prompt-convention.md §10（型スライド=厳格 / 自由形=ブランド適合のみ）。
**ループ上限5回**（skill-config.json）。合格したら早期終了・2巡目以降は**変更したスライドのみ再検証**・
前巡より差分が減らなければ人間へエスカレーション（発散ガード）。**ペア比較サブエージェントは並列で投げる**
（互いに独立・1体1ペアの画像規律とも整合）。

**免罪符の禁止（チェッカーにもオーケストレーターにも・2026-07-11 差し戻し実証）**: チェッカーに
「この差分は想定内」という免除リストを渡さない。さらに、チェッカーが報告した差分をオーケストレーターが
「装飾差・規約由来だから意図的」と一括却下するのも同じ失敗——却下は個別に理由を明文化できるものだけ、
迷ったら修正側に倒す。構図一致に加えて「**人間が見て洗練されて見えるか（品質印象）**」を独立の
合格基準として問う（タイトルの迫力・アイコンの存在感・空パネルの未完成感は構図一致では検出されない）。
詳細は `references/pipeline.md` ④。

**④の完了定義**: (a) 決定的チェック合格 **かつ** (b) preview↔mockup 全ペア比較で乖離が解消
（または人間が明示承認）——**両方**を満たして初めて④.5（人間の画像確認ゲート）へ進める。欠陥QA（あふれ・豆腐・衝突）だけ
回して④を済ませたことにしない。欠陥QAは「壊れていないこと」しか保証せず、「承認したデザインで
あること」はペア比較でしか保証されない（2026-07-10: 欠陥QAのみで⑤に進め、18枚の構図乖離が
人間レビューまで素通りした実失敗）。
**トークン規律**: 検収・目視検分=Haiku、VLM比較/ビルド=Sonnetの使い捨てサブエージェントで行い、**メインループは画像PNGを
原則Readしない**。**④.5の送付前検分・⑤注入後の目視確認もHaikuに任せ**、メインループは報告を受けて添付送付するだけ。
複数枚の検分は**1体あたり最大4枚で分割並列**（1体に全部読ませると文脈累積で総入力が約N²/2に膨らむ）。
（詳細は `references/pipeline.md` の「モデル・トークン運用規則」）。

## 前提環境

- `pptxgenjs`（npm）
- フォント: BIZ UDPGothic（Windows 10+標準。macはGoogle Fontsから導入・`fc-list | grep -i "biz ud"`で確認）
- `soffice`（LibreOffice）・`pdftoppm`（poppler）
- Codex CLI ログイン済み（`codex features list` で `image_generation stable true`）

## リファレンス

- `references/pipeline.md` — ⓪〜⑤各ステップの詳細（入出力・コマンド・チェックポイント・検証ループ・コスト効率）
- `references/edit-mode.md` — **既存デッキの部分修正・スライド追加**（正の所在判定→差分ビルド/外科的編集の分岐、人間の手編集を消さないガード、python-pptxレシピ、ミニデッキ合流）
- `references/persona-eval.md` — **②.5 ペルソナレビュー**（任意工程・人間指示時のみ）: 評価計画の人間承認・eval_deck の組み立て・出力スキーマ・正典照合ルール（gate）・ループ規律
- `references/image_gen.md` — image_gen の全詳細（保存先の罠・透過クロマキー・権限ルール・モデル仕様）
- `references/pptxgenjs-gotchas.md` — ③ネイティブビルドの落とし穴集（ShapeType・chevron/homePlateの文字切れ・フッター下端clip・ST.line矢じり・画像フィット・addChart限界・アイコン正規化・LibreOffice特有の罠）
- `references/builder_brief.template.md` — 各スライドビルダーへ渡す共通ブリーフの雛形（正本優先順位）
- `style-guide/` — tokens / layout-grammar / imagegen-prompt-convention / chart-rules
- `scripts/build_slide.example.js` — PptxGenJS 1枚もの参照実装（v2時代・構成の参考のみ。座標/色はv3が正）
- `scripts/deck_helpers.js` — 複数枚デッキ用の共通ヘルパ v3（③の土台に使う。masterクローム/numUnit含む）
- `scripts/split_deck.py` — build_deck.js を slides/sNN.js に分割（ビルダー並列化用）
- `scripts/normalize_icons.py` — image_gen産アイコンの正規化（白背景透過＋#404040単色化。埋め込み前に必須）
- `scripts/set_fonts.py` — **ビルド後必須**のフォント統一（全XMLの a:latin/a:ea/a:cs を BIZ UDPGothic に書き換え。ビルド直後と inject後の最終ファイルの2回実行・冪等）
- `scripts/sanitize_pptx.py` — **ビルド後必須**のPowerPoint修復エラー矯正（チャートXMLスキーマ違反・負extentの修正。gotchas §15-17）
- `scripts/inject_template.py` — **④.5（人間の画像確認OK）後の最終成果物化**: template_v3.pptx を土台に生成スライドを移植（人間がマスター準拠スライドを追加できる形にする）
- `scripts/check_feedback.py` — **③開始前の必須プリフライト**: `feedback.json` の verdict が
  全て ok か確認する（②の人間承認を飛ばして③へ進むのを機械的に防ぐ）
- `scripts/cleanup_deck.py` — **⑤承認後の必須クリーンアップ**（dry-run既定）。役割を終えたmockup/preview/ログ等を削除し、再構築ソース（outline/slides/assets/output.pptx）は残す。承認前に実行しない
- `scripts/build_template_v3.py` — template_v3.pptx の生成スクリプト（旧会社テンプレ→間引き・v3化。テンプレ更新時に再実行）
- `scripts/build_chart_svg.py` — 実データチャートのネイティブ描画（SVG→soffice PNG）雛形
