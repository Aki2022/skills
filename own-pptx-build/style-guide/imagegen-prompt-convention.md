---
title: Imagegen Prompt Convention
version: 4.0.0
---

# Imagegen Prompt Convention v4

> v4 (2026-07-13): スタイル体系を v3トークン（グレーグラデーション・Noto Sans CJK JP・960×540・
> title=tracker+1行キーメッセージ）に全面更新。プロンプト構造・運用は v3 のまま。

`tokens.json` ＋ `layout-grammar.md` の型 ＋ スライド内容（outline.mdの1枚分）から、imagegen用プロンプトを
組み立てる規約。OpenAI公式のプロンプトガイド（GPT Image Prompting Guide / Codex imagegen SKILL.mdの
スキーマ）に準拠する。生成物は「レビュー用デザインモックアップ」または「文字なしアセット」であり、
最終テキストは常に outline.md からネイティブに流し込む。

> **v3の核心（2026-07-07 A/Bテストで実証）**: 文章によるStyle記述だけでは実際のハウススタイルは
> 再現されない。**`anchors/` の実スライド参照画像をedit-modeで渡すこと**が最も効果的で、
> Style DNAはその補助。また、レイアウトを固定するのではなく**「この内容に最も効果的な図解
> デバイスを選べ」と図解の選択をモデルに委ねる**ことで創造性が引き出される。

> **創造性の原則**: Style DNA（§2）とハウススタイル構造（§2.5）は常時適用。図解デバイス（§4）は
> モデルに選択させる。自由形（Free）ではさらに制約を減らし構図全体を委ねる。規約はモデルを縛る
> ためではなく、**大量生成時の一貫性と後工程（ネイティブ再現）の成立**のためにある。

## 0. スタイルアンカー（最重要・毎回使う）

`style-guide/anchors/` に**実デッキから採った様式見本**が保存されている:

- `anchor_v3_graytheme-3col-cards.png` — **v3様式の正（2026-07-13、人間手直しデッキ由来）**:
  tracker＋赤1行キーメッセージ・3列カード（極薄グレー角丸・枠線なし）・グレー単色アイコン・
  数値バンド（数字+単位50%）・赤枠吹き出し・薄灰フッター・右端縦書きコピーライト
- `anchor_v3_flow-heronode.png` — 3段フロー＋中央強調ノード（#44546A塗り白文字）＋グレー太矢印
  （2026-07-14 宇宙リスクファイナンスデッキS9・人間承認済み）
- `anchor_v3_contenttext-chart.png` — ネイティブ縦棒チャート＋赤枠吹き出しコールアウト
  （同デッキS7・人間承認済み）
- `anchor_v3_agenda-nav.png` — **現在地ナビ章扉**: 左端の紺 #44546A 縦バンド（Contents）＋
  目次再掲・現在セクションのみ黒太大（2026-07-16 北浦メディアデッキ・人間手直し由来）
- `anchor_v3_contenttext-insightbox.png` — ContentText型の右カラム見本: 項目数スケールの
  箇条書き（4項目=18pt）＋右カラム下部の示唆ボックス（白地・細枠・太字紺）
  （同デッキ・人間手直し由来。layout-grammar §8 の右カラム規則とセット）
- `_deprecated_v2/` — 旧navy様式のアンカー（v2互換の再生成時のみ使用。**新規デッキでは使わない**）

新規デッキで人間承認されたスライドから様式見本になる構図（マトリクス・人物ユースケース等）を
随時 `anchor_v3_*.png` として追加していくこと。

**フルスライドモックアップ生成時は、必ずv3アンカー（＋承認済み自デッキスライド）をedit-modeの参照画像として渡す**:

```
First view style_ref images with view_image, then generate in edit-mode:
Image 1-3: style references (real slides from our house deck).
Match their visual style exactly — same title treatment, same annotation devices,
same density, same color roles. New slide content: <内容>
```

## 1. プロンプト構造 — ラベル付きフィールドの固定順序

公式スキーマに従い、**この順序のラベル付き短文ブロック**で組む（長い一段落にしない）:

```
Use case: <infographic-diagram | ui-mockup | productivity-visual 等の分類>
Asset type: <full slide mockup (16:9) / icon / illustration / background>
Style references: <§0 anchors/ の参照画像（edit-mode）>
Style: <§2 Style DNA を逐語貼付>
House structure: <§2.5 ハウススタイル構造>
Composition: <§4 図解デバイス指示（内容＋候補提示、選択はモデル）>
Content: <§5 スライド固有の内容（日本語は「」引用）>
Color palette: <§3 色指示（hex＋言語表現）>
Constraints: <§6 テキスト量・余白等>
Avoid: <§7 負の制約>
```

- 用途宣言（Use case / Asset type）を必ず先頭に。モデルの「モード」と仕上げ水準が決まる
- 反復修正時は「1ターン1変更」とし、**変えない要素を毎回明示的に再宣言**する（ドリフト防止）

## 2. Style DNA（固定・デッキ内で逐語再利用）

デッキ内の全スライド・全アセットで**一字一句同じもの**を使う。書き換え禁止。

```
Information-rich Japanese consulting-style presentation slide, in the exact house style
of the reference images. GRAY-GRADATION palette: the slide is almost entirely grayscale —
medium gray #7F7F7F body text, dark charcoal gray #404040 for headings, emphasis and any
colored strokes/icons; very light gray #F2F2F2 rounded panels WITHOUT borders (filled
objects never get outlines); light gray #D9D9D9 footer text and hairlines. Only two hues
are allowed as accents: muted slate blue-gray #44546A for positive/recommended emphasis,
and strong red #C00000 for negative key messages, warning annotations, red-bordered
callouts and inline emphasis of critical words. Icons and illustrations are monochrome
dark-gray line art. Typography: Noto Sans JP style gothic. Dense but organized —
multiple visual devices per slide are welcome (numbered badges, underlined sub-headers,
thin gray connector arrows, red-bordered speech-bubble callouts, small gray line icons,
flat illustration characters for human situations, dotted section dividers).
No gradients, no glossy effects, no 3D renders, no photographic stock imagery.
This is one slide in a consistent multi-slide deck — keep the visual language identical
to the reference images and prior slides.
```

- hexは近似的にしか守られないため、**hex＋言語表現の冗長ペア**で書く（上記の形）
- ブランド最重要色（背景・帯・チャート系列）はどのみちネイティブ層で正確に塗る。モックアップの色ズレは構図承認に影響しない範囲で許容

## 2.5 ハウススタイル構造（毎回のプロンプトに含める）

実デッキ（test3.pptx）から抽出した頁構造。**塗りつぶしタイトルバンドは使わない**（過去の生成で頻出した誤り）:

```
Page structure: a SMALL gray tracker label at top-left — a short chapter noun-phrase or a
question (e.g. 「課題は？」). Below it, ONE key-message line (larger text): if the tracker
is a question this line answers it; colored dark gray #404040 for neutral, slate blue-gray
#44546A for positive, or #C00000 (red) for negative messages. NO filled title band, NO
header bar. Body: one rich diagram area — ONE message per slide, NO takeaway box or summary
band at the bottom (if explanation is needed, use a right-third bullet column instead).
Footer: small very-light-gray date bottom-left, page number bottom-right, and a tiny
vertical copyright line on the right edge.
```

- tone修飾子はここで効く: neutral=キーメッセージが濃灰 / positive=青灰 #44546A / negative=赤 #C00000

## 3. 色指示

tokens.json のロールを自然文に落とす。新しい色相の導入を明示的に禁止する:

```
The slide is grayscale by default: #7F7F7F body text, #404040 dark gray for headings,
emphasis, icons and strokes; #F2F2F2 light gray panels with NO borders; #D9D9D9 footer.
Neutral key messages in #404040; positive emphasis ONLY in #44546A (slate blue-gray);
negative key messages and warnings ONLY in #C00000 (strong red).
Do not introduce any other hue (no greens, no oranges, no bright blues, no purples).
```

## 4. 図解デバイス指示 — 内容を渡し、デバイスの選択はモデルに委ねる

**%座標や数値座標は渡さない**（モデルは守らない）。また**「箇条書きカード」に自分で決め打ちしない** —
退屈な出力の最大要因は、プロンプトがカード列しか語彙として与えていないことだった（A/Bテスト実証）。

基本形: 内容の論理構造（因果・対比・循環・構成）を説明し、**「choose the MOST EFFECTIVE consulting
diagram device for this content」と選択を委ねる**。必要なら候補を添える:

**図解デバイス語彙**（実デッキから抽出。プロンプトで名指し可能）:

- 因果/プロセスフロー（番号バッジ＋下線小見出し＋薄灰色矢印、各ステップに補足行とアイコン）
- 2×2マトリクス（薄いグラデ背景、軸ラベル、○×/記号、象限注釈）
- 中心ハブ地図（当社/対象を中央に、左右にステークホルダー、矢印に小さな取引ラベル）
- 循環ループ（3〜4ノード＋回転矢印）
- 対比カラム（下線付きカラム見出し、before/after矢印）
- 引用グリッド（アイコン＋出所ラベル＋「」引用文、キーフレーズを赤強調）
- 赤枠吹き出しコールアウト（重要注記）、点線区切りによる上下段構成
- 人物イラスト（フラット・いらすとや調）— 人間の状況描写に有効な場合

個数の規律は維持する: 「**exactly three** steps」「exactly two arrows」（"some cards" は禁止）。
余白: 「no element touches the frame edge」。

- **自由形（Free）**: デバイス語彙も省略し「Compose freely to express: <スライドの意図・感情>」とだけ書く

### レイアウト厳守が必要な場合の上級レバー: ワイヤーフレームアンカー

定性表現で構図が安定しない場合、layout-grammar.md の座標表からネイティブ描画した粗いレイアウトPNG
（灰色の矩形ブロックのみ）を作り、**edit-modeの入力画像として渡す**:

```
Image 1: layout wireframe (gray blocks show where each element goes).
Follow this layout exactly: replace each gray block with the corresponding element described
below. Keep block positions and sizes.
```

## 5. スライド内容

- 日本語の必須テキストは**「」で引用**し、実ラベルを使う（lorem ipsum・ダミー英文は禁止 — 文字化けの誘発源）
- モックアップの日本語は構図確認用の近似でよい（誤字許容）が、**項目の個数・行数は正確に**
- 例: `Slide title: 「導入までの3ステップ」 / Card 1: 「申込み」 with one supporting line 「オンラインで数分」`

## 5.5 事前指定画像（input/ 規約）

アプリのスクリーンショット・書影・実物写真など、**生成ではなく実物を使うべき画像**は、
デッキ作業ディレクトリの `input/`（SKILL.md「作業ディレクトリ規約」参照）に置き、
outline.md の該当スライドでファイル名を参照する
（例: `画像: input/app_dashboard.png をボディ右側に配置`）。

- **②モックアップ生成時**: 該当画像を view_image で読み込み、**edit-modeの入力画像**として渡す:
  ```
  Image N: content image (real screenshot) — place this INSIDE the slide composition at
  <位置>, framed with a thin gray hairline. Do not redraw or stylize it.
  ```
  実デッキのUIスクショ埋め込み（構造図の中に実画面を配置する様式）と同じ扱い。
- **③ネイティブビルド時**: モックアップ内の再描画版ではなく**assets/の原本**を `addImage` で
  埋め込む（品質と正確性の保証。モックアップ内のスクショ描画はあくまで配置指定）。
- 検収チェックに追加: assets指定画像が構図に含まれ、置き換え・省略されていないこと。

## 6. テキスト量・その他制約

- layout-grammar.md のテキスト量目安（ノード見出し8字、補足14字×2行等）を一文で添える
- 画像内テキスト総量は最小限に。文字が多い・小さいほど日本語の字形崩れリスクが上がる
- フルスライドモックアップは 16:9。サイズはパラメータ/コマンドで指定し、プロンプト文では述べない

## 7. 負の制約（毎回必ず付ける）

```
Avoid: watermarks, logos, photographic stock imagery, 3D bevels, neon or saturated colors
outside the palette, filled title bands, takeaway boxes or summary bands at the bottom,
outlines on filled panels, more columns/nodes than specified,
decorative borders around the frame, unrelated extra elements.
```

（v2にあった「photographic people 全面禁止」「minimal iconography」「dense paragraphs禁止」は撤廃 —
ハウススタイルは人物イラスト・高密度図解を歓迎する。禁止するのは写真調の人物・様式外の要素のみ）

## 8. 大量生成の一貫性 — スタイルアンカー方式

シード値は存在しない。N枚デッキの一貫性は次の運用で作る:

1. **常に `anchors/` の実スライド見本2〜3枚を参照画像に含める**（§0。ゼロから文章だけで様式を
   伝えようとしない）
2. デッキ1枚目を生成 → 人間承認 → 以降は **anchors/ ＋ 承認済み1枚目**を併せて参照画像に渡す:
   ```
   Image 1-2: house style references (anchors/). Image 3: approved slide from this deck.
   Match their visual style exactly — same title treatment, same annotation devices,
   same density, same color roles. New slide content: <このスライドの内容>
   ```
3. 反復修正が3〜4ターン続いたら、ベースプロンプト＋アンカー参照に**再アンカー**する（ドリフト防止）
4. アイコン等の小アセットも同様: 最初の1個を「基準アイコン」とし、以降は参照画像＋同一スタイルトークンで1個ずつ生成（シート一括生成は探索時のみ。量産は1呼び出し1アセット）

## 9. 品質2段階運用

- **quality=low（下書き）**: 構図・レイアウトの承認を取るまで。速く安い
- **quality=high（最終）**: 承認後の清書と、文字を含む・細部が問われるアセットのみ
- **例外（最初からhigh）**: 人物・キャラクターのイラストとスマホUIモック等、密度が品質に直結する
  アセット。lowの人物はスライド上で60-70ptに縮小されると表情が潰れて無個性なシルエット化する
  （2026-07-11差し戻し実証）。人物は「close-up bust portrait (shoulders up), large expressive face」
  を必ずプロンプトに含める
- 量産時のコスト・時間はこの2段階で制御する

## 10. 出力検収チェックリスト（1枚ごと）

**型スライド（厳格）**:

- [ ] 要素の個数が指示と一致（列数・ノード数・矢印数）
- [ ] 余分な要素・余分なテキスト・透かしがない
- [ ] 端に要素が接触していない（セーフマージン確保）
- [ ] 色がパレット内（新しい色相が混入していない）
- [ ] tone色が正しい（neutral=#404040系 / positive=#44546A系 / negative=#C00000系）
- [ ] キーメッセージが1行・下段にtakeaway box/まとめ帯がない
- [ ] 塗りパネルに枠線が付いていない
- [ ] テキストの個数・行数がoutlineと一致（字形の正確さは不問）

**自由形（ブランド適合のみ）**:

- [ ] 色・トーンがStyle DNAに適合
- [ ] 16:9・端の接触なし
- [ ] 意図した感情・メッセージが伝わる構図か（人間判断）

## 11. チャートについて

実データのチャートをimagegenで描かせない（数値・軸・比率の捏造が起きる）。詳細は **chart-rules.md**。

**②のモックアップでも、実データチャートは数値を描かせない。** チャート領域は「空のラベル付き
プレースホルダ枠」にする。プロンプト例:

```
Composition: reserve a large empty chart area (a light gray framed box) labeled in Japanese
「チャート領域（7系列 折れ線） / 数値は最終版でネイティブ描画」. Do NOT draw any data points,
axes values, or numbers inside it — it is a placeholder to be replaced by a native chart.
```

**理由（2026-07-10 実証）**: ②で"それらしい"数値をimagegenに描かせると、人間が構図レビュー時に
捏造数値をそのまま見てしまい「数字が全く違う」とrejectされる。モックアップは構図の仕様であって
数値の仕様ではない。実データの図は③で `scripts/build_chart_svg.py`（SVG→soffice）等でネイティブ
描画し、数値は outline.md を唯一の真実源とする（画像からの読み取り禁止）。

---

## 実行手順（Codex image_gen）

モデル: gpt-image-2（Codex CLI組み込み、ChatGPTサブスクリプションOAuth、`OPENAI_API_KEY`不要・従量課金なし）。

### 呼び出し

```bash
CODEX_BIN=$(command -v codex2 || command -v codex)   # skill-config.json imageGen で解決（codex2=別シートラッパー優先）
"$CODEX_BIN" features list | grep image_generation   # stable true を確認
"$CODEX_BIN" exec --sandbox workspace-write -c sandbox_workspace_write.network_access=true --cd "$PWD" "<組み立てたプロンプト>"
```

- Codex既定サンドボックスはネットワーク遮断のため、`-c sandbox_workspace_write.network_access=true` で
  ネットワークのみ許可する（image_gen 動作実証済み・2026-07-26）。この安全モードは **AI が直接実行できる**
- `--dangerously-bypass-approvals-and-sandbox` は**使わない**（Claude Code 分類器がハードブロック。
  安全モードで不足がある場合の最終手段としてのみ人間が手で実行する）。詳細は `references/image_gen.md` の
  AUTHORIZATION 節

### ⚠️ 保存先の罠（find-and-copy 必須）

image_gen は生成物を**プロンプトで指定したパスに保存しない**。既定で
`$CODEX_HOME/generated_images/<session-id>/*.png`（`CODEX_HOME`既定 `~/.codex`）に保存される。
実行直後に最新ファイルを探して目的パスへコピーする:

```bash
find "${CODEX_HOME:-$HOME/.codex}/generated_images" -type f -name '*.png' -print0 \
  | xargs -0 ls -t | head -1
# → 見つけたファイルを目的パスへ cp
```

### 透過アセット（クロマキー方式）

組み込みimage_genはネイティブ透過非対応。切り出しアセットは:

1. フラットな `#00ff00` 背景で生成（被写体が緑系なら `#ff00ff`）。「flat solid #00ff00 background,
   no gradients or shadows on the background, no text」を明示
2. 同梱ヘルパーでアルファ化:
   ```bash
   python3 "$CODEX_HOME/skills/.system/imagegen/scripts/remove_chroma_key.py" \
     --input <src.png> --out <final.png> --auto-key border --soft-matte --despill
   ```

白背景スライドに白背景アセットを載せる場合は透過不要（シーム最小化のため背景色を正確に一致させる）。
