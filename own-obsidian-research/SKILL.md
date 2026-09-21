---
name: own-obsidian-research
description: >-
  Obsidian vault（手元の論文・記事クリップ・スキャン資料のノート群）を vector 検索し、
  指定テーマの調査結果を出典リンク付きで整理して返す。発火するのは情報源が「手元の vault」だと
  明示された時だけ: 「obsidian で調査／調べて」「vault を検索」「vector で調べて」「ノート／蓄積資料／
  保存した論文から調べて」など。これらに続けてテーマや観点（医療と公共 等）が指定されたら使う。
  発火しない（重要）: 単に「〇〇について調べて／まとめて」とだけ言われた一般的な調査依頼、最新情報や
  事実確認を求める依頼（Web 検索が適切）、リポジトリのコード調査。obsidian/vault/ノート等の語が
  無ければ起動せず、必要なら「vault を検索する？」と確認する。
---

# Obsidian Research

vault を vector 検索して、テーマ・観点で整理し出典リンク付きで返す。

## 手順

1. `vector.sh` のパスを取得する（絶対パスは秘匿情報。応答やログに出さない）:

   ```bash
   CODE="$(cat "<skill-dir>/code_path.local" 2>/dev/null || echo "$HOME/code/obsidian_code")"
   ```

   **`vector.sh` はコード側リポジトリ `obsidian_code` にある。** vault（データ側）の
   `script/` は 2 リポジトリ分離で空になっており、そこを叩くと
   `No such file or directory` で空振りする。データ root は `vector.sh` が
   `.env` の `OBSIDIAN_DATA_ROOT` から自力で解決するため、vault のパスは渡さなくてよい。

   取れない場合は `script/vector.sh` を含むリポジトリの root をユーザーに聞き、
   `<skill-dir>/code_path.local` に1行で書く（gitignore 済み）。

2. 検索する（出典リンクを返すので `--dedupe --abs-links` を常に付ける）:

   ```bash
   bash "$CODE/script/vector.sh" search "大腸がん 診療 課題" --dedupe --abs-links
   ```

   観点が複数で1クエリに収まらなければ複数回呼ぶ。絞るなら `--limit N`。

3. stdout の JSON `results[]` を読む。各要素は
   `title` / `relative_path` / `category` / `score`（高いほど近い）/ `preview` / `abs_link`。

4. テーマと観点で構造化して答える。観点ごとに見出しを立て、各論点に根拠ノートを添える。
   出典リンクは `abs_link` をそのまま使う（エンコード済み。組み立て直さない）:

   ```markdown
   → [ノートのわかりやすい名前](abs_link)
   ```

   関連の薄いヒット（score 低・話題ずれ）は含めない。

## 制約（必要に応じてユーザーに伝える）

- ノート内の画像はローカル参照のため表示されない（本文テキストは読める）。
- リンクは mac ではクリックで開けるが、スマホからは開けない。スマホで中身を読みたい時は
  「このノートを読んで」と頼めばセッションがファイルを読んで返す。
- 索引が古く最新ノートが出ない時は `bash "$CODE/script/vector.sh" update` を案内する。
- vault はクラウド同期ストレージ上にあるため、ヒットしたノートの本文読み込みが
  `ETIMEDOUT` / `Operation timed out` で失敗することがある（`ls` は通るのに read だけ落ちる）。
  数回のリトライで復帰しなければ、preview だけで答えを作り、**本文を読めなかったノートを
  明示して報告する**。読めた前提で書かない。
