# Project contract

この参照は `own-project-update` が扱う vault 内の最小契約だよ。実際の現在仕様は
`obsidian_code/docs/specs/project-update.md` と recorder 側の project 連携 guide を優先する。

## Paths

```text
vault/project/project_<key>.md
vault/record/<meeting>.md
vault/setting/list/list_project.md
vault/setting/template/template_project.md
```

`project/<key>.md` の `<key>` は project frontmatter の `project` と一致する。

## Project note

最低限、次の frontmatter を持つ。値は実際の案件に合わせ、推測で埋めない。

```yaml
---
title: project_<key>
project: <key>
client: <client>
client_aliases: []
status: active
started: <YYYY-MM-DD>
last_updated: <timestamp>
tags:
  - project
---
```

標準見出しは `overview`、`people`、`stakeholders`、`next actions`、
`open issues / risks`、`minutes`、`related notes`、`documents`。既存ノートの見出し、
手書き領域、既存リンクを削除しない。

## Meeting record

project に属する record は YAML と本文の両方で紐付ける。`project` は**常にリスト**（1 件でもリスト）で、主／副の区別は持たない。
出所は `project_source`（`manual` / `model` / `legacy`）。手動の指定は model に勝つ。

```yaml
project:
  - <key>
project_source: manual
summary: <1 行要約>   # 任意。backfill の --summary は、record に summary が無いときだけ書く
```

frontmatter の直後、本文の先頭に、所属する project ごとに 1 行ずつ次を置く（`project` のリストと同じ project 集合）。

```markdown
> project: [project_<key>](../project/project_<key>.md)
```

- 読み手は旧来の文字列 `project: <key>` を 1 要素のリストとして読む。既に同じ key があれば何も書かない（旧形式のまま触らず、
  リスト化は `migrate` の仕事）。
- 書き手（`--mode backfill`）は既存の所属を**保持して追加**する。追加時は `project` をリストで書き直し、`project_source: manual`、
  backlink 行を project ごとに作り直す。他の frontmatter キー・本文・改行（CRLF）は保つ。
- `project` のリストと backlink の集合が食い違う（崩れた backlink・リストに無い project の backlink・backlink 欠落）record は、
  どちらが正しいかを推測せず停止する（backfill/bootstrap は CONFLICT、validate は ERROR）。
- 明示的な backfill で、既に `project_source: model` の key が載っている record は `manual` に昇格する（人の採用は model に勝つ。
  `legacy`・未記載は触らない）。
- `project` / `project_source` の行（リストの中を含む）に YAML コメントがあると、書き直しで消えるので CONFLICT で止まる（人間が直す）。
  引用符つきの `"project":` も同じキーとして置き換え、重複キーを作らない。書く前に書いた結果を読み直して検査する。
- `summary` は既存と違う値なら上書きせず停止する。1 行・ローカルパスなし。`--summary` は backfill だけ・`--record` が 1 件のときだけ
  （1 会議の要約を複数 record に書かない）。key が既にある record に summary だけを足す場合は、backlink ブロックの空行も整える。
- 食い違いの停止は backfill/bootstrap では CONFLICT（2）、validate では ERROR（3）。backlink の重複と「frontmatter 直後でない」は常に ERROR（3）。
- **書き手の注意**: 書き手はリスト形式で書く。recorder など旧形式しか読めない読み手が残る vault では、読み手が両形式に対応してから使う
  （recorder 側 WS の ISSUE-03）。

## Project list

`list_project.md` の `name` は project key と一意に一致し、`path` は project ノートを指す。
`client` は project ノート frontmatter の表示コピーであり、既存の異なる表記を黙って
上書きしない。差分を示し、確認後に合わせる。

## Document note

生成資料（最初の対象は `own-pptx-build` のプレゼンテーション）のテキストだけを vault に実体 Markdown として置く。
契約は SPEC-document-publish。`attach-document` が書く。

```text
vault/<kind>/<name>.md     # kind は置き場ディレクトリ名（project・setting・record は使えない）
vault/<kind>/<name>.<ext>  # 原本への相対 symlink（artifact link。任意・機械ローカル）
```

```yaml
---
title: <name>
kind: <kind>
project:            # 常にリスト（1 件でもリスト）。未紐付けなら project / project_source ごと省略
  - <key>
project_source: manual
date: <YYYY-MM-DD>
summary: <`--summary` の 1 行、無ければダイジェストの 1 行目>
source_repo: <repository name>
source_path: <repository-relative path>
source_hash: <sha256 of digest(任意) + outline + notes>
content_hash: <sha256 of the published body, backlinks excluded>
published_at: <timestamp>
---
```

本文は frontmatter 直後の backlink 行（project ごとに 1 行）、artifact link、`## digest`（`--digest-file` を渡したときだけ）、`## outline`、`## speaker notes`。
`content_hash` は backlink 行を除いた本文のハッシュで、人間が本文を編集したかの判定に使う。一致しなければ上書きせず停止する。
`source_hash` が同じなら本文は書き換えない。素材・title・hint の Unicode の行区切り（垂直タブ・改ページ・U+0085・U+2028/2029 等。PowerPoint の Shift+Enter は垂直タブ）は取り込み時に改行へ揃える（揃えないと再読込で行構造が変わり `content_hash` が合わず偽の CONFLICT になる）。絶対パス・`file://`・クラウド同期フォルダの実体パスは素材・title・hint・既存ノートの summary に入れない（検出したら何も書かず停止）。
この検査は入口の best-effort で、確定判定は vibe-guard の `scan-text`（pre-commit・夜間 vault_doctor）が持つ。止める対象は
「普通に起きる形」: `/Users/<name>`（大文字小文字を区別。小文字の `/users/` は REST ルートとして通す）・`/home/<name>`・`/Volumes/`・
`/mnt/<drive>/`・マウント接頭辞つきのホーム（`/System/Volumes/Data/Users/…` 等）・`~/…`（`~/.config` `~/.local` `~/.cache` 以外）・`~user/`・
`$HOME`・`%USERPROFILE%`・ドライブ文字つき/UNC/WSL の Windows パス・`file:`・`smb://`・クラウド同期フォルダ（`Library/CloudStorage`・
`GoogleDrive-<account>@`・`OneDrive - `・`iCloud Drive/`・`Dropbox (`・共有ドライブ等）。URL エンコード・HTML 実体・JSON エスケープ・
2 重エンコード・改行分断・重複スラッシュ・全角は正規化してから検査する。`/Users/{id}` のようなプレースホルダと `/Users/Shared` は通す。
安全側の誤拒否（`/home/app`、`/Users/123` など）は書き換えて回避する。
`--artifact` の拡張子を変えて再実行すると本文の link 行は 1 本に置き換わるが、旧拡張子の symlink は vault に残る（手で消す）。
`--artifact` が linked worktree 内なら main チェックアウトの同じ相対パスへ解決し、中身が一致しなければ（未マージ）link を作らない。
既存 note の改行（CRLF）は保つ。書き込み直前に対象ファイルが計画時から変わっていれば停止する（別セッションの書き込みを上書きしない）。
読み手は旧来の文字列 `project:` を 1 要素のリストとして読み、既存の所属は保持して追加する。

## Generated region (render)

project ノートの関連ノート表は `<!-- generated:associated-notes begin -->` と `<!-- generated:associated-notes end -->`
で囲んだ領域 1 つ。`render` が frontmatter の `project` に key を持つ vault 内ノート（`project/` と `setting/` を除く）を集め、
kind 昇順・日付降順の表（kind・日付・リンク・`summary`）に全体を書き直す。領域外の本文は 1 byte も変えない（CRLF も保つ）。BOM 付きの note も読む。閉じ `---` の無い frontmatter が `project:` を名乗る場合と、`project:` を名乗るのに YAML が読めない場合は停止する。
`attach-document` は紐付け先に加えて、まだその note を載せている project（手で外した所属）も同じ batch で render する。領域が無く関連ノートが
あれば末尾に新設する。マーカーが壊れている・`project:` を持つノートの YAML が読めない・summary にローカルパスがある場合は何も書かずに停止する。
存在しない project key を指すノートは stderr に警告する。

## Evidence and uncertainty

- 議事録の相対リンクを、project note の主張の出典として残す。
- Deep Research、推測、参加者名の補完などは `未検証` / `要確認` と明示する。
- 会議に内容がない場合は `minutes` 行だけを追加し、状況・課題・アクションを創作しない。

## Collision behavior

- bootstrap 先の project ノートが存在する場合は、ノート・一覧・record を変更せず停止。
- 一覧に同じ project key がある場合も重複登録せず停止。
- 対象 record が別 project を指す場合は、既存の所属を保持して追加する（旧: 停止）。
- 解除・再割当・明示上書きは別操作で、bootstrap の暗黙動作にしない。
