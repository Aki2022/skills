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
`open issues / risks`。旧 `minutes`・`related notes`・`documents` は migrate で generated 領域へ置き換わる（移行前のノートには残っていてよい）。既存ノートの見出し、
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
存在しない project key を指すノートは stderr に警告する。UTF-8 として読めないノート（実 vault には途中で切れたマルチバイトを含む clip がある）は、project を名乗らなければ警告して飛ばし、
名乗る可能性がある（frontmatter が読めて `project:` を持つ）場合は止まる。

## Generated outputs: list_project.md and the Raycast dropdown (render)

`render` は上の generated 領域に加えて、同じ入力（project ノートの frontmatter と各ノートの association）から次の 2 つを全体生成する。
3 出力は全部計画してから 1 batch で書き、どれか 1 つでも作れなければ何も書かない（書き込み中の I/O 失敗は全部元に戻す）。

- `setting/list/list_project.md`: 列は `name | client | partner | status | last meeting | path`。`client`・`partner`（文字列またはリスト。リストは `, ` 連結）・
  `status` は project ノートの frontmatter、`last meeting` は**record だけ**の最新日付（資料の公開では更新しない）、`path` は project ノートへの相対リンク。
  行は active（`status` が無い場合も active 扱い）が先、最終会議日の新しい順、名前順。`setting/list` が無ければ停止する（勝手に作らない）。
  **情報を黙って落とさない（fail-closed）**: 既存の一覧に、frontmatter に無い `client`・`status`・`partner`（`,` と `、` 区切りの集合として比較）、
  より新しい／日付として再現できない `last meeting`、表に無い列（手書きのメモ列など）、project ノートの無い行があるとき、また空でない既存ファイルが
  render の読み戻せる表でないときは CONFLICT で停止する（先に `migrate` で frontmatter へ移す。意図して落とすときだけ `--accept-list-changes`）。
  `|` を含む key・client・partner は表に入れられないので ERROR で停止する。UTF-8 でない・symlink の一覧も ERROR。
- Raycast 一覧（`--raycast-script <path>`）: そのスクリプトにちょうど 1 行ある `# @raycast.argument2 {...}` を、active な project の
  `{"title": "<key> (<client>)" または "<key>", "value": "<key>"}` に置き換える（recorder の `render_raycast_dropdown_line` と同じ形式。title は client が key と同じなら key だけ）。
  管理行が 0 本・2 本以上、active が 0 件、スクリプトが symlink（実体のパスを渡す）のときは停止する。実行ビットと CRLF を保つ。`--raycast-script` が無いときは `RAYCAST skipped` と報告する。

`--project-key` は generated 領域の対象だけを絞る。list_project.md と Raycast 一覧は常に vault 全体から作るので、他の project ノートが壊れていれば `--project-key` でも停止する。
書き込み中の失敗は、読み取り専用のファイルも含めて元のバイトに戻す。

## Migration (migrate)

既存 vault を新契約へ 1 回で移す。夜間 job では行わない。`migrate` は dry-run が既定で、`--apply` は dry-run が出した `PLAN_ID` を
`--plan-id` で渡したときだけ書く（人間が見た差分と書く内容を一致させる。vault が変わっていれば CONFLICT）。`--diff-dir <path>`（vault の外。中や、vault を含む場所、前回の migrate dry-run が作ったものでない空でないディレクトリは拒否。前回の dry-run が書いた `.diff` だけを消す）に
ファイルごとの unified diff（改行だけの変更も出る）と `SUMMARY.md` を出す。`--plan-id` は `--apply` と一緒のときだけ使える。変更が無ければ `NOOP`（再実行も no-op）。書く batch は全部元に戻せる（render と同じ）。

- **record**（vault 内の `.md` で `project` を持つもの。`record/` 直下に限らず、`presentation/` や `record/` の下位ディレクトリも）: `project` を常にリストにし（出所があっても、スカラーなら書き換える）、出所が無ければ `project_source: legacy` を付ける。project ノートの旧 3 節
  （`## minutes` の表・`## documents` の vault 内リンク）に載っている record・資料は、その project に属する証拠として `legacy` で紐付ける（既存の所属は保持して追加）。
  minutes の topics は、record に summary が無いときだけ `summary` に移す（既存の summary は上書きせず `SUMMARY_KEPT` と報告）。行（minutes の行・documents の項目）は、`[ラベル](リンク)` だけの純粋な形（リンクに `#anchor`・`?query` が無い）で、ラベルが record のファイル名または title と同じで、topics が summary になったときだけ消える。
  minutes の表は見出しの名前（`date`・`minutes`・`topics`）で列を読む（列順が違っても可）。見出しが想定外の表は、見出しも行も原文のまま残し紐付けしない。
  それ以外（summary にならなかった topics・topics 以外の列・表の日付が record と違う・リンクの前後に文字や title 属性がある・リンクが複数・ラベルが違う・項目に字下げした子や続きの行が付く・documents の表の行）は、リンクが解決できる record は紐付けたうえで、
  情報を落とさないよう**行ごと** `## migrated notes` に残し `ROW_KEPT` と報告する（表の見出し行・区切り行も、その表の行が残るときは付ける）。
  backlink は所属 project ごとに 1 行・本文の先頭へ正規形で 1 本化する（本文の途中にあるもの・`project\_<key>` のようにエスケープされたものも取り込む）。
  frontmatter の `project` と backlink が食い違う record（旧運用では backlink だけの record がある）は和集合を `legacy` で補修し、`REPAIRED_RECORD` として必ず報告する（backlink が足りないだけなら `ADDED_BACKLINK`）。
  崩れた backlink・YAML コメントのある `project` 行は止まる。
- **related notes**: `## related notes` の項目は「言及」であって所属ではないので、**既定では紐付けず原文のまま残す**（紐付けると言及しただけの record の日付が
  last meeting に入る）。`--related-notes associate` を選んだときだけ `legacy` で紐付ける。
- **project ノート**: 旧 3 節を取り除く。表現できない内容（引用メモ・外部リンク・related notes の項目・実体や frontmatter の無い record の行）は消さずに
  `## migrated notes`（`### from <旧節>`）へ原文のまま移す（空行・コードブロックも原文のまま。既に `## migrated notes` があればその末尾に足す）。旧節の見出しは
  `## minutes (2026)`・`### minutes`・大文字小文字違い・閉じ `##`・3 桁までのインデントも旧節（別の旧節の中に入れ子のものも、自分の題で分類する）。`# 見出し`（H1）は節を終わらせる。
  空リンクだけの行は、label が template の旧 3 節の空リンクと同じとき（`[local]()` など）と、中身のない `-` だけ placeholder として捨てる（template に無い label の空リンク・空リンクに他の文字やリンクが付いた行は残す）。見出しが想定外の minutes の表は原文のまま残し、`UNRECOGNIZED_TABLE` と報告する。HTML コメントは、template の旧 3 節にあるものと
  同じ文面のときだけ placeholder として捨て、それ以外（人が書いたもの）は残す。旧 3 節に既にある generated 領域は render が作り直す（begin/end が片方だけなら止まる）。
  リンクの `#anchor`・`?query`・`<…>` 形・ファイル名の `(1)` は解決する。解決できない行（`.trash` など隠しディレクトリへのリンクを含む）は `UNLINKED_ROW` として残す。コードブロックの中の backlink 風の行は所属ではなく、そのまま残す。record の `project`（と紐付ける project）に対応する project ノートが無い note は書き換えず `UNKNOWN_PROJECT` と報告し、書き換える対象（project ノートがある key）が backlink に書けない key（`[]()|/` など）なら dry-run で止める。
  generated 領域は render が作る。`list_project.md` にだけある `partner`・`status` は frontmatter へ移す（key が無い、または空のとき）。frontmatter に別の値があって食い違うときは止まり、`--accept-list-changes` で list 側の値を捨てると決めたときだけ進み、捨てる値は dry-run に `LIST_VALUE_DROPPED` として列挙する。BOM 付きの project ノートは止まる（BOM を外してから）。旧 3 節は `validate` の必須見出しではなくなった。
- **template_project.md**: 旧 3 節にコメント・placeholder 以外の内容（注意書きの本文等）があるときは自動では書き換えず `TEMPLATE_MANUAL` と報告する（人間が置き場所を決める）。
  コメントだけなら旧 3 節を外し、`partner:`・`scope:` と空の generated 領域を足す。捨てるコメントは `TEMPLATE_COMMENTS_DROPPED` として dry-run に出す。
- **render**: 上の結果（generated 領域・`list_project.md`・`--raycast-script`）を同じ batch に含める。適用後に `render` は no-op になる。
- **書かないもの**: frontmatter の無い note には frontmatter を作らない（行は `## migrated notes` に残す）。project を名乗らない UTF-8 でない note は警告して飛ばす（project を名乗るものは止まる）。

## Evidence and uncertainty

- 議事録の相対リンクを、project note の主張の出典として残す。
- Deep Research、推測、参加者名の補完などは `未検証` / `要確認` と明示する。
- 会議に内容がない場合は `minutes` 行だけを追加し、状況・課題・アクションを創作しない。

## Collision behavior

- bootstrap 先の project ノートが存在する場合は、ノート・一覧・record を変更せず停止。
- 一覧に同じ project key がある場合も重複登録せず停止。
- 対象 record が別 project を指す場合は、既存の所属を保持して追加する（旧: 停止）。
- 解除・再割当・明示上書きは別操作で、bootstrap の暗黙動作にしない。
