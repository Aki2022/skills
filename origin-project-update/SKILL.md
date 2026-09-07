---
name: origin-project-update
description: >-
  Maintain and bootstrap Obsidian project notes: safely sync explicitly selected meeting records,
  create a new project note from related records, backfill project YAML/backlinks, and audit project
  list consistency. Use for project-note organization or repair; recorder's automatic existing-project
  pipeline remains the default for newly recorded meetings.
---

# Origin Project Update

Obsidian の `vault/project/`、`vault/record/`、`vault/setting/list/list_project.md` を
整合させる skill だよ。新規 project の初期化、過去議事録の backfill、project/list の
監査・手動再同期で使う。録音後に既存 project を更新する recorder の自動経路は置き換えない。

## 使い分け

### existing-project sync（自動）

録音開始時に recorder が `vault_project` を解決した会議は、recorder が次を行う正規経路だよ。

- 議事録 YAML に `project: <key>` を入れる
- frontmatter 直後に record backlink を入れる
- publish 後に既存 project ノートを更新する
- `list_project.md` の最終会議日を決定論的に更新する

この skill で同じ LLM 更新を再実行して二重反映させない。欠落や失敗を調べるときは
監査・検証として扱い、`session.project`（Slack/Notion 用）と `session.vault_project`
（vault 紐付け用）を混同しない。

### project bootstrap（任意）

新しい project を過去議事録から作るときの手順だよ。project key と対象 record は
ユーザーが明示し、候補検索や LLM の順位付けは補助にとどめる。対象 record を無断で
増やしたり、client 名・aliases・ファイル名だけで紐付けを確定しない。

### backfill / audit（任意）

既存 project に選択した record の YAML/backlink を付与するときは backfill、project・
record・一覧の契約を確認するだけなら audit として扱う。いずれも対象パスを明示する。

## 正典と形式

詳細なフィールド、見出し、backlink の形式は [references/project-contract.md](references/project-contract.md)
を読む。要点は次のとおりだよ。

- project key の正典は `vault/project/project_<key>.md` の frontmatter `project`
- project ノートの frontmatter `client` が正典で、`list_project.md` は表示用の写像
- record は YAML `project` と、次の backlink を両方持つ
  `> project: [project_<key>](../project/project_<key>.md)`
- `client_aliases` は候補発見用で、曖昧な候補を自動確定しない

## 共通ワークフロー

1. 対象 vault を確定する。対象が曖昧なら書き込まず、候補を示して確認する。
2. モードを選ぶ。既存録音の自動経路なら recorder に任せ、bootstrap/backfill/audit なら
   project key と record パスを明示的に受け取る。
3. すべての対象を先に preflight する。project ノート、一覧、record の衝突を検査し、
   一つでも不合格なら全ファイルを書き込まず停止する。
4. semantic な作業だけを LLM に依頼する。選択済み record から事実・出典・未検証情報を
   抽出し、project 各節の下書き、用語候補、矛盾、変更レビューを作らせる。
5. dry-run と差分を提示し、人間確認後に機械的な frontmatter/backlink/list 更新を行う。
6. 更新後に machine-verifiable checks と human review を実行し、実物の出力と exit code を
   確認してから完了と報告する。

## bootstrap の手順

- `template_project.md` を基に、project key、client、status、started、tags などの
  frontmatter を決める。client や参加者を推測で確定しない。
- LLM にはユーザーが選択した record だけを渡し、`overview`、`people`、`stakeholders`、
  `next actions`、`open issues / risks`、`minutes` などの草案を作らせる。各主張に元 record
  の相対リンクを付け、未検証の情報を明示する。
- project ノートが既に存在する場合は、上書き・自動マージせず事前検査で停止する。既存
  ノート、record、一覧、入力内容は変更しない。
- `list_project.md` に既存の project 行がある場合も重複作成せず停止する。client 差分は
  project ノートを正典として示し、確認を得てから一覧を合わせる。
- 対象 record に別 project の YAML または backlink があれば、既存紐付けを保持して停止する。
  解除・再割当・明示上書きは別操作であり、bootstrap に含めない。
- 内容のない会議は `minutes` へのリンク追加に限定し、overview / next actions / risks に
  空振りの「確認する」項目を生成しない。

## backfill と決定論的ヘルパー

同梱の `scripts/project_update.py` は YAML/backlink の preflight、検証、record backfill を
担当する。PyYAML が無い場合は手作業に切り替えず停止する。

```text
python3 scripts/project_update.py --vault <vault-root> --mode bootstrap \
  --project-key <key> --record record/<meeting>.md

python3 scripts/project_update.py --vault <vault-root> --mode backfill \
  --project-key <key> --record record/<meeting>.md
# dry-run が既定。差分確認後だけ --apply を追加する。

python3 scripts/project_update.py --vault <vault-root> --mode validate \
  --project-key <key> --record record/<meeting>.md
```

`bootstrap` は新規 project ノートが無いことを確認するだけで、ノート本文は生成しない。
LLM 草案と人間確認を終えて project ノートを作成した後、`backfill` と `validate` を実行する。
ヘルパーは preflight を全件終えてから書き込み、書き込み失敗時は変更済み record を復元する。

## 必須の検査

- YAML parse 成功、指定 key との一致、backlink の相対リンクと重複なし
- project ノートの必須見出し、`client`、project/list の重複なし
- 既存本文・手動メモ・既存 frontmatter の保持
- ローカル絶対パスの追加なし、明示した対象以外のファイル変更なし
- 衝突・事前検査エラー時の部分適用なし
- 同じ入力の再実行が no-op
- `git diff --check` が成功し、変更ファイルが想定リストだけであること

LLM の本文は機械的な文字列一致だけで合格にしない。人間が事実性、出典、未検証表示、
既存内容との整合性、client 差分をレビューする。検査を実行していない範囲は完了報告に
明記する。

## 安全境界

- project key と対象 record を明示しない再帰的な書き込みをしない。
- raw な `record/<meeting_id>/metadata` の `session.project` を vault project の代わりに
  書き換えない。
- 既存ノートや record の削除、別 project への自動移管、外部送信、秘密情報の追加をしない。
- vault 内リンクは相対パスにし、ローカル絶対パス、`file://`、クラウド同期フォルダの実体パスを
  project ノートや record に書かない。
- recorder の実装や prompt を変更する必要が出たら、この skill の範囲を越えるため、変更対象を
  明示して別途承認を得る。
