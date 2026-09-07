# Project contract

この参照は `origin-project-update` が扱う vault 内の最小契約だよ。実際の現在仕様は
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

project に属する record は YAML と本文の両方で紐付ける。

```yaml
project: <key>
```

frontmatter の直後、本文の先頭に次を一つだけ置く。

```markdown
> project: [project_<key>](../project/project_<key>.md)
```

既に別 project の YAML または backlink がある record は移管せず停止する。既に同じ key が
ある場合は重複させず no-op にする。

## Project list

`list_project.md` の `name` は project key と一意に一致し、`path` は project ノートを指す。
`client` は project ノート frontmatter の表示コピーであり、既存の異なる表記を黙って
上書きしない。差分を示し、確認後に合わせる。

## Evidence and uncertainty

- 議事録の相対リンクを、project note の主張の出典として残す。
- Deep Research、推測、参加者名の補完などは `未検証` / `要確認` と明示する。
- 会議に内容がない場合は `minutes` 行だけを追加し、状況・課題・アクションを創作しない。

## Collision behavior

- bootstrap 先の project ノートが存在する場合は、ノート・一覧・record を変更せず停止。
- 一覧に同じ project key がある場合も重複登録せず停止。
- 対象 record が別 project を指す場合は既存紐付けを保持して停止。
- 解除・再割当・明示上書きは別操作で、bootstrap の暗黙動作にしない。
