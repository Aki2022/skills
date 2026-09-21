# Trouble status ledger

`entries/` の entry 本体は、発生時点の証拠として append-only に保つ。後から変わる
「トリアージ済みか」「どの対策に対応するか」「対策が効いたか」は、同じ保管ルートの
`triage/status.tsv` を唯一の可変な正典として管理する。entry 本体へ status を書き戻さない。

## Ledger schema

`status.tsv` は次の列を持つ TSV で、entry 名を一意キーにする。

| 列 | 意味 |
| --- | --- |
| `entry` | `entries/` 配下の basename |
| `triage_status` | `untriaged` または `triaged` |
| `response_status` | 下記の対策状態 |
| `response_ids` | `triage/responses.tsv` の `response_id`。複数は `;` 区切り |
| `last_triaged` | 最後に対象欄へ列挙したレポートの日付 |
| `status_updated_at` | 現在の status を更新した日。`last_triaged` とは別に管理する |
| `status_basis` | 状態を決めた根拠の短文。実測・レポート名・判定を分ける |
| `next_action` | 次回に行う評価。実装の再実行を書くのではなく、必要な測定を書く |

`responses.tsv` は対策単位の台帳で、`response_id`, `owner`, `state`, `implemented_at`,
`evidence`, `next_evaluation` を持つ。entry の `response_ids` と response registry の
ID は一致させる。ID は短い小文字 kebab-case にし、日付や一時的なセッション番号を入れない。

## Status meanings

`triage_status` と `response_status` は別の軸である。トリアージ済みでも、対策が効いたとは
限らない。

| `response_status` | 意味 | 次回の扱い |
| --- | --- | --- |
| `unknown` | 未トリアージ、または対策との関係が未確認 | 対象を読み、対応する response を決める |
| `legacy_unrecorded` | 過去レポートにはあるが、当時は status/evidence 欄が無かった履歴枠 | 現役 backlog に数えない。再発または新しい一次証拠が出た場合だけ response へ紐付ける |
| `none` | 対策を採らないと決めた | 決定理由と再発監視だけ残す |
| `proposed` | 候補を提示したが、人間承認または実装が未完了 | 承認待ち。実装済みとは報告しない |
| `implemented_unverified` | 対策の実装と局所テストは確認したが、運用上の有効性をまだ測っていない | 同型の発生率・見逃し・誤警告を測る。再実装しない |
| `effective_provisional` | 定義した観測期間で再発・見逃し・誤警告が確認されなかった | 次回も測定し、恒久的な成功とは断定しない |
| `recurred` | 対策が存在する状態で同型が再発した | 対策を再実行せず、coverage gap / 運用逸脱 / 別原因を切り分ける |
| `ineffective` | 対策を実行しても目的の失敗を防げなかった、または誤警告が許容外 | 変更案を再評価し、人間承認後に改訂する |
| `manual_only` | 自然文の真偽・意図・一次証拠の妥当性など、汎用機械化しないと決めた | review checklist と実測を使う |
| `repo_specific_pending` | リポジトリ固有の guide / issue へ委譲し、横断対策はしない | 所有リポジトリの issue 状態を確認する |

`effective_provisional` は測定済みの暫定状態であり、「もう起きない」や「対応完了」と
同義にしない。`recurred` は、発生時点で response が存在したことを根拠にできる場合だけ
使う。レポートに「既存 hook が稼働中」と書かれているだけで実行経路が確認できない場合は
`legacy_unrecorded` のままにする。

## State transitions

```text
new entry
  -> untriaged / unknown
  -> triaged / legacy_unrecorded       (過去レポートだけでバックフィルした場合)
  -> triaged / proposed                (候補を提示)
  -> triaged / implemented_unverified  (実装と局所テストを確認)
  -> triaged / effective_provisional    (観測期間の測定が通過)
                         \-> recurred / ineffective
```

`manual_only` と `repo_specific_pending` は、汎用 hook の再実装を避けるための終端的な
振り分けであり、必要なら根拠を更新する。状態遷移は既存 entry 本体へ追記せず、次回の
トリアージレポートの `## 過去対策の評価` と ledger の現在行に記録する。

## Backfill policy

過去の対象欄を使った backfill では、次の順序を守る。

1. 全 entry を列挙し、全過去レポートの `## 対象 entry 一覧` の和集合だけで
   `triage_status` を決める。本文中の参照名は使わない。
2. 過去レポートに列挙済みで、実装・有効性の一次証拠が無い行は
   `triaged / legacy_unrecorded` とする。`effective` や `implemented` を推測せず、
   一括した遡及評価の backlog にもしない。
3. 既存対策の稼働と検知漏れがレポートに明記された場合だけ `recurred` とし、該当する
   response ID と引用可能な根拠を付ける。
4. snapshot 後に増えた entry は `untriaged / unknown` とする。現在のレポートに入れず、
   次回対象へ残す。
5. バックフィルの判定日・件数・保守的な未確定理由は、新規のバックフィルレポートへ書く。
   過去レポートと entry 本体は書き換えない。

`status_ledger.py summary` は `legacy_unrecorded` を `historical_unlinked_entries` として
表示し、`evaluation_pending_entries` には `implemented_unverified` だけを数える。
過去の証拠本文は削除せず、存在しない evidence を後付けしない。

## Historical analysis index

`legacy_unrecorded` を横断分析するときは、status ledger を検索索引の代用にしない。
`triage/patterns.tsv` と `triage/entry-pattern-links.tsv` に分析結果を置き、実行時 snapshot と
判断は `triage/history/YYYY-MM-DD-legacy-mining.md` に新規作成する。entry 本体と既存レポートは
書き換えない。

`patterns.tsv` の列:

| 列 | 意味 |
| --- | --- |
| `pattern_id` | 小文字 kebab-case の安定ID |
| `title` | 人間向けの短い名称 |
| `failure_shape` | 観測可能な失敗形 |
| `detection_mode` | `syntax` / `runtime_state` / `manual_review` / `repo_specific` |
| `owner` | 形式化する場合の所有先 |
| `formalization_state` | `existing_response` / `candidate` / `manual_only` / `repo_specific` |
| `basis` | パターンを分離した根拠 |

`entry-pattern-links.tsv` の列:

| 列 | 意味 |
| --- | --- |
| `entry` | entry basename |
| `pattern_id` | 登録済み pattern ID |
| `response_id` | 候補 response。無ければ空欄 |
| `confidence` | `direct` / `mechanism` / `thematic` / `unclassified` |
| `basis` | その確度を選んだ根拠。`unclassified` では不足証拠 |
| `analyzed_at` | 分析日 |

`mechanism` と `thematic` は候補検索専用で、ledger の `response_ids` や status を更新しない。
status を変更できるのは同じ entry と response に対応する `direct` 行があり、状態固有の一次証拠も
説明できる場合だけである。索引へ載せたこと自体は、実装・再発・有効性の証拠ではない。
重複patternを統合する場合も、旧patternを削除して既存responseへ寄せるだけならstatus遷移ではない。
統合前後のlink集合と、退役pattern・新しいpattern定義は、別の日付付きhistory reportで追跡する。

## Next triage procedure

対象集合と status は別々に扱う。

1. `status_ledger.py sync` で新規 entry の不足行だけを `untriaged / unknown` として追加し、
   `status_ledger.py validate` で一対一対応を確認する。
2. 既存レポートの対象欄の和集合から新規対象を snapshot する。ledger の `triage_status` を
   上書きして「未読」に戻したり、同じ entry を実装作業の理由にしたりしない。
3. `response_status` が `implemented_unverified` または `effective_provisional` の entry は、
   過去文と response registry を参照して有効性だけを測る。同じ修正を再適用しない。
4. 同型が再発したら `recurred`、測定で防げなかったら `ineffective`、誤警告や見逃しを
   数えられない場合は `implemented_unverified` のままにする。
5. レポートに `## 過去対策の評価` を作り、少なくとも次を記録する。

   | response_id | 前回状態 | 今回の証拠 | 再発 | 見逃し/誤警告 | 次状態 | 次の測定 |
   | --- | --- | --- | ---: | ---: | --- | --- |
   | `hook-h1-pipeline-evidence` | implemented_unverified | 実行経路とテスト | 0 | 0 | effective_provisional | 次回も同型を数える |

   「実装した」「テストが通った」だけでは `effective_provisional` に進めない。
   `recurred` / `ineffective` は、同じ response を何度も作り直す代わりに、coverage・
   運用・仕様のどこを変えるかを決める入力にする。

## Commands

```bash
python3 ~/.agents/skills/own-trouble-log/scripts/status_ledger.py \
  --root "$ROOT" sync
python3 ~/.agents/skills/own-trouble-log/scripts/status_ledger.py \
  --root "$ROOT" summary
python3 ~/.agents/skills/own-trouble-log/scripts/status_ledger.py \
  --root "$ROOT" validate
python3 ~/.agents/skills/own-trouble-log/scripts/historical_index.py \
  --root "$ROOT" validate
python3 ~/.agents/skills/own-trouble-log/scripts/status_ledger.py \
  --root "$ROOT" update --entry 2026-08-28-example.md \
  --triage-status triaged --response-status implemented_unverified \
  --response-ids hook-h1-pipeline-evidence --last-triaged 2026-08-28 \
  --status-basis '局所回帰テストを確認、運用効果は未測定' \
  --next-action '次回トリアージで再発と誤警告を測定'
```

`sync` は既存行を上書きしない。status を変えるときは、先に response registry と
トリアージレポート側の根拠を整え、`update` 後に `validate` を通す。
