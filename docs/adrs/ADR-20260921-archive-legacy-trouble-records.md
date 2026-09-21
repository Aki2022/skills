---
id: ADR-20260921-archive-legacy-trouble-records
status: accepted
scope: development
created_at: 2026-09-21
updated_at: 2026-09-21
source_workstreams: []
source_issues: [ISSUE-20260921-archive-legacy-trouble-records]
related_specs: []
related_guides: []
supersedes: ""
superseded_by: ""
---

# Decision: 履歴未紐付けを現役評価負債から分離する

## Context

`own-trouble-log` の最初の2回のトリアージからbackfillした226件には、response IDや
実装・有効性の一次証拠が無い。保守的に `legacy_unrecorded` とした点は正しいが、集計が
これを `implemented_unverified` 30件と合算し、256件すべてを現役の評価待ちとして表示していた。
過去に存在しなかった証拠は作れず、全件評価を義務にすると恒久的な負債になる一方、entry本文を
削除すると再発照合と監査の材料を失う。

## Decision

`legacy_unrecorded` は監査用の履歴枠として保存し、現役 backlog には数えない。
`status_ledger.py summary` は履歴を `historical_unlinked_entries`、実装後の現役評価待ちを
`evaluation_pending_entries` として分け、後者には `implemented_unverified` だけを数える。
履歴entryは同型の再発または新しい一次証拠が得られた場合だけresponseへ紐付け、別statusへ進める。
既存226行の `next_action` もこの方針へ合わせる。

## Alternatives Considered

- 226件を遡及して評価する: 当時存在しなかった一次証拠を推測または捏造するため採らない。
- 226件のentryを削除する: 監査証跡と再発時の比較材料を失うため採らない。
- 従来どおり評価待ちへ合算する: 解消条件のないbacklogを維持し、現役30件を埋没させるため採らない。

## Consequences

### Positive

- 現在行動すべき評価件数が30件として見える。
- append-onlyの証拠本文と過去レポートを維持できる。
- 新しい証拠が出た履歴だけを需要駆動で再開できる。

### Negative or Follow-up

- 履歴226件は `response_id` 未紐付けのまま残るため、再発時には対応する形を人間またはトリアージが判断する。
- `historical_unlinked_entries` はゼロを目標とするKPIではなく、監査用件数として解釈する必要がある。

## Links

- Issue: `../issues/archive/ISSUE-20260921-archive-legacy-trouble-records.md`
- Skill: `../../own-trouble-log/SKILL.md`
- Status reference: `../../own-trouble-log/references/status.md`
