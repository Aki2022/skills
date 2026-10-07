---
schema_version: 2
id: ISSUE-20261005-branch-checks-scoped-to-in-progress
status: archived
workstream: none
priority: medium
due: none
created_at: 2026-10-05
updated_at: 2026-10-05
branch: ""
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "own-doc-update / own-git-clean の挙動修正で、SKILL.md が正典。docs/guides は使わない運用。SKILL.md には in_progress の規約を 1 行足した"
---

# branch の 2 つの検査が未着手 issue で同時に黙らず、MISSING が母集団の全件を指していた

## Goal

`branch:` に関する 2 つの検査を、**作業中（`status: in_progress`）の issue だけ**を対象にして、
未着手・完了の issue で両方が黙る状態にする。

### 何が起きていたか

yorisoi_kaigo（消費側）で、2 つの検査が互いに矛盾し、どちらを満たしても片方が鳴った
（消費側の `ISSUE-20260903-improve-loop-branch-record-checks-disagree`）:

| `branch:` の値 | `check_active_issue_branches.py` | `validate_repo_docs.py` |
| --- | --- | --- |
| 空 | 黙る | `missing branch` 警告 |
| 存在しない枝名 | `MISSING` | 黙る |

さらに、枝は merge 後に削除する運用なので、**枝を記録した issue は完了・保留を含めて全件が MISSING**
になり、母集団の 100% を指す検査は何も指していなかった（消費側 2026-10-02: OK 0 件 / MISSING 95 件）。

### 変更

- `validate_repo_docs.py`: `missing branch` 警告を **`status: in_progress` の issue だけ**に出す
- `check_active_issue_branches.py`: `MISSING` を **`in_progress` の issue だけ**に出す。
  それ以外で枝が無いものは `(not checked: N issue(s) …)` として**件数を出力に残す**（黙って消さない）
- `own-doc-update/SKILL.md`: 作業が始まったら `status: in_progress` にする規約を 1 行足した

**枝の集合（孤児候補の判定・重複の検出）は全 status から作るまま**にした。MISSING の判定だけ絞る。
枝の集合まで絞ると、`active` の issue が記録している**生きた枝が「孤児候補」と報告**される
（テストで固定。対象を壊して赤を確認した）。

## Acceptance

- verify: machine — 消費側 yorisoi_kaigo の docs に修正版を当て、MISSING が in_progress の issue だけになり、missing branch 警告が in_progress の枝無しだけになること。生きた枝を孤児候補と誤判定しないこと
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-10-05 — 実装・検証済み。消費側で MISSING が 30 → 3、`missing branch` 警告が 92 → 1 になった。

**消費側 yorisoi_kaigo（`origin/main` の同じ docs）で、原版と修正版を当てて実測した:**

| 指標 | 変更前 | 変更後 |
| --- | --- | --- |
| MISSING | 30 | **3** |
| OK（枝が実在） | 6 | 6（不変） |
| `missing branch` 警告 | 92 | **1** |

残った MISSING の 3 件（`lp-deploy-backend-adoption` / `slack-approval-bypasses-apply-gate` /
`partial-file-set-promote`）は、いずれも `in_progress` で枝が消えている。同日に一次証拠で
「本当に未了」と判定した issue と一致し、**検査の本来の目的（作業中のはずの枝が消えた異常）どおり**。
警告の 1 件も `in_progress` で枝が無いもの。

**検証:** 新規 9 テスト（own-doc-update 328+10 → 347 件）。赤を先に確認し、`in_progress` の絞り込みを
外すと 3 件、枝の集合を絞ると孤児テストが赤になることも確認した。既存 338 件は 1 件も落ちていない。

**限界:**

- 作業中なのに `in_progress` へ更新し忘れた issue（`active` のまま）は、枝が無くても**警告されない**。
  `active` は事実上「未着手〜作業中の既定値」で、消費側は 111 件ある。この限界は運営者に事実を示して
  承認を得た（警告 94 → 1 に減る代わりに、更新し忘れは検出されない）
- **滞留（長く動いていない issue）の指標は入れていない。** `updated_at` は一括処理が書き換えるため
  使えない（消費側で経過日数が 2 日・11 日の 2 点に 129 件中 111 件が集中した）。
  消費側の `ISSUE-20260927-improve-loop-archive-sieve-age-gate-never-fires` が扱う

<!-- Snapshot for the next session, overwritten each time: first line `as of 2026-10-05 — <state in one sentence>`. History goes to ## Log, not here. -->

## Next Actions

1. なし（実装・検証済み）

<!-- First bullet = the very next command or step (or what unblocks a blocked issue). Required once work starts; the validator rejects an empty section. -->

## Guide Impact

- Decision: none
- Target or reason:

## Notes

## Log

<!-- Append-only, dated: `- 2026-10-05 — what changed / what was learned`. The only place chronology belongs. -->

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
