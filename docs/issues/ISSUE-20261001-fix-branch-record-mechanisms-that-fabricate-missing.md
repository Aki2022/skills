---
schema_version: 2
id: ISSUE-20261001-fix-branch-record-mechanisms-that-fabricate-missing
status: active
workstream: none
priority: high
due: none
created_at: 2026-10-01
updated_at: 2026-10-01
branch: ""
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "own-doc-update / own-git-clean の挙動修正であり、スキル本体の SKILL.md が正典。docs/guides は使わない運用（00_index.md の Guides 節の注記どおり）。SKILL.md の記述更新は不要（内部実装の修正で、契約文は変わらない）"
---

# create_issue.py の branch プレースホルダと front matter の # コメント剥がし漏れが MISSING を捏造する

## Goal

消費側リポジトリ（yorisoi_kaigo 等、own-doc-update / own-git-clean を正典として使う repo）で
2026-09-27 に実測した、**MISSING を製造・放置する 2 つのバグを直す**。

### (1) create_issue.py が起票するたびに MISSING を 1 件作る

`own-doc-update/scripts/create_issue.py:130`:

```python
f"branch: {issue_id}\npr: \"\"\n"
```

新規 issue の front matter に、**存在しない枝名**を無条件でプレースホルダとして書く。
`references/issue.template.md:8` の `branch: ISSUE-YYYYMMDD-short-slug` も同じ規約。

`own-git-clean/scripts/check_active_issue_branches.py` は「記録された branch が
local にも remote にも無い」を MISSING と呼ぶので、**起票 1 件 = MISSING 1 件**になる。

実測（yorisoi_kaigo、2026-09-27）: 本件を調査中に issue を 2 件起票しただけで
MISSING が 93 → 95 に増えた。同リポジトリの MISSING 95 件のうち **59 件**が
「`branch` == 自分の id」で、全てこのプレースホルダに由来する。

### (2) front matter パーサが `#` コメントを剥がさない

対象は 2 箇所、同型の実装:

- `own-git-clean/scripts/check_active_issue_branches.py:17-39`（`parse_front_matter`）
- `own-doc-update/scripts/validate_repo_docs.py` の同型パーサ

どちらも `key, _, val = line.partition(":")` の後 `val.strip().strip('"').strip("'")`
しかせず、`#` 以降のコメントを落とさない。

`branch: "" # ISSUE-…` のように**意図的に空にした上でコメントを添えた**行を書くと、
`branch` の値が `"" # ISSUE-…"` ではなく（`""` を剥がした後の）`# ISSUE-…` という
**非空文字列**になり、MISSING に数えられる。

この書き方自体が own-doc-update の推奨する回避策である
（`own-doc-update/scripts/create_issue.py` 等が branch 不明時に案内する「再開時に
新規で切る」運用は、消費側では `branch: "" # <旧枝名の記録>` の形で書かれることが多い）。

実測（yorisoi_kaigo）: 該当 2 件。うち 1 件は `docs/issues/ISSUE-20260803-issue-branch-record-drift.md`
——  **この MISSING 問題そのものを記録している issue 自身**が、自分の Next Actions が
処方した `branch: ""` 運用を適用した結果、偽 MISSING になっていた。
処方された治療が黙って失敗している。

## Acceptance

- verify: machine — 修正後、消費側リポジトリ（yorisoi_kaigo）で新規 1 件を create_issue.py で起票し、check_active_issue_branches.py の MISSING が増えないこと。branch: "" # <コメント> と書いた issue が MISSING に数えられないこと
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-10-01 — 起票のみ。未着手。

消費側 yorisoi_kaigo の WS-20260926-missing-issue-archive-sieve（archive 候補を
Jev で篩う試み）が 2026-09-27 に調査する過程で発見した。その WS は別の理由
（校正が原理的に成立しない）で不採用になったが、**本件 2 つのバグは独立に実在し、
その WS の成否と無関係に直す価値がある**。詳細は yorisoi_kaigo の
`docs/log/archive-sieve-rejection-20260927.md` と
`docs/issues/ISSUE-20260803-issue-branch-record-drift.md`。

<!-- Snapshot for the next session, overwritten each time: first line `as of 2026-10-01 — <state in one sentence>`. History goes to ## Log, not here. -->

## Next Actions

1. `create_issue.py:130` の `branch: {issue_id}` を `branch: ""` に変える。
2. `references/issue.template.md:8` の `branch: ISSUE-YYYYMMDD-short-slug` を
   `branch: ""` に変える（テンプレートがプレースホルダの書き方を教えてしまっている）。
3. `check_active_issue_branches.py:17-39`・`validate_repo_docs.py` の front matter
   パーサに `#` コメント剥がしを足す（`val.split("#", 1)[0]` 相当、ただし値が
   クォートで囲まれている場合はクォート内の `#` を剥がさないこと）。
4. 消費側リポジトリで、本 Acceptance の verify を実際に回して確認する。

<!-- First bullet = the very next command or step (or what unblocks a blocked issue). Required once work starts; the validator rejects an empty section. -->

## Guide Impact

- Decision: none
- Target or reason:

## Notes

本 issue の `branch:` は意図的に空にしてある。プレースホルダを書くと
bug (1) を自分で再現してしまうため。

## Log

<!-- Append-only, dated: `- 2026-10-01 — what changed / what was learned`. The only place chronology belongs. -->

## Completion

- [ ] Implementation completed or intentionally not needed
- [ ] Specs updated if direction or requirements changed
- [ ] Guide impact classified before implementation
- [ ] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [ ] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
