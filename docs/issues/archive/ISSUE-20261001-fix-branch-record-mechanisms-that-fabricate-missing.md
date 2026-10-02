---
schema_version: 2
id: ISSUE-20261001-fix-branch-record-mechanisms-that-fabricate-missing
status: archived
workstream: none
priority: high
due: none
created_at: 2026-10-01
updated_at: 2026-10-02
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

as of 2026-10-02 — 実装・検証済み（commit 8f30056）。issue の記述と実装で違った点がある。

**issue の記述との差分（実装して初めて分かったこと）:**

1. **主経路の原因はテンプレート側だった。** 本 issue は `create_issue.py:130` を原因と書いたが、
   そこはテンプレートが読めない時のフォールバック。主経路は
   `content.replace("ISSUE-YYYYMMDD-short-slug", issue_id)` で、`issue.template.md:8` の
   `branch:` もこの置換で実 ID になる。130 行だけ直しても主経路は直らなかった。
2. **`workstream.template.md` の `branch: WS-YYYYMMDD-short-slug` も同型**なので直した
   （範囲外だが同じ原因。消費側 WS-20260926 が `branch_note` で自力回避していた罠）。
3. **パーサの不具合は issue の記述より重かった。** `branch: "feat/x" # memo` は
   閉じクォートが残って `feat/x" # memo` になっていた（空値だけの話ではなかった）。

**やったこと:** `strip_inline_comment()` を `validate_repo_docs.py` と
`check_active_issue_branches.py` に置いた。YAML 仕様どおり、`#` の前に空白がある時だけ
コメントで、クォート内は対象外（`feat/x#frag`・`"has # inside"` は保つ）。
`own-git-clean` は `own-doc-update` に依存しない独立 skill なので import せず同じ関数を
置き、`StripInlineCommentMirrorTest` でソース一致を機械的に縛った。

**検証（Acceptance の verify を消費側で実測）:**

| 検査 | 結果 |
| --- | --- |
| 赤を先に確認 | 新規 10 テスト中 6 が落ち 14 件失敗 |
| 既存テスト | own-doc-update 328 → 338 件（差分は追加分ちょうど）全緑、own-git-clean 4 件緑 |
| yorisoi_kaigo の MISSING | 93 → 91。消えたのは `ISSUE-20260803-issue-branch-record-drift` と `ISSUE-20260803-core-capability-request-pr`（事前に特定した 2 件）、増加 0 |
| 対照実験（同じ docs に 1 件起票） | 原版の生成器 98 → **99**（+1）、修正後 98 → **98**（±0） |

**未対応（本 issue の範囲外）:** 消費側の既存 `branch == 自分の id` 59 件の一括
`branch: ""` 化は各 repo の作業。

## Next Actions

1. 消費側リポジトリで `branch == 自分の id` かつ実在しない枝の issue を `branch: ""` にする
   （yorisoi_kaigo は `ISSUE-20260803-issue-branch-record-drift` の Next Actions 2）。

## Guide Impact

- Decision: none
- Target or reason:

## Notes

本 issue の `branch:` は意図的に空にしてある。プレースホルダを書くと
bug (1) を自分で再現してしまうため。

## Log

<!-- Append-only, dated: `- 2026-10-01 — what changed / what was learned`. The only place chronology belongs. -->

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
