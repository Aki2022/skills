---
schema_version: 2
id: WS-20261108-third-party-skill-mirror-sync
status: active
priority: medium
due: none
created_at: 2026-11-08
updated_at: 2026-11-08
branch: "WS-20261108-third-party-skill-mirror-sync"
pr: ""
human_boundary_confirmed_at: 2026-11-08
next_human_gate: wip 枝の削除と、各 mirror の同期実行の確認
related_specs: [docs/specs/third-party-skill-mirror-update.md]
related_guides: []
---

# 第三者 skill を plugin から正典へ同期し、全 AI で共有する

## Goal

どの AI からも同じ第三者 skill を同じ版で使えるようにする。plugin の導入版を正典へ同期する道具と、ずれを知らせる検査を main に入れ、wip 枝を片付ける。

## Success Criteria

- `mirror_skill.py --update-existing` と `check_mirrors.py` の drift 検査が main に入り、テストとゲートが緑。
- wip/stash-mirror-skill-20261001 が所有者の確認後に削除されている。

## Authorization Envelope

- Approved scope: 第三者 skill は Claude Code の plugin で取得・固定・更新し、その版を正典へ同期して全エージェントへ配る。自作は plugin の版から正典への同期だけ。wip/stash-mirror-skill-20261001 の --update-existing を部品として移植し、移植後に wip 枝を削除する。既存 mirror の取り直しの実行は含めない（同期の実行は所有者の確認ごと）。spec: docs/specs/third-party-skill-mirror-update.md
- Autonomous actions allowed: main から新しい枝を切る、テスト先行の実装、SKILL.md の更新、ローカル検査、ws の更新、作業枝の push と PR、ゲート通過後の merge
- Confirm first: ネットワーク取得、plugin の導入と更新、正典の実ファイルの置き換え（同期の実行）、wip 枝の削除
- Merge policy: ローカルのゲート（mise run verify と validate_repo_docs.py）通過で自動 merge（CI なし）
- Cost or usage ceiling: 従量課金なし
- Out of scope: 上流から直接取る更新の自作、自動同期、既存 mirror の同期の実行

## Human Gates

- Start gate: confirmed on 2026-11-08
- Next gate: wip 枝の削除と、各 mirror の同期実行の確認
  <!-- a named human checkpoint, or `none-workstream-complete` when nothing gates completion -->
- Stop conditions: ゲート（skill_lint を作業場所に向けた実行、validate_repo_docs.py）が赤

## Issue Queue

| Issue               | Status  | Depends on | Outcome            |
| ------------------- | ------- | ---------- | ------------------ |
| ISSUE-01-mirror-sync-spec-approval | complete | none       | spec を active にした（所有者の「実装して完成させましょう」、2026-11-08） |

### ISSUE-01-mirror-sync-spec-approval

- status: complete
- depends_on: []
- runnability: ready
  <!-- `ready`, or `gated on <the human decision / missing input>` — executors treat a missing value as gated and stop -->
- guide_impact: none
- related_guides: []
- guide_impact_reason: "spec の確定だけで実装の挙動は変えない。skill の挙動は各 SKILL.md が正典（docs/00_index.md）"

#### Goal

#### Acceptance

- verify: human-review — 所有者が spec を読み、status を active にする承認を返す
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

#### Current Status

as of 2026-11-08 — 所有者の指示で spec を active にした。残っていた判断は推奨どおり（必須検証＝差分と name 一致、RESULT 行に SKIP 件数）。

#### Next Actions

1. なし（完了）

## Split Issues

Issue files in docs/issues/ that declare `workstream: <this id>`. Generated from their front matter
(create_issue.py, docs_hygiene.py --fix); do not edit between the markers.

<!-- own-doc-update:generated split-issues begin -->
- [ISSUE-20261108-mirror-sync-update-existing](../issues/ISSUE-20261108-mirror-sync-update-existing.md) — plugin の版から正典へ置き換える同期（--update-existing の移植と修正）
- [ISSUE-20261108-retire-wip-mirror-branch](../issues/ISSUE-20261108-retire-wip-mirror-branch.md) — 移植後に wip/stash-mirror-skill-20261001 を削除する
<!-- own-doc-update:generated split-issues end -->

## Decisions

Record only decisions that are difficult to reverse or surprising without context.

## Completion

- [ ] Every issue meets its acceptance criteria
- [ ] Every issue records guide impact as required or none
- [ ] Required guides describe current implemented behavior
- [ ] Specs reflect any changed direction or requirements
- [ ] Next human gate reached or the workstream intentionally stopped
- [ ] 00_index.md updated
- [ ] Workstream archived when complete
