---
schema_version: 2
id: ISSUE-20260926-global-skill-plugin-commonization
status: archived
workstream: none
priority: medium
due: none
created_at: 2026-09-26
updated_at: 2026-09-26
branch: codex/global-skill-commonization-20260926
pr: "https://github.com/Aki2022/skills/pull/16"
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "own-skill-commonize/SKILL.md と references/plugin-commonization-2026-09-26.md がこの手順・実測の正典で、docs/guides は本 repo では使用しない"
---

# グローバルスキルの共通化とプラグイン監査

## Goal

Claude・Codex・AGY のグローバル skill 供給を `~/.agents/skills` に共通化し、retired の再配線を防ぎ、プラグインの導入記録と有効設定の不整合を検知する。認証状態と project scope は共有・変更しない。

## Acceptance

- verify: machine — `bash ~/.agents/skills/own-skill-commonize/scripts/skill_lint.sh` が `OK: all skill checks passed`、かつ `own-skill-commonize` 99 tests と `own-seat2-build` 12 tests が成功
- verify: machine — `python3 ~/.agents/skills/own-skill-commonize/scripts/plugin_inventory.py --codex main=$HOME/.codex --codex private=$HOME/.codex-private --codex seat2=$HOME/.codex-seat2` が `RESULT: OK` / exit 0、3 account の enabled config と installation record の mismatch が0件。marketplace asset は exact-version cache を解決する
- verify: machine — Claude `/eli5`、Codex `$eli5`、AGY `/eli5` の native canary が一時 token を読み戻し、正典ファイルが元バイト列へ復元済み
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-09-26 — PR #16 merged as `1b1f1d7`; active mirrors, the three-account plugin inventory, native ELI5 canaries, and full skill lint pass. The task branch has been deleted.

## Next Actions

- None — PR #16 is merged, the task branch is cleaned up, and this issue is archived.

## Guide Impact

- Decision: none
- Target or reason: `own-skill-commonize/SKILL.md` and its dated inventory reference are the operational source of truth; this repository does not use `docs/guides` for skill procedures.

## Notes

- Before the settings fix, the live inventory surfaced 8 private and 1 seat2 enabled-config-only entries but incorrectly returned `RESULT: OK`; the regression test now makes this mismatch return REVIEW / exit 1.
- The nine config flags were changed to `enabled = false`; the plugin integrations themselves were not newly installed. The provider-level ELI5 license discrepancy remains recorded as unresolved in the inventory reference.
- An earlier re-audit found `google-drive@openai-curated` enabled in Codex main config without an installed record and returned `RESULT: REVIEW`. The latest native CLI listing now reports it installed and enabled; the three-account inventory returns `RESULT: OK` with zero mismatches. Current counts and limits are recorded in the dated plugin inventory reference.
- The current provider is AGY / Antigravity CLI. Its live skills root `~/.gemini/antigravity-cli/skills` is a whole-root symlink directly to `~/.agents/skills`; the separate `~/.gemini/config/skills` path is the Gemini CLI per-skill alias and is not used as evidence for the AGY canary.
- Codex marketplace records without `source.path` were falsely reported unresolved by the first inventory implementation. The inventory now resolves the same-account cache using marketplace, plugin ID, and exact version; a regression fixture covers this shape.
- `plugin_inventory.py` regression suite: 8 passed. Final `skill_lint.sh`: `OK: all skill checks passed`; this included 99 own-skill-commonize tests, 4 own-git-clean tests, 12 own-seat2-build tests, and the other enabled skill suites.

## Log

- 2026-09-26 — Corrected sync and audit to use active `mirrors:` entries and exclude `retired:`; mirrored and wired ELI5; documented plugin skill/MCP/command/hook routes and license caveat; reconciled config-only Codex flags; verified test suites, lint, and native canaries for Claude, Codex, and AGY.
- 2026-09-26 — Re-read the live account-specific Codex plugin inventory after the main Drive record appeared; confirmed all three inventories return `RESULT: OK` with zero config/install mismatches and updated the dated reference.
- 2026-09-26 — PR #16 merged; the task branch was deleted after syncing main, and the issue was closed during session cleanup.

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
