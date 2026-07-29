---
name: origin-doc-update
description: Keep repository documentation aligned with implementation during development. Use when starting or resuming features, bugs, refactors, API/schema/config/command changes, docs work, workstream or issue creation, unfinished work, or docs reorganization. Trigger especially when the user asks to create a workstream, when implemented behavior may require docs/guides updates, or when docs/00_index.md exists. Do not use for pure operational checks, log inspection, status reporting, or read-only explanations with no behavior or documentation change.
---

# origin-doc-update

Treat `docs/` as persistent AI context. Keep each fact in one layer only.

| Path                | Owns                                            |
| ------------------- | ----------------------------------------------- |
| `docs/00_index.md`  | Small routing index; read first                 |
| `docs/specs/`       | Intent, requirements, and design policy         |
| `docs/workstreams/` | Multi-issue autonomous work between human gates |
| `docs/issues/`      | Standalone one-off work only                    |
| `docs/guides/`      | Current implemented behavior; source of truth   |
| `*/archive/`        | Historical work context, not current truth      |

Do not copy implementation history into guides or current behavior into workstreams. Link instead.

## Start every task

1. Read `docs/00_index.md` if present.
2. Read the active workstream or issue.
3. Read only its related specs and guides.
4. Choose one work unit:
   - Use a **workstream** when multiple vertical-slice issues can run under the same authorization envelope until the same next human gate.
   - Use a **standalone issue** for one bounded change, an unrelated blocker, or work with an independent lifecycle.
   - Use **docs only** when no implementation changes.
5. Reuse the recorded branch or worktree. Do not create a second branch for resumed work.

Default to one workstream file with embedded issue blocks. Split it only when independent branches, parallel ownership, or file size makes one file materially harder to resume.

## Create a workstream

Do not create the file immediately when the user asks for a workstream. First establish the human boundary.

1. Inspect the code and existing docs. Answer anything discoverable without asking the user.
2. Identify decisions that change authorization or the path of work.
3. Ask one question at a time, in dependency order, and include a recommended answer.
4. Confirm these items before writing the workstream:
   - goal, success criteria, and out-of-scope work;
   - actions the agent may take autonomously;
   - actions requiring confirmation, especially external writes, sends, submissions, deploys, destructive changes, data migrations, auth/secrets, dependencies, network access, and metered services;
   - cost or usage ceiling when metered work is possible;
   - test and quality gates;
   - the next human checkpoint and early stop conditions.
5. State that `origin-doc-update` is pausing creation until these boundaries are confirmed.
6. After confirmation, create from `references/workstream.template.md` or run `create_workstream.py` with the confirmed boundary fields.

Record only decisions that are hard to reverse or surprising without context. Do not add a separate glossary or ADR unless the repository already uses one and the decision belongs there.

## Implement a workstream or issue

For each issue, complete one vertical slice:

1. Set `guide_impact` before implementation:
   - `required`: name every guide that must describe the resulting behavior.
   - `none`: write a concrete reason, such as internal refactor with unchanged behavior.
2. Define acceptance criteria and dependencies.
3. Implement and test, preferring red-green-refactor where practical.
4. Update the target guide in the same slice, before marking the issue complete.
5. Update current status and next actions.
6. Continue automatically while inside the authorization envelope.
7. Stop at the next human gate or any recorded stop condition.

Never defer all guide work to workstream close. A guide is part of the definition of done for the issue that changed behavior.

## Guide contract

Update or create a guide when behavior observable by a user, operator, integrator, or future agent changes, including:

- commands, configuration, schemas, APIs, supported workflows, and defaults;
- operational procedures, safety constraints, failure handling, and known limitations;
- behavior needed to use, maintain, debug, or extend the implementation correctly.

Do not update a guide for a pure internal refactor with identical behavior. Record `guide_impact: none` and why.

Write guides as current truth, not as a changelog. Include what the system does, how to use it, guarantees or constraints, maintenance/verification notes, and known limitations. Add the source workstream or issue ID in front matter.

## Update specs and history

- Update a spec only when intent, requirements, architecture, or design policy changes.
- Keep chronological investigation and abandoned approaches in the active work unit, then archive it.
- Keep `docs/00_index.md` as links plus one-line routing descriptions. Do not add a second progress dashboard unless ordering across many workstreams cannot fit in the index.

## Complete work

Before archive:

1. Verify every issue acceptance criterion.
2. Verify every issue has `guide_impact: required` or `none`.
3. Verify required guides describe the implemented behavior and reference the source work.
4. Update specs if direction changed.
5. Reach the recorded human gate or record why the workstream stopped.
6. Run `validate_repo_docs.py`.
7. Archive the work unit and update `docs/00_index.md`.
8. Hand merged branch cleanup to `origin-git-cleanup`.

## Resume and onboard

When resuming, follow `docs/00_index.md` to the active work unit, reuse its branch, then continue from `Next Actions` without rescanning the repository.

When onboarding scattered docs, initialize the scaffold, classify each file by the ownership table, preserve history with `git mv`, confirm ambiguous removals, update cross-references, and validate.

## Hooks

Keep hooks best-effort and non-blocking:

- `session_start.sh` may inject only `docs/00_index.md`.
- `stop_nudge.sh` may emit one short reminder when non-doc changes lack docs changes.
- Do not force-load this full skill from a hook and do not block commits or task completion from a semantic guess.

Use template fields and `validate_repo_docs.py` for deterministic enforcement. Hooks cannot reliably infer whether behavior changed and hard enforcement creates false positives and repeated token cost.

## Resources

Templates in `references/`:

- `00_index.template.md`
- `workstream.template.md`
- `issue.template.md`
- `guide.template.md`
- `spec.template.md`

Scripts in `scripts/`:

- `init_repo_docs.py [repo]`
- `create_workstream.py <slug> --issue <slug> --scope <text> --confirmed-at YYYY-MM-DD --next-human-gate <name> (--guide <GUIDE-id> | --no-guide-reason <text>) [--repo <repo>]`
- `create_issue.py <slug> (--guide <GUIDE-id> | --no-guide-reason <text>) [--title <title>] [--repo <repo>]`
  Pass a slug, not a full issue id — the `ISSUE-<date>-` prefix is added for you.
  The guide decision is required, as it is for `create_workstream.py`: name the guide
  this issue must update, or state why it changes no implemented behavior. Without it
  the generated file cannot pass `validate_repo_docs.py`.
- `archive_workstream.py <workstream> [--repo <repo>]`
- `archive_issue.py <issue> [--repo <repo>]`
- `validate_repo_docs.py [repo]`

Read the matching template before creating a file manually. Keep legacy archives in place; promote useful current knowledge into a guide or spec instead of renaming history.
