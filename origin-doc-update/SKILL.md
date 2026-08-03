---
name: origin-doc-update
description: Keep repository documentation aligned with implementation during development, including recording architectural and implementation decisions as ADRs. Use when starting or resuming features, bugs, refactors, API/schema/config/command changes, docs work, workstream or issue creation, unfinished work, docs reorganization, or a hard-to-reverse or surprising design/development decision. Trigger especially when the user asks to create a workstream, when implemented behavior may require docs/guides updates, when an ADR should be created or superseded, or when docs/00_index.md exists. Do not use for pure operational checks, log inspection, status reporting, or read-only explanations with no behavior or documentation change.
---

# origin-doc-update

Treat `docs/` as persistent AI context. Keep each fact in one layer only.

| Path                | Owns                                                                           |
| ------------------- | ------------------------------------------------------------------------------ |
| `docs/00_index.md`  | Small routing index; read first                                                |
| `docs/specs/`       | Intent, requirements, and design policy                                        |
| `docs/workstreams/` | Multi-issue autonomous work between human gates                                |
| `docs/issues/`      | Standalone one-off work only                                                   |
| `docs/guides/`      | Current implemented behavior; source of truth                                  |
| `docs/adrs/`        | Decision rationale, alternatives, and consequences; historical decision record |
| `*/archive/`        | Historical work context, not current truth                                     |

Do not copy implementation history into guides or current behavior into workstreams. Keep decision rationale in an ADR and link to it from the relevant spec, workstream/issue, or guide.

## Start every task

1. Read `docs/00_index.md` if present.
2. Read the active workstream or issue.
3. Read only its related specs, guides, and ADRs when present.
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
   - merge policy: continuous delivery is the default — a PR whose recorded quality gates pass (CI when present, otherwise the recorded local gates) is merged autonomously. Record a human merge gate only as a named exception with its reason (e.g. live/production impact, spend, new dependencies); an unexplained merge gate silently kills autonomous runs downstream. Note the choice as one line in the Authorization Envelope or Human Gates section;
   - the next human checkpoint and early stop conditions.
5. Classify runnability for every planned issue: `ready` (completable with only current permissions and currently available information) or `gated` (a human decision, missing input, or new permission is foreseeable). Surface every gated point as a question now, during creation — a question asked here costs one interview turn, while the same question discovered mid-execution stops an entire autonomous run. As part of the same pass, judge whether each issue's acceptance is machine-verifiable (counts, thresholds, passing tests) or needs human review: push subjective acceptance toward a quantifiable restatement, and where human judgment is genuinely required, record that the issue ends at a review gate — an agent cannot self-verify a subjective goal and will either stall or overclaim. Record the outcome as one line in the workstream's Human Gates section, e.g. `- Runnability: all issues ready (YYYY-MM-DD)` or `- Runnability: ISSUE-02 gated on <decision>`.
6. State that `origin-doc-update` is pausing creation until these boundaries are confirmed.
7. After confirmation, create from `references/workstream.template.md` or run `create_workstream.py` with the confirmed boundary fields.

Record only decisions that are hard to reverse or surprising without context. Do not create ADRs for routine implementation choices, temporary investigation notes, or ordinary history already captured by the workstream/issue.

Do not audit or revoke IAM roles, API scopes, bucket bindings, or other session-granted access here. `origin-permission-audit` owns that lifecycle; `origin-close-session` invokes it before this skill when applicable and passes the verified outcome into the documentation slice.

## Record an ADR

Use the repository-level `docs/adrs/` directory for both kinds of decision:

- `scope: spec` — intent, requirements, architecture, design policy, security posture, or other decisions that change what the system is meant to be. Update the related `docs/specs/` document with the current policy and link back to the ADR.
- `scope: development` — implementation, integration, migration, operational, dependency, or tooling tradeoffs made while delivering a workstream or issue. Link the ADR from that work unit and update a guide when the resulting behavior is user-, operator-, integrator-, or agent-visible.

Create the record when the decision is made, not only at session close. Use `references/adr.template.md` or `scripts/create_adr.py <slug> --scope <spec|development> --status <proposed|accepted|rejected>`. The generated filename is `ADR-YYYYMMDD-<slug>.md`. Keep accepted and rejected records in `docs/adrs/`; when a decision changes, create a new ADR and mark the old one `superseded` with a link instead of rewriting its decision history.

An ADR must state the context/problem, the decision, alternatives considered, consequences, and links to its source workstream/issue and affected specs/guides. Use `status: proposed` while a human gate is pending and `status: accepted` or `rejected` after the decision is settled. Record only the rationale here; current policy belongs in specs and current behavior belongs in guides.

## Improvement issues

An improvement issue is a standalone issue in `docs/issues/` that records a friction observation from development work — a stuck point, a repeated manual step, an inefficiency worth fixing later — rather than a requested change. Name it `ISSUE-YYYYMMDD-improve-<slug>` when the improvement targets the repository itself, and `ISSUE-YYYYMMDD-improve-loop-<slug>` when it targets the development-loop machinery (skills canonical under `~/.agents/skills/`). The two scopes have different owners and approval paths, so the name must reveal the scope at a glance. Autonomous runs (e.g. `origin-ws-loop`) file observations here instead of interrupting their work; humans triage them later. Create with `create_issue.py` as usual.

## Implement a workstream or issue

For each issue, complete one vertical slice:

1. Set `guide_impact` before implementation:
   - `required`: name every guide that must describe the resulting behavior.
   - `none`: write a concrete reason, such as internal refactor with unchanged behavior.
2. Define acceptance criteria and dependencies.
3. Capture each qualifying decision in `docs/adrs/` during the slice and link it from the issue/workstream. Mark a human-gated decision as `proposed` until the gate is resolved.
4. Implement and test, preferring red-green-refactor where practical.
5. Update the target guide in the same slice, before marking the issue complete.
6. Update current status and next actions. An issue's `status` must be exactly
   one of `pending`, `in_progress`, `blocked`, `complete` — `validate_repo_docs.py`
   rejects anything else, and plausible words like `done` are the usual way to
   find that out the hard way.

7. Continue automatically while inside the authorization envelope.
8. Stop at the next human gate or any recorded stop condition.

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
- When that change follows a qualifying decision, update the related spec with the current policy and link the ADR; do not duplicate the full rationale in the spec.
- Keep chronological investigation and abandoned approaches in the active work unit, then archive it.
- Keep `docs/00_index.md` as links plus one-line routing descriptions. Do not add a second progress dashboard unless ordering across many workstreams cannot fit in the index.

## Complete work

Before archive:

1. Verify every issue acceptance criterion.
2. Verify every issue has `guide_impact: required` or `none`.
3. Verify required guides describe the implemented behavior and reference the source work.
4. Verify qualifying decisions have an ADR in `docs/adrs/`, with a settled status or an explicit proposed human gate, and that related specs/guides/work units link to it.
5. Update specs if direction changed.
6. Reach the recorded human gate or record why the workstream stopped.
7. Run `validate_repo_docs.py <repo path>`. Name the repository rather than
   relying on the current directory: reached through an orchestrator, the current
   directory is a different repository, whose docs would validate clean and be
   reported as this one's result. Check the `validated:` line it prints.
8. Archive the work unit and update `docs/00_index.md`.
9. Hand merged branch cleanup to `origin-git-cleanup`.

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
- `adr.template.md`
- `spec.template.md`

Scripts in `scripts/`:

- `init_repo_docs.py [repo]`
- `create_workstream.py <slug> --issue <slug> --scope <text> --confirmed-at YYYY-MM-DD --next-human-gate <name> (--guide <GUIDE-id> | --no-guide-reason <text>) [--repo <repo>]`
- `create_issue.py <slug> (--guide <GUIDE-id> | --no-guide-reason <text>) [--title <title>] [--repo <repo>]`
  Pass a slug, not a full issue id — the `ISSUE-<date>-` prefix is added for you.
  The guide decision is required, as it is for `create_workstream.py`: name the guide
  this issue must update, or state why it changes no implemented behavior. Without it
  the generated file cannot pass `validate_repo_docs.py`.
- `create_adr.py <slug> --scope <spec|development> [--status <proposed|accepted|rejected>] [--title <title>] [--repo <repo>]`
- `archive_workstream.py <workstream> [--repo <repo>]`
- `archive_issue.py <issue> [--repo <repo>]`
- `validate_repo_docs.py <repo>` (prints the repository it validated)

Read the matching template before creating a file manually. Keep legacy archives in place; promote useful current knowledge into a guide or spec instead of renaming history.
