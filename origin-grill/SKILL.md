---
name: origin-grill
description: Shape, create, or materially revise repository specs through a rigorous interview that stress-tests material assumptions one question at a time, grounded in existing docs and code. Use at the start of greenfield development, when a feature or product direction is still ambiguous, when the user asks to create or grill a spec, or when an existing docs/specs file needs a substantive change. Create the expected docs scaffold when absent, persist resolved decisions into a draft spec during the dialogue, and finalize only after explicit shared-understanding approval. Do not use for mechanical spec edits, typo fixes, read-only summaries, or implementation after a spec is already settled.
---

# Origin Grill

Turn an ambiguous development direction into an approved `docs/specs/` source of intent. Interview the user, stress-test material assumptions, maintain the domain language, and write the spec as decisions resolve. Do not implement the result.

## Prepare the repository

1. Locate the repository root and read its agent instructions.
2. Read `docs/00_index.md` first when it exists.
3. Read only the relevant specs, guides, active workstream or issue, and code needed to understand the request.
4. If the docs scaffold is absent, use `origin-doc-update` to initialize `docs/00_index.md`, `docs/specs/`, `docs/workstreams/`, `docs/issues/`, and `docs/guides/` before grilling.
5. Determine the mode:
   - **Greenfield**: shape a new spec from an undeveloped idea.
   - **Revision**: update one existing spec while preserving unrelated settled intent.

Answer factual questions from the repository, tools, or authoritative sources. Ask the user only for decisions, priorities, domain knowledge, and acceptable tradeoffs.

## Stress-test assumptions and delegate fact-finding

- Keep facts and decisions separate. For each decision that materially affects scope, constraints, architecture, or acceptance, challenge the current assumption with a concrete counterexample, boundary case, or failure scenario.
- Keep the exploration bounded: stop probing a branch once its outcome cannot change the stated goals, constraints, or acceptance criteria. Do not turn stress-testing into an exhaustive list of hypothetical branches.
- When a factual prerequisite is not readily available and can be isolated, delegate a bounded, read-only fact-finding task to a sub-agent. Keep independent user decisions moving; do not wait for delegation before asking unrelated questions.
- Treat a sub-agent report as a discovery lead, not primary evidence. Verify material findings against the repository, a tool result, or another authoritative source before recording them in the draft.
- If delegation is unavailable or its latency or cost is disproportionate, use direct tools or record the factual gap explicitly rather than blocking the grill.

## Establish the draft

For greenfield work, first resolve enough of the purpose and canonical name to choose a stable slug. Then create `docs/specs/<slug>.md` from `references/spec.template.md` with `status: draft` and add a clearly marked draft entry to `docs/00_index.md`.

For revisions, edit the existing spec in place. Set `status: draft` while the material change is unresolved. Do not create a second spec for the same intent merely to preserve history; Git owns history.

Create files lazily. Do not invent a spec filename before its subject is understood.

## Run the grill

Build a **bounded decision tree** privately. Mark each material branch `open`, `settled`, `deferred`, or `out-of-scope`, then walk it in dependency order. Ask exactly one question per turn and wait for the answer.

For every question:

1. State the relevant evidence or current assumption briefly.
2. Ask one decision question.
3. Give a recommended answer and its main tradeoff.
4. Wait. Do not include a second question, even as a postscript.

Start with upstream decisions because they reshape downstream questions:

1. desired outcome and problem;
2. actors and canonical domain language;
3. scope, non-goals, and observable behavior;
4. concrete scenarios, edge cases, and failure behavior;
5. constraints, compatibility, migration, and tradeoffs;
6. design direction only where requirements or existing architecture constrain it;
7. acceptance criteria and explicitly deferred decisions.

Adapt the tree to the answers. Do not recite this list as a questionnaire.

## Sharpen the domain model

- Challenge vague or overloaded terms and propose one canonical term.
- Compare new terminology with existing specs, guides, and code. Surface contradictions immediately.
- Probe relationships with concrete scenarios, especially boundary and failure cases.
- Distinguish domain concepts that the user has collapsed into one word.
- Record resolved terms in the spec's `Domain Language` section immediately.
- Keep definitions conceptual. Put implementation choices in `Design Direction`, not in domain definitions.

Do not create a separate glossary or ADR structure unless the repository already uses one. In the standard docs layout, the spec owns its domain language and hard-to-reverse design decisions.

## Persist decisions inline

After each answer, update the draft before asking the next question:

- replace superseded intent instead of appending a changelog;
- preserve unrelated settled requirements;
- record only the resolved outcome and important rationale, not the interview transcript;
- keep unresolved branches in `Open Questions`;
- mark each material branch as `open`, `settled`, `deferred`, or `out-of-scope`; do not silently collapse a branch into an assumption;
- store the single next question in `Next Question` so an interrupted session can resume;
- update `updated_at`.

When revising an existing spec, compare the proposed intent with current guides and active work. Record mismatches under `Impact on Existing System`; do not rewrite guides to pretend the behavior is already implemented.

## Reach convergence

Treat the grill as complete only when the spec has:

- a clear problem and desired outcome;
- explicit goals and non-goals;
- stable domain terms;
- observable requirements and representative edge cases;
- relevant constraints and tradeoffs;
- testable acceptance criteria;
- every material decision-tree branch has a terminal status (`settled`, `deferred`, or `out-of-scope`);
- no hidden open decision, only explicitly deferred ones;
- an identified impact on existing guides, code, and active work.

Then summarize the resulting decisions, stress-tested assumptions, and deferred or out-of-scope branches, and ask one final question: whether shared understanding has been reached and the spec is approved.

If the user says no, continue the grill. If the user approves:

1. Set `status: active` and remove `Next Question`.
2. Remove resolved items from `Open Questions`.
3. Update the spec and `docs/00_index.md` descriptions as current intent.
4. Run the repository docs validator from `origin-doc-update` when available.
5. Report affected guides and workstreams without changing their implementation state.
6. Hand off any implementation planning or workstream creation to `origin-doc-update`, which must establish its own human authorization boundary.

## Stop safely

If the session ends before approval, leave the spec as `draft`, preserve `Open Questions` and `Next Question`, and state that implementation must not start from it yet.

Never during this skill:

- ask multiple questions in one turn;
- ask for facts the repository can answer;
- claim exhaustive coverage of hypothetical branches; close every material branch explicitly instead;
- start implementation, create a branch, or create a workstream;
- mark a spec active without explicit approval;
- silently overwrite a contradiction;
- turn the spec into a chronological scratchpad.

## Resource

Read `references/spec.template.md` when creating a new spec.

Principles adapted from:

- `mattpocock/skills` `grill-with-docs`
- `mattpocock/skills` `grilling`
- `mattpocock/skills` `domain-modeling`
