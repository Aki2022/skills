# 判定 subagent への指示（sonnet）

置換: `{REPO}` = 作業ディレクトリ（main その場、または hygiene worktree）、`{SWEEP}` = `docs/log/sweep-YYYYMMDD.md`、
`{REVIEW}` = `docs/log/review-YYYYMMDD.md`。

---

You are deciding, with primary evidence, which open documentation items in the repository at {REPO} can be
closed (archived) and which must stay. Work ONLY inside that directory. Do not commit, do not create branches,
do not run archive scripts, do not touch other repositories. Read excerpts (grep, sed -n), not whole files.

Two checklists: {SWEEP} (closure candidates for issues/workstreams: R1 = recorded branch gone, R2 = no commit
for 60+ days, R7 = its text already says done) and the "Demote candidates" section of {REVIEW} (guides/specs/ADRs
no living document links to, older than 90 days). Each candidate is a line `- [ ] archive `<path>`` with
evidence bullets under it.

For each issue/workstream candidate: read frontmatter `status`/`branch`/`pr`, `## Acceptance` (`- verify:`),
`## Current Status`, `## Next Actions`, `## Completion`; check `git log --all --oneline --grep='<id>' | head`,
`git log main --oneline -- <file> | head -3`, whether the branch exists (`git branch -a | grep`), whether a PR in
`pr:` was merged; when a concrete artifact (script, spec, guide section, config key) is named, grep once that it
exists as described. Tick `[x]` ONLY when the acceptance criterion is demonstrably met — finished work missing only
bookkeeping. Otherwise append ` — keep: <one sentence with the evidence>`. Never tick because the text says 完了,
never tick because a branch is gone. Missing evidence → ` — keep: undecided: <what is missing>`.

For each guide/spec/ADR demote candidate: grep the whole repository (code and docs, excluding docs/log and
archive/) for its path, its frontmatter `id`, and its title. Tick `[x]` when nothing outside archive/log refers to
it AND its content is either duplicated by a living guide/spec (name it) or describes something that no longer
exists in the code (name the grep that showed absence). Otherwise ` — keep: <reason>`.

Edit both files in place: only the checkbox state and the appended suffix. Change nothing else.

Finish with one table per file: item | decision | one-line evidence; then the counts ticked / kept / undecided.
