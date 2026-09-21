# 分割 subagent への指示（sonnet）

置換: `{REPO}`、`{DOC}` = 分割対象（例 `docs/guides/deployment.md`）、`{KIND}` = guide|spec、
`{SECTIONS}` = review が出した H2 と大きさの一覧、`{OTHERS}` = 同時に他の subagent が触っているファイル。

---

Restructure ONE oversized {KIND} at {REPO}/{DOC}. Work ONLY inside {REPO}. Do NOT commit, do NOT create branches,
do NOT run the validator or docs_hygiene, do NOT edit docs/00_index.md (the orchestrator integrates). Other agents
are working on {OTHERS} in the same tree — do not touch those files or their new siblings.

Governing rule (~/.agents/skills/origin-doc-update/SKILL.md, "Write for the next session"): a guide is CURRENT
TRUTH (what the system does now, how to use it, constraints, verification, known limitations); a spec is CURRENT
POLICY. Chronology — dated sections, review cycles, measurement diaries, superseded approaches, "訂正" trails — is
history: move it VERBATIM to docs/log/<slug>-history.md (front matter `---\nupdated_at: <today>\nkind: {KIND}-history\n---`,
heading levels may be demoted by one). Hard-to-reverse decisions with rationale become ADRs
(`python3 ~/.agents/skills/origin-doc-update/scripts/create_adr.py <slug> --scope development --status accepted --repo {REPO}`;
front matter lists in block style, one `  - item` per line, never `[a, b]`); reuse an existing ADR when one covers
the decision.

Sections and sizes measured by the review: {SECTIONS}

Procedure:
1. Read the outline, then sections in slices; classify each: keep / history / decision.
2. Split the current-truth remainder by task or area into files of at most 32 KB, named <slug>-<task>.md next to
   the original, each with the template front matter (`id: <ID>-<slug>-<task>`, `updated_at: <today>`,
   `source_*`/`related_*` copied from the original in block style) and the template sections.
3. Turn {DOC} into a pointer under 3 KB: KEEP ITS ENTIRE ORIGINAL FRONT MATTER VERBATIM (same id and every
   `source_*` list — the validator requires the id to keep listing them), one paragraph on the area, one line per
   new file, a link to the history file.
4. Do not invent content; do not delete facts (everything lands in a guide, the history, or an ADR). Rewrite
   `](../x)` links in moved text so they still resolve from docs/log/. Fix in-page anchors that now cross files.
5. In every NEW file, replace the literal company mail domain with `<company-domain>` and any org-specific label
   prefix (e.g. `com.<org>.`) with a masked form; say so in one line at the top of the history file. Never write
   absolute local paths.
6. Verify programmatically that every line of the original body landed in exactly one destination; report coverage.

Report: original → new files with sizes, history size, ADRs created, dropped dead links, sections you were unsure about.
