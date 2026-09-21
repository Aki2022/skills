# 圧縮 subagent への指示（sonnet）

置換: `{REPO}`、`{DOC}`、`{KIND}` = guide|spec、`{OTHERS}`。対象は 32 KB 超だが task 分割に向かない（一つの
話題で、経緯・重複・死んだ参照が膨らませている）文書。

---

Condense ONE oversized {KIND} at {REPO}/{DOC} into current truth without losing facts. Work ONLY inside {REPO}.
Do NOT commit, create branches, run the validator/docs_hygiene, or edit docs/00_index.md. Other agents are working
on {OTHERS} — leave those alone.

What "thin" looks like here, measured across 26 repositories: the same fact stated in three sections written
on three dates; a resolved incident narrated where a one-line current rule belongs; commands or paths that no
longer exist (`npm run x`, `.github/workflows/y.yml`, files removed from the tree); dated headings
("### 2026-08-27 revision"); status prose ("完了", "未着手") that belongs in a work unit's frontmatter.

Procedure:
1. Read the outline and every section. For each paragraph decide: current truth (keep, once), history (move
   verbatim to docs/log/<slug>-history.md with front matter `---\nupdated_at: <today>\nkind: {KIND}-history\n---`),
   duplicate (keep the best statement, note the merge in the history file), dead reference (verify with one grep
   or `ls` that the target is absent; then remove it from the guide and list it under "Dead references removed"
   in the history file), decision rationale (create or link an ADR via create_adr.py, block-style lists).
2. Rewrite {DOC} in place as one file under 32 KB, same front matter (keep id and every `source_*` list; set
   `updated_at: <today>`), template sections (What It Does / How To Use / Maintenance Notes / Known Limitations
   for a guide; Purpose / Current Policy / Requirements / Design Direction for a spec), present tense, no dated
   headings, no status words.
3. Do not invent content. Every removed sentence is either in the history file verbatim, merged into a kept
   sentence (say which), or a verified-dead reference (listed). Mask the company mail domain and org labels in
   new files; never write absolute local paths.
4. Verify programmatically: every original line is accounted for (kept / moved / merged / dead-listed); report the
   counts and the before/after size.

Report: before → after bytes, history size, ADRs, dead references removed, merges made, sections you were unsure about.
