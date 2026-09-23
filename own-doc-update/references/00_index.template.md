---
updated_at: YYYY-MM-DD
current_focus:
  - docs/workstreams/WS-YYYYMMDD-short-slug.md
---

# 00 Index

## Read Policy

Read this file first. Do not scan all docs unless needed. Read the active workstream or issue, then only related specs and guides.

## Current Focus

- docs/workstreams/WS-YYYYMMDD-short-slug.md — one-line purpose

## Active Workstreams

<!-- own-doc-update:generated active-workstreams begin -->
- docs/workstreams/WS-YYYYMMDD-short-slug.md — autonomous work until the next human gate
<!-- own-doc-update:generated active-workstreams end -->

## Specs

- docs/specs/product.md — product direction and requirements
- docs/specs/architecture.md — architecture and design policy

## ADRs

- docs/adrs/ADR-YYYYMMDD-short-slug.md — decision rationale, alternatives, and consequences

## Active Issues

<!-- own-doc-update:generated active-issues begin -->
- docs/issues/ISSUE-YYYYMMDD-short-slug.md — standalone one-off work only
<!-- own-doc-update:generated active-issues end -->

## Unassigned Issues

<!-- own-doc-update:generated unassigned-issues begin -->
<!-- own-doc-update:generated unassigned-issues end -->

## Guides

- docs/guides/feature-or-domain.md — current implemented behavior

## Archive Policy

Completed workstreams and issues are stored in their respective archive directories. Archive files are historical context, not current truth.

## Index Policy

This file stays under 32 KB and no line exceeds 500 characters (`validate_repo_docs.py` enforces both). One link and one line of routing per entry, under 200 characters after the link, with no status words (the frontmatter `status` is the only status). Current Focus holds at most five entries. Progress narrative goes to the work unit or `docs/log/`. `docs_hygiene.py --fix` moves violations out.

Rows between `own-doc-update:generated` markers are generated from front matter (`workstream`, `priority`, `due`) by `create_issue.py`, `create_workstream.py` and `docs_hygiene.py --fix`; do not edit them by hand. An issue that belongs to a workstream is routed from that workstream's Split Issues, not from Active Issues. Unassigned Issues lists issues whose declared workstream is not active.
