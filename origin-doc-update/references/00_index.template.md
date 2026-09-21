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

- docs/workstreams/WS-YYYYMMDD-short-slug.md — autonomous work until the next human gate

## Specs

- docs/specs/product.md — product direction and requirements
- docs/specs/architecture.md — architecture and design policy

## ADRs

- docs/adrs/ADR-YYYYMMDD-short-slug.md — decision rationale, alternatives, and consequences

## Active Issues

- docs/issues/ISSUE-YYYYMMDD-short-slug.md — standalone one-off work only

## Guides

- docs/guides/feature-or-domain.md — current implemented behavior

## Archive Policy

Completed workstreams and issues are stored in their respective archive directories. Archive files are historical context, not current truth.

## Index Policy

This file stays under 32 KB and no line exceeds 500 characters (`validate_repo_docs.py` enforces both). One link and one line of routing per entry, under 200 characters after the link, with no status words (the frontmatter `status` is the only status). Current Focus holds at most five entries. Progress narrative goes to the work unit or `docs/log/`. `docs_hygiene.py --fix` moves violations out.
