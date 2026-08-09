---
name: origin-vibe-security-setup
description: Set up or update vibe-guard security guardrails for an AI-assisted coding repository (git pre-commit/pre-push hooks, gitleaks secret scanning, local-info detection, agent bypass rules). Use when bootstrapping a repo's guardrails, running `vibe-guard doctor`/`repo-bootstrap`, or when a security check blocks a commit/push and the user asks how to remediate without bypassing it.
---

# Vibe Security Setup Skill

Use this skill when setting up or updating security guardrails for an AI-assisted coding repository.

## Goal

Apply lightweight, reproducible guardrails without blocking normal development unnecessarily.

## Steps

1. Confirm the current directory is a git repository.
2. Run:

```bash
vibe-guard doctor
```

3. If global setup is missing, tell the user to run:

```bash
bash setup-vibe-guard.sh
```

4. Bootstrap the repository:

```bash
vibe-guard repo-bootstrap
```

5. Review generated files:

```bash
git diff -- AGENTS.md SECURITY.md .vibe-guard/README.md .github/workflows/vibe-guard.yml
```

6. Do not bypass failed hooks or CI checks.
7. If a check blocks work, fix the root cause. Only add a narrow allowlist rule after human review.

## Agent rules

- Never use `git commit --no-verify` or `git commit -n`.
- Never disable or rewrite hooks to continue.
- Never remove security workflows to make CI pass.
- Never expose local paths, personal emails, credentials, private URLs, drive names, or tokens in code, docs, PRs, issues, logs, or commit messages.
- Ask before adding dependencies, enabling network access, changing auth, changing CI/CD, or editing secret handling.

## Output

When finished, report:

- Whether global hook is active.
- Whether repo bootstrap files exist.
- What was changed.
- Any blocked items and how they were remediated.
