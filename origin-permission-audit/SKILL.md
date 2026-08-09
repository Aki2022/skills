---
name: origin-permission-audit
description: Audit and safely settle elevated access granted during the current or resumed session. Use when a task granted IAM roles, API scopes, bucket bindings, or other temporary access and the session is closing, or when the caller explicitly asks to audit or revoke session-granted permissions. Classify grants as temporary or standing, revoke temporary grants, verify the removal, and report the result; do not audit pre-existing access by guesswork.
---

# origin-permission-audit

Audit only access granted by the current session or a resumed session whose history is visible. Do not infer that an older grant belongs to this session. This skill owns permission discovery, classification, revocation, and verification; it does not own repository documentation or ADR creation.

## Workflow

1. List every grant this session made: principal, project or resource, exact role/scope, and reason. Include IAM roles, API scopes, bucket-level bindings, and equivalent elevated access.
2. Classify each grant as `temporary` or `standing`.
   - Treat one-off investigation, remediation, or debugging access as temporary by default, including read-only access.
   - Treat access as standing only when a documented recurring workstream or automation depends on it going forward.
   - Do not relitigate pre-existing grants.
3. For each temporary grant, revoke it now using the narrowest available command. Do not defer the revocation to documentation or a later session.
4. Re-query the exact principal/resource and verify that the grant is gone and the documented baseline remains. If revocation is blocked by an access boundary, return the exact command for the human, wait for confirmation, then verify.
5. Leave standing grants unchanged and record why they are required.

## Safety boundaries

- Resolve the exact principal, resource, role, and scope before changing anything.
- Never use a broad project-wide or wildcard removal when a narrower binding can be targeted.
- Never revoke access that predates the session based only on suspicion.
- Never claim success from a removal command alone; verification is required.
- If the caller has not authorized the relevant external operation or the action is outside the session's authorization envelope, stop and report the gate.

## Handoff contract

Return a concise per-grant report containing: grant identity, temporary/standing classification, action taken, verification result, and any human gate or exact command. Report "no session-granted elevated access" when none exists.

Pass this report to `origin-doc-update` through the caller. That skill decides whether an operational guide or ADR needs updating. Do not create a second permission-history document here. `origin-close-session` invokes this skill before documentation and git cleanup when the audit applies.
