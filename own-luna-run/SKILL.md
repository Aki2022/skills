---
name: own-luna-run
description: Route explicitly requested Luna code investigation, technical research, and independent review through the native GPT-6 Luna read-only role in Codex. Use for bounded read-only delegation when Luna is requested; never launch a CLI bridge, select another model silently, or use this role for implementation.
---

# Native Luna read-only delegation

Keep the primary agent on its current model. Use the native `luna_readonly`
role for bounded code investigation, technical research, and independent review.
Implementation stays with the primary agent; this skill does not provide a
write-capable Luna role.

## Preflight

1. Inspect the installed model catalog without a model turn: `codex debug models`.
   Require the exact `gpt-6-luna` entry with `multi_agent_version` equal to `v2`.
   Missing or unparseable metadata means unavailable; do not substitute a model.
2. Confirm the current runtime exposes `luna_readonly`. Its shared role definition
   must specify `gpt-6-luna`, reasoning `high`, and sandbox `read-only`.
   Registration in a configuration file is placement evidence, not runtime proof.
3. If the role is absent, stop delegation and report the unavailable route.
   Role definitions are maintained in the reinstall repository's
   `home/agent-config/codex/agents/` and deployed through Nix to
   `~/.agents/codex/agents/`. Each account registers the role against its own
   `agents/` alias. Do not edit a Nix-rendered file or broaden activation scope.
4. After changing routing, verify one small read-only task in a fresh session.
   Confirm the actual child model ID, completed final response, and absence of
   changes before scheduling the real investigation. A role file, successful
   spawn acknowledgement, or the child's self-reported model name is insufficient.

## Delegate and collect

- Spawn with native `agent_type: luna_readonly`, fresh context (`fork_turns: none`),
  and a self-contained prompt. The fixed role supplies the model and reasoning;
  do not override them or reuse a Terra role with a model argument.
- State the exact question, permitted read roots, evidence requirements, output
  format, and time/call budget. Require file and line references, verified
  guarantees, concrete gaps, and unresolved uncertainty.
- Prohibit writes, installs, configuration changes, external mutations, recursive
  delegation, and additional model processes. Do not include credentials or
  unnecessary conversation history.
- Use at most two concurrent children unless the user explicitly chooses another
  limit. Include model usage in the shared budget; unknown cost is not zero.
- Wait for the final response through native lifecycle tools and inspect failures.
  An acknowledgement or completed status without substantive output is not success.
  The primary agent verifies material claims against the referenced source.
- On an unavailable role, model mismatch, or failed canary, stop and report the
  exact error. Do not retry automatically or fall back to another model or CLI.

The expired GPT-5.6 CLI bridge has been retired. There is no date-based bypass,
replacement runner, or implicit fallback path.
