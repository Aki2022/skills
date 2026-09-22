---
name: own-luna-run
description: Temporarily run bounded Codex implementation, code exploration, technical research, and independent review as separate GPT-5.6 Luna CLI processes while Luna cannot be selected as a native Sol/V2 subagent. Use only for Codex delegation that would normally route to worker, explorer, researcher, or reviewer. Expires after 2026-09-30 JST or earlier when Luna becomes Multi-Agent V2 compatible; never use from Claude.
---

# August Luna loop

Use this Codex-only bridge instead of native `spawn_agent` when the requested
child role requires Luna. Keep the primary orchestrator on its current model.

## Run a leaf agent

1. Select exactly one role:
   - `worker`: bounded implementation or fixes; Luna/high; workspace-write.
   - `explorer`: codebase investigation; Luna/high; read-only.
   - `researcher`: source-oriented technical research; Luna/high; read-only.
   - `reviewer`: independent review with fresh context; Luna/max; read-only.
2. Give the leaf a self-contained prompt containing its objective, relevant
   paths, constraints, expected output, and verification requirements. For
   `worker`, include a disjoint Write set. Do not pass unnecessary parent chat.
3. Invoke `scripts/run_luna.py --role <role>`. Pass the prompt on stdin, as a
   positional argument, or with `--prompt-file <path>`. Prefer stdin or a prompt
   file when the text contains shell metacharacters.
4. Consume the JSONL output and check the process exit status. Integrate the
   leaf's final result in the primary thread. Do not describe it as a native
   subagent or expect native wait/steer/thread controls.
5. Run independent writers only in separate worktrees or with non-overlapping
   Write sets.

The runner uses `--ephemeral`, prevents recursive bridge calls, and refuses to
run after 2026-09-30 JST. It also stops early when the installed model catalog
reports Luna as Multi-Agent V2, signaling that this bridge should be removed and
native routing restored.

## Check without model usage

Run `scripts/run_luna.py --check`. This inspects the CLI and model catalog but
does not start a model turn.

Do not bypass the expiration or native-support checks. Extend or remove the
bridge only through an explicit configuration update.
