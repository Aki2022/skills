# Global plugin commonization snapshot — 2026-09-26

This is a dated decision record, not a live inventory. Re-run `scripts/plugin_inventory.py` before
changing plugin state. Inventory covers global account installations only; Claude project-scope plugins
are excluded from the counts. Disabling matching user-scope entries also made three same-ID
project-scope rows appear disabled in Claude’s CLI output. Project settings remained out of scope;
those skill-only plugins now use the canonical skill route. Cache directories without an
installed-and-enabled product record are not installations. Account authentication, sessions,
and history remain account-local.

## Installed and enabled counts before commonization

| Product account | Installed and enabled plugins | Scope note |
| --- | ---: | --- |
| Claude main | 4 | User scope only |
| Claude seat2 | 10 | 10 enabled project-scope records were excluded |
| Claude private | 0 | No user plugin settings file |
| Codex main | 20 | Eight other Codex plugin families are product runtime integrations |
| Codex seat2 | 19 | One enabled config row had no installed record |
| Codex private | 12 | Eight enabled config rows had no installed record |

`CONFIG-ONLY` entries do not count as active plugins or active function assets. They remain unresolved
configuration mismatches and make the live inventory return `RESULT: REVIEW` / exit 1 until reconciled.
Disabled plugin caches likewise do not count. Project-scope settings are out of scope; the observed
cross-scope CLI coupling is recorded below.

## Result after user-scope changes

| Product account | Enabled user plugins after change | Disabled installed plugins | Project-scope state |
| --- | ---: | ---: | --- |
| Claude main | 3 | 1 | None reported |
| Claude seat2 | 7 | 3 | 7 rows still appear enabled; 3 same-ID skill-only rows appear disabled |
| Claude private | 0 | 0 | No user plugin settings file |
| Codex main | 17 | 3 | Not applicable |
| Codex seat2 | 16 | 3 | Not applicable |
| Codex private | 10 | 2 | Not applicable |

Claude CLI's `plugin disable --scope user` made the three matching project-scope rows appear disabled
as well. The project settings were left out of scope and no repository plugin settings were retained.
These are skill-only plugins, whose skills are now served by the shared canonical route. This is an
observed CLI coupling for identical plugin IDs; do not add a project-scope override merely to restore
the display value.

## Skills moved to the shared canonical route

| Skill | Source route | Decision |
| --- | --- | --- |
| `eli5` | Claude community plugin, skill-only | Mirror into the canonical skill root and disable the Claude user plugin. Its plugin manifest says MIT while the upstream repository LICENSE says Apache-2.0; record both claims and leave the license conflict unresolved pending upstream clarification. |
| `karpathy-guidelines` | Claude and Codex skill-only plugin | Mirror once; disable the installed skill-only plugin per applicable global account. |
| `claude-automation-recommender` | Official Claude/Codex skill-only plugin | Mirror once; disable the skill-only plugin per applicable global account. |
| `frontend-design` | Official Claude/Codex skill-only plugin | Reactivate its retired canonical entry from the currently installed upstream skill copy; disable the dedicated skill-only plugin per applicable global account. The mixed `example-skills` collection still has a duplicate copy in Claude seat2. |
| `claude-md-improver` | Skill within the hybrid `claude-md-management` plugin | Mirror the independent skill. Keep the plugin enabled because its `/revise-claude-md` command is a separate active feature; the skill remains duplicated inside that plugin. |

All imported third-party skill bodies remain byte-identical to the selected local source. The mirror ledger
records upstream identity/version, observed license, and comparison source where available.

## Functional exceptions kept enabled

| Plugin family | Observed assets | Reason and alternative |
| --- | --- | --- |
| `example-skills@anthropic-agent-skills` (Claude seat2) | 20 skills, 3 agents; per-skill licenses: 14 Apache-2.0, 4 Proprietary, 2 unspecified | Keep the collection for its non-skill agent assets. Do not bulk mirror a mixed-license collection. Reconsider only individual skills after checking their own license and provider/tool dependencies. |
| `google-drive@openai-curated` (Codex main) | 5 skills, MCP server, app | The skills explicitly route through the Codex `google-drive` MCP/app tools and account authorization. Claude has no equivalent connector installed in this scope; copying the instructions would not create equivalent function. Keep the Codex plugin and its account auth local. |
| `claude-md-management` (Claude seat2 and Codex accounts) | `claude-md-improver` skill plus command | Keep the command feature. The independently readable skill is mirrored, with plugin duplication recorded above. |
| `codex@openai-codex` | Skills, commands, agent, hooks | Keep Codex runtime support and hook/command behavior; do not disable the hybrid plugin. |
| `browser`, `chrome`, `computer-use`, `unified-computer-use`, `codex-app-tools`, `sites`, `visualize`, document and spreadsheet runtime plugins | Product skills and/or hooks, MCP, or app integrations | These connect to Codex desktop/runtime capabilities or proprietary app behavior. Keep enabled where installed; no equivalent Claude/Gemini feature was verified in this inventory. |
| `pyright-lsp`, `security-guidance`, `feature-dev`, `code-simplifier` | LSP, hooks, commands, or agents | These are integrations or workflow components, not standalone skills. Keep enabled. |

The inventory records only observed installed routes. A similar product name or a cached directory does not
prove that another provider has the same connector or feature.

The Drive row above describes the initial snapshot only. The later live re-audit below found
that Codex main no longer had an installed record for that plugin ID. Because the initial snapshot
identified it as an intended MCP/app integration, its `enabled = true` config flag remains visible and
unresolved pending an explicit restore-or-retire decision. Codex seat2's separately installed
remote-marketplace Drive plugin remains active.


## Initial config-only Codex entries found on 2026-09-26 (resolved below)

The live inventory found eight enabled config entries without installed records in Codex/private:

- andrej-karpathy-skills@personal
- computer-use@openai-bundled
- documents@openai-primary-runtime
- google-drive@openai-curated
- pdf@openai-primary-runtime
- presentations@openai-primary-runtime
- spreadsheets@openai-primary-runtime
- template-creator@openai-primary-runtime

Codex/seat2 has one: google-drive@openai-curated. These are account-local config/installation
mismatches, not evidence by themselves that the plugins were once installed and removed. The
pre-commonization inventory recorded the same counts (eight private, one seat2), but did not save the
IDs, so it does not prove these are the exact same entries. Older private config backups omit the
eight IDs while the current private config includes them; the seat2 Drive flag is present in an
August 31 backup. No Codex plugin removal is performed by plugin_inventory.py; it only reads the
account CLI list and config. Plugin cache directories are not install-history evidence. These nine
flags were later set to `enabled = false` after the user requested reconciliation when no installed
record exists. That did not install, remove, or reauthenticate any plugin. Spreadsheet plugin records
remain installed and active where the CLI reports them.

## Live re-audit after reconciliation — 2026-09-26

The current CLI inventory later surfaced one additional Codex main flag,
`google-drive@openai-curated`, with no installed record. Its config remains `enabled = true` because
the initial snapshot recorded an active MCP/app integration; no plugin was installed, removed, or
reauthenticated. Codex seat2 still has an installed and enabled
`google-drive@openai-curated-remote` plugin with five skills and the app integration, so that account's
feature remains active. Codex main currently has no `google-drive` installed record, so the seat2
integration is not evidence of a main-account connection. The mismatch is deliberately retained as
`CONFIG-ONLY` / `RESULT: REVIEW` until the user authorizes a restore or confirms retirement.

The current non-Claude/Codex agent is AGY / Antigravity CLI. Its live skill root
`~/.gemini/antigravity-cli/skills` is a whole-root symlink directly to
`~/.agents/skills`; the separate `~/.gemini/config/skills` per-skill aliases belong to Gemini CLI and
are not used as evidence for the AGY canary.

| Product account | Installed + enabled | Disabled installed | Enabled config without installed record | Enabled config with disabled record |
| --- | ---: | ---: | ---: | ---: |
| Codex main | 19 | 3 | 1 | 0 |
| Codex private | 12 | 2 | 0 | 0 |
| Codex seat2 | 26 | 3 | 0 | 0 |

Current inventory result: `RESULT: REVIEW (1 unresolved inventory entries)`, exit 1, due to the main
`google-drive` mismatch. The counts differ from the earlier post-change table
because the current Codex CLI now lists additional marketplace plugins. For installed marketplace
records that omit `source.path`, the inventory resolves only the exact-version cache under that same
account's `plugins/cache/<marketplace>/<plugin>/<version>` directory. Cache presence by itself still
does not count as an installation. A regression test covers this current CLI record shape.

The first re-audit classified 15 remote marketplace records as unresolved because their records used
`marketplaceName` / `pluginId` / `version` rather than `source.path`. The exact-version resolver fixed
that parser gap and the final inventory resolves the asset sources for all installed+enabled records.
The only remaining unresolved entry is Codex main's config-only `google-drive` plugin.
