---
name: own-jev-research
description: >-
  Research and answer questions about Jev and TypeSafe, including what Jev is,
  availability, installation or setup, model selection, API or SDK use, and
  troubleshooting. Use whenever Jev or TypeSafe is the subject; distinguish the
  Codex Skill from installed software and service access, and verify each from
  relevant evidence before saying it is unavailable.
---

# Jev and TypeSafe questions

<!-- discovery-canary: jev-layer-check-20261022-7f3c -->

Use this skill for general Jev or TypeSafe questions as well as setup and
troubleshooting. Establish which layer the user means before answering:

- **Codex guidance:** the `typesafe-ai` Skill provides integration instructions.
  Its presence does not prove that a local package, account, or API is installed
  or configured.
- **Local software:** verify SDK or CLI installation from the relevant project
  manifest, lockfile, executable, or runtime. A missing local SDK alone does not
  establish that a hosted model or API is unavailable.
- **Service access and model settings:** inspect the relevant project config and
  current official TypeSafe documentation. Never reveal credential values or
  claim credentials were verified without evidence.

Do not repeat questions already answered by the conversation or project
configuration. Ask only for a missing value that is necessary to proceed, or
identify a real conflict in the available settings. Preserve the configured
provider and model unless the user asks to change them. Get authorization before
installing packages, enabling network access, changing authentication, or
modifying model settings.

For current product behavior, use the official documentation index at
<https://docs.typesafe.ai/llms.txt>. For application architecture and API/SDK
integration, follow the `typesafe-ai` Skill and consult the relevant current
documentation page.
