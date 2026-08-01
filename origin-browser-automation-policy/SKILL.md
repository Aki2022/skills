---
name: origin-browser-automation-policy
description: Cross-repo default policy for AI-driven browser automation — when to use playwright-cli vs Chrome, headless vs headed, which Chrome profile to use, and how to clean up browser sessions when done. Use whenever a task involves opening a browser, automating a website, logging into a site, or any browser-based action, regardless of which browser tool ends up being invoked.
---

# Browser Automation Policy

## Default tool

- Use `playwright-cli` by default (see the `playwright-cli` Skill for command usage).
- Default to headless. Only open a visible window (`--browser=chrome` without headless, or a manual Chrome launch) when the task genuinely needs a human to see or interact with the screen.

## When to use Chrome instead

- Use Chrome only when authentication requires it (login sessions, cookies that must persist, SSO, 2FA/reCAPTCHA hand-off) or when the site bot-detects headless playwright.
- Bot detection often lets the page load and only rejects the submit: watch for a 403/ERR_EMPTY_RESPONSE on the form-submit request while the page silently stays put. Check the console/network log before concluding "the click didn't work" — a 403 on the POST means switch tools, not retry.
- When Chrome is needed, always use a dedicated automation profile — never the user's default/daily profile (`--profile-directory=Default` or equivalent is prohibited).
- If the current repo has its own automation-profile launcher (e.g. biz_ops's `./bin/chrome-agent-open` / `chrome-agent-quit`, profile at `~/.config/agent-browser/chrome-agent`), use that. Otherwise use `playwright-cli open --browser=chrome --persistent --profile=~/.config/agent-browser/chrome-agent` (create the profile dir on first use) so the same dedicated profile is reused across repos.
- If a CDP port (e.g. 9222) is already bound by something other than the dedicated profile, stop and report it before connecting — do not attach blindly.

## Session lifecycle — no orphaned browsers

Every browser session you open is yours to close. Orphaned automation browsers linger invisibly (headless has no window) and keep holding memory, ports, and profile locks after the task looks "done".

- **On tool switch**: when you abandon a session mid-task (e.g. playwright-cli hit bot detection, so you switch to Chrome CDP), close the abandoned session _at the moment you switch_ — `playwright-cli close` for the default session — not "later". Later gets forgotten.
- **If a close command is denied or fails**: the session is still alive. Retry with an allowed equivalent, or tell the user it is still running — never silently move on. A denied cleanup that goes unmentioned becomes an orphaned process the user discovers hours later.
- **At end of work**: verify, don't assume. `playwright-cli list` should show no browsers; if the repo has a doctor script (e.g. `./bin/chrome-agent-doctor`), run it and confirm it reports no automation processes. Include the final browser state in your report to the user.

## Escalation

- Falling back to the user's normal browser/profile requires explicit user approval first: state the reason, target site, scope of actions, and how you'll clean up.
- If closing browser processes risks killing the user's normal working session, confirm before doing so.

## Repo-specific detail

Repos with their own established browser workflow take precedence over this default — check for a project doc/skill first (e.g. biz_ops's `docs/guides/browser-automation-policy.md`, referenced from the `origin-bizops-daily-mail` skill). This Skill only fills the gap when no repo-specific policy exists.
