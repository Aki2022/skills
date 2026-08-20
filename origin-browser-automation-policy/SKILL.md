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
- **Once you've named a persistent session (e.g. `-s=chrome-agent`), never issue a bare/default-session `playwright-cli` command for the rest of that task — no exceptions.** Omitting `-s=<name>` doesn't reuse the named session; it silently opens a _different_, usually throwaway, in-memory-profile browser with no relationship to the profile you just set up. Any login performed there evaporates the moment it's closed, while the real dedicated profile looks untouched — this produces the confusing symptom of "the dedicated profile should be logged in, but it isn't," when the actual cause is that the work happened in the wrong browser entirely. This has caused real incidents (leftover `playwright_chromiumdev_profile-*` Chrome processes discovered only by accident during unrelated cleanup checks). Treat a missing `-s=` flag as a typo-class bug, not a style choice.
- **Before starting browser work, check for stray sessions, not just the dedicated one.** Run `playwright-cli list` in addition to the profile's own doctor/status check. Any session other than the one you intend to use (including `default`) is very likely a leftover from an earlier task that never got closed — close it without inspecting its contents (a throwaway session holds nothing worth preserving) before proceeding. Don't wait to discover this opportunistically; make it a standing pre-flight step.

## CDP readiness

Launching a headed Chrome (`chrome-agent-open` or equivalent) and immediately attaching over CDP is a race: the debugging port isn't listening yet in the first instant after launch, and `attach` fails with `ECONNREFUSED`. Poll the doctor/status check (or the CDP `json/version` endpoint) until it reports ready, then attach — don't treat the first failure as a real error and don't guess a fixed `sleep` duration.

## Verify login state directly — don't trust cookies or memory

A dedicated profile that was logged in during an earlier session can be logged out later without any action on your part (site-side session expiry, cleared server state, etc.). A session cookie's mere presence doesn't mean it's still valid server-side. Before relying on a profile being authenticated:

- Load a page whose content differs by auth state (account name in a header, a dashboard, "My Tickets") and read that content — don't infer login from a cookie existing or from "it worked last time."
- If a login button routes through a third-party OAuth provider (Google, Apple, etc.) and lands on an identifier/email prompt instead of auto-consenting, that means there's no existing session with that provider in this profile either — don't type credentials into it; hand off to the user.

## Account creation and credential entry — never perform, even on request

Creating an account or setting a password is prohibited regardless of who asks — this holds even when the user explicitly says to do it. When a form's completion requires setting a new password for future login (not just re-entering an already-known one), that's account creation in disguise, however it's labeled ("register a ticket", "set up your profile"). Recognize it and hand off:

- A password field isn't always genuinely required just because it's presented next to required fields — some sites make it optional and only needed for future self-service login. Fill every other field, leave the password blank, and see whether the submit control enables. If it does, proceed without a password. If clicking submit surfaces a validation error specifically on the password field, it's genuinely required — stop and hand off just that one field.
- When handing off, fill in every field you can supply yourself first, so the human only has to touch the credential field(s). But have them carry the action all the way through to the final submit click before returning control to automation — a filled-but-unsubmitted form does not survive closing the browser or reattaching a new automation session; the human's input is lost and the step must be redone. "Just type the password" followed by you resubmitting later only works if the tab stays open and attached the whole time.

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
