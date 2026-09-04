---
name: lmarena-direct
description: Chat with arena.ai (LMArena) direct mode through the user's own logged-in Chrome tab. Use when the user says arena-direct, direct mode, or wants an answer from a specific arena.ai model.
---

# arena.ai direct chat via your browser

Drive **arena.ai direct mode** (https://arena.ai/text/direct) through the
user's own Chrome tab, where they are already logged in. The site has **no
public API**; this skill operates the real UI as the user would. Never ask
for or handle their Google/email credentials — login stays in their browser.

## One-time setup (user does this, not the agent)

```sh
omp browser-relay install
```

Then load the unpacked extension from `~/.omp/browser-relay/extension` in
their Chrome and log into https://arena.ai once. Verify with a screenshot
before the first real run.

## Verified UI map (arena.ai/text/direct, dark UI)

- Mode combobox reads `Direct`; model button shows the slug (e.g. `Max`).
  Deep-link a model with `?model_a=<slug>` (verified: `?model_a=max`).
- Prompt box: textbox with placeholder `Ask anything...`; send via the
  `Send message` button (or Enter).
- First-visit dialogs, dismiss in order: `Accept Cookies`, then Terms
  `Agree`. If a `Continue with Google` / `Continue with email` wall appears,
  **stop and tell the user to log in** — do not attempt credentials or
  CAPTCHAs.

## Workflow (`/arena-direct <model> <prompt>`)

1. Open the user's tab on the model URL (relay, dedicated tab — never the
   visible tab without asking):
   `browser.open({ name: "arena-direct", app: { relay: true }, url: "https://arena.ai/text/direct?model_a=<slug>" })`
2. `observe()` → dismiss cookie/terms dialogs if present → confirm the model
   button shows the requested slug.
3. Fill the prompt box, send, wait ~20-40s (re-observe until the answer
   stabilises; streaming text changes between polls).
4. Return the final assistant text verbatim plus the model slug. Never vote,
   never open battle mode — direct only.

## Hard limits (verified 2026-09-04)

- Headless/anonymous use stops at the login wall — relay + user login required.
- The site states reCAPTCHA protection: if a challenge appears, stop and hand
  back to the user.
- No battle mode, no voting, no API keys: direct chat only, as the user.
