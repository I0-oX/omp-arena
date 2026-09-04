# /arena-direct — Ask an arena.ai model via your browser

Arguments: `$ARGUMENTS` — `<model-slug> <prompt...>`. Example:
`/arena-direct max Explain recursion in one paragraph`.

Requires the one-time relay setup from the `lmarena-direct` skill (user
logged into arena.ai in their own Chrome). Without it, stop and explain
the setup — headless use cannot pass the login wall.

1. Open relay tab `arena-direct` on
   `https://arena.ai/text/direct?model_a=<slug>` (default slug: `max`).
2. Dismiss `Accept Cookies` / Terms `Agree` if shown. Login wall
   (`Continue with Google`) → stop, ask user to log in.
3. Confirm the model button matches `<slug>`; fill `Ask anything...`,
   send, wait for the answer to stabilise.
4. Reply with the model slug + full answer text. Direct mode only — no
   battles, no votes.
