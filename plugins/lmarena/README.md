# lmarena plugin for OMP

Chat with **arena.ai (LMArena) direct mode** from OMP through the user's own
logged-in Chrome tab. arena.ai exposes **no public API**, so this plugin
drives the real UI instead of faking a provider.

## Install

```
/marketplace add I0-oX/omp-arena
/marketplace install lmarena@omp-arena
```

One-time browser setup (the user, once):

```sh
omp browser-relay install
```

Load the unpacked extension from `~/.omp/browser-relay/extension`, log into
https://arena.ai in that Chrome.

## Use

```
/arena-direct max Explain recursion in one paragraph
```

Or invoke the `lmarena-direct` skill. See the skill for the verified UI map
(model deep-links, dialogs, limits).

## Limits

- Relay + user login required; anonymous/headless stops at the login wall.
- reCAPTCHA-protected: on any challenge, stop and hand back to the user.
- Direct chat only — no battle mode, no voting, no credentials handling.
