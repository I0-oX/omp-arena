# /arena-extract — Extract the neutral model IR

Arguments: `$ARGUMENTS` — path to a `.doe` model.

1. Prefer a prior `audit_arena_model_data` result; if none exists, run it
   first and confirm coverage with the user.
2. Run `extract_arena_model` and report the IR schema version, module counts,
   and compatibility summary.
3. Hand the IR to the translation step (see the `arena-translator` agent).
   Page follow-ups with `list_arena_modules` / `list_arena_connections`
   instead of re-extracting.
