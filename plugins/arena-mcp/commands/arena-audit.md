# /arena-audit — Pre-translation coverage gate

Arguments: `$ARGUMENTS` — path to a `.doe` model (and optionally flags like
`--include-vba-source`, `--include-siman-source`, `--include-binary-payloads`).

1. Run `audit_arena_model_data` on the model with the requested flags.
2. Summarize per-surface status: extracted / metadata_only / not_present.
3. Call out `translation_readiness`, `manual_or_unmapped_definitions`, and
   anything that needs widening flags or manual translation. Do not run
   `extract_arena_model` until the user accepts the coverage picture.
