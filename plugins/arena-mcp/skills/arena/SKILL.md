---
name: arena
description: Read-only extraction of Rockwell Automation Arena (.doe) simulation models for translation to Python. Use when the user mentions Arena, .doe models, SIMAN, or simulation translation.
---

# Arena model extraction

Read-only access to Rockwell **Arena** simulation models (`.doe`) through
Arena's own COM object model. Nothing is ever saved, run, or edited.
Extracted data is a versioned JSON intermediate representation (IR) for
driving a Python simulation translator.

## Requirements

- **Windows** with a licensed Arena install (`Arena.Application` ProgID) for
  every tool except the ones marked portable below.
- Python 3.10+ on that machine (`ARENA_PYTHON` overrides `python3`).
- `pip install -r server/requirements.txt` for MCP-server mode.

Portable tools (work on any OS, no Arena needed): `inspect_arena_native`,
`arena_status` (without `live_check`), `list_arena_models`,
`inspect_arena_compound_file`, `inspect_arena_results`, `read_arena_results`.

## Environment

- `ARENA_MODEL_ROOTS` — `os.pathsep`-separated search roots (default:
  Documents/Desktop/OneDrive + public Rockwell folder).
- `ARENA_ALLOW_ANY_PATH=1` — allow paths outside the roots.
- `ARENA_PYTHON` — interpreter used by the extension tools.
- `ARENA_BRIDGE_PATH` — override for `server/omp_bridge.py`.

## Workflow (same as upstream arena-mcp)

1. `arena_status` — confirm Arena is reachable (`live_check: true` opens Arena).
2. `list_arena_models` — find the `.doe` file.
3. `audit_arena_model_data` — the coverage gate: what the model contains and
   what cannot be represented yet. Widen with `include_vba_source`,
   `include_siman_source`, `include_binary_payloads` as needed.
4. `extract_arena_model` — the neutral IR to translate (IR schema `0.3.0`,
   audit schema `2.0.0`, reported in every response).

Then zoom in: `inspect_arena_model` (summary), `list_arena_modules` /
`list_arena_connections` (paged detail), `analyze_arena_model_compatibility`
(automatic / assisted / manual per definition), `inspect_arena_project_bar`,
`extract_arena_submodels`, `extract_arena_visual_model`,
`extract_arena_material_handling`, `inspect_arena_compound_file`,
`extract_arena_siman_source`, and `inspect_arena_results` /
`read_arena_results` (sections: `project`, `output`, `continuous`,
`counter`, `discrete`, `frequency`) for result databases.

## Rules

- Read-only: never ask Arena to save, run, or modify a model.
- Pin downstream behavior on the reported schema versions.
- Large models: page `list_arena_modules` (`offset`/`limit`) and cap repeat
  rows / audit items instead of re-running full extracts.
