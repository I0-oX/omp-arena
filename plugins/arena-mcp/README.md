# Arena plugin for OMP

Read-only extractor for Rockwell Automation **Arena** simulation models
(`.doe`) — an OMP port of
[jofongang/arena-mcp](https://github.com/jofongang/arena-mcp) (MIT, see
`LICENSE.upstream`). The server asks Arena itself to open models through its
Windows COM object model. It **never saves, runs, or edits a model**.

Same 16 tools plus Linux-native `inspect_arena_native`, same workflow (`arena_status` → `audit_arena_model_data` →
`extract_arena_model`), same schema versions (IR `0.3.0`, audit `2.0.0`).

## Layout

```
.omp-plugin/marketplace.json      marketplace catalog (add with /marketplace add I0-oX/omp-arena)
plugins/arena-mcp/
  package.json                    omp.extensions → src/extension.ts
  .mcp.json                       MCP server declaration (Windows host)
  src/extension.ts                16 LLM tools + /arena-setup (zero-config)
  server/arena_extractor.py       vendored upstream + lazy Windows imports
  server/omp_bridge.py            JSON bridge used by the extension tools
  server/requirements.txt
  skills/arena/SKILL.md           agent workflow
  commands/arena-{status,audit,extract}.md
  agents/arena-translator.md
```

## Install

**Remote (recommended):**

```
/marketplace add I0-oX/omp-arena
/marketplace install arena@omp-arena
```

CLI equivalent:

```sh
omp plugin marketplace add I0-oX/omp-arena
omp plugin install arena@omp-arena
```

**Local dev without installing:**

```sh
omp --extension ./plugins/arena-mcp
```

Restart the session after installing so new tools/hooks/extensions load
(`/reload-plugins` refreshes skills, commands, and MCP).

## Requirements

- Windows + licensed Arena (`Arena.Application`) for the COM tools.
- Python 3.10+; `pip install -r plugins/arena-mcp/server/requirements.txt`
  in MCP-server mode. The extension tools work without the `mcp` package
  (they use `server/omp_bridge.py`).
- These also work without Arena (and `inspect_arena_native` works on any OS): `arena_status` (without `live_check`),
  `list_arena_models`, `inspect_arena_compound_file`, `inspect_arena_results`,
  `read_arena_results`.

## Configuration

- `ARENA_MODEL_ROOTS` — search roots separated by `os.pathsep`.
- `ARENA_ALLOW_ANY_PATH=1` — allow paths outside the roots.
- `ARENA_PYTHON` — Python interpreter for the tools (default `python3`).
- `ARENA_BRIDGE_PATH` — override for `server/omp_bridge.py`.

**Manual MCP** (if you prefer the MCP server over extension mode): `/arena-setup`
prints the snippet with absolute paths for `.omp/mcp.json`. Or edit
`.mcp.json` (`${CLAUDE_PLUGIN_ROOT}` resolves in compatible clients).

## Upstream

Extraction logic vendored from `arena-mcp` with a single change: the Windows
imports (`winreg`, `pywin32`) and `mcp` are lazy so the module imports on
Linux/macOS and the non-COM tools stay useful. No behavior change on
Windows.
