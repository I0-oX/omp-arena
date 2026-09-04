# omp-arena · Arena for OMP (Linux-compatible)

OMP marketplace with a plugin that exposes Rockwell Automation **Arena**
simulation models (`.doe`) as 16 read-only tools. A port of the MCP server
[jofongang/arena-mcp](https://github.com/jofongang/arena-mcp) (MIT; original
license in `LICENSE`, untouched copy in
`plugins/arena-mcp/LICENSE.upstream`).

Same tools, same workflow: `arena_status` → `audit_arena_model_data` →
`extract_arena_model` (IR `0.3.0`, audit `2.0.0`). Nothing is ever saved,
run, or edited.

## Linux compatibility

Upstream only imported on Windows (`winreg`, `pywin32`). Here those imports
are lazy: the module imports on any OS and each tool takes only what it
needs.

| Tool | Linux | Windows + Arena |
| --- | :---: | :---: |
| `arena_status` (without `live_check`) | ✅ | ✅ |
| `list_arena_models` | ✅ | ✅ |
| `inspect_arena_compound_file` (needs `olefile`) | ✅ | ✅ |
| `inspect_arena_results`, `read_arena_results` | ✅ | ✅ |
| Rest (COM: inspect, modules, connections, extract, audit, …) | ❌ clear error | ✅ |

Off Windows, COM tools return an `ArenaExtractorError` explaining that
Windows + licensed Arena is required, instead of breaking the import.

## Requirements

- Python 3.10+ (`ARENA_PYTHON` to pick the interpreter, `python3`→`python`
  fallback).
- Extension mode: no dependencies (uses `server/omp_bridge.py`).
- MCP-server mode: `pip install -r plugins/arena-mcp/server/requirements.txt`
  (`mcp`, `olefile`; `pywin32` installs on Windows only).
- COM tools: Windows + licensed Arena (`Arena.Application`).

## Use with a model

Arena is not a model and can never appear in `omp models` or `--model` —
it has no chat API. The plugin attaches its 16 read-only tools
(`loadMode: essential`, `approval: read`) to whichever real model you run:

```sh
omp --model anthropic/claude-sonnet-4-5
# then: /arena-status, or "audit this .doe model ..."
```

`examples/models.yml.example` shows a typical model setup to copy into
`~/.omp/agent/models.yml`.

### Selectable `arena` provider

A provider must terminate in a real LLM chat API, and Arena has none — so a
standalone `arena` provider entry would fail every call. The honest
equivalent: `examples/arena-provider.models.yml.example` defines an `arena`
provider id backed by a real backend (local Ollama by default). Copy it into
`~/.omp/agent/models.yml` and its models appear as `arena/<id>` in `/model`
and `--model`, with the 16 Arena tools attached.

## Install from GitHub

```
/marketplace add I0-oX/omp-arena
/marketplace install arena@omp-arena
```

CLI equivalent:

```sh
omp plugin marketplace add I0-oX/omp-arena
omp plugin install arena@omp-arena
```

For local development without installing:

```sh
git clone https://github.com/I0-oX/omp-arena.git
omp --extension ./omp-arena/plugins/arena-mcp
```

Restart the session after installing (`/reload-plugins` refreshes skills,
commands, and MCP; new extensions need a restart).

## Try it on Linux (no Arena needed)

The portable tools work end-to-end here. A sample results DB is included —
regenerate it anytime with `python3 tests/fixtures/make_sample_results.py`:

```sh
export ARENA_ALLOW_ANY_PATH=1
python3 plugins/arena-mcp/server/omp_bridge.py \
  --json-call '{"tool":"inspect_arena_results","args":{"database_path":"tests/fixtures/sample_results.db"}}'
python3 plugins/arena-mcp/server/omp_bridge.py \
  --json-call '{"tool":"read_arena_results","args":{"database_path":"tests/fixtures/sample_results.db","section":"project"}}'
ARENA_MODEL_ROOTS=/tmp/arena_demo python3 plugins/arena-mcp/server/omp_bridge.py \
  --json-call '{"tool":"list_arena_models","args":{"limit":50}}'
```

COM tools (`audit_arena_model_data`, `extract_arena_model`, …) need a
Windows host with licensed Arena and a real `.doe` path — on Linux they
answer with a clear `ArenaExtractorError` instead of failing silently.

## Environment variables

- `ARENA_MODEL_ROOTS` — `os.pathsep`-separated search roots (default:
  Documents/Desktop/OneDrive + Rockwell public folder).
- `ARENA_ALLOW_ANY_PATH=1` — allow paths outside the roots.
- `ARENA_PYTHON` — interpreter used by the tools.
- `ARENA_BRIDGE_PATH` — override for `server/omp_bridge.py`.

## Verify (Linux)

```sh
python3 -m py_compile plugins/arena-mcp/server/arena_extractor.py plugins/arena-mcp/server/omp_bridge.py
python3 -m unittest discover -s tests
python3 plugins/arena-mcp/server/arena_extractor.py --status
python3 plugins/arena-mcp/server/omp_bridge.py --json-call '{"tool":"arena_status","args":{}}'
```

## Layout

```
.omp-plugin/marketplace.json     catalog (marketplace "omp-arena")
plugins/arena-mcp/
  package.json                   omp.extensions → src/extension.ts
  .mcp.json                      MCP server declaration
  src/extension.ts               16 LLM tools + /arena-setup command
  server/arena_extractor.py      vendored upstream + lazy Windows imports
  server/omp_bridge.py           JSON dispatcher used by the extension
  server/requirements.txt
  skills/arena/SKILL.md          agent workflow
  commands/arena-{status,audit,extract}.md
  agents/arena-translator.md
tests/test_arena_extractor.py    upstream tests (pure logic, no COM)
```
