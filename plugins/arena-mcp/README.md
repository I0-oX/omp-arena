# Arena plugin for OMP

Read-only extractor for Rockwell Automation **Arena** simulation models
(`.doe`) — an OMP port of
[jofongang/arena-mcp](https://github.com/jofongang/arena-mcp) (MIT, see
`LICENSE.upstream`). The server asks Arena itself to open models through its
Windows COM object model. It **never saves, runs, or edits a model**.

Same 16 tools, same workflow (`arena_status` → `audit_arena_model_data` →
`extract_arena_model`), same schema versions (IR `0.3.0`, audit `2.0.0`).

## Layout

```
.omp-plugin/marketplace.json      marketplace catalog (add with /marketplace add ./...)
plugins/arena-mcp/
  package.json                    omp.extensions → src/extension.ts
  .mcp.json                       MCP server declaration (Windows host)
  src/extension.ts                16 LLM tools + /arena-setup (zero-config)
  server/arena_extractor.py       vendored upstream + lazy Windows imports
  server/omp_bridge.py            JSON bridge used by the extension tools
  server/requirements.txt
  skills/arena/SKILL.md           workflow for the agent
  commands/arena-{status,audit,extract}.md
  agents/arena-translator.md
```

## Install

**Marketplace (recomendado):**

```
/marketplace add ./path/to/omp-things
/marketplace install arena@omp-arena
```

**Dev local sin instalar:**

```sh
omp --extension ./plugins/arena-mcp
```

Reinicia la sesión tras instalar para cargar tools/hooks/extensiones nuevas
(`/reload-plugins` refresca skills, comandos y MCP).

## Requisitos

- Windows + Arena con licencia (`Arena.Application`) para las tools COM.
- Python 3.10+; `pip install -r plugins/arena-mcp/server/requirements.txt`
  en modo MCP-server. Las tools de la extensión funcionan sin el paquete
  `mcp` (usan `server/omp_bridge.py`).
- Sin Arena también funcionan: `arena_status` (sin `live_check`),
  `list_arena_models`, `inspect_arena_compound_file`, `inspect_arena_results`,
  `read_arena_results`.

## Configuración

- `ARENA_MODEL_ROOTS` — raíces de búsqueda separadas por `os.pathsep`.
- `ARENA_ALLOW_ANY_PATH=1` — permite rutas fuera de las raíces.
- `ARENA_PYTHON` — intérprete Python de las tools (defecto `python3`).
- `ARENA_BRIDGE_PATH` — override de `server/omp_bridge.py`.

**MCP manual** (si prefieres el servidor MCP al modo extensión): `/arena-setup`
imprime el snippet con rutas absolutas para `.omp/mcp.json`. O edita
`.mcp.json` (`${CLAUDE_PLUGIN_ROOT}` se resuelve en clientes compatibles).

## Upstream

Lógica de extracción vendida de `arena-mcp` con un único cambio: los imports
de Windows (`winreg`, `pywin32`) y `mcp` son perezosos para que el módulo
importe en Linux/Mac y las tools no-COM sigan útiles. Sin cambios de
comportamiento en Windows.
