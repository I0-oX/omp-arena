# omp-things · Arena para OMP (compatible con Linux)

Marketplace OMP con un plugin que expone modelos de simulación Rockwell
Automation **Arena** (`.doe`) como 16 tools de solo lectura. Port del
servidor MCP [jofongang/arena-mcp](https://github.com/jofongang/arena-mcp)
(MIT; licencia original en `LICENSE`, copia intacta en
`plugins/arena-mcp/LICENSE.upstream`).

Mismo flujo, mismos esquemas: `arena_status` → `audit_arena_model_data` →
`extract_arena_model` (IR `0.3.0`, audit `2.0.0`). Nada guarda, ejecuta ni
edita modelos.

## Compatible con Linux

El extractor original solo importaba en Windows (`winreg`, `pywin32`). Aquí
esos imports son perezosos: el módulo importa en cualquier SO y cada tool
elige qué necesita.

| Tool | Linux | Windows + Arena |
| --- | :---: | :---: |
| `arena_status` (sin `live_check`) | ✅ | ✅ |
| `list_arena_models` | ✅ | ✅ |
| `inspect_arena_compound_file` (requiere `olefile`) | ✅ | ✅ |
| `inspect_arena_results`, `read_arena_results` | ✅ | ✅ |
| Resto (COM: inspect, modules, connections, extract, audit, …) | ❌ error claro | ✅ |

Fuera de Windows, las tools COM devuelven `ArenaExtractorError` explicando
que hace falta Windows + Arena con licencia, en vez de romper el import.

## Requisitos

- Python 3.10+ (`ARENA_PYTHON` para elegir intérprete, `python3`→`python` con fallback).
- Modo extensión: sin dependencias (usa `server/omp_bridge.py`).
- Modo servidor MCP: `pip install -r plugins/arena-mcp/server/requirements.txt`
  (`mcp`, `olefile`; `pywin32` solo se instala en Windows).
- Tools COM: Windows + Arena con licencia (`Arena.Application`).

## Instalación en OMP

```sh
/marketplace add ./path/to/omp-things
/marketplace install arena@omp-arena
```

Sin instalar, para desarrollo:

```sh
omp --extension ./plugins/arena-mcp
```

Reinicia la sesión tras instalar (`/reload-plugins` refresca skills,
comandos y MCP; las extensiones nuevas piden reinicio).

## Variables de entorno

- `ARENA_MODEL_ROOTS` — raíces de búsqueda separadas por `os.pathsep`
  (defecto: Documents/Desktop/OneDrive + carpeta pública de Rockwell).
- `ARENA_ALLOW_ANY_PATH=1` — permite rutas fuera de las raíces.
- `ARENA_PYTHON` — intérprete de las tools.
- `ARENA_BRIDGE_PATH` — override de `server/omp_bridge.py`.

## Verificar (Linux)

```sh
python3 -m py_compile plugins/arena-mcp/server/arena_extractor.py plugins/arena-mcp/server/omp_bridge.py
python3 -m unittest discover -s tests
python3 plugins/arena-mcp/server/arena_extractor.py --status
python3 plugins/arena-mcp/server/omp_bridge.py --json-call '{"tool":"arena_status","args":{}}'
```

## Estructura

```
.omp-plugin/marketplace.json     catálogo (marketplace "omp-arena")
plugins/arena-mcp/
  package.json                   omp.extensions → src/extension.ts
  .mcp.json                      declaración del servidor MCP
  src/extension.ts               16 tools LLM + comando /arena-setup
  server/arena_extractor.py      upstream vendored + imports perezosos
  server/omp_bridge.py           dispatcher JSON para la extensión
  server/requirements.txt
  skills/arena/SKILL.md          workflow del agente
  commands/arena-{status,audit,extract}.md
  agents/arena-translator.md
tests/test_arena_extractor.py    tests upstream (lógica pura, sin COM)
```
