"""JSON bridge between the OMP Arena extension and the vendored extractor.

Protocol: ``omp_bridge.py --json-call '{"tool": "<name>", "args": {...}}'``
prints exactly one JSON document to stdout::

    {"ok": true, "result": {...}}
    {"ok": false, "error": "..."}

Only the 16 read-only arena tools are exposed. Unknown tools and unexpected
argument names are rejected before anything runs. Arena/COM tools raise a
clear ``ArenaExtractorError`` on machines without Arena instead of an
import-time crash (see the lazy Windows imports in ``arena_extractor``).
"""

from __future__ import annotations

import inspect
import json
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

import arena_extractor as extractor

_TOOLS: dict[str, str] = {
    "arena_status": "get_arena_status",
    "list_arena_models": "discover_models",
    "inspect_arena_model": "inspect_model",
    "list_arena_modules": "extract_modules",
    "list_arena_connections": "extract_connections",
    "extract_arena_model": "extract_model_ir",
    "analyze_arena_model_compatibility": "analyze_compatibility",
    "audit_arena_model_data": "audit_model_data",
    "inspect_arena_project_bar": "extract_project_bar_catalog",
    "extract_arena_submodels": "extract_submodel_tree",
    "extract_arena_visual_model": "extract_visual_model",
    "extract_arena_material_handling": "extract_material_handling",
    "inspect_arena_compound_file": "inspect_compound_file",
    "extract_arena_siman_source": "extract_siman_source",
    "inspect_arena_results": "inspect_results_database",
    "read_arena_results": "read_results",
}


def _call(tool: str, args: dict[str, Any]) -> Any:
    if tool not in _TOOLS:
        raise ValueError(
            f"Unknown tool {tool!r}; choose from {', '.join(sorted(_TOOLS))}."
        )
    fn = getattr(extractor, _TOOLS[tool])
    if not isinstance(args, dict):
        raise ValueError("'args' must be an object.")
    accepted = inspect.signature(fn).parameters
    unexpected = [key for key in args if key not in accepted]
    if unexpected:
        raise ValueError(
            f"Unexpected argument(s) for {tool}: {', '.join(unexpected)}."
        )
    return fn(**args)


def _main(argv: list[str]) -> int:
    if len(argv) != 2 or argv[0] != "--json-call":
        sys.stderr.write(
            "Usage: omp_bridge.py --json-call '{\"tool\": \"<name>\", \"args\": {...}}'\n"
        )
        return 2
    try:
        payload = json.loads(argv[1])
        result = _call(payload.get("tool", ""), payload.get("args", {}))
        sys.stdout.write(json.dumps({"ok": True, "result": result}, ensure_ascii=True))
    except Exception as error:  # bridge contract: errors are data, not tracebacks
        text = str(error).replace("\r", " ").replace("\n", " ").strip()
        sys.stdout.write(
            json.dumps(
                {"ok": False, "error": (text[:2000] or type(error).__name__)},
                ensure_ascii=True,
            )
        )
        return 0
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main(sys.argv[1:]))
