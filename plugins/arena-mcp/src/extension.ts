import type { ExtensionAPI } from "@oh-my-pi/pi-coding-agent";
import { fileURLToPath } from "node:url";

const IR_SCHEMA = "0.3.0";
const AUDIT_SCHEMA = "2.0.0";

function pythonCandidates(): string[] {
  const override = process.env.ARENA_PYTHON?.trim();
  if (override) return [override];
  return ["python3", "python"];
}

function bridgePath(): string {
  const override = process.env.ARENA_BRIDGE_PATH?.trim();
  if (override) return override;
  return fileURLToPath(new URL("../server/omp_bridge.py", import.meta.url));
}

async function callArena(
  pi: ExtensionAPI,
  tool: string,
  args: Record<string, unknown>,
  signal: AbortSignal | undefined,
  onUpdate: unknown,
): Promise<{ content: Array<{ type: "text"; text: string }>; details: unknown; isError?: boolean }> {
  (onUpdate as ((u: unknown) => void) | undefined)?.({
    content: [{ type: "text", text: `Arena: running ${tool}…` }],
  });
  let proc: { code: number; stdout: string; stderr: string } | undefined;
  let launchedWith = "";
  let lastError = "";
  for (const bin of pythonCandidates()) {
    try {
      proc = (await pi.exec(bin, [bridgePath(), "--json-call", JSON.stringify({ tool, args })], {
        signal,
      })) as { code: number; stdout: string; stderr: string };
      launchedWith = bin;
      break;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
  }
  if (!proc) {
    const tried = pythonCandidates().join(", ");
    return {
      content: [
        {
          type: "text",
          text: `Arena bridge failed to launch (tried ${tried}): ${lastError}. Set ARENA_PYTHON to a Python 3.10+ interpreter with the plugin requirements installed.`,
        },
      ],
      details: { tool, launched: false, launchedWith, error: lastError },
      isError: true,
    };
  }
  const raw = (proc.stdout ?? "").trim();
  let payload: { ok?: boolean; result?: unknown; error?: string };
  try {
    payload = JSON.parse(raw || "{}");
  } catch {
    const stderr = (proc.stderr ?? "").trim().slice(0, 500);
    return {
      content: [{ type: "text", text: `Arena bridge returned non-JSON output. ${stderr}` }],
      details: { tool, launched: true, code: proc.code, stdout: raw.slice(0, 500), stderr },
      isError: true,
    };
  }
  if (payload.ok !== true) {
    return {
      content: [{ type: "text", text: `Arena: ${tool} failed: ${payload.error ?? "unknown error"}` }],
      details: { tool, launched: true, error: payload.error },
      isError: true,
    };
  }
  return {
    content: [{ type: "text", text: JSON.stringify(payload.result, null, 2) }],
    details: { tool, result: payload.result },
  };
}

export default function arenaExtension(pi: ExtensionAPI) {
  const z = pi.zod;

  pi.setLabel("Arena");
  pi.on("session_start", async (_event, ctx) => {
    ctx.ui.notify("Arena plugin loaded (IR 0.3.0 / audit 2.0.0). Needs Windows + licensed Arena for COM tools.", "info");
  });

  const str = (desc: string) => z.string().describe(desc);
  const num = (desc: string) => z.number().describe(desc);
  const flag = (desc: string) => z.boolean().describe(desc);
  const tool = (
    name: string,
    description: string,
    parameters: ReturnType<typeof z.object>,
    argNames: string[],
  ) => {
    pi.registerTool({
      name,
      label: name,
      description,
      parameters,
      approval: "read",
      loadMode: "essential",
      async execute(_id, params, signal, onUpdate) {
        const args: Record<string, unknown> = {};
        for (const key of argNames) {
          const value = (params as Record<string, unknown>)[key];
          if (value !== undefined) args[key] = value;
        }
        return callArena(pi, name, args, signal, onUpdate);
      },
    });
  };

  tool("arena_status", "Check Arena registration, model roots, and optionally the live COM connection.", z.object({
    live_check: flag("Open Arena over COM to verify it answers. Default false.").optional(),
  }), ["live_check"]);

  tool("list_arena_models", "Discover .doe models under the configured read-only model roots.", z.object({
    root: str("Directory to search. Must be inside ARENA_MODEL_ROOTS unless ARENA_ALLOW_ANY_PATH=1.").optional(),
    include_backups: flag("Include *.backup.doe files. Default false.").optional(),
    limit: num("Max models. 1-2000, default 200.").optional(),
  }), ["root", "include_backups", "limit"]);

  tool("inspect_arena_model", "Model metadata, run settings, module counts, and compatibility.", z.object({
    model_path: str("Path to a .doe model file."),
  }), ["model_path"]);

  tool("list_arena_modules", "One page of Arena modules with scalar and repeat-group operands.", z.object({
    model_path: str("Path to a .doe model file."),
    offset: num("Page offset. Default 0.").optional(),
    limit: num("Page size. Default 100.").optional(),
    include_operands: flag("Include operand data. Default true.").optional(),
    max_repeat_rows: num("Max rows per repeat group. Default 250.").optional(),
  }), ["model_path", "offset", "limit", "include_operands", "max_repeat_rows"]);

  tool("list_arena_connections", "Directed connections with resolved source/destination modules.", z.object({
    model_path: str("Path to a .doe model file."),
  }), ["model_path"]);

  tool("extract_arena_model", `Full versioned neutral IR for translation (IR schema ${IR_SCHEMA}).`, z.object({
    model_path: str("Path to a .doe model file."),
    max_modules: num("Max modules. Default 1000.").optional(),
    max_repeat_rows: num("Max rows per repeat group. Default 250.").optional(),
  }), ["model_path", "max_modules", "max_repeat_rows"]);

  tool("analyze_arena_model_compatibility", "Classify module definitions as automatic / assisted / manual.", z.object({
    model_path: str("Path to a .doe model file."),
  }), ["model_path"]);

  tool("audit_arena_model_data", `Pre-translation coverage gate over every data source (audit schema ${AUDIT_SCHEMA}).`, z.object({
    model_path: str("Path to a .doe model file."),
    include_vba_source: flag("Capture VBA source (line-capped). Default false.").optional(),
    include_siman_source: flag("Generate SIMAN from a temp copy. Default false.").optional(),
    include_binary_payloads: flag("Include base64 OLE payloads. Default false.").optional(),
    max_modules: num("Max modules. Default 1000.").optional(),
    max_repeat_rows: num("Max rows per repeat group. Default 250.").optional(),
    max_audit_items: num("Max items per audit surface. Default 5000.").optional(),
    max_vba_lines: num("Max VBA lines. Default 10000.").optional(),
    max_siman_chars: num("Max SIMAN chars. Default 2000000.").optional(),
    max_binary_bytes: num("Max binary bytes. Default 2000000.").optional(),
  }), ["model_path", "include_vba_source", "include_siman_source", "include_binary_payloads", "max_modules", "max_repeat_rows", "max_audit_items", "max_vba_lines", "max_siman_chars", "max_binary_bytes"]);

  tool("inspect_arena_project_bar", "Full attached Project Bar panel and operand-definition schemas.", z.object({
    model_path: str("Path to a .doe model file."),
  }), ["model_path"]);

  tool("extract_arena_submodels", "Recursive submodel modules, operands, connections, and boundaries.", z.object({
    model_path: str("Path to a .doe model file."),
    max_items: num("Max items. Default 5000.").optional(),
    max_modules: num("Max modules. Default 1000.").optional(),
    max_repeat_rows: num("Max rows per repeat group. Default 250.").optional(),
  }), ["model_path", "max_items", "max_modules", "max_repeat_rows"]);

  tool("extract_arena_visual_model", "Shape geometry, animation records, pictures, optional raw payloads.", z.object({
    model_path: str("Path to a .doe model file."),
    max_items: num("Max items. Default 5000.").optional(),
    include_binary_payloads: flag("Include raw payloads. Default false.").optional(),
    max_binary_bytes: num("Max binary bytes. Default 2000000.").optional(),
  }), ["model_path", "max_items", "include_binary_payloads", "max_binary_bytes"]);

  tool("extract_arena_material_handling", "Material-handling collections and their readable properties.", z.object({
    model_path: str("Path to a .doe model file."),
    max_items: num("Max items. Default 5000.").optional(),
  }), ["model_path", "max_items"]);

  tool("inspect_arena_compound_file", "Every .doe OLE stream with hashes and optional base64 payloads.", z.object({
    model_path: str("Path to a .doe model file."),
    include_payloads: flag("Include base64 payloads. Default false.").optional(),
    max_payload_bytes: num("Max payload bytes. Default 2000000.").optional(),
  }), ["model_path", "include_payloads", "max_payload_bytes"]);

  tool("extract_arena_siman_source", "SIMAN text generated from an isolated temporary copy of the model.", z.object({
    model_path: str("Path to a .doe model file."),
    max_chars: num("Max chars. Default 2000000.").optional(),
  }), ["model_path", "max_chars"]);

  tool("inspect_arena_results", "Schema and row counts of an Arena SQLite results database.", z.object({
    database_path: str("Path to an Arena results .db/.sqlite/.sqlite3 file."),
  }), ["database_path"]);

  tool("read_arena_results", "Read one statistics section from an Arena SQLite database.", z.object({
    database_path: str("Path to an Arena results .db/.sqlite/.sqlite3 file."),
    section: str("project | output | continuous | counter | discrete | frequency. Default project.").optional(),
    limit: num("Max rows. 1-10000, default 100.").optional(),
  }), ["database_path", "section", "limit"]);

  pi.registerCommand("arena-setup", {
    description: "Print the .omp/mcp.json snippet that wires the Arena MCP server",
    handler: async (_args, ctx) => {
      const server = fileURLToPath(new URL("../server/arena_extractor.py", import.meta.url));
      const snippet = JSON.stringify(
        { mcpServers: { arena: { command: pythonCandidates()[0], args: [server] } } },
        null,
        2,
      );
      pi.sendMessage(
        { customType: "arena-setup", content: `Arena MCP server config (merge into .omp/mcp.json):\n\`\`\`json\n${snippet}\n\`\`\``, display: true, attribution: "user" },
        { triggerTurn: false },
      );
      ctx.ui.notify("Arena MCP snippet sent to the conversation.", "info");
    },
  });
}
