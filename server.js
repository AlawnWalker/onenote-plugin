/**
 * onenote-mcp — stdio MCP server for the local Microsoft OneNote desktop app.
 *
 * Wraps the onenote bridge: every tool call writes a JSON request, spawns a
 * short-lived Windows PowerShell 5.1 process that drives OneNote over COM, and
 * reads the JSON response back. The MCP layer uses the low-level Server API so
 * our own error codes surface directly as isError results.
 */
import { pathToFileURL } from "node:url";

import { Server } from "@modelcontextprotocol/server";
import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";

import { OneNoteBridge } from "./lib/bridge.js";
import { makeToolHandlers, buildIdList } from "./lib/tools.js";
import { readConfig } from "./lib/config.js";
import { TOOL_NAMES, TOOL_DESCRIPTIONS, buildJsonSchema } from "./lib/schemas.js";

const SERVER_INFO = { name: "onenote", version: "0.1.0" };

/** Render a handler result into an MCP CallToolResult. */
function toCallResult(args, result) {
  if (!result || result.ok !== true) {
    const code = result?.code ?? "ONENOTE_ERROR";
    const message = result?.message ?? "OneNote operation failed";
    return { content: [{ type: "text", text: `${code}: ${message}` }], isError: true };
  }
  const data = result.data ?? {};
  if (typeof data.markdown === "string") {
    const header = data.title ? `# ${data.title}\n\n` : "";
    const note = data.truncated ? `\n> ⚠ truncated: only the first ${data.blockCount} blocks shown.\n` : "";
    const footer = `\n---\npageId: ${data.pageId ?? ""}\nlastModified: ${data.lastModified ?? ""}`;
    return { content: [{ type: "text", text: `${header}${data.markdown}${note}${footer}${buildIdList(args, result)}` }] };
  }
  return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
}

/**
 * Build a configured MCP Server. `handlers` is the result of
 * makeToolHandlers({ bridge }); keeping it injectable lets the tests drive the
 * full protocol over an in-memory transport without touching OneNote.
 */
export function createMcpServer({ handlers }) {
  const server = new Server(SERVER_INFO, { capabilities: { tools: {} } });

  server.setRequestHandler("tools/list", async () => ({
    tools: TOOL_NAMES.map((name) => ({
      name,
      description: TOOL_DESCRIPTIONS[name],
      inputSchema: buildJsonSchema(name),
    })),
  }));

  server.setRequestHandler("tools/call", async (request) => {
    const name = request.params?.name;
    const args = request.params?.arguments ?? {};
    const handler = handlers[name];
    if (typeof handler !== "function") {
      return { content: [{ type: "text", text: `Unknown tool: ${name}` }], isError: true };
    }
    const result = await handler(args);
    return toCallResult(args, result);
  });

  return server;
}

async function main() {
  const config = readConfig(process.env);
  const bridge = new OneNoteBridge(config);
  const handlers = makeToolHandlers({
    bridge,
    renderOptions: {
      indentedChildren: config.indentedChildren,
      codePrefix: config.codePrefix,
      alignCommentColumn: config.alignCommentColumn,
      alignCommentMax: config.alignCommentMax,
    },
  });
  const server = createMcpServer({ handlers });
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

const isEntry = process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href;
if (isEntry) {
  main().catch((err) => {
    // eslint-disable-next-line no-console
    console.error(err && err.stack ? err.stack : String(err));
    process.exit(1);
  });
}
