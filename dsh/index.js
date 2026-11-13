/**
 * onenote-plugin â€?DSH/Cordis adapter.
 *
 * Thin host layer over the shared core in ../lib. It registers the same 11
 * tools the MCP adapter exposes, but through the DSH author DSL so DSH can
 * render results and validate activation natively.
 */
import { OneNoteBridge } from "../lib/bridge.js";
import { makeToolHandlers, renderResult } from "../lib/tools.js";
import { TOOL_NAMES, TOOL_DESCRIPTIONS, toDshParameters } from "../lib/schemas.js";
import { normalizeConfig } from "../lib/config.js";

export const name = "onenote-plugin";
export const version = "0.6.6";
export const inject = ["tools"];

export const OUTPUT = Object.freeze({
  schema: { type: "object", additionalProperties: true },
  render: renderResult,
});

/** Read-only tools can run concurrently with anything else. */
const CONCURRENCY_SAFE = new Set(["onenote_list", "onenote_read", "onenote_search"]);

/** Wire the plugin against an injected `defineTool` and a Cordis-like `ctx`. */
export function applyWith({ defineTool, ctx, config = {}, bridge } = {}) {
  if (typeof defineTool !== "function") throw new Error("applyWith: defineTool is required");
  if (!ctx || !ctx.tools || typeof ctx.tools.register !== "function") {
    throw new Error("applyWith: ctx.tools.register is required");
  }

  const cfg = normalizeConfig(config);
  const instance = bridge ?? new OneNoteBridge(cfg);
  const handlers = makeToolHandlers({
    bridge: instance,
    renderOptions: {
      indentedChildren: cfg.indentedChildren,
      codePrefix: cfg.codePrefix,
      alignCommentColumn: cfg.alignCommentColumn,
      alignCommentMax: cfg.alignCommentMax,
    },
  });

  const registered = [];
  for (const toolName of TOOL_NAMES) {
    const def = defineTool({
      name: toolName,
      description: TOOL_DESCRIPTIONS[toolName],
      parameters: toDshParameters(toolName),
      output: OUTPUT,
      execute: (args = {}) => handlers[toolName](args),
      timeoutMs: cfg.timeoutMs,
      isConcurrencySafe: CONCURRENCY_SAFE.has(toolName),
    });
    ctx.tools.register(def);
    registered.push(toolName);
  }

  return { config: cfg, tools: registered };
}

/**
 * Real Cordis entry point. Cordis treats a plugin's return value as an effect,
 * so this must stay `undefined`.
 */
export async function apply(ctx, config = {}) {
  const dshTools = await import("@deepseek-ai/dsh-tools");
  applyWith({ defineTool: dshTools.defineTool, ctx, config });
}

export { OneNoteBridge };
