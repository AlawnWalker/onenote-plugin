import assert from "node:assert/strict";
import test from "node:test";
import { applyWith, name } from "../dsh/index.js";
import { TOOL_NAMES, TOOL_REQUIRED } from "../lib/schemas.js";

const bridge = { invoke: async () => ({}) };

test("package identity and tool count", () => {
  assert.equal(name, "onenote-plugin");
  assert.equal(TOOL_NAMES.length, 11);
});

test("applyWith registers all tools with DSH-shaped definitions", () => {
  const registered = [];
  const result = applyWith({
    defineTool: (def) => def,
    ctx: { tools: { register: (def) => registered.push(def) } },
    config: {},
    bridge,
  });
  assert.equal(registered.length, 11);
  assert.deepEqual(result.tools, [...TOOL_NAMES]);
  for (const def of registered) {
    assert.equal(typeof def.execute, "function");
    assert.equal(typeof def.output.render, "function");
    assert.equal(def.output.schema.type, "object");
    const required = TOOL_REQUIRED[def.name] ?? [];
    for (const [key, param] of Object.entries(def.parameters)) {
      assert.equal(param.required === true, required.includes(key), `${def.name}.${key}`);
    }
  }
});

test("execute delegates to the shared handlers", async () => {
  const registered = [];
  applyWith({
    defineTool: (def) => def,
    ctx: { tools: { register: (def) => registered.push(def) } },
    config: {},
    bridge,
  });
  const list = registered.find((def) => def.name === "onenote_list");
  const result = await list.execute({});
  assert.equal(result.ok, true);
});
