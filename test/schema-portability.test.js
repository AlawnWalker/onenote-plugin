import assert from "node:assert/strict";
import test from "node:test";
import { TOOL_NAMES, TOOL_REQUIRED, buildJsonSchema, toDshParameters, assertPortableSchema } from "../lib/schemas.js";

test("every tool input schema passes the DSH subset with an object root", () => {
  for (const name of TOOL_NAMES) {
    assert.doesNotThrow(() => assertPortableSchema(buildJsonSchema(name), { requireObjectRoot: true }), name);
  }
});

test("toDshParameters marks required properties", () => {
  const params = toDshParameters("onenote_create");
  assert.equal(params.markdown.required, true);
  assert.notEqual(params.title?.required, true);
});

test("toDshParameters covers every parameter of every tool", () => {
  for (const name of TOOL_NAMES) {
    const params = toDshParameters(name);
    assert.deepEqual(Object.keys(params).sort(), Object.keys(buildJsonSchema(name).properties).sort());
    for (const [key, value] of Object.entries(params)) {
      if ((TOOL_REQUIRED[name] ?? []).includes(key)) assert.equal(value.required, true, `${name}.${key}`);
    }
  }
});

test("unsupported keywords are rejected", () => {
  assert.throws(() => assertPortableSchema({ type: "string", pattern: "^a" }), /pattern/);
  assert.throws(() => assertPortableSchema({ type: "number", minimum: 0 }), /minimum/);
  assert.throws(() => assertPortableSchema({ type: "array", items: [{ type: "string" }] }), /items/);
});

test("type arrays, type+oneOf, and bad oneOf siblings are rejected", () => {
  assert.throws(() => assertPortableSchema({ type: ["string", "null"] }), /type/);
  assert.throws(() => assertPortableSchema({ type: "object", oneOf: [{ type: "string" }, { type: "number" }] }), /oneOf/);
  assert.throws(() => assertPortableSchema({
    oneOf: [{ type: "object", properties: { a: { type: "string" } } }, { type: "string" }],
    properties: { b: { type: "string" } },
  }), /properties/);
});

test("required must reference declared properties", () => {
  assert.throws(() => assertPortableSchema({ type: "object", properties: { a: { type: "string" } }, required: ["b"] }), /required/);
});

test("tool input root must be an object when required", () => {
  assert.throws(() => assertPortableSchema({ type: "string" }, { requireObjectRoot: true }), /object/);
});
