import assert from "node:assert/strict";
import test from "node:test";

import { TOOL_NAMES, TOOL_PARAMETERS, TOOL_REQUIRED, TOOL_DESCRIPTIONS, buildJsonSchema } from "../lib/schemas.js";

test("the catalog exposes all 13 onenote tools", () => {
  assert.deepEqual(TOOL_NAMES, [
    "onenote_list",
    "onenote_read",
    "onenote_search",
    "onenote_create",
    "onenote_update",
    "onenote_append",
    "onenote_edit",
    "onenote_insert_image",
    "onenote_normalize_tables",
    "onenote_publish",
    "onenote_delete",
    "onenote_create_section",
    "onenote_rename_section",
  ]);
});

test("every tool has parameters, required names and a description", () => {
  for (const name of TOOL_NAMES) {
    assert.ok(TOOL_PARAMETERS[name], `${name} has parameters`);
    assert.ok(Array.isArray(TOOL_REQUIRED[name]), `${name} has a required list`);
    assert.ok(typeof TOOL_DESCRIPTIONS[name] === "string", `${name} has a description`);
  }
});

test("buildJsonSchema emits standard JSON Schema", () => {
  const schema = buildJsonSchema("onenote_search");
  assert.equal(schema.type, "object");
  assert.equal(typeof schema.properties.query, "object");
  assert.equal(schema.properties.query.type, "string");
  assert.deepEqual(schema.required, ["query"]);
});

test("tools with no required params advertise an empty required list", () => {
  const schema = buildJsonSchema("onenote_list");
  assert.equal(schema.type, "object");
  assert.deepEqual(schema.required, []);
  assert.equal(schema.properties.includePages.type, "boolean");
  assert.equal(schema.properties.maxSections.type, "integer");
});

test("page locator params are shared across read/update/append/publish/delete", () => {
  for (const name of ["onenote_read", "onenote_update", "onenote_append", "onenote_publish", "onenote_delete"]) {
    for (const key of ["pageId", "pageTitle", "sectionId", "sectionName"]) {
      assert.ok(TOOL_PARAMETERS[name][key], `${name} has ${key}`);
    }
  }
});

test("known limitations are folded into the tool descriptions", () => {
  assert.match(TOOL_DESCRIPTIONS.onenote_create, /Number|bullet/i);
  assert.match(TOOL_DESCRIPTIONS.onenote_update, /outlines|images/i);
  assert.match(TOOL_DESCRIPTIONS.onenote_publish, /backslash/i);
});
