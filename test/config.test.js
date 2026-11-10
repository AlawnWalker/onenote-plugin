import assert from "node:assert/strict";
import test from "node:test";
import { normalizeConfig, readConfig } from "../lib/config.js";

test("normalizeConfig defaults the 0.5.8 switches to false", () => {
  const cfg = normalizeConfig({});
  assert.equal(cfg.nestUnderHeadings, false);
  assert.equal(cfg.indentedChildren, false);
});

test("normalizeConfig keeps explicit true switches", () => {
  const cfg = normalizeConfig({ nestUnderHeadings: true, indentedChildren: true });
  assert.equal(cfg.nestUnderHeadings, true);
  assert.equal(cfg.indentedChildren, true);
});

test("readConfig maps the new ONENOTE_* env switches", () => {
  const cfg = readConfig({ ONENOTE_NEST_UNDER_HEADINGS: "1", ONENOTE_INDENTED_CHILDREN: "true" });
  assert.equal(cfg.nestUnderHeadings, true);
  assert.equal(cfg.indentedChildren, true);
});
