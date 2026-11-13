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

test("headingDelta defaults to [2,1] and normalises arrays/objects/JSON", () => {
  assert.deepEqual(normalizeConfig({}).headingDelta, [2, 1]);
  assert.deepEqual(normalizeConfig({ headingDelta: [4, 2] }).headingDelta, [4, 2]);
  assert.deepEqual(normalizeConfig({ headingDelta: { h2: 3, h3: 1 } }).headingDelta, [3, 1]);
  assert.deepEqual(normalizeConfig({ headingDelta: "[1,0]" }).headingDelta, [1, 0]);
  // out-of-range / non-numeric values fall back per entry
  assert.deepEqual(normalizeConfig({ headingDelta: ["x", -1] }).headingDelta, [2, 1]);
});

test("bulletFontSize defaults to body and accepts positive point sizes", () => {
  assert.equal(normalizeConfig({}).bulletFontSize, "body");
  assert.equal(normalizeConfig({ bulletFontSize: 12 }).bulletFontSize, 12);
  assert.equal(normalizeConfig({ bulletFontSize: "11" }).bulletFontSize, 11);
  assert.equal(normalizeConfig({ bulletFontSize: 0 }).bulletFontSize, "body");
});

test("readConfig maps ONENOTE_HEADING_DELTA / ONENOTE_BULLET_FONT_SIZE", () => {
  const cfg = readConfig({ ONENOTE_HEADING_DELTA: "[3,1]", ONENOTE_BULLET_FONT_SIZE: "12" });
  assert.deepEqual(cfg.headingDelta, [3, 1]);
  assert.equal(cfg.bulletFontSize, 12);
});
