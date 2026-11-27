import assert from "node:assert/strict";
import test from "node:test";

import { OneNoteBridge } from "../lib/bridge.js";

test("bridge normalises per-level bullet glyphs", () => {
  const bridge = new OneNoteBridge({ bulletValues: [20, "2", -1, "x"] });
  assert.deepEqual(bridge.bulletValues, [20, 2]);
});

test("bridge falls back to a single bullet when no list is given", () => {
  const bridge = new OneNoteBridge({ bulletValue: 6 });
  assert.deepEqual(bridge.bulletValues, []);
  assert.equal(bridge.bulletValue, 6);
});

test("bridge accepts bullet indexes above 9", () => {
  const bridge = new OneNoteBridge({ bulletValue: 20 });
  assert.equal(bridge.bulletValue, 20);
});

test("bridge normalises headingDelta and bulletFontSize", () => {
  const custom = new OneNoteBridge({ headingDelta: [4, 2], bulletFontSize: 11 });
  assert.deepEqual(custom.headingDelta, { h2: 4, h3: 2 });
  assert.equal(custom.bulletFontSize, 11);

  const fallback = new OneNoteBridge({});
  assert.deepEqual(fallback.headingDelta, { h2: 2, h3: 1 });
  assert.equal(fallback.bulletFontSize, "body");

  const object = new OneNoteBridge({ headingDelta: { h2: 3, h3: 1 } });
  assert.deepEqual(object.headingDelta, { h2: 3, h3: 1 });
});
