import assert from "node:assert/strict";
import test from "node:test";

import { blocksToMarkdown, markdownToBlocks, normalizeTreeIndent } from "../lib/markdown.js";
import { normalizeConfig } from "../lib/config.js";

test("normalizeTreeIndent turns a flat number of spaces into U+3000 levels", () => {
  assert.equal(normalizeTreeIndent("   ├─ a", 3), "\u3000├─ a");
  assert.equal(normalizeTreeIndent("      └─ a1", 3), "\u3000\u3000└─ a1");
  assert.equal(normalizeTreeIndent("         └─ a2", 3), "\u3000\u3000\u3000└─ a2");
  // Already-normalised lines are untouched (idempotent).
  assert.equal(normalizeTreeIndent("\u3000\u3000└─ x", 3), "\u3000\u3000└─ x");
  assert.equal(normalizeTreeIndent("root", 3), "root");
});

test("normalizeTreeIndent honours a 2-space source", () => {
  assert.equal(normalizeTreeIndent("  ├─ a", 2), "\u3000├─ a");
  assert.equal(normalizeTreeIndent("    └─ b", 2), "\u3000\u3000└─ b");
});

test("tree blocks recognise the indentation step of the whole fence", () => {
  // Widths 2 and 6 -> step gcd(2, 6) = 2 -> one and three levels.
  const md = "```tree\nroot\n  ├─ a\n      └─ a1\n```\n";
  const blocks = markdownToBlocks(md);
  assert.deepEqual(blocks.map((b) => b.text), ["root", "\u3000├─ a", "\u3000\u3000\u3000└─ a1"]);
});

test("tree blocks keep their fence on read-back", () => {
  const blocks = [
    { type: "tree", text: "\u3000├─ a", level: 0 },
    { type: "tree", text: "\u3000\u3000└─ a1", level: 0 },
  ];
  const md = blocksToMarkdown(blocks, { codePrefix: "" });
  assert.match(md, /^```tree\n/);
  assert.ok(md.includes("\u3000├─ a"));
  assert.deepEqual(markdownToBlocks(md).map((b) => b.text), ["\u3000├─ a", "\u3000\u3000└─ a1"]);
});

test("config: treeStyle defaults to paragraph, treeIndentStep to 3", () => {
  const defaults = normalizeConfig({});
  assert.equal(defaults.treeStyle, "paragraph");
  assert.equal(defaults.treeIndentStep, 3);
  assert.equal(normalizeConfig({ treeStyle: "mono" }).treeStyle, "mono");
  assert.equal(normalizeConfig({ treeStyle: "weird" }).treeStyle, "paragraph");
  assert.equal(normalizeConfig({ treeIndentStep: "2" }).treeIndentStep, 2);
  assert.equal(normalizeConfig({ treeIndentStep: 99 }).treeIndentStep, 3);
});
