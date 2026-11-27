import test from "node:test";
import assert from "node:assert/strict";

import { markdownToBlocks, blocksToMarkdown } from "../lib/markdown.js";

test("markdownToBlocks keeps code blocks when codePrefix is empty", () => {
  const blocks = markdownToBlocks("```js\nint a = 1;  // first\nlong b = 2; // second\n```\n", { codePrefix: "" });
  const code = blocks.filter((block) => block.type === "code");
  assert.equal(code.length, 2);
  assert.equal(code[0].text, "int a = 1;  // first");
  assert.equal(code[1].text, "long b = 2; // second");
});

test("blocksToMarkdown fences code blocks that carry no gutter marker", () => {
  const md = blocksToMarkdown([
    { type: "code", text: "int a = 1;  // first", level: 0 },
    { type: "code", text: "long b = 2; // second", level: 0 },
    { type: "para", text: "after", level: 0 },
  ]);
  assert.match(md, /```/);
  assert.match(md, /int a = 1;  \/\/ first/);
  assert.ok(!/^## /m.test(md), "code lines must not be rendered as headings");
});

test("blocksToMarkdown still strips the gutter when it is present", () => {
  const md = blocksToMarkdown([
    { type: "para", text: "│ int a = 1;  // first", level: 0 },
    { type: "para", text: "│ long b = 2; // second", level: 0 },
  ]);
  assert.match(md, /```/);
  assert.match(md, /int a = 1;/);
  assert.ok(!md.includes("│ int a"));
});
