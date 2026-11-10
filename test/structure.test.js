import assert from "node:assert/strict";
import test from "node:test";

import { markdownToBlocks } from "../lib/markdown.js";

const list = [
  "### 1. 标题",
  "",
  "- 第一级内容",
  "  数据流（四步）：",
  "- 第二条",
  "",
].join("\n");

test("indentedChildren: an indented line after a list item becomes its child", () => {
  const blocks = markdownToBlocks(list, { indentedChildren: true });
  const [heading, bullet, para] = blocks;
  assert.equal(heading.type, "h3");
  assert.equal(bullet.type, "bullet");
  assert.equal(bullet.level, 0);
  assert.equal(para.type, "para");
  assert.equal(para.level, 1, "nested paragraph is one level below the list item");
  assert.equal(para.text, "数据流（四步）：");
});

test("indentedChildren: deeper list items keep their own level", () => {
  const markdown = ["- 一级", "    1. 步骤一", "    2. 步骤二", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { indentedChildren: true });
  const numbers = blocks.filter((block) => block.type === "number");
  assert.equal(numbers.length, 2);
  assert.deepEqual(numbers.map((block) => block.level), [2, 2]);
});

test("indentedChildren: off keeps the old flattened paragraph", () => {
  const blocks = markdownToBlocks(list);
  const paras = blocks.filter((block) => block.type === "para");
  assert.equal(paras.length, 1);
  assert.equal(paras[0].level, 0);
  assert.equal(paras[0].text, "　数据流（四步）：", "indentation is kept as an EM SPACE");
});

test("indentedChildren: a heading resets the list context", () => {
  const markdown = ["- 一级", "", "### 新小节", "", "  后续段落", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { indentedChildren: true });
  const para = blocks.find((block) => block.type === "para");
  assert.equal(para.level, 0, "no list item above it, so it stays top level");
  assert.equal(para.text, "　后续段落");
});

test("indentedChildren: tree drawings keep the EM SPACE treatment", () => {
  const markdown = ["- 一级", "  ├─ 子节点", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { indentedChildren: true });
  const para = blocks.find((block) => block.type === "para");
  assert.equal(para.level, 0);
  assert.equal(para.text, "　├─ 子节点");
});
