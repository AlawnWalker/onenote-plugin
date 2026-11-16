import assert from "node:assert/strict";
import test from "node:test";

import { alignCodeComments, blocksToMarkdown, displayWidth, findInlineComment, markdownToBlocks } from "../lib/markdown.js";

/**
 * Emulate what the bridge does to a block before OneNote stores it:
 * numbered items become bullets whose text carries the number, which is the
 * workaround for OneNote refusing `one:Number` imports on Office 2024.
 */
function throughOneNote(blocks) {
  return blocks.map((block) => {
    if (block.type === "number") {
      return { ...block, type: "bullet", text: `${block.ordinal ?? 1}. ${block.text}` };
    }
    return block;
  });
}

test("headings become typed blocks and round-trip", () => {
  const markdown = "# 页面标题\n\n正文一段。\n\n## 二级标题\n\n### 三级标题\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => block.type),
    ["title", "para", "h2", "h3"],
  );
  assert.equal(blocks[2].text, "二级标题");
  assert.equal(blocks[3].text, "三级标题");
  assert.equal(blocksToMarkdown(blocks), markdown);
});

test("legacy symbol headings still read back as markdown", () => {
  const blocks = [
    { type: "para", text: "━━━ 旧二级 ━━━", level: 0 },
    { type: "para", text: "▌ 旧三级", level: 0 },
  ];
  assert.equal(blocksToMarkdown(blocks), "## 旧二级\n\n### 旧三级\n");
});

test("nested bullet lists keep their levels", () => {
  const markdown = "- 一\n- 二\n  - 二点一\n    - 二点一之甲\n- 三\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => [block.type, block.level]),
    [
      ["bullet", 0],
      ["bullet", 0],
      ["bullet", 1],
      ["bullet", 2],
      ["bullet", 0],
    ],
  );
  assert.equal(blocksToMarkdown(blocks), markdown);
});

test("numbered lists survive the bullet workaround", () => {
  const markdown = "1. 第一步\n2. 第二步\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => [block.type, block.ordinal]),
    [
      ["number", 1],
      ["number", 2],
    ],
  );
  assert.equal(blocksToMarkdown(throughOneNote(blocks)), markdown);
});

test("tables render as a markdown table with separators", () => {
  const markdown = "| 列A | 列B |\n| --- | --- |\n| a1 | b1 |\n| a2 | b2 |\n";
  const blocks = markdownToBlocks(markdown);
  assert.equal(blocks.length, 1);
  assert.deepEqual(blocks[0].rows, [
    ["列A", "列B"],
    ["a1", "b1"],
    ["a2", "b2"],
  ]);
  const back = blocksToMarkdown([
    { type: "table", rows: [["列A", "列B"], ["a1", "b1"], ["a2", "b2"]], level: 0 },
  ]);
  assert.equal(back, markdown);
});

test("fenced code becomes gutter lines and back", () => {
  const markdown = "```js\nconst x = 1;\nconsole.log(x);\n```\n";
  const blocks = markdownToBlocks(markdown);
  assert.deepEqual(
    blocks.map((block) => [block.type, block.text]),
    [
      ["code", "│ const x = 1;"],
      ["code", "│ console.log(x);"],
    ],
  );
  // OneNote has nowhere to store the fence language, so it comes back bare.
  assert.equal(blocksToMarkdown(blocks), "```\nconst x = 1;\nconsole.log(x);\n```\n");
});

test("blank lines separate lists from the blocks that follow", () => {
  const blocks = [
    { type: "bullet", text: "一", level: 0 },
    { type: "bullet", text: "二", level: 0 },
    { type: "para", text: "━━━ 标题 ━━━", level: 0 },
    { type: "table", rows: [["a", "b"], ["1", "2"]], level: 0 },
    { type: "para", text: "正文", level: 0 },
  ];
  const markdown = blocksToMarkdown(blocks);
  assert.equal(markdown, "- 一\n- 二\n\n## 标题\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\n正文\n");
});

test("quotes and paragraph breaks are preserved", () => {
  const markdown = "> 一句话\n\n普通段落。\n";
  const blocks = markdownToBlocks(markdown);
  assert.equal(blocks[0].text, "▎ 一句话");
  assert.equal(blocksToMarkdown(blocks), markdown);
});

test("pipes inside table cells are escaped on the way out", () => {
  const markdown = blocksToMarkdown([{ type: "table", rows: [["a|b", "c"]], level: 0 }]);
  assert.match(markdown, /a\\\|b/);
});

test("indented tree lines become separate paragraphs with EM SPACE indents", () => {
  const blocks = markdownToBlocks(
    "kernel_init_freeable()\n  ├─ do_pre_smp_initcalls()\n  └─ do_basic_setup()\n        │    note\n",
  );
  assert.deepEqual(blocks.map((b) => b.type), ["para", "para", "para", "para"]);
  assert.equal(blocks[0].text, "kernel_init_freeable()");
  assert.equal(blocks[1].text, "　├─ do_pre_smp_initcalls()");
  assert.equal(blocks[2].text, "　└─ do_basic_setup()");
  assert.equal(blocks[3].text, "　　　　│    note");
});
test("displayWidth counts CJK as two columns", () => {
  assert.equal(displayWidth("abc"), 3);
  assert.equal(displayWidth("中文"), 4);
  assert.equal(displayWidth("a中b"), 4);
});

test("findInlineComment follows the trailing-comment rule", () => {
  assert.equal(findInlineComment("// whole line"), -1);
  assert.equal(findInlineComment("   // indented whole line"), -1);
  assert.equal(findInlineComment('url = "http://x"; '), -1);
  assert.equal(findInlineComment("x = 1; // note"), 7);
  assert.equal(findInlineComment("a//b"), -1, "needs whitespace before //");
  // last whitelisted // wins
  assert.equal(findInlineComment("a = b; // first // second"), 16, "last // wins");
});

test("alignCodeComments: auto column includes indentation and CJK width", () => {
  const lines = ["  int a = 1; // 一", "      long_call(x); // 二", "no comment"];
  const aligned = alignCodeComments(lines, 0);
  const widths = aligned.slice(0, 2).map((line) => displayWidth(line.slice(0, line.indexOf("//"))));
  assert.equal(widths[0], widths[1], "both comments start at the same column");
  assert.equal(aligned[2], "no comment", "lines without comments are untouched");
});

test("alignCodeComments: fixed column and the maxColumn cap", () => {
  const fixed = alignCodeComments(["a; // x", "longer_name; // y"], 30);
  assert.deepEqual(fixed.map((line) => line.indexOf("//")), [30, 30]);
  const capped = alignCodeComments(["a".repeat(120) + " // z"], 0, { maxColumn: 80 });
  assert.equal(capped[0].indexOf("//"), 121, "a too-long line keeps exactly one space");
});

test("alignCodeComments is idempotent", () => {
  const once = alignCodeComments(["a; // one", "longer_call(); // two"], 0);
  const twice = alignCodeComments(once, 0);
  assert.deepEqual(once, twice);
});

test("markdownToBlocks: alignCommentColumn auto inside a fence only", () => {
  const markdown = ["```c", "int a = 1; // one", "long_call(a, b); // two", "```", "", "int z = 9; // outside", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { alignCommentColumn: "auto" });
  const codeLines = blocks.filter((block) => block.type === "code").map((block) => block.text);
  assert.equal(codeLines.length, 2);
  const columns = codeLines.map((line) => displayWidth(line.slice(0, line.indexOf("//"))));
  assert.equal(columns[0], columns[1]);
  const outside = blocks.find((block) => block.type === "para");
  assert.equal(outside.text, "int z = 9; // outside", "paragraphs are never aligned");
});

test("markdownToBlocks: alignCommentColumn 0 leaves code untouched", () => {
  const markdown = ["```c", "int a = 1; // one", "long_call(a, b); // two", "```", ""].join("\n");
  const blocks = markdownToBlocks(markdown, { alignCommentColumn: 0 });
  const codeLines = blocks.filter((block) => block.type === "code").map((block) => block.text);
  assert.deepEqual(codeLines, ["│ int a = 1; // one", "│ long_call(a, b); // two"]);
});

test("markdownToBlocks: codePrefix can be empty", () => {
  const blocks = markdownToBlocks(["```", "int a = 1;", "```", ""].join("\n"), { codePrefix: "" });
  assert.equal(blocks[0].type, "code");
  assert.equal(blocks[0].text, "int a = 1;");
});

test("markdownToBlocks: default output is unchanged", () => {
  const markdown = ["```c", "int a = 1; // one", "```", ""].join("\n");
  assert.deepEqual(markdownToBlocks(markdown, { codePrefix: "│ ", alignCommentColumn: 0 }), markdownToBlocks(markdown));
});
test("tree fenced blocks keep leading spaces and are not realigned", () => {
  const md = "```tree\nroot()          // r.c:1\n  \u251C\u2500 child()   // c.c:2\n  \u2502   \u2514\u2500 leaf()\n```\n";
  const blocks = markdownToBlocks(md, { codePrefix: "" });
  assert.deepEqual(blocks.map((block) => block.type), ["tree", "tree", "tree"]);
  assert.equal(blocks[1].text, "  \u251C\u2500 child()   // c.c:2");
  const back = blocksToMarkdown(blocks, { codePrefix: "" });
  assert.match(back, /^```tree\n/);
  assert.ok(back.includes("  \u251C\u2500 child()   // c.c:2"));
});

test("an empty codePrefix keeps a leading tree bar; a configured gutter is stripped", () => {
  const tree = [{ type: "code", text: "│   └─ leaf()" }];
  assert.equal(blocksToMarkdown(tree, { codePrefix: "" }), "```tree\n│   └─ leaf()\n```\n");
  const code = [{ type: "code", text: "│ int x;" }];
  assert.equal(blocksToMarkdown(code, { codePrefix: "│ " }), "```\nint x;\n```\n");
});