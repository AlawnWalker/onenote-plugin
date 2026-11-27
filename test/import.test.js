import assert from "node:assert/strict";
import test from "node:test";

import { cleanBody, preambleIntro, splitChapters, stripInline } from "../lib/import.js";

test("splitChapters splits the document by ## headings", () => {
  const md = "# Title\n\nintro\n\n## A\n\nbody A\n\n### A.1\n\n## B\n\nbody B\n";
  const { preamble, chapters } = splitChapters(md);
  assert.equal(chapters.length, 2);
  assert.equal(chapters[0].title, "A");
  assert.equal(chapters[1].title, "B");
  assert.match(chapters[0].bodyLines.join("\n"), /### A\.1/);
  assert.match(preamble.join("\n"), /# Title/);
});

test("stripInline removes links, bold and inline code", () => {
  assert.equal(stripInline("[文本](#anchor)"), "文本");
  assert.equal(stripInline("**粗体** 和 `code`"), "粗体 和 code");
});

test("cleanBody strips inline formatting but leaves fenced code intact", () => {
  const out = cleanBody(["**加粗**", "```", "**keep** `as`", "```", "`after`"]);
  assert.match(out, /\*\*keep\*\*/);
  assert.match(out, /加粗/);
  assert.match(out, /after/);
});

test("preambleIntro drops the H1 title and horizontal rules", () => {
  assert.equal(preambleIntro(["# Title", "", "> line1", "> line2", "", "---"]), "> line1\n> line2");
});
