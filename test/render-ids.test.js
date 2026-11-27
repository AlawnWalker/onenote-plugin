import test from "node:test";
import assert from "node:assert/strict";

import { renderResult } from "../lib/tools.js";

const VALUE = {
  ok: true,
  data: {
    title: "page",
    pageId: "{PAGE}",
    lastModified: "2026-01-01T00:00:00Z",
    markdown: "  - first bullet",
    blocks: [
      { index: 0, type: "bullet", objectId: "{OID-1}", text: "first bullet" },
      { index: 1, type: "h2", objectId: "{OID-2}", text: "heading" },
    ],
  },
};

test("renderResult surfaces objectIds when withIds is set", () => {
  const out = renderResult({ withIds: true }, VALUE);
  const text = out[0].text;
  assert.match(text, /objectId=\{OID-1\}/);
  assert.match(text, /objectId=\{OID-2\}/);
  assert.match(text, /first bullet/);
});

test("renderResult omits the id list without withIds", () => {
  const out = renderResult({}, VALUE);
  assert.ok(!/objectId=/.test(out[0].text));
});
