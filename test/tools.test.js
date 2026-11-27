import assert from "node:assert/strict";
import path from "node:path";
import test from "node:test";

import { makeToolHandlers, renderResult } from "../lib/tools.js";

function fakeBridge(responder) {
  const calls = [];
  return {
    calls,
    async invoke(action, params) {
      calls.push({ action, params });
      return responder(action, params);
    },
  };
}

test("onenote_create lifts a leading '# Title' out of the markdown", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", title: "标题", sectionId: "{SEC}" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_create({ sectionId: "{SEC}", markdown: "# 标题\n\n正文\n" });

  assert.equal(result.ok, true);
  const call = bridge.calls[0];
  assert.equal(call.action, "create");
  assert.equal(call.params.title, "标题");
  assert.equal(call.params.blocks[0].type, "para");
  assert.equal(call.params.blocks[0].text, "正文");
});

test("onenote_create rejects markdown without a title", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_create({ sectionId: "{SEC}", markdown: "只有正文" });
  assert.equal(result.ok, false);
  assert.equal(result.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);
});

test("onenote_read renders the page as markdown", async () => {
  const bridge = fakeBridge(() => ({
    pageId: "{PAGE}",
    title: "页面",
    lastModified: "2026-10-29T00:00:00.000Z",
    blocks: [
      { type: "para", text: "━━━ 小节 ━━━", level: 0 },
      { type: "bullet", text: "条目", level: 0 },
    ],
  }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_read({ pageId: "{PAGE}" });
  assert.equal(result.ok, true);
  assert.match(result.data.markdown, /## 小节/);
  assert.match(result.data.markdown, /- 条目/);

  const rendered = renderResult({}, result);
  assert.match(rendered[0].text, /^# 页面/);
  assert.match(rendered[0].text, /pageId: \{PAGE\}/);
});

test("onenote_update forwards an explicit title and strips it from the body", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", updated: true }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_update({
    pageId: "{PAGE}",
    markdown: "# 新标题\n\n正文\n",
  });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].params.title, "新标题");
  assert.deepEqual(
    bridge.calls[0].params.blocks.map((block) => block.type),
    ["para"],
  );
});

test("onenote_publish validates the format and resolves relative paths", async () => {
  const outPath = path.resolve("out", "p.one");
  const bridge = fakeBridge(() => ({ path: outPath, bytes: 1 }));
  const handlers = makeToolHandlers({ bridge });
  const bad = await handlers.onenote_publish({ pageId: "{PAGE}", path: path.resolve("out", "p.exe"), format: "exe" });
  assert.equal(bad.ok, false);
  assert.equal(bad.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);

  const good = await handlers.onenote_publish({ pageId: "{PAGE}", path: path.resolve("out", "p.pdf"), format: "pdf" });
  assert.equal(good.ok, true);
  assert.equal(bridge.calls[0].params.format, "pdf");
});

test("bridge failures surface as structured tool errors", async () => {
  const bridge = {
    async invoke() {
      const err = new Error("page not found: 不存在");
      err.code = "PAGE_NOT_FOUND";
      throw err;
    },
  };
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_read({ pageTitle: "不存在" });
  assert.deepEqual(result, { ok: false, code: "PAGE_NOT_FOUND", message: "page not found: 不存在" });
});

test("onenote_list passes filters and limits through", async () => {
  const bridge = fakeBridge(() => ({ notebooks: [], localSections: [], counts: {} }));
  const handlers = makeToolHandlers({ bridge });
  await handlers.onenote_list({ notebook: "Work", includePages: false, maxSections: 50, maxPages: 10 });
  assert.deepEqual(bridge.calls[0].params, {
    notebook: "Work",
    section: undefined,
    includePages: false,
    includeRecycleBin: false,
    maxSections: 50,
    maxPages: 10,
  });
});

test("onenote_list rejects negative limits", async () => {
  const bridge = fakeBridge(() => ({ notebooks: [] }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_list({ maxSections: -1 });
  assert.equal(result.ok, false);
  assert.equal(bridge.calls.length, 0);
});

test("onenote_read forwards maxBlocks and surfaces truncation", async () => {
  const bridge = fakeBridge(() => ({
    pageId: "{PAGE}",
    title: "T",
    lastModified: "2026-11-03T00:00:00.000Z",
    blockCount: 100,
    truncated: true,
    blocks: [{ type: "para", text: "first", level: 0 }],
  }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_read({ pageId: "{PAGE}", maxBlocks: 20 });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].params.maxBlocks, 20);
  assert.equal(result.data.blockCount, 100);
  assert.equal(result.data.truncated, true);
  const rendered = renderResult({}, result);
  assert.match(rendered[0].text, /截断/);
});

test("onenote_create_section requires an absolute path and passes keepBlankPage", async () => {
  const bridge = fakeBridge(() => ({ sectionId: "{SEC}", path: path.resolve("out", "a.one"), created: true, blankPageRemoved: false }));
  const handlers = makeToolHandlers({ bridge });

  const missing = await handlers.onenote_create_section({});
  assert.equal(missing.ok, false);
  assert.equal(missing.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);

  const result = await handlers.onenote_create_section({ path: path.resolve("out", "a.one") });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].action, "createsection");
  assert.equal(bridge.calls[0].params.path, path.resolve("out", "a.one"));
  assert.equal(bridge.calls[0].params.keepBlankPage, false);
});

test("onenote_rename_section needs a target and a name", async () => {
  const bridge = fakeBridge(() => ({ sectionId: "{SEC}", name: "新名" }));
  const handlers = makeToolHandlers({ bridge });

  const noTarget = await handlers.onenote_rename_section({ name: "新名" });
  assert.equal(noTarget.ok, false);
  assert.equal(noTarget.code, "BAD_ARGUMENT");

  const noName = await handlers.onenote_rename_section({ sectionId: "{SEC}" });
  assert.equal(noName.ok, false);
  assert.equal(noName.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);

  const result = await handlers.onenote_rename_section({ sectionId: "{SEC}", name: "新名" });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].action, "renamesection");
  assert.equal(bridge.calls[0].params.name, "新名");
});

test("onenote_edit forwards replaceText to the bridge", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", mode: "replaceText", changed: 1 }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({
    pageId: "{PAGE}",
    mode: "replaceText",
    objectId: "{OE}",
    search: "old",
    replacement: "new",
  });
  assert.equal(result.ok, true);
  const call = bridge.calls[0];
  assert.equal(call.action, "edit");
  assert.equal(call.params.mode, "replaceText");
  assert.equal(call.params.objectId, "{OE}");
  assert.equal(call.params.search, "old");
  assert.equal(call.params.replacement, "new");
  assert.equal(call.params.all, false);
});

test("onenote_edit rejects an unknown mode before touching the bridge", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({ pageId: "{PAGE}", mode: "rename", objectId: "{OE}" });
  assert.equal(result.ok, false);
  assert.equal(result.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);
});

test("onenote_edit requires an objectId", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({ pageId: "{PAGE}", mode: "deleteBlock" });
  assert.equal(result.ok, false);
  assert.equal(result.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);
});

test("onenote_edit insertAfter defaults the block kind to para", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", mode: "insertAfter", changed: 1 }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({
    pageId: "{PAGE}",
    mode: "insertAfter",
    objectId: "{OE}",
    text: "inserted",
  });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].params.kind, "para");
  assert.equal(bridge.calls[0].params.text, "inserted");
});
test("onenote_read withIds returns objectIds per block", async () => {
  const bridge = fakeBridge(() => ({
    pageId: "{PAGE}",
    title: "T",
    lastModified: "2026-11-06T00:00:00.000Z",
    blockCount: 2,
    blocks: [
      { type: "para", text: "a", level: 0, objectId: "{A}" },
      { type: "para", text: "b", level: 0, objectId: "{B}" },
    ],
  }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_read({ pageId: "{PAGE}", withIds: true });
  assert.equal(result.ok, true);
  assert.deepEqual(result.data.blocks.map((block) => block.objectId), ["{A}", "{B}"]);
});

test("onenote_edit can target a block by unique text", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", mode: "replaceText", changed: 1 }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({
    pageId: "{PAGE}",
    mode: "replaceText",
    find: "beta",
    search: "beta",
    replacement: "BETA",
  });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].params.find, "beta");
  assert.equal(bridge.calls[0].params.objectId, "");
});

test("onenote_edit rejects a call without objectId or find", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({ pageId: "{PAGE}", mode: "deleteBlock" });
  assert.equal(result.ok, false);
  assert.equal(result.code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);
});

/**
 * DSH rejects a tool result that JSON cannot carry unchanged ("value is not
 * lossless JSON"). An enumerable `undefined` property is the classic trap: the
 * body looks fine, but the call is reported as failed *after* its side effect.
 */
function assertLossless(value, path = "result") {
  if (value === null || typeof value === "string" || typeof value === "boolean") return;
  if (typeof value === "number") {
    assert.ok(Number.isFinite(value), `${path} is not finite`);
    return;
  }
  assert.equal(typeof value, "object", `${path} is not JSON-safe (${typeof value})`);
  for (const key of Object.keys(value)) {
    const child = value[key];
    assert.notEqual(child, undefined, `${path}.${key} is undefined`);
    assertLossless(child, `${path}.${key}`);
  }
}

test("onenote_append returns lossless JSON when markdown has no title", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", appended: true, chunks: [{ type: "para", text: "x", level: 0 }] }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_append({ pageId: "{PAGE}", markdown: "只有正文，无标题" });

  assert.equal(result.ok, true);
  assertLossless(result);
  assert.equal(result.data.appendedFromTitle, undefined);
  assert.deepEqual(JSON.parse(JSON.stringify(result)), result);
});

test("onenote_append reports appendedFromTitle when markdown starts with '# Title'", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", appended: true, chunks: [] }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_append({ pageId: "{PAGE}", markdown: "# 标题\n\n正文" });

  assert.equal(result.ok, true);
  assertLossless(result);
  assert.equal(result.data.appendedFromTitle, "标题");
  assert.deepEqual(JSON.parse(JSON.stringify(result)), result);
});

test("onenote_create_section registers the opened section for later list calls", async () => {
  const bridge = fakeBridge((action) => {
    if (action === "createsection") {
      return { sectionId: "{SEC}", path: "C:\\Notes\\Quick.one", created: true, blankPageRemoved: false };
    }
    if (action === "list") return { notebooks: [], localSections: [], counts: {}, truncated: false };
    return {};
  });
  bridge.extraSections = [];
  const handlers = makeToolHandlers({ bridge });

  const created = await handlers.onenote_create_section({ path: "C:\\Notes\\Quick.one" });
  assert.equal(created.ok, true);
  assert.equal(bridge.extraSections.length, 1);
  assert.deepEqual(bridge.extraSections[0], { name: "Quick", path: "C:\\Notes\\Quick.one", id: "{SEC}" });

  await handlers.onenote_create_section({ path: "C:/Notes/Quick.one" });
  assert.equal(bridge.extraSections.length, 1, "re-opening the same section must not duplicate it");

  await handlers.onenote_list({ includePages: false });
  const listCall = bridge.calls.find((call) => call.action === "list");
  assert.ok(listCall, "list was invoked");
});

test("onenote_edit forwards exact matching to the bridge", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", changed: 1, applied: true }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_edit({
    pageId: "{PAGE}",
    mode: "replaceText",
    find: "beta",
    exact: true,
    search: "beta",
    replacement: "BETA",
  });

  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].params.find, "beta");
  assert.equal(bridge.calls[0].params.exact, true);
});

test("onenote_edit defaults exact matching to false", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", changed: 1, applied: true }));
  const handlers = makeToolHandlers({ bridge });
  await handlers.onenote_edit({ pageId: "{PAGE}", mode: "deleteBlock", find: "beta" });
  assert.equal(bridge.calls[0].params.exact, false);
});