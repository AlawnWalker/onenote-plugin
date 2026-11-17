import assert from "node:assert/strict";
import test from "node:test";

import { blocksToMarkdown, markdownToBlocks } from "../lib/markdown.js";
import { makeToolHandlers, renderResult } from "../lib/tools.js";
import { normalizeConfig } from "../lib/config.js";
import { buildJsonSchema } from "../lib/schemas.js";

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

test("a standalone markdown image becomes an image block with its source path", () => {
  const blocks = markdownToBlocks("![示意图](C:\\pics\\mmu.png)\n");
  assert.deepEqual(blocks, [{ type: "image", src: "C:\\pics\\mmu.png", alt: "示意图", level: 0 }]);
});

test("an image whose alt text is a number sets the width in points", () => {
  const [block] = markdownToBlocks("![560](C:\\pics\\mmu.png)");
  assert.equal(block.type, "image");
  assert.equal(block.width, 560);
  assert.equal(block.alt, "");
});

test("indented image lines keep their level", () => {
  const [, image] = markdownToBlocks("- 条目\n\n  ![图](D:\\p.png)\n");
  assert.equal(image.type, "image");
  assert.equal(image.level, 1);
});

test("image blocks render back to markdown (angle brackets when the path has spaces)", () => {
  const markdown = blocksToMarkdown([{ type: "image", src: "D:\\a b\\p.png", alt: "图", level: 0 }]);
  assert.match(markdown, /!\[图\]\(<D:\\a b\\p\.png>\)/);
});

test("an image read from OneNote (no source path) becomes the placeholder", () => {
  const markdown = blocksToMarkdown([{ type: "image", text: "", level: 0, width: 240, height: 120 }]);
  assert.match(markdown, /!\[图片\]/);
});

test("the rendered read footer carries revision and media so the model sees them", () => {
  const rendered = renderResult({}, {
    ok: true,
    data: {
      title: "页面",
      pageId: "{PAGE}",
      lastModified: "2026-11-16T08:36:10.000Z",
      revision: "ab63dce549e1feba",
      markdown: "正文\n",
      media: { images: 2, ink: 0, files: 0, total: 2 },
    },
  });
  const text = rendered[0].text;
  assert.match(text, /revision: ab63dce549e1feba/);
  assert.match(text, /media: images=2 ink=0 files=0/);
});

test("onenote_read forwards images, media counts and the revision hash", async () => {
  const bridge = fakeBridge(() => ({
    pageId: "{PAGE}",
    title: "页面",
    lastModified: "2026-11-16T08:19:16.000Z",
    revision: "abc1234567890def",
    blockCount: 2,
    truncated: false,
    blocks: [
      { type: "para", text: "正文", level: 0 },
      { type: "image", text: "", level: 1, width: 240, height: 120, objectId: "{OE}", outlineId: "{OL}" },
    ],
    images: [{ type: "image", width: 240, height: 120, objectId: "{OE}", outlineId: "{OL}" }],
    media: { images: 1, ink: 0, files: 0, total: 1 },
  }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_read({ pageId: "{PAGE}" });

  assert.equal(result.ok, true);
  assert.equal(result.data.revision, "abc1234567890def");
  assert.equal(result.data.images.length, 1);
  assert.equal(result.data.media.total, 1);
  assert.match(result.data.markdown, /!\[图片\]/);
});

test("onenote_insert_image forwards path/width/position/find to the bridge", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", objectId: "{IMG}", width: "300", height: "200" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_insert_image({
    pageId: "{PAGE}",
    path: "D:\\pics\\mmu.png",
    width: 300,
    position: "before",
    find: "小节标题",
    backup: false,
  });

  assert.equal(result.ok, true);
  const call = bridge.calls[0];
  assert.equal(call.action, "insertimage");
  assert.equal(call.params.path, "D:\\pics\\mmu.png");
  assert.equal(call.params.width, 300);
  assert.equal(call.params.position, "before");
  assert.equal(call.params.find, "小节标题");
  assert.equal(call.params.backup, false);
});

test("onenote_insert_image leaves the backup default to the config", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}" }));
  const handlers = makeToolHandlers({ bridge });
  await handlers.onenote_insert_image({ pageId: "{PAGE}", path: "D:\\pics\\mmu.png" });
  assert.equal("backup" in bridge.calls[0].params, false);
  assert.equal(bridge.calls[0].params.position, "after");
});

test("onenote_insert_image validates its arguments before calling OneNote", async () => {
  const bridge = fakeBridge(() => ({}));
  const handlers = makeToolHandlers({ bridge });

  assert.equal((await handlers.onenote_insert_image({ pageId: "{PAGE}" })).code, "BAD_ARGUMENT");
  assert.equal((await handlers.onenote_insert_image({ pageId: "{PAGE}", path: "p.png", width: -5 })).code, "BAD_ARGUMENT");
  assert.equal((await handlers.onenote_insert_image({ pageId: "{PAGE}", path: "p.png", position: "left" })).code, "BAD_ARGUMENT");
  assert.equal(bridge.calls.length, 0);
});

test("onenote_update forwards backup/preserveMedia only when set, so config wins", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", updated: true, backupPath: "C:\\bak\\p.xml" }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_update({ pageId: "{PAGE}", markdown: "正文" });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].action, "update");
  assert.equal("backup" in bridge.calls[0].params, false);
  assert.equal("preserveMedia" in bridge.calls[0].params, false);
  assert.equal(result.data.backupPath, "C:\\bak\\p.xml");

  await handlers.onenote_update({ pageId: "{PAGE}", markdown: "正文", preserveMedia: false });
  assert.equal(bridge.calls[1].params.preserveMedia, false);
});

test("config: backup dir, media preservation and image width defaults", () => {
  const defaults = normalizeConfig({});
  assert.equal(defaults.backupDir, "");
  assert.equal(defaults.preserveMedia, true);
  assert.equal(defaults.imageMaxWidth, 600);

  const custom = normalizeConfig({ backupDir: " D:\\bak ", preserveMedia: "false", imageMaxWidth: "480" });
  assert.equal(custom.backupDir, "D:\\bak");
  assert.equal(custom.preserveMedia, false);
  assert.equal(custom.imageMaxWidth, 480);

  assert.equal(normalizeConfig({ imageMaxWidth: 5000 }).imageMaxWidth, 600);
  assert.equal(normalizeConfig({ imageMaxWidth: -1 }).imageMaxWidth, 600);
});

test("onenote_normalize_tables forwards page locator and flags", async () => {
  const bridge = fakeBridge(() => ({ pageId: "{PAGE}", baseFontSize: 12, tables: 1, cellsChanged: 4, outlinesUpdated: 1 }));
  const handlers = makeToolHandlers({ bridge });
  const result = await handlers.onenote_normalize_tables({ pageId: "{PAGE}", dryRun: true });
  assert.equal(result.ok, true);
  assert.equal(bridge.calls[0].action, "normalizetables");
  assert.equal(bridge.calls[0].params.pageId, "{PAGE}");
  assert.equal(bridge.calls[0].params.dryRun, true);
  assert.equal("backup" in bridge.calls[0].params, false);
});

test("onenote_insert_image is advertised with a required path and typed options", () => {
  const schema = buildJsonSchema("onenote_insert_image");
  assert.deepEqual(schema.required, ["path"]);
  assert.equal(schema.properties.path.type, "string");
  assert.equal(schema.properties.width.type, "integer");
  assert.equal(schema.properties.position.type, "string");
});
