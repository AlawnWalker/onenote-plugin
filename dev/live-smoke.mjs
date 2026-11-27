/**
 * Live smoke test — talks to the real local OneNote application over COM.
 *
 * NOT part of `node --test` (it lives outside test/ on purpose). Run it in a
 * NON-sandboxed terminal with OneNote available:
 *
 *   npm run smoke
 *
 * It resolves a writable section (from ONENOTE_SMOKE_SECTION_ID /
 * ONENOTE_SMOKE_SECTION / ONENOTE_EXTRASECTIONS, or by auto-discovery through
 * onenote_list), creates one page, reads it back, verifies the markdown
 * round-trip, and deletes the page again. Nothing else is touched.
 *
 * The COM server cannot start inside a file sandbox (0x80080005), so this must
 * run as the user in a full-permission shell.
 */
import assert from "node:assert/strict";

import { OneNoteBridge } from "../lib/bridge.js";
import { makeToolHandlers } from "../lib/tools.js";
import { readConfig } from "../lib/config.js";

const TITLE = `ONENOTE-MCP-SMOKE-${new Date().toISOString().replace(/[:.]/g, "-")}`;
const BODY = `冒烟测试正文。

## 小节标题

- 条目甲
- 条目乙
  - 子条目

1. 第一步
2. 第二步

| 列1 | 列2 |
| --- | --- |
| x | y |

\`\`\`
line one
line two
\`\`\`
`;

const checks = [];

function check(name, fn) {
  return Promise.resolve()
    .then(fn)
    .then(() => {
      checks.push({ name, ok: true });
      console.log(`PASS  ${name}`);
    })
    .catch((err) => {
      checks.push({ name, ok: false, error: err.message });
      console.log(`FAIL  ${name}: ${err.message}`);
      throw err;
    });
}

function assertOk(result, action) {
  if (!result || result.ok !== true) {
    throw new Error(`${action} failed: ${result?.code ?? "UNKNOWN"} ${result?.message ?? ""}`);
  }
  return result.data;
}

async function resolveSection(handlers, config) {
  if (process.env.ONENOTE_SMOKE_SECTION_ID) return { sectionId: process.env.ONENOTE_SMOKE_SECTION_ID };
  if (process.env.ONENOTE_SMOKE_SECTION) return { sectionName: process.env.ONENOTE_SMOKE_SECTION };
  if (config.extraSections.length > 0) return { sectionName: config.extraSections[0].name };
  const list = assertOk(await handlers.onenote_list({ includePages: false, maxSections: 50 }), "onenote_list");
  const sections = [
    ...(list.localSections ?? []),
    ...(list.notebooks ?? []).flatMap((n) => n.sections ?? []),
  ];
  const pick = sections.find((section) => !section.isRecycle && !section.error);
  if (!pick) throw new Error("no writable section found; set ONENOTE_EXTRASECTIONS or ONENOTE_SMOKE_SECTION");
  return pick.id ? { sectionId: pick.id } : { sectionName: pick.name };
}

const config = readConfig(process.env);
const bridge = new OneNoteBridge(config);
const handlers = makeToolHandlers({ bridge });

let pageId = null;
try {
  const listing = await handlers.onenote_list({ includePages: false, maxSections: 20 });
  await check("onenote_list returns the notebook tree", () => {
    const data = assertOk(listing, "onenote_list");
    const total = (data.counts?.sections ?? 0) + (data.localSections?.length ?? 0);
    assert.ok(total >= 0, "counts parsed");
    console.log(`  info: ${data.counts?.notebooks ?? 0} notebooks, ${data.counts?.sections ?? 0} sections, ${data.counts?.pages ?? 0} pages`);
  });

  const section = await resolveSection(handlers, config);
  console.log(`  info: targeting section ${JSON.stringify(section)}`);

  const created = await handlers.onenote_create({ ...section, title: TITLE, markdown: BODY });
  await check("onenote_create returns a page id", () => {
    const data = assertOk(created, "onenote_create");
    pageId = data.pageId;
    assert.ok(pageId && pageId.startsWith("{"), `unexpected page id: ${pageId}`);
  });

  const readBack = await handlers.onenote_read({ pageId });
  await check("onenote_read preserves title and content", () => {
    const data = assertOk(readBack, "onenote_read");
    assert.equal(data.title, TITLE);
    assert.match(data.markdown, /冒烟测试正文/);
    assert.match(data.markdown, /## 小节标题/);
    assert.match(data.markdown, /- 条目甲/);
    assert.match(data.markdown, /  - 子条目/);
    assert.match(data.markdown, /1\. 第一步/);
    assert.match(data.markdown, /\| 列1 \| 列2 \|/);
    assert.match(data.markdown, /line one/);
  });
} finally {
  if (pageId) {
    try {
      const deleted = await handlers.onenote_delete({ pageId, permanent: true });
      assertOk(deleted, "onenote_delete");
      console.log("PASS  onenote_delete removed the test page");
    } catch (err) {
      console.log(`WARN  cleanup failed for ${pageId}: ${err.message}`);
    }
  }
}

const failed = checks.filter((entry) => !entry.ok);
console.log(`\n${checks.length - failed.length}/${checks.length} checks passed`);
if (failed.length > 0) process.exit(1);
