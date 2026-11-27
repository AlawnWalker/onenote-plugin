/**
 * Import a markdown file into OneNote as a new standalone section, one page per
 * `##` chapter. Run in a NON-sandboxed shell (COM) with OneNote available:
 *
 *   node dev/import-md.mjs [md-path] [section.one-path]
 *
 * Defaults target D:\work\Temp\KVM\KVM_FRAMEWORK_NOTES.md -> KVM.one. The
 * section is created if missing and the import is idempotent: blank pages the
 * section creation leaves behind are removed, and chapters whose page already
 * exists (by title) are skipped.
 *
 * Page size follows ONENOTE_PAGESIZE and defaults to OneNote's Automatic size,
 * matching the DSH plugin; set the env var when a document needs wider lines.
 */
import { readFile } from "node:fs/promises";

import { OneNoteBridge } from "../lib/bridge.js";
import { markdownToBlocks } from "../lib/markdown.js";
import { cleanBody, preambleIntro, splitChapters } from "../lib/import.js";
import { readConfig } from "../lib/config.js";

const MD_PATH = process.argv[2] ?? "D:/work/Temp/KVM/KVM_FRAMEWORK_NOTES.md";
const SECTION_PATH = process.argv[3] ?? "D:/work/Temp/KVM/KVM.one";
const BLANK_PAGE_NAMES = new Set(["未命名", "Untitled", "Untitled Page"]);

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** bridge.invoke throws BridgeError on failure; on success it returns `data`. */

/** Read this section's pages through the bridge (path-keyed probe). */
async function listSectionPages(bridge) {
  for (let attempt = 1; ; attempt += 1) {
    try {
      const listing = await bridge.invoke("list", { includePages: true });
      const section = (listing.localSections ?? []).find((entry) => entry.error === null);
      if (section) return section.pages ?? [];
      throw new Error(`section not visible yet: ${SECTION_PATH}`);
    } catch (err) {
      if (attempt >= 3) throw err;
      await sleep(3000);
    }
  }
}

const raw = await readFile(MD_PATH, "utf8");
const { preamble, chapters } = splitChapters(raw);

// Fold the intro quote into the first page so the document context is kept.
const intro = preambleIntro(preamble);
if (intro && chapters.length > 0) {
  chapters[0].bodyLines = [...intro.split("\n"), "", ...chapters[0].bodyLines];
}

const config = readConfig(process.env);
const bridge = new OneNoteBridge(config);
const section = await bridge.invoke("createsection", { path: SECTION_PATH });
console.log(`section: ${section.path}`);
console.log(`  id: ${section.sectionId}\n`);

// Creating a section leaves an auto-generated blank page behind; drop it so the
// page list is exactly one page per chapter.
const probe = new OneNoteBridge({ extraSections: [{ name: "__probe__", path: SECTION_PATH }] });
const existing = await listSectionPages(probe);
for (const page of existing) {
  const read = await probe.invoke("read", { pageId: page.id });
  if (read.blockCount === 0 && BLANK_PAGE_NAMES.has(page.name)) {
    await probe.invoke("delete", { pageId: page.id, permanent: true });
    console.log(`removed blank page: [${page.name}]`);
  }
}
const existingTitles = new Set(existing.map((page) => page.name));

let created = 0;
let skipped = 0;
for (const chapter of chapters) {
  if (existingTitles.has(chapter.title)) {
    skipped += 1;
    console.log(`[skip] ${chapter.title} (already imported)`);
    continue;
  }
  const body = cleanBody(chapter.bodyLines);
  const page = await bridge.invoke("create", {
    sectionId: section.sectionId,
    title: chapter.title,
    blocks: markdownToBlocks(body),
  });
  created += 1;
  console.log(`[${created}/${chapters.length}] ${chapter.title}`);
  console.log(`    pageId: ${page.pageId}`);
}

console.log(`\ndone: ${created} created, ${skipped} skipped, ${chapters.length} chapters total`);
