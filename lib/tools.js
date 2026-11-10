/**
 * dsh-onenote — tool handlers.
 *
 * Every handler validates its arguments, converts between the markdown subset
 * and the bridge's block model, and returns `{ ok, data }` / `{ ok, code,
 * message }` so the DSH layer always gets a JSON-serialisable result.
 */
import {
  blocksToMarkdown,
  extractTitle,
  markdownToBlocks,
  stripLeadingTitle,
} from "./markdown.js";

export function ok(data) {
  return { ok: true, data };
}

export function fail(code, message) {
  return { ok: false, code, message };
}

function bad(message) {
  return fail("BAD_ARGUMENT", message);
}

function optionalString(value, name, { maxLength = 4096 } = {}) {
  if (value === undefined || value === null) return { ok: true, value: undefined };
  if (typeof value !== "string") return bad(`${name} must be a string`);
  const trimmed = value.trim();
  if (trimmed.length === 0) return { ok: true, value: undefined };
  if (trimmed.length > maxLength) return bad(`${name} is too long (max ${maxLength})`);
  return { ok: true, value: trimmed };
}

function requiredString(value, name, { maxLength = 65536 } = {}) {
  if (value === undefined || value === null) return bad(`${name} is required`);
  if (typeof value !== "string") return bad(`${name} must be a string`);
  const trimmed = value.trim();
  if (trimmed.length === 0) return bad(`${name} must not be empty`);
  if (trimmed.length > maxLength) return bad(`${name} is too long (max ${maxLength})`);
  return { ok: true, value };
}

function asBoolean(value, fallback) {
  if (value === undefined || value === null) return fallback;
  return Boolean(value);
}

function asInt(value, fallback) {
  if (value === undefined || value === null || value === "") return fallback;
  const n = Number.parseInt(String(value), 10);
  return Number.isFinite(n) ? n : fallback;
}

/** Shared page-locator arguments (pageId, or section + pageTitle). */
function pageLocator(args, { requireLocator = true } = {}) {
  const pageId = optionalString(args.pageId, "pageId", { maxLength: 512 });
  if (!pageId.ok) return pageId;
  const sectionId = optionalString(args.sectionId, "sectionId", { maxLength: 512 });
  if (!sectionId.ok) return sectionId;
  const sectionName = optionalString(args.sectionName, "sectionName", { maxLength: 256 });
  if (!sectionName.ok) return sectionName;
  const pageTitle = optionalString(args.pageTitle, "pageTitle", { maxLength: 512 });
  if (!pageTitle.ok) return pageTitle;
  if (requireLocator && !pageId.value && !pageTitle.value) {
    return bad("pageId or pageTitle is required");
  }
  return {
    ok: true,
    value: {
      pageId: pageId.value,
      sectionId: sectionId.value,
      sectionName: sectionName.value,
      pageTitle: pageTitle.value,
    },
  };
}

function sectionLocator(args) {
  const sectionId = optionalString(args.sectionId, "sectionId", { maxLength: 512 });
  if (!sectionId.ok) return sectionId;
  const sectionName = optionalString(args.sectionName, "sectionName", { maxLength: 256 });
  if (!sectionName.ok) return sectionName;
  const notebook = optionalString(args.notebook, "notebook", { maxLength: 256 });
  if (!notebook.ok) return notebook;
  if (!sectionId.value && !sectionName.value) {
    return bad("sectionId or sectionName is required (use onenote_list to discover sections)");
  }
  return {
    ok: true,
    value: { sectionId: sectionId.value, sectionName: sectionName.value, notebook: notebook.value },
  };
}

function blocksFromMarkdown(markdown, renderOptions = {}) {
  const blocks = markdownToBlocks(markdown, renderOptions);
  const titleBlock = blocks.find((block) => block.type === "title");
  return { blocks, titleFromMarkdown: titleBlock ? titleBlock.text : null };
}

function removeTitleBlock(blocks) {
  const index = blocks.findIndex((block) => block.type === "title");
  if (index >= 0) blocks.splice(index, 1);
  return blocks;
}

export function makeToolHandlers({ bridge, renderOptions = {} } = {}) {
  const call = (action, params) => bridge.invoke(action, params);

  async function guard(fn) {
    try {
      return await fn();
    } catch (err) {
      return fail(err?.code ? String(err.code) : "ONENOTE_ERROR", err?.message ? String(err.message) : "OneNote operation failed");
    }
  }

  return {
    /** List notebooks -> sections -> pages. */
    onenote_list(args = {}) {
      return guard(async () => {
        const notebook = optionalString(args.notebook, "notebook", { maxLength: 256 });
        if (!notebook.ok) return notebook;
        const section = optionalString(args.section, "section", { maxLength: 256 });
        if (!section.ok) return section;
        const maxSections = asInt(args.maxSections, 0);
        if (maxSections < 0) return bad("maxSections must be >= 0");
        const maxPages = asInt(args.maxPages, 0);
        if (maxPages < 0) return bad("maxPages must be >= 0");
        const data = await call("list", {
          notebook: notebook.value,
          section: section.value,
          includePages: asBoolean(args.includePages, true),
          includeRecycleBin: asBoolean(args.includeRecycleBin, false),
          maxSections,
          maxPages,
        });
        return ok(data);
      });
    },

    /** Read a page and render it as markdown. */
    onenote_read(args = {}) {
      return guard(async () => {
        const locator = pageLocator(args);
        if (!locator.ok) return locator;
        const maxBlocks = asInt(args.maxBlocks, 0);
        if (maxBlocks < 0) return bad("maxBlocks must be >= 0");
        const data = await call("read", { ...locator.value, maxBlocks });
        const markdown = blocksToMarkdown(data.blocks ?? [], renderOptions);
        const withIds = asBoolean(args.withIds, false);
        const items = withIds
          ? (data.blocks ?? []).map((block, index) => ({
              index,
              type: block.type,
              objectId: block.objectId ?? "",
              text: block.text ?? "",
            }))
          : undefined;
        return ok({
          pageId: data.pageId,
          title: data.title,
          lastModified: data.lastModified,
          blockCount: data.blockCount,
          truncated: data.truncated,
          markdown,
          ...(items ? { blocks: items } : {}),
        });
      });
    },

    /** Search pages by title, optionally scanning page text. */
    onenote_search(args = {}) {
      return guard(async () => {
        const query = requiredString(args.query, "query", { maxLength: 512 });
        if (!query.ok) return query;
        const scopeId = optionalString(args.scopeId, "scopeId", { maxLength: 512 });
        if (!scopeId.ok) return scopeId;
        const limit = asInt(args.limit, 20);
        if (limit < 1) return bad("limit must be >= 1");
        const maxScan = asInt(args.maxScan, 200);
        if (maxScan < 1) return bad("maxScan must be >= 1");
        const data = await call("search", {
          query: query.value,
          scopeId: scopeId.value,
          limit,
          maxScan,
          includeContent: asBoolean(args.includeContent, false),
          includeRecycleBin: asBoolean(args.includeRecycleBin, false),
        });
        return ok(data);
      });
    },

    /** Create a page from markdown. */
    onenote_create(args = {}) {
      return guard(async () => {
        const section = sectionLocator(args);
        if (!section.ok) return section;
        const markdown = requiredString(args.markdown, "markdown", { maxLength: 1024 * 1024 });
        if (!markdown.ok) return markdown;
        const { blocks, titleFromMarkdown } = blocksFromMarkdown(markdown.value, renderOptions);
        const explicitTitle = optionalString(args.title, "title", { maxLength: 512 });
        if (!explicitTitle.ok) return explicitTitle;
        const title = explicitTitle.value ?? titleFromMarkdown;
        if (!title) return bad("title is required (pass title, or start markdown with '# Title')");
        removeTitleBlock(blocks);
        const data = await call("create", { ...section.value, title, blocks });
        return ok(data);
      });
    },

    /** Replace the text content of an existing page. */
    onenote_update(args = {}) {
      return guard(async () => {
        const locator = pageLocator(args);
        if (!locator.ok) return locator;
        const markdown = requiredString(args.markdown, "markdown", { maxLength: 1024 * 1024 });
        if (!markdown.ok) return markdown;
        const { blocks, titleFromMarkdown } = blocksFromMarkdown(markdown.value, renderOptions);
        const explicitTitle = optionalString(args.title, "title", { maxLength: 512 });
        if (!explicitTitle.ok) return explicitTitle;
        const title = explicitTitle.value ?? titleFromMarkdown;
        removeTitleBlock(blocks);
        const data = await call("update", { ...locator.value, title, blocks });
        return ok(data);
      });
    },

    /** Append markdown to the end of an existing page. */
    onenote_append(args = {}) {
      return guard(async () => {
        const locator = pageLocator(args);
        if (!locator.ok) return locator;
        const markdown = requiredString(args.markdown, "markdown", { maxLength: 1024 * 1024 });
        if (!markdown.ok) return markdown;
        const { blocks, titleFromMarkdown } = blocksFromMarkdown(markdown.value, renderOptions);
        removeTitleBlock(blocks);
        const explicitTitle = optionalString(args.title, "title", { maxLength: 512 });
        if (!explicitTitle.ok) return explicitTitle;
        const data = await call("append", {
          ...locator.value,
          title: explicitTitle.value,
          blocks,
        });
        return ok({ ...data, appendedFromTitle: titleFromMarkdown ?? undefined });
      });
    },

    /**
     * Edit one block of a page in place, leaving every other block untouched.
     *
     * The bridge re-sends only the outline that owns the block; OneNote replaces
     * that object by objectID and keeps the rest of the page as it was.
     */
    onenote_edit(args = {}) {
      return guard(async () => {
        const locator = pageLocator(args);
        if (!locator.ok) return locator;
        const mode = requiredString(args.mode, "mode", { maxLength: 32 });
        if (!mode.ok) return mode;
        const objectId = optionalString(args.objectId, "objectId", { maxLength: 512 });
        if (!objectId.ok) return objectId;
        const find = optionalString(args.find, "find", { maxLength: 4096 });
        if (!find.ok) return find;
        if (!objectId.value && !find.value) {
          return bad("objectId or find is required (onenote_read with withIds=true gives objectIds)");
        }
        const validModes = ["replaceText", "replaceBlock", "insertAfter", "deleteBlock"];
        if (!validModes.includes(mode.value)) {
          return bad(`mode must be one of: ${validModes.join(", ")}`);
        }

        const params = { ...locator.value, mode: mode.value, objectId: objectId.value ?? "", ...(find.value ? { find: find.value } : {}) };

        if (mode.value === "replaceText") {
          const search = requiredString(args.search, "search", { maxLength: 4096 });
          if (!search.ok) return search;
          if (args.replacement !== undefined && typeof args.replacement !== "string") {
            return bad("replacement must be a string");
          }
          params.search = search.value;
          params.replacement = typeof args.replacement === "string" ? args.replacement : "";
          params.all = asBoolean(args.all, false);
        } else if (mode.value === "replaceBlock" || mode.value === "insertAfter") {
          if (typeof args.text !== "string") return bad("text must be a string");
          if (args.text.length > 1024 * 1024) return bad("text is too long");
          params.text = args.text;
          if (mode.value === "insertAfter") {
            const kind = optionalString(args.kind, "kind", { maxLength: 16 });
            if (!kind.ok) return kind;
            const kinds = ["para", "h2", "h3", "bullet", "number"];
            if (kind.value && !kinds.includes(kind.value)) {
              return bad(`kind must be one of: ${kinds.join(", ")}`);
            }
            params.kind = kind.value ?? "para";
          }
        }

        const data = await call("edit", params);
        return ok({ ...data });
      });
    },

    /** Export a page to a file (.one/.pdf/.docx/...). */
    onenote_publish(args = {}) {
      return guard(async () => {
        const locator = pageLocator(args);
        if (!locator.ok) return locator;
        const target = requiredString(args.path, "path", { maxLength: 1024 });
        if (!target.ok) return target;
        const format = optionalString(args.format, "format", { maxLength: 16 });
        if (!format.ok) return format;
        const allowed = ["one", "onepkg", "mht", "pdf", "xps", "docx", "doc", "emf", "html"];
        if (format.value && !allowed.includes(format.value.toLowerCase())) {
          return bad(`format must be one of: ${allowed.join(", ")}`);
        }
        const data = await call("publish", {
          ...locator.value,
          path: target.value,
          format: format.value ?? "one",
          overwrite: asBoolean(args.overwrite, true),
        });
        return ok(data);
      });
    },

    /** Delete a page (recycle bin by default). */
    onenote_delete(args = {}) {
      return guard(async () => {
        const locator = pageLocator(args);
        if (!locator.ok) return locator;
        const data = await call("delete", {
          ...locator.value,
          permanent: asBoolean(args.permanent, false),
        });
        return ok(data);
      });
    },

    /** Create (or open) a standalone .one section file. */
    onenote_create_section(args = {}) {
      return guard(async () => {
        const target = requiredString(args.path, "path", { maxLength: 1024 });
        if (!target.ok) return target;
        const data = await call("createsection", {
          path: target.value,
          keepBlankPage: asBoolean(args.keepBlankPage, false),
        });
        return ok(data);
      });
    },

    /** Rename a section (notebook sections only; standalone ones are read-only). */
    onenote_rename_section(args = {}) {
      return guard(async () => {
        const sectionId = optionalString(args.sectionId, "sectionId", { maxLength: 256 });
        if (!sectionId.ok) return sectionId;
        const path = optionalString(args.path, "path", { maxLength: 1024 });
        if (!path.ok) return path;
        if (!sectionId.value && !path.value) return bad("sectionId or path is required");
        const name = requiredString(args.name, "name", { maxLength: 256 });
        if (!name.ok) return name;
        const data = await call("renamesection", {
          sectionId: sectionId.value ?? "",
          path: path.value ?? "",
          name: name.value,
        });
        return ok(data);
      });
    },
  };
}

/**
 * Build the model-visible `blocks (objectId)` list for a withIds read.
 * Both hosts use this: DSH's renderResult and the MCP adapter's text result.
 */
export function buildIdList(args, value) {
  const data = value && value.ok === true ? value.data : null;
  if (!data || args?.withIds !== true || !Array.isArray(data.blocks) || data.blocks.length === 0) return "";
  const rows = data.blocks.map((block, index) => {
    const kind = block?.type ?? "block";
    const text = String(block?.text ?? "").replace(/\s+/g, " ").slice(0, 60);
    return `${index + 1}. ${kind}  objectId=${block?.objectId || "(none)"}  ${text}`;
  });
  return `\n\nblocks (objectId):\n${rows.join("\n")}`;
}

/** Text rendering: markdown reads stay readable, everything else is JSON. */
export function renderResult(args, value) {
  if (value && value.ok === true && value.data && typeof value.data.markdown === "string") {
    const { title, pageId, lastModified, markdown, truncated, blockCount } = value.data;
    const header = `# ${title ?? ""}`.trimEnd();
    const note = truncated ? `\n> ⚠ 内容已截断：仅返回前 ${blockCount ?? "部分"} 个块，用 maxBlocks 调整。` : "";
    const footer = `\n---\npageId: ${pageId ?? ""}\nlastModified: ${lastModified ?? ""}`;
    return [{ type: "text", text: `${header}\n\n${markdown}${note}${footer}${buildIdList(args, value)}` }];
  }
  return [{ type: "text", text: JSON.stringify(value ?? null, null, 2) }];
}

export { markdownToBlocks, blocksToMarkdown, extractTitle, stripLeadingTitle };
