---
name: onenote
description: Read and write the local Microsoft OneNote desktop app through the bundled onenote MCP tools. Use when the user asks to list notebooks/sections/pages, read or search notes, create/update/append/publish/delete OneNote pages, or import Markdown into OneNote.
---

# OneNote (local desktop)

This plugin bundles the `onenote` MCP server, which drives the **OneNote desktop
app on Windows via COM automation**. It works on the real local notebooks and
`.one` section files open on this machine, not on cloud Graph copies.

## Requirements

- Windows with the OneNote desktop app installed.
- Node.js >= 18. The plugin launcher installs the server's npm dependencies on
  first run.

## Tools

| Tool | Use |
|---|---|
| `onenote_list` | Notebook -> section -> page tree, with IDs |
| `onenote_read` | Read a page as Markdown |
| `onenote_search` | Search titles (fast) or page content |
| `onenote_create` | Create a page from Markdown |
| `onenote_update` | Rewrite an existing page from Markdown |
| `onenote_append` | Append an outline to a page |
| `onenote_edit` | Edit one block in place (`replaceText` / `replaceBlock` / `insertAfter` / `deleteBlock`) |
| `onenote_publish` | Export a page to `.one` / `.pdf` / `.docx` / `.html` / ... |
| `onenote_delete` | Delete a page (recycle bin by default) |
| `onenote_create_section` | Create or open a standalone `.one` section |
| `onenote_rename_section` | Rename a section (notebook sections only) |

Page locator: `pageId` first; otherwise `sectionId`/`sectionName` plus
`pageTitle` (exact, then substring).

## Workflow

1. Use `onenote_list` to find the target section/page.
2. Read before rewriting: `onenote_update` replaces all text outlines on the
   page, so capture anything the user wants preserved first.
3. Confirm with the user before `onenote_delete` or a destructive
   `onenote_update`.
4. For chaptered Markdown import, use `dev/import-md.mjs` (one page per `##`
   chapter) instead of creating pages one by one.

## Markdown subset

`#` title, `##`/`###` headings (real font sizes), bullets, numbered lists
(stored as bullet + number), tables, fenced code (language tag lost), quotes.
Inline bold/italic/code/links are stripped. `onenote_update` applies the
configured page background.

## Configuration

The server reads environment variables. The plugin launcher also reads
`~/.cc-onenote.json` (override with `CC_ONENOTE_CONFIG`); keys are the same
environment-variable names, already-set variables win, and arrays/objects are
serialized as JSON.

- `ONENOTE_PAGECOLOR` — default `#C7EDCC`; `automatic` or `none` to inherit.
- `ONENOTE_PAGESIZE` — JSON object, or `automatic`.
- `ONENOTE_EXTRASECTIONS` — JSON array `[{ "name", "path" }]` for standalone
  `.one` sections. Standalone sections are invisible to `onenote_list` unless
  added here; they cannot be renamed or deleted through COM.
- `ONENOTE_TIMEOUT_MS` — per-call timeout, default 180000.

## Known limitations

- First call may take 1-3 minutes if OneNote is not already running.
- COM automation fails inside a file sandbox (`0x80080005`); run non-sandboxed.
- `onenote_publish` target paths must use backslashes; PDF exports have a white
  background.
- Long ASCII diagrams wrap at the page width (A4 body ~468pt, about 61
  half-width characters per line).
- Standalone `.one` sections cannot be renamed or deleted (`0x8004200B
  SECTION_READ_ONLY`); remove the file and restart OneNote instead.
- `onenote_create_section` may leave one blank page; the tool deletes it when
  the section is otherwise empty.
