/**
 * onenote-mcp — markdown import helpers.
 *
 * Pure functions used by dev/import-md.mjs to split a markdown document into
 * OneNote pages (one per `##` chapter) and to strip inline formatting that the
 * OneNote markdown subset cannot represent (links, bold, inline code).
 */

/** Split a markdown document into chapters by `##` headings. */
export function splitChapters(markdown) {
  const lines = String(markdown ?? "").split(/\r?\n/);
  const chapters = [];
  const preamble = [];
  let current = null;

  for (const line of lines) {
    const heading = line.match(/^##\s+(.*)$/);
    if (heading) {
      if (current) chapters.push(current);
      current = { title: heading[1].trim(), bodyLines: [] };
    } else if (current) {
      current.bodyLines.push(line);
    } else {
      preamble.push(line);
    }
  }
  if (current) chapters.push(current);
  return { preamble, chapters };
}

/** Strip inline formatting the OneNote subset cannot represent. */
export function stripInline(text) {
  return String(text ?? "")
    .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1") // [text](url) -> text
    .replace(/\*\*([^*]+)\*\*/g, "$1") // **bold** -> bold
    .replace(/`([^`]+)`/g, "$1"); // `code` -> code
}

/** Strip inline formatting from body lines, leaving fenced code intact. */
export function cleanBody(bodyLines) {
  const out = [];
  let inFence = false;
  for (const line of bodyLines) {
    if (/^\s*```/.test(line)) {
      inFence = !inFence;
      out.push(line);
      continue;
    }
    out.push(inFence ? line : stripInline(line));
  }
  return out.join("\n");
}

/** Extract the intro quote from the preamble (drop the H1 title and `---`). */
export function preambleIntro(preamble) {
  return preamble
    .filter((line) => {
      const t = line.trim();
      return t !== "" && !/^---\s*$/.test(t) && !/^#\s+/.test(t);
    })
    .join("\n");
}
