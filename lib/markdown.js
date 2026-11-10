/**
 * Markdown subset <-> OneNote block model.
 *
 * OneNote 2013+ refuses `one:Style` / `QuickStyles` imports, and numbered-list
 * import (`one:Number`) fails as well, but a per-paragraph `style` attribute
 * plus `spaceBefore`/`spaceAfter` on `one:OE` DO import, so headings are real
 * formatted paragraphs:
 *
 *   ## Heading      ->  16pt paragraph, 8pt before / 5pt after
 *   ### Heading     ->  14pt paragraph, 5pt before / 5pt after (body is 12pt)
 *   > quote         ->  ▎ quote
 *   ```code```      ->  │ code line      (one outline element per line)
 *   1. item         ->  • 1. item         (bullet that carries its own number)
 *
 * The reverse conversion recognises both the font sizes (new pages) and the
 * legacy symbol markers (`━━━`, `▌`) written by earlier versions.
 */

/** Legacy symbols written by plugin versions <= 0.1.1; still read back. */
export const H2_PREFIX = "━━━";
export const H3_MARK = "▌";
export const CODE_MARK = "│";
export const QUOTE_MARK = "▎";

const H2_RE = /^━{2,}\s*(.*?)\s*━{2,}$/;
const H3_RE = /^▌\s*(.*)$/;
const QUOTE_RE = /^▎\s?(.*)$/;
const CODE_RE = /^│\s?(.*)$/;
const NUMBER_RE = /^(\d+)[.)]\s+(.*)$/;

function indentLevel(prefix) {
  // Tabs count as one level, every two spaces count as one level.
  let level = 0;
  for (const ch of prefix) {
    if (ch === "\t") level += 1;
    else if (ch === " ") level += 0.5;
  }
  return Math.max(0, Math.floor(level));
}

function splitTableRow(line) {
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  const cells = [];
  let current = "";
  let escaped = false;
  for (const ch of trimmed) {
    if (escaped) {
      current += ch;
      escaped = false;
      continue;
    }
    if (ch === "\\") {
      escaped = true;
      continue;
    }
    if (ch === "|") {
      cells.push(current.trim());
      current = "";
      continue;
    }
    current += ch;
  }
  cells.push(current.trim());
  return cells;
}

function isTableSeparator(line) {
  return /^\s*\|?[\s:|-]*-[\s:|-]*\|?\s*$/.test(line) && /-/.test(line);
}

/** Convert a markdown subset into the block model the bridge renders. */
/**
 * Index of the trailing `//` comment on a code line, or -1.
 * Skips whole-line comments, `://` inside URLs and `//` inside strings.
 */
/**
 * Display width of a string: ASCII = 1, CJK/full-width = 2 (wcwidth semantics).
 * Used so Chinese comments line up with ASCII ones.
 */
export function displayWidth(text) {
  let width = 0;
  for (const ch of String(text)) {
    const code = ch.codePointAt(0);
    const wide =
      (code >= 0x1100 && code <= 0x115f) ||
      (code >= 0x2e80 && code <= 0xa4cf) ||
      (code >= 0xac00 && code <= 0xd7a3) ||
      (code >= 0xf900 && code <= 0xfaff) ||
      (code >= 0xfe30 && code <= 0xfe4f) ||
      (code >= 0xff00 && code <= 0xff60) ||
      (code >= 0xffe0 && code <= 0xffe6) ||
      (code >= 0x20000 && code <= 0x3fffd);
    width += wide ? 2 : 1;
  }
  return width;
}

/**
 * Index of the trailing `//` comment on a code line, or -1.
 *
 * Rules: the last `//` that is preceded by whitespace; `http://`, `file://` and
 * any scheme:// are skipped (the character before `//` is a colon); a
 * whole-line comment is not a trailing comment.
 */
export function findInlineComment(line) {
  const text = String(line);
  if (text.trimStart().startsWith("//")) return -1;
  let found = -1;
  for (let i = 1; i < text.length - 1; i += 1) {
    if (text[i] !== "/" || text[i + 1] !== "/") continue;
    const before = text[i - 1];
    if (before === ":") continue;              // scheme://
    if (!/\s/.test(before)) continue;          // needs whitespace before it
    found = i;
  }
  return found;
}

/**
 * Align trailing `//` comments inside one fenced block.
 *
 * column > 0 aligns to that column; otherwise the column is the widest code
 * part (indentation included, trailing blanks trimmed) + 2, capped by
 * maxColumn. Every commented line keeps at least one space before `//`.
 */
export function alignCodeComments(lines, column = 0, options = {}) {
  const source = Array.isArray(lines) ? lines : [];
  const maxColumn = Number.isFinite(options.maxColumn) && options.maxColumn > 0 ? options.maxColumn : 80;
  const hits = source
    .map((line) => ({ line, index: findInlineComment(line) }))
    .filter((entry) => entry.index > 0);
  if (hits.length === 0) return source.slice();
  const codePart = (entry) => entry.line.slice(0, entry.index).replace(/\s+$/, "");
  const widest = Math.max(...hits.map((entry) => displayWidth(codePart(entry))));
  let target = column > 0 ? column : widest + 2;
  if (column <= 0) target = Math.min(target, maxColumn);
  return source.map((line) => {
    const index = findInlineComment(line);
    if (index <= 0) return line;
    const code = line.slice(0, index).replace(/\s+$/, "");
    const padding = Math.max(1, target - displayWidth(code));
    return `${code}${" ".repeat(padding)}${line.slice(index)}`;
  });
}
export function markdownToBlocks(markdown, options = {}) {
  const codePrefix = options.codePrefix === undefined ? `${CODE_MARK} ` : String(options.codePrefix);
  // 0 / undefined = off; "auto" = longest code + 2 (capped); N > 0 = fixed column.
  const rawColumn = options.alignCommentColumn;
  const alignMode = rawColumn === "auto" ? "auto" : Number.isFinite(rawColumn) && rawColumn > 0 ? rawColumn : 0;
  const alignMax = Number.isFinite(options.alignCommentMax) && options.alignCommentMax > 0 ? options.alignCommentMax : 80;
  // Indented lines that follow a list item become nested paragraphs instead of
  // being flattened into one paragraph with EM SPACE indentation.
  const indentedChildren = options.indentedChildren === true;
  const text = String(markdown ?? "").replace(/\r\n?/g, "\n");
  const lines = text.split("\n");
  const blocks = [];
  let paragraph = [];
  let inFence = false;
  let table = null;
  const counters = [];
  let fenceLines = [];
  // Level of the most recent list item, so an indented line right after it can
  // become its child. -1 means "no list item in context".
  let listContext = -1;

  // A fenced block becomes one "code" block per source line, each prefixed with
  // codePrefix (default the box-drawing bar). alignCommentColumn > 0 pushes the
  // trailing // comment of every line to one column.
  const emitCodeBlock = (sourceLines) => {
    const aligned = alignMode === 0 ? sourceLines : alignCodeComments(sourceLines, alignMode === "auto" ? 0 : alignMode, { maxColumn: alignMax });
    for (const sourceLine of aligned) {
      blocks.push({ type: "code", text: `${codePrefix}${sourceLine}`, level: 0 });
    }
  };

  const flushParagraph = () => {
    if (paragraph.length === 0) return;
    // Leading indentation is meaningful for ASCII trees (├─ / └─ drawings):
    // it survives into OneNote as EM SPACE characters, because regular spaces
    // would be collapsed or trimmed. Everything else is a normal paragraph.
    const first = paragraph[0];
    const indentMatch = first.match(/^( +)/);
    const indent = indentMatch ? "　".repeat(Math.floor(indentMatch[1].length / 2)) : "";
    const body = paragraph
      .map((line, index) => (index === 0 ? line.replace(/^ +/, "") : line.trim()))
      .join(" ")
      .trim();
    blocks.push({ type: "para", text: `${indent}${body}`, level: 0 });
    paragraph = [];
  };
  const flushTable = () => {
    if (table && table.length > 0) blocks.push({ type: "table", rows: table });
    table = null;
  };

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const fence = line.match(/^\s*```(.*)$/);
    if (fence) {
      flushParagraph();
      flushTable();
      if (inFence) {
        emitCodeBlock(fenceLines);
        fenceLines = [];
      }
      inFence = !inFence;
      continue;
    }
    if (inFence) {
      fenceLines.push(line);
      continue;
    }
    if (/^\s*$/.test(line)) {
      flushParagraph();
      flushTable();
      continue;
    }
    if (line.trim().startsWith("|") && !table) {
      // A table only starts when the following line is a separator row.
      const next = lines[i + 1] ?? "";
      if (next.trim().startsWith("|") && isTableSeparator(next)) {
        flushParagraph();
        table = [splitTableRow(line)];
        i += 1; // consume separator
        continue;
      }
    }
    if (table) {
      if (line.trim().startsWith("|")) {
        table.push(splitTableRow(line));
        continue;
      }
      flushTable();
    }
    const heading = line.match(/^\s*(#{1,6})\s+(.*)$/);
    if (heading) {
      flushParagraph();
      const depth = heading[1].length;
      const title = heading[2].trim();
      listContext = -1;
      if (depth === 1) {
        blocks.push({ type: "title", text: title, level: 0 });
      } else if (depth === 2) {
        blocks.push({ type: "h2", text: title, level: 0 });
      } else {
        blocks.push({ type: "h3", text: title, level: 0 });
      }
      continue;
    }
    const bullet = line.match(/^(\s*)[-*+]\s+(.*)$/);
    if (bullet) {
      flushParagraph();
      const bulletLevel = indentLevel(bullet[1]);
      blocks.push({ type: "bullet", text: bullet[2].trim(), level: bulletLevel });
      listContext = bulletLevel;
      continue;
    }
    const numbered = line.match(/^(\s*)(\d+)[.)]\s+(.*)$/);
    if (numbered) {
      flushParagraph();
      const level = indentLevel(numbered[1]);
      counters[level] = (counters[level] ?? 0) + 1;
      for (let deeper = level + 1; deeper < counters.length; deeper += 1) counters[deeper] = 0;
      const ordinal = Number.parseInt(numbered[2], 10) || counters[level];
      blocks.push({ type: "number", text: numbered[3].trim(), level, ordinal });
      listContext = level;
      continue;
    }
    const quote = line.match(/^\s*>\s?(.*)$/);
    if (quote) {
      flushParagraph();
      blocks.push({ type: "para", text: `${QUOTE_MARK} ${quote[1].trim()}`.trim(), level: 0 });
      continue;
    }
    // A line prefixed with U+200B is a tree line written the way OneNote keeps
    // it: the zero-width space stops trim() from eating the indentation.
    if (/^\u200b/.test(line)) {
      flushParagraph();
      blocks.push({ type: "para", text: line.trim(), level: 0 });
      continue;
    }
    // Tree drawings (├─ / └─ / │ with indentation) are one paragraph per line,
    // with the indentation preserved as EM SPACEs; OneNote would drop plain
    // leading spaces.
    if (/^[ \t]*[├└│]/.test(line) || /^[ \t]{2,}\S/.test(line)) {
      flushParagraph();
      const isTree = /^[ \t]*[├└│]/.test(line);
      // A plain indented line that follows a list item is a child of that item,
      // so the renderer can nest it (OneNote indents it one level) instead of
      // faking the indentation with EM SPACEs.
      if (indentedChildren && !isTree && listContext >= 0) {
        blocks.push({ type: "para", text: line.trim(), level: listContext + 1 });
        continue;
      }
      const lead = (line.match(/^ +/) ?? [""])[0].length;
      const indent = "　".repeat(Math.floor(lead / 2));
      blocks.push({ type: "para", text: `${indent}${line.trim()}`, level: 0 });
      continue;
    }
    paragraph.push(line.trim());
  }
  if (inFence && fenceLines.length > 0) {
    emitCodeBlock(fenceLines);
    fenceLines = [];
  }
  flushParagraph();
  flushTable();
  return blocks;
}

function renderTable(rows) {
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  if (width === 0) return [];
  const escape = (cell) => String(cell ?? "").replace(/\|/g, "\\|").replace(/\n/g, " ");
  const normalised = rows.map((row) => {
    const cells = [];
    for (let i = 0; i < width; i += 1) cells.push(escape(row[i] ?? ""));
    return cells;
  });
  const lines = [];
  lines.push(`| ${normalised[0].join(" | ")} |`);
  lines.push(`| ${normalised[0].map(() => "---").join(" | ")} |`);
  for (let i = 1; i < normalised.length; i += 1) {
    lines.push(`| ${normalised[i].join(" | ")} |`);
  }
  return lines;
}

/** Convert the bridge's block model back into markdown. */
export function blocksToMarkdown(blocks) {
  const input = Array.isArray(blocks) ? blocks : [];
  const lines = [];
  // Tracks what the previous block was so blank lines land in the right place:
  // list items stay adjacent, everything else is separated by one blank line.
  let previous = null; // null | "list" | "block" | "code"

  const blank = () => {
    if (lines.length > 0 && lines[lines.length - 1] !== "") lines.push("");
  };
  const closeFence = () => {
    if (previous !== "code") return;
    lines.push("```");
    lines.push("");
    previous = "block";
  };

  for (const block of input) {
    const type = block?.type ?? "para";
    const level = Number.isFinite(block?.level) ? block.level : 0;
    const raw = String(block?.text ?? "");
    const indent = "  ".repeat(Math.max(0, level));

    // Code lines coming back from OneNote are paragraphs that carry the gutter
    // marker, so normalise both sources into one fenced block.
    const codeMatch = raw.match(CODE_RE);
    if (type === "code" || (type === "para" && codeMatch)) {
      if (previous !== "code") {
        blank();
        lines.push("```");
        previous = "code";
      }
      lines.push(codeMatch ? codeMatch[1] : raw);
      continue;
    }
    closeFence();

    if (type === "title") {
      blank();
      lines.push(`# ${raw}`.trimEnd());
      previous = "block";
      continue;
    }
    if (type === "h2") {
      blank();
      lines.push(`## ${raw}`.trimEnd());
      previous = "block";
      continue;
    }
    if (type === "h3") {
      blank();
      lines.push(`### ${raw}`.trimEnd());
      previous = "block";
      continue;
    }
    if (type === "table") {
      blank();
      lines.push(...renderTable(Array.isArray(block.rows) ? block.rows : []));
      previous = "block";
      continue;
    }
    if (type === "image" || type === "ink" || type === "file") {
      if (!String(previous).startsWith("list")) blank();
      const marker = type === "image" ? "![图片]" : type === "ink" ? "![墨迹]" : "[嵌入文件]";
      lines.push(`${indent}${marker}`);
      previous = "block";
      continue;
    }
    if (type === "bullet" || type === "number") {
      const listKind = `list-${type}`;
      if (previous !== listKind) blank();
      const numbered = raw.match(NUMBER_RE);
      if (numbered) lines.push(`${indent}${numbered[1]}. ${numbered[2]}`.trimEnd());
      else if (type === "number") lines.push(`${indent}${block.ordinal ?? 1}. ${raw}`.trimEnd());
      else lines.push(`${indent}- ${raw}`.trimEnd());
      previous = listKind;
      continue;
    }

    // Paragraphs, and the symbol-based headings/quotes written by the bridge.
    blank();
    const h2 = raw.match(H2_RE);
    const h3 = raw.match(H3_RE);
    const quote = raw.match(QUOTE_RE);
    if (h2) lines.push(`## ${h2[1]}`.trimEnd());
    else if (h3) lines.push(`### ${h3[1]}`.trimEnd());
    else if (quote) lines.push(`> ${quote[1]}`.trimEnd());
    else if (raw.trim() !== "") lines.push(`${indent}${raw}`.trimEnd());
    previous = "block";
  }
  closeFence();

  // Collapse the trailing blank line and runs of blank lines.
  const cleaned = [];
  for (const line of lines) {
    if (line === "" && cleaned.length > 0 && cleaned[cleaned.length - 1] === "") continue;
    cleaned.push(line);
  }
  while (cleaned.length > 0 && cleaned[cleaned.length - 1] === "") cleaned.pop();
  return `${cleaned.join("\n")}\n`;
}

/** Pull a leading `# Title` out of a markdown document. */
export function extractTitle(markdown) {
  const match = String(markdown ?? "").match(/^\s*#\s+(.+?)\s*$/m);
  return match ? match[1] : null;
}

/** Drop the leading `# Title` line so it is not repeated as body text. */
export function stripLeadingTitle(markdown) {
  return String(markdown ?? "").replace(/^\s*#\s+(.+?)\s*(\r?\n|$)/, "");
}
