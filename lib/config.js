/**
 * onenote-plugin — configuration normalization.
 *
 * `normalizeConfig` is the host-agnostic entry point: it accepts a flat
 * camelCase object (DSH `cordis.patch.yml`, a test, or the env mapping) and
 * returns the complete, validated config used by the bridge and renderers.
 *
 * `readConfig` maps the ONENOTE_* environment variables used by the MCP
 * adapter (Claude Code) onto that same shape.
 */

const DEFAULTS = Object.freeze({
  powershellPath: "powershell.exe",
  timeoutMs: 180000,
  // Matches the DSH plugin: generated pages get the eye-friendly green
  // (199,237,204). Set ONENOTE_PAGECOLOR="" to inherit OneNote's default.
  pageColor: "#C7EDCC",
  pageSize: null,
  // Code blocks: each source line gets this prefix. The bar is a readable gutter
  // marker; read-back also recognises code by its pinned monospace style, so ""
  // only drops the gutter text and no longer loses the code type.
  codePrefix: "│ ",
  // 0 = leave comments where they are; "auto" aligns trailing // comments of
  // one fenced block to longest code + 2; N > 0 = fixed column.
  alignCommentColumn: 0,
  // Upper bound for the automatic comment column.
  alignCommentMax: 80,
  // OneNote bullet glyph index: 0 big black circle, 1 U+2022 (solid dot), ...
  bulletValue: 1,
  // Optional per-level glyphs, e.g. [20, 2]; the last entry repeats.
  bulletValues: [],
  // Heading sizes relative to the body: [h2, h3]. Default [2, 1] so a 12pt
  // notebook gets 14/13; pages written by older versions ([4, 2], 16/14) still
  // read back correctly.
  headingDelta: [2, 1],
  // Bullet glyph size: "body" follows the body size; a number is explicit pt.
  bulletFontSize: "body",
  // 0.5.8 switches: nest content under headings, and treat indented lines
  // after a list item as its children.
  nestUnderHeadings: false,
  indentedChildren: false,
  extraSections: [],
});

function parseIntStrict(raw, fallback) {
  if (raw === undefined || raw === null || raw === "") return fallback;
  const n = Number.parseInt(raw, 10);
  return Number.isFinite(n) ? n : fallback;
}

function parseJson(raw, fallback) {
  if (raw === undefined || raw === null || raw === "") return fallback;
  try {
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

/** Accept booleans, 0/1 numbers, and "1"/"true" strings from host configs. */
function asSwitch(value) {
  return value === true || value === 1 || value === "1" || value === "true";
}

/**
 * Page size in OneNote points (1/72 in). OneNote only accepts either
 * <Automatic/> or Orientation+Dimensions+Margins, so width/height are required.
 * The string "automatic" forces OneNote's Automatic size; null leaves the page
 * size untouched. Mirrors the DSH plugin's normalizePageSize.
 */
function normalizePageSize(value) {
  if (typeof value === "string") {
    return value.trim().toLowerCase() === "automatic" ? "automatic" : null;
  }
  if (!value || typeof value !== "object") return null;
  const width = Number(value.width);
  const height = Number(value.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
  const out = { width, height, landscape: Boolean(value.landscape) };
  if (value.margins && typeof value.margins === "object") {
    const margins = {};
    for (const key of ["top", "bottom", "left", "right"]) {
      const n = Number(value.margins[key]);
      if (Number.isFinite(n) && n >= 0) margins[key] = n;
    }
    if (Object.keys(margins).length > 0) out.margins = margins;
  }
  return out;
}

/**
 * Standalone sections (not part of an open notebook) get a fresh OneNote ID
 * after every restart, so the reliable key is the .one file path, not a cached
 * ID. The bridge resolves each path to its current ID per call.
 */
function normalizeExtraSections(value) {
  if (!Array.isArray(value)) return [];
  return value
    .filter((entry) => entry && typeof entry === "object" && (typeof entry.path === "string" || typeof entry.id === "string"))
    .map((entry) => {
      const path = typeof entry.path === "string" ? entry.path.trim() : "";
      const id = typeof entry.id === "string" ? entry.id.trim() : "";
      return {
        name: typeof entry.name === "string" && entry.name.trim() !== "" ? entry.name : path || id,
        path,
        id,
      };
    })
    .filter((entry) => entry.path !== "" || entry.id !== "");
}

/**
 * Heading sizes relative to the body: `[h2, h3]` (also accepts `{h2, h3}` or a
 * JSON string from an env var). Invalid values fall back to [2, 1].
 */
function normalizeHeadingDelta(value) {
  const parsed = typeof value === "string" ? parseJson(value, null) : value;
  let h2 = DEFAULTS.headingDelta[0];
  let h3 = DEFAULTS.headingDelta[1];
  if (Array.isArray(parsed)) {
    h2 = Number(parsed[0]);
    h3 = Number(parsed[1]);
  } else if (parsed && typeof parsed === "object") {
    h2 = Number(parsed.h2);
    h3 = Number(parsed.h3);
  }
  const clamp = (n, fallback) => (Number.isFinite(n) && n >= 0 && n <= 20 ? n : fallback);
  return [clamp(h2, DEFAULTS.headingDelta[0]), clamp(h3, DEFAULTS.headingDelta[1])];
}

/** Bullet glyph size: "body" (default) or a positive point size. */
function normalizeBulletFontSize(value) {
  if (value === undefined || value === null) return DEFAULTS.bulletFontSize;
  const text = String(value).trim().toLowerCase();
  if (text === "" || text === "body") return DEFAULTS.bulletFontSize;
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? n : DEFAULTS.bulletFontSize;
}

/** Normalize a flat camelCase config object from any host. */
export function normalizeConfig(raw = {}) {
  const input = raw && typeof raw === "object" ? raw : {};

  const powershellPath =
    typeof input.powershellPath === "string" && input.powershellPath.trim() !== ""
      ? input.powershellPath.trim()
      : DEFAULTS.powershellPath;

  const timeoutRaw = parseIntStrict(input.timeoutMs, DEFAULTS.timeoutMs);
  const timeoutMs = Number.isFinite(timeoutRaw) && timeoutRaw > 0 ? timeoutRaw : DEFAULTS.timeoutMs;

  // Unset falls back to the green default; an explicit empty string inherits
  // OneNote's own page color.
  const pageColorRaw = input.pageColor;
  const pageColor = (pageColorRaw === undefined || pageColorRaw === null ? DEFAULTS.pageColor : String(pageColorRaw)).trim();
  const pageColorSafe =
    /^#[0-9a-f]{6}$/i.test(pageColor) || pageColor === "automatic" || pageColor === "none" ? pageColor : "";

  const pageSize = normalizePageSize(
    typeof input.pageSize === "string" ? parseJson(input.pageSize, null) : input.pageSize,
  );

  // Code-block rendering (JS side): prefix and trailing-comment alignment.
  const codePrefix = typeof input.codePrefix === "string" ? input.codePrefix : DEFAULTS.codePrefix;
  const alignRaw = input.alignCommentColumn;
  let alignCommentColumn = 0;
  if (alignRaw === "auto" || alignRaw === "true" || alignRaw === true) {
    alignCommentColumn = "auto";
  } else {
    const column = parseIntStrict(alignRaw, 0);
    if (column > 0) alignCommentColumn = column;
  }
  const alignMax = parseIntStrict(input.alignCommentMax, DEFAULTS.alignCommentMax);
  const alignCommentMax = alignMax > 0 ? alignMax : DEFAULTS.alignCommentMax;

  // Bullet glyphs (PowerShell side).
  const bulletRaw = parseIntStrict(input.bulletValue, DEFAULTS.bulletValue);
  const bulletValue = bulletRaw >= 0 ? bulletRaw : DEFAULTS.bulletValue;
  const bulletParsed = typeof input.bulletValues === "string" ? parseJson(input.bulletValues, []) : input.bulletValues;
  const bulletValues = Array.isArray(bulletParsed)
    ? bulletParsed.map((value) => Number(value)).filter((value) => Number.isFinite(value) && value >= 0)
    : [];

  const extraParsed =
    typeof input.extraSections === "string" ? parseJson(input.extraSections, []) : input.extraSections;

  return {
    powershellPath,
    timeoutMs,
    pageColor: pageColorSafe,
    pageSize,
    codePrefix,
    alignCommentColumn,
    alignCommentMax,
    bulletValue,
    bulletValues,
    headingDelta: normalizeHeadingDelta(input.headingDelta),
    bulletFontSize: normalizeBulletFontSize(input.bulletFontSize),
    nestUnderHeadings: asSwitch(input.nestUnderHeadings),
    indentedChildren: asSwitch(input.indentedChildren),
    extraSections: normalizeExtraSections(extraParsed),
  };
}

export function readConfig(env = process.env) {
  return normalizeConfig({
    powershellPath: env.ONENOTE_POWERSHELL,
    timeoutMs: env.ONENOTE_TIMEOUT_MS,
    pageColor: env.ONENOTE_PAGECOLOR,
    pageSize: env.ONENOTE_PAGESIZE,
    codePrefix: env.ONENOTE_CODE_PREFIX,
    alignCommentColumn: env.ONENOTE_ALIGN_COMMENT_COLUMN,
    alignCommentMax: env.ONENOTE_ALIGN_COMMENT_MAX,
    bulletValue: env.ONENOTE_BULLET_VALUE,
    bulletValues: env.ONENOTE_BULLET_VALUES,
    headingDelta: env.ONENOTE_HEADING_DELTA,
    bulletFontSize: env.ONENOTE_BULLET_FONT_SIZE,
    nestUnderHeadings: env.ONENOTE_NEST_UNDER_HEADINGS,
    indentedChildren: env.ONENOTE_INDENTED_CHILDREN,
    extraSections: env.ONENOTE_EXTRASECTIONS,
  });
}
