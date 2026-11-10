/**
 * dsh-onenote — bridge to the local OneNote desktop application.
 *
 * OneNote only exposes a COM automation API, so every operation runs through a
 * short-lived Windows PowerShell process (the plugin itself stays dependency
 * free). The request/response protocol is JSON on temp files, which keeps
 * quoting, code pages and long CDATA payloads out of the command line.
 */
import { spawn } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
export const BRIDGE_SCRIPT = path.resolve(HERE, "..", "scripts", "onenote-bridge.ps1");

export class BridgeError extends Error {
  constructor(code, message, detail) {
    super(message);
    this.name = "BridgeError";
    this.code = code;
    if (detail !== undefined) this.detail = detail;
  }
}

function runProcess(command, args, timeoutMs) {
  return new Promise((resolve, reject) => {
    let child;
    try {
      child = spawn(command, args, { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    } catch (err) {
      reject(new BridgeError("POWERSHELL_NOT_FOUND", `cannot start ${command}: ${err.message}`));
      return;
    }
    let stdout = "";
    let stderr = "";
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      try {
        child.kill();
      } catch {
        /* ignore */
      }
      reject(new BridgeError("TIMEOUT", `OneNote did not answer within ${timeoutMs} ms`));
    }, timeoutMs);

    child.stdout?.on("data", (chunk) => {
      if (stdout.length < 65536) stdout += chunk.toString("utf8");
    });
    child.stderr?.on("data", (chunk) => {
      if (stderr.length < 65536) stderr += chunk.toString("utf8");
    });
    child.on("error", (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      reject(new BridgeError("POWERSHELL_FAILED", `cannot start ${command}: ${err.message}`));
    });
    child.on("close", (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve({ code: code ?? 0, stdout, stderr });
    });
  });
}

export class OneNoteBridge {
  constructor({ powershellPath = "powershell.exe", timeoutMs = 180000, extraSections = [], pageColor = "", pageSize = null, bulletValue = 1, bulletValues = [], nestUnderHeadings = false } = {}) {
    this.powershellPath = powershellPath || "powershell.exe";
    this.timeoutMs = Number.isFinite(timeoutMs) && timeoutMs > 0 ? timeoutMs : 180000;
    this.extraSections = Array.isArray(extraSections) ? extraSections : [];
    this.pageColor = typeof pageColor === "string" ? pageColor.trim() : "";
    this.pageSize = pageSize && typeof pageSize === "object" ? pageSize : null;
    // OneNote bullet glyph index; 1 = U+2022 (solid dot).
    const bullet = Number(bulletValue);
    this.bulletValue = Number.isFinite(bullet) && bullet >= 0 ? bullet : 1;
    // Per-level bullet glyphs: index 0 is the top level, the last entry repeats.
    this.nestUnderHeadings = nestUnderHeadings === true;
    this.bulletValues = Array.isArray(bulletValues)
      ? bulletValues.map((value) => Number(value)).filter((value) => Number.isFinite(value) && value >= 0)
      : [];
  }

  /**
   * Run one bridge action. Resolves to the `data` payload; throws BridgeError
   * with the plugin's error code when OneNote rejects the operation.
   */
  async invoke(action, params = {}) {
    const dir = await mkdtemp(path.join(os.tmpdir(), "dsh-onenote-"));
    const requestPath = path.join(dir, "request.json");
    const responsePath = path.join(dir, "response.json");
    try {
      const payload = {
        action,
        params: {
          pageColor: this.pageColor,
          pageSize: this.pageSize,
          bulletValue: this.bulletValue,
          bulletValues: this.bulletValues,
          nestUnderHeadings: this.nestUnderHeadings,
          ...params,
          extraSections: this.extraSections,
        },
      };
      await writeFile(requestPath, JSON.stringify(payload), "utf8");
      const run = await runProcess(
        this.powershellPath,
        [
          "-NoProfile",
          "-NonInteractive",
          "-ExecutionPolicy",
          "Bypass",
          "-File",
          BRIDGE_SCRIPT,
          "-Request",
          requestPath,
          "-Response",
          responsePath,
        ],
        this.timeoutMs,
      );

      let raw;
      try {
        raw = await readFile(responsePath, "utf8");
      } catch {
        const detail = (run.stderr || run.stdout || "").trim().slice(0, 2000);
        throw new BridgeError(
          "BRIDGE_CRASHED",
          detail
            ? `OneNote bridge exited (${run.code}) without a response: ${detail}`
            : `OneNote bridge exited (${run.code}) without a response`,
        );
      }
      let parsed;
      try {
        parsed = JSON.parse(raw);
      } catch {
        throw new BridgeError("BAD_RESPONSE", `OneNote bridge returned invalid JSON: ${raw.slice(0, 500)}`);
      }
      if (!parsed || parsed.ok !== true) {
        const code = parsed && parsed.code ? String(parsed.code) : "ONENOTE_ERROR";
        const message = parsed && parsed.message ? String(parsed.message) : "OneNote operation failed";
        throw new BridgeError(code, message);
      }
      return parsed.data;
    } finally {
      await rm(dir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
