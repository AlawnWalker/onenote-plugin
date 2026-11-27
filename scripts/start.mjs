/**
 * Plugin launcher: install runtime dependencies on first run, then start the
 * MCP server as a child process on inherited stdio. npm output is kept off
 * stdout so it cannot corrupt the JSON-RPC channel.
 */
import { spawn, spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = dirname(here);
const sdkMarker = join(root, "node_modules", "@modelcontextprotocol", "server", "package.json");

// Optional per-user settings: ~/.cc-onenote.json. Values apply only when the
// corresponding environment variable is not already set. Arrays/objects are
// serialized as JSON, which is what the server expects for list/object options.
const configPath = process.env.CC_ONENOTE_CONFIG || join(homedir(), ".cc-onenote.json");
if (existsSync(configPath)) {
  try {
    const values = JSON.parse(readFileSync(configPath, "utf8"));
    for (const [key, value] of Object.entries(values)) {
      if (process.env[key] === undefined) {
        process.env[key] = typeof value === "string" ? value : JSON.stringify(value);
      }
    }
  } catch (err) {
    console.error("onenote-plugin: ignoring invalid config " + configPath + ": " + err.message);
  }
}

if (!existsSync(sdkMarker)) {
  const npmArgs = ["install", "--omit=dev", "--no-audit", "--no-fund"];
  const result =
    process.platform === "win32"
      ? spawnSync(process.env.ComSpec || "cmd.exe", ["/d", "/s", "/c", "npm " + npmArgs.join(" ")], {
          cwd: root,
          stdio: ["ignore", "ignore", "inherit"],
        })
      : spawnSync("npm", npmArgs, {
          cwd: root,
          stdio: ["ignore", "ignore", "inherit"],
        });
  if (result.status !== 0) {
    console.error("onenote-plugin: dependency installation failed; run `npm install` in " + root);
    process.exit(1);
  }
}

const child = spawn(process.execPath, [join(root, "server.js")], {
  stdio: "inherit",
  env: process.env,
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => {
    if (!child.killed) child.kill(signal);
  });
}

child.on("exit", (code, signal) => {
  process.exit(code ?? (signal ? 1 : 0));
});
