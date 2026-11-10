/**
 * Host contract check — validates the plugin against the REAL
 * `@deepseek-ai/dsh-tools`.defineTool that ships inside the DSH app.
 *
 * Why this exists: DSH does not accept arbitrary JSON Schema. Its parameters
 * spec is a small DSL (string/boolean/integer/array/object + descriptions and
 * enums); unsupported keywords such as `minimum`, `maximum` or `pattern` throw
 * `JsonSchemaError` at activation time, which means the plugin silently never
 * registers a single tool — and you only find out after restarting the app.
 *
 * Run it after any change to the shared tool metadata:
 *
 *   node dev/host-contract.mjs [path-to-asar-dsh-node_modules]
 *
 * It needs Electron (the app binary) because `@deepseek-ai/dsh-tools` lives
 * inside app.asar. Set DSH_ELECTRON to override the detected executable.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import process from "node:process";

// No absolute paths in the repo: point these at your install through the env.
//   DSH_ELECTRON       e.g. <install dir>/DeepSeek Harness.exe
//   DSH_ASAR_MODULES   e.g. <install dir>/resources/app.asar/dsh/node_modules
const ELECTRON = process.env.DSH_ELECTRON ?? "";
const ASAR_MODULES = process.argv[2] ?? process.env.DSH_ASAR_MODULES ?? "";
const TOOLS = `${ASAR_MODULES}/@deepseek-ai/dsh-tools`;
const PLUGIN = new URL("../dsh/index.js", import.meta.url).href;

if (ELECTRON === "" || !existsSync(ELECTRON)) {
  console.error("SKIP: set DSH_ELECTRON to your DSH executable and DSH_ASAR_MODULES to <install>/resources/app.asar/dsh/node_modules");
  process.exit(0);
}
// app.asar is a file to plain Node, so only require the value (the child
// Electron process resolves paths inside the archive).
if (ASAR_MODULES === "") {
  console.error("SKIP: set DSH_ASAR_MODULES to <install>/resources/app.asar/dsh/node_modules");
  process.exit(0);
}

const probe = `
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const tools = require(${JSON.stringify(TOOLS)});
const plugin = await import(${JSON.stringify(PLUGIN)});
const registered = [];
const ctx = { tools: { register: (def) => registered.push(def) } };
plugin.applyWith({ defineTool: tools.defineTool, ctx, config: {} });
for (const def of registered) {
  if (typeof def.execute !== "function") throw new Error(def.name + ": execute missing");
  if (!def.parameters || typeof def.parameters !== "object") throw new Error(def.name + ": parameters missing");
  if (typeof def.output?.render !== "function") throw new Error(def.name + ": output.render missing");
}
console.log("HOST-CONTRACT OK:", registered.map((d) => d.name).join(","));
`;

const result = spawnSync(ELECTRON, ["--expose-internals", "--input-type=module", "-e", probe], {
  env: { ...process.env, ELECTRON_RUN_AS_NODE: "1" },
  encoding: "utf8",
});

const output = `${result.stdout ?? ""}${result.stderr ?? ""}`.trim();
console.log(output);
if (result.status !== 0 || !output.includes("HOST-CONTRACT OK")) {
  console.error("\nhost contract check failed");
  process.exit(1);
}
