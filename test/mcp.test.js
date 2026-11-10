import assert from "node:assert/strict";
import test from "node:test";

import { InMemoryTransport, LATEST_PROTOCOL_VERSION } from "@modelcontextprotocol/server";

import { createMcpServer } from "../server.js";
import { makeToolHandlers } from "../lib/tools.js";
import { TOOL_NAMES } from "../lib/schemas.js";

function fakeBridge(responder) {
  return { async invoke(action, params) { return responder(action, params); } };
}

function failingBridge(code, message) {
  return {
    async invoke() {
      const err = new Error(message);
      err.code = code;
      throw err;
    },
  };
}

async function startServer(handlers) {
  const server = createMcpServer({ handlers });
  const [client, serverTransport] = InMemoryTransport.createLinkedPair();
  const received = [];
  client.onmessage = (message) => received.push(message);
  await server.connect(serverTransport);
  await client.start();

  await client.send({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      protocolVersion: LATEST_PROTOCOL_VERSION,
      capabilities: {},
      clientInfo: { name: "test", version: "1.0.0" },
    },
  });
  await client.send({ jsonrpc: "2.0", method: "notifications/initialized" });

  return { client, server, received };
}

const send = (client, id, method, params) =>
  client.send({ jsonrpc: "2.0", id, method, params });
const flush = () => new Promise((resolve) => setTimeout(resolve, 50));

test("tools/list advertises all 8 tools with standard JSON Schema", async () => {
  const handlers = makeToolHandlers({ bridge: fakeBridge(() => ({})) });
  const { client, server, received } = await startServer(handlers);

  await send(client, 2, "tools/list", {});
  await flush();

  const response = received.find((m) => m.id === 2);
  assert.ok(response, "tools/list response received");
  assert.equal(response.error, undefined);
  const tools = response.result.tools;
  assert.equal(tools.length, 11);
  assert.deepEqual(tools.map((t) => t.name).sort(), [...TOOL_NAMES].sort());
  for (const tool of tools) {
    assert.equal(tool.inputSchema.type, "object");
    assert.ok(typeof tool.description === "string");
  }
  await server.close();
});

test("tools/call returns markdown content for onenote_read", async () => {
  const bridge = fakeBridge((action) => {
    if (action === "read") {
      return {
        pageId: "{PAGE}",
        title: "页面",
        lastModified: "2026-11-03T00:00:00.000Z",
        blockCount: 1,
        truncated: false,
        blocks: [{ type: "para", text: "正文", level: 0 }],
      };
    }
    return {};
  });
  const handlers = makeToolHandlers({ bridge });
  const { client, server, received } = await startServer(handlers);

  await send(client, 2, "tools/call", { name: "onenote_read", arguments: { pageId: "{PAGE}" } });
  await flush();

  const response = received.find((m) => m.id === 2);
  assert.equal(response.result.isError, undefined);
  const text = response.result.content[0].text;
  assert.match(text, /^# 页面/);
  assert.match(text, /正文/);
  assert.match(text, /pageId: \{PAGE\}/);
  await server.close();
});

test("tools/call exposes objectIds for withIds reads", async () => {
  const bridge = fakeBridge((action) => {
    if (action === "read") {
      return {
        pageId: "{PAGE}",
        title: "页面",
        lastModified: "2026-11-03T00:00:00.000Z",
        blockCount: 2,
        truncated: false,
        blocks: [
          { type: "para", text: "正文", level: 0, objectId: "{OID-1}" },
          { type: "bullet", text: "项目", level: 0, objectId: "{OID-2}" },
        ],
      };
    }
    return {};
  });
  const handlers = makeToolHandlers({ bridge });
  const { client, server, received } = await startServer(handlers);

  await send(client, 2, "tools/call", {
    name: "onenote_read",
    arguments: { pageId: "{PAGE}", withIds: true },
  });
  await flush();

  const response = received.find((m) => m.id === 2);
  assert.equal(response.result.isError, undefined);
  const text = response.result.content[0].text;
  assert.match(text, /objectId=\{OID-1\}/);
  assert.match(text, /objectId=\{OID-2\}/);
  await server.close();
});

test("tools/call surfaces bridge errors as isError with the code", async () => {
  const bridge = failingBridge("PAGE_NOT_FOUND", "page not found: 不存在");
  const handlers = makeToolHandlers({ bridge });
  const { client, server, received } = await startServer(handlers);

  await send(client, 2, "tools/call", { name: "onenote_read", arguments: { pageTitle: "不存在" } });
  await flush();

  const response = received.find((m) => m.id === 2);
  assert.equal(response.result.isError, true);
  assert.match(response.result.content[0].text, /PAGE_NOT_FOUND/);
  await server.close();
});

test("tools/call returns an isError result for an unknown tool", async () => {
  const handlers = makeToolHandlers({ bridge: fakeBridge(() => ({})) });
  const { client, server, received } = await startServer(handlers);

  await send(client, 2, "tools/call", { name: "nope", arguments: {} });
  await flush();

  const response = received.find((m) => m.id === 2);
  assert.equal(response.result.isError, true);
  assert.match(response.result.content[0].text, /Unknown tool/);
  await server.close();
});
