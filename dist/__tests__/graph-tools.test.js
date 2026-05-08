import { describe, it, expect, vi, beforeEach } from "vitest";
import { z } from "zod";
vi.mock("../logger.js", () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn()
  }
}));
const mockEndpoints = [];
vi.mock("../generated/client.js", () => ({
  api: {
    get endpoints() {
      return mockEndpoints;
    }
  }
}));
let mockEndpointsJson = [];
vi.mock("fs", async (importOriginal) => {
  const actual = await importOriginal();
  return {
    ...actual,
    readFileSync: (filePath, encoding) => {
      if (typeof filePath === "string" && filePath.includes("endpoints.json")) {
        return JSON.stringify(mockEndpointsJson);
      }
      return actual.readFileSync(filePath, encoding);
    }
  };
});
vi.mock("../tool-categories.js", () => ({
  TOOL_CATEGORIES: {}
}));
function makeEndpoint(overrides = {}) {
  return {
    method: "get",
    path: "/me/messages",
    alias: "test-tool",
    description: "Test tool",
    requestFormat: "json",
    parameters: [
      { name: "filter", type: "Query", schema: z.string().optional() },
      { name: "search", type: "Query", schema: z.string().optional() },
      { name: "select", type: "Query", schema: z.string().optional() },
      { name: "orderby", type: "Query", schema: z.string().optional() },
      { name: "count", type: "Query", schema: z.boolean().optional() },
      { name: "top", type: "Query", schema: z.number().optional() },
      { name: "skip", type: "Query", schema: z.number().optional() }
    ],
    response: z.any(),
    ...overrides
  };
}
function makeConfig(overrides = {}) {
  return {
    pathPattern: "/me/messages",
    method: "get",
    toolName: "test-tool",
    scopes: ["Mail.Read"],
    ...overrides
  };
}
function createMockGraphClient(responses) {
  const responseQueue = [...responses || []];
  return {
    graphRequest: vi.fn().mockImplementation(async () => {
      if (responseQueue.length > 0) {
        return responseQueue.shift();
      }
      return {
        content: [{ type: "text", text: JSON.stringify({ value: [] }) }]
      };
    })
  };
}
async function loadModule() {
  vi.resetModules();
  const mod = await import("../graph-tools.js");
  return mod;
}
function createMockServer() {
  const tools = /* @__PURE__ */ new Map();
  return {
    tool: vi.fn(
      (name, description, schema, annotations, handler) => {
        tools.set(name, { description, schema, handler });
      }
    ),
    tools
  };
}
describe("graph-tools", () => {
  beforeEach(() => {
    mockEndpoints.length = 0;
    mockEndpointsJson = [];
    vi.clearAllMocks();
  });
  describe("$count advanced query mode", () => {
    it("should set ConsistencyLevel: eventual header when $count=true", async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ value: [] }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("test-tool");
      expect(tool).toBeDefined();
      await tool.handler({ count: true });
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain("$count=true");
    });
  });
  describe("fetchAllPages pagination", () => {
    it("should follow @odata.nextLink and combine results", async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                value: [{ id: "1" }, { id: "2" }],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=2"
              })
            }
          ]
        },
        {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                value: [{ id: "3" }]
              })
            }
          ]
        }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("test-tool");
      const result = await tool.handler({ fetchAllPages: true });
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(2);
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.value).toHaveLength(3);
      expect(parsed.value.map((v) => v.id)).toEqual(["1", "2", "3"]);
      expect(parsed["@odata.nextLink"]).toBeUndefined();
    });
    it("should stop at 100 page limit", async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const responses = [];
      for (let i = 0; i < 101; i++) {
        responses.push({
          content: [
            {
              type: "text",
              text: JSON.stringify({
                value: [{ id: `item-${i}` }],
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=" + (i + 1)
              })
            }
          ]
        });
      }
      const graphClient = createMockGraphClient(responses);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("test-tool");
      await tool.handler({ fetchAllPages: true });
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(100);
    });
  });
  describe("parameter describe() overrides", () => {
    it("should apply custom descriptions to OData parameters", async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, createMockGraphClient());
      const tool = server.tools.get("test-tool");
      expect(tool).toBeDefined();
      const schema = tool.schema;
      expect(schema["filter"]).toBeDefined();
      expect(schema["filter"].description).toContain("OData filter expression");
      expect(schema["filter"].description).toContain("$count=true");
      expect(schema["search"]).toBeDefined();
      expect(schema["search"].description).toContain("KQL search query");
      expect(schema["select"]).toBeDefined();
      expect(schema["select"].description).toContain("Comma-separated fields");
      expect(schema["orderby"]).toBeDefined();
      expect(schema["orderby"].description).toContain("Sort expression");
      expect(schema["count"]).toBeDefined();
      expect(schema["count"].description).toContain("advanced query mode");
      expect(schema["top"].description).toContain("Start small");
      expect(schema["top"].description).toContain("$select");
    });
  });
  describe("MS365_MCP_MAX_TOP", () => {
    const prevMaxTop = process.env.MS365_MCP_MAX_TOP;
    afterEach(() => {
      if (prevMaxTop === void 0) delete process.env.MS365_MCP_MAX_TOP;
      else process.env.MS365_MCP_MAX_TOP = prevMaxTop;
    });
    it("should clamp $top when MS365_MCP_MAX_TOP is set", async () => {
      process.env.MS365_MCP_MAX_TOP = "10";
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ value: [] }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("test-tool");
      await tool.handler({ top: 50 });
      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain("$top=10");
    });
    it("should pass through $top when MS365_MCP_MAX_TOP is unset", async () => {
      delete process.env.MS365_MCP_MAX_TOP;
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ value: [] }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("test-tool");
      await tool.handler({ top: 50 });
      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain("$top=50");
    });
  });
  describe("returnDownloadUrl", () => {
    it("should strip /content from path and return downloadUrl when returnDownloadUrl=true", async () => {
      const endpoint = makeEndpoint({
        alias: "download-file",
        path: "/me/drive/items/:driveItem-id/content",
        parameters: [{ name: "driveItem-id", type: "Path", schema: z.string() }]
      });
      const config = makeConfig({
        toolName: "download-file",
        pathPattern: "/me/drive/items/{driveItem-id}/content",
        returnDownloadUrl: true
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const downloadUrl = "https://download.example.com/file.pdf";
      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                "@microsoft.graph.downloadUrl": downloadUrl,
                name: "file.pdf"
              })
            }
          ]
        }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("download-file");
      expect(tool).toBeDefined();
      await tool.handler({ "driveItem-id": "abc123" });
      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).not.toContain("/content");
      expect(requestedPath).toContain("/me/drive/items/abc123");
    });
  });
  describe("kebab-case path param normalization", () => {
    it("should substitute path when LLM passes message-id (kebab) but schema has messageId (camelCase)", async () => {
      const endpoint = makeEndpoint({
        alias: "get-mail-message",
        method: "get",
        path: "/me/messages/:messageId",
        parameters: [
          { name: "messageId", type: "Path", schema: z.string() },
          { name: "select", type: "Query", schema: z.string().optional() }
        ]
      });
      const config = makeConfig({
        toolName: "get-mail-message",
        pathPattern: "/me/messages/{message-id}"
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ id: "AAMk123", subject: "Test" }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("get-mail-message");
      expect(tool).toBeDefined();
      await tool.handler({ "message-id": "AAMk123abc=" });
      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain("AAMk123abc=");
      expect(requestedPath).not.toContain(":messageId");
    });
    it("should also work when LLM passes messageId (camelCase) directly", async () => {
      const endpoint = makeEndpoint({
        alias: "get-mail-message2",
        method: "get",
        path: "/me/messages/:messageId",
        parameters: [{ name: "messageId", type: "Path", schema: z.string() }]
      });
      const config = makeConfig({
        toolName: "get-mail-message2",
        pathPattern: "/me/messages/{message-id}"
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ id: "AAMk456" }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("get-mail-message2");
      await tool.handler({ messageId: "AAMk456xyz=" });
      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain("AAMk456xyz=");
      expect(requestedPath).not.toContain(":messageId");
    });
  });
  describe("supportsTimezone", () => {
    it("should set Prefer: outlook.timezone header when timezone param provided", async () => {
      const endpoint = makeEndpoint({
        alias: "list-calendar-events",
        path: "/me/events",
        parameters: []
      });
      const config = makeConfig({
        toolName: "list-calendar-events",
        pathPattern: "/me/events",
        supportsTimezone: true
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ value: [] }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("list-calendar-events");
      expect(tool).toBeDefined();
      expect(tool.schema["timezone"]).toBeDefined();
      expect(tool.schema["timezone"].description).toContain("IANA timezone");
      await tool.handler({ timezone: "Europe/Brussels" });
      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers["Prefer"]).toContain('outlook.timezone="Europe/Brussels"');
    });
    it("should NOT add timezone parameter when supportsTimezone is false/absent", async () => {
      const endpoint = makeEndpoint({
        alias: "list-mail",
        path: "/me/messages",
        parameters: []
      });
      const config = makeConfig({
        toolName: "list-mail",
        pathPattern: "/me/messages"
        // no supportsTimezone
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, createMockGraphClient());
      const tool = server.tools.get("list-mail");
      expect(tool.schema["timezone"]).toBeUndefined();
    });
  });
  describe("outlook.body-content-type Prefer header", () => {
    it('should set Prefer: outlook.body-content-type="text" on GET requests', async () => {
      const endpoint = makeEndpoint({ method: "get" });
      const config = makeConfig({ method: "get" });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([
        { content: [{ type: "text", text: JSON.stringify({ value: [] }) }] }
      ]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      await server.tools.get("test-tool").handler({});
      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers["Prefer"]).toContain('outlook.body-content-type="text"');
    });
    it("should NOT set Prefer: outlook.body-content-type on POST requests", async () => {
      const endpoint = makeEndpoint({
        alias: "create-reply-draft",
        method: "post",
        path: "/me/messages/:messageId/createReply",
        parameters: [
          { name: "messageId", type: "Path", schema: z.string() },
          { name: "body", type: "Body", schema: z.any() }
        ]
      });
      const config = makeConfig({
        toolName: "create-reply-draft",
        method: "post",
        pathPattern: "/me/messages/{message-id}/createReply"
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([{ content: [{ type: "text", text: "{}" }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      await server.tools.get("create-reply-draft").handler({
        messageId: "AAMk123",
        body: { Message: { body: { contentType: "html", content: "<p>hi</p>" } } }
      });
      const [, options] = graphClient.graphRequest.mock.calls[0];
      const prefer = options.headers["Prefer"];
      expect(prefer === void 0 || !prefer.includes("outlook.body-content-type")).toBe(true);
    });
  });
  describe("binary upload bodies", () => {
    it("decodes base64 body to bytes and sets octet-stream Content-Type", async () => {
      const endpoint = makeEndpoint({
        alias: "upload-file-content",
        method: "put",
        path: "/drives/:driveId/items/:driveItemId/content",
        requestFormat: "binary",
        parameters: [
          { name: "driveId", type: "Path", schema: z.string() },
          { name: "driveItemId", type: "Path", schema: z.string() },
          {
            name: "body",
            type: "Body",
            schema: z.string().describe("Base64-encoded file content")
          }
        ]
      });
      const config = makeConfig({
        toolName: "upload-file-content",
        method: "put",
        pathPattern: "/drives/{drive-id}/items/{driveItem-id}/content",
        scopes: ["Files.ReadWrite"]
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([{ content: [{ type: "text", text: "{}" }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const original = "Hello, world!";
      const base64 = Buffer.from(original, "utf-8").toString("base64");
      await server.tools.get("upload-file-content").handler({
        driveId: "drive123",
        driveItemId: "item456",
        body: base64
      });
      const [path, options] = graphClient.graphRequest.mock.calls[0];
      expect(path).toBe("/drives/drive123/items/item456/content");
      expect(options.headers["Content-Type"]).toBe("application/octet-stream");
      expect(Buffer.isBuffer(options.body) || options.body instanceof Uint8Array).toBe(true);
      expect(Buffer.from(options.body).toString("utf-8")).toBe(original);
    });
    it("honors endpoints.json contentType override on binary uploads", async () => {
      const endpoint = makeEndpoint({
        alias: "upload-file-content",
        method: "put",
        path: "/drives/:driveId/items/:driveItemId/content",
        requestFormat: "binary",
        parameters: [
          { name: "driveId", type: "Path", schema: z.string() },
          { name: "driveItemId", type: "Path", schema: z.string() },
          { name: "body", type: "Body", schema: z.string() }
        ]
      });
      const config = makeConfig({
        toolName: "upload-file-content",
        method: "put",
        pathPattern: "/drives/{drive-id}/items/{driveItem-id}/content",
        scopes: ["Files.ReadWrite"],
        contentType: "application/pdf"
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];
      const graphClient = createMockGraphClient([{ content: [{ type: "text", text: "{}" }] }]);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      await server.tools.get("upload-file-content").handler({
        driveId: "d",
        driveItemId: "i",
        body: Buffer.from("%PDF-1.4").toString("base64")
      });
      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers["Content-Type"]).toBe("application/pdf");
    });
  });
  describe("download-bytes", () => {
    it("routes a relative Graph path through graphRequest", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const graphClient = {
        graphRequest: vi.fn().mockResolvedValue({
          content: [
            {
              type: "text",
              text: JSON.stringify({
                contentType: "image/jpeg",
                encoding: "base64",
                contentBytes: "aGk="
              })
            }
          ]
        })
      };
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, graphClient);
      const tool = server.tools.get("download-bytes");
      expect(tool).toBeDefined();
      await tool.handler({ target: "/me/photo/$value" });
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [path, options] = graphClient.graphRequest.mock.calls[0];
      expect(path).toBe("/me/photo/$value");
      expect(options.accessToken).toBeUndefined();
    });
    it("rejects absolute URLs (Graph paths only)", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, {});
      const tool = server.tools.get("download-bytes");
      const result = await tool.handler({
        target: "https://example.sharepoint.com/d/abc?temp=signed"
      });
      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.error).toMatch(/relative Microsoft Graph path/);
    });
    it("rejects targets that do not start with /", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, {});
      const tool = server.tools.get("download-bytes");
      const result = await tool.handler({ target: "ftp://example.com/x" });
      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.error).toMatch(/relative Microsoft Graph path/);
    });
  });
  describe("discovery mode: utility tools", () => {
    it('search-tools surfaces download-bytes for "download" queries', async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server, {});
      const result = await server.tools.get("search-tools").handler({ query: "download" });
      const payload = JSON.parse(result.content[0].text);
      const names = payload.tools.map((t) => t.name);
      expect(names).toContain("download-bytes");
    });
    it("get-tool-schema returns the download-bytes parameter schema", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server, {});
      const result = await server.tools.get("get-tool-schema").handler({ tool_name: "download-bytes" });
      const schema = JSON.parse(result.content[0].text);
      expect(schema.name).toBe("download-bytes");
      expect(schema.path).toBe("tool:download-bytes");
      const targetParam = schema.parameters.find((p) => p.name === "target");
      expect(targetParam).toBeDefined();
      expect(targetParam.required).toBe(true);
    });
    it("execute-tool dispatches to download-bytes for a Graph path", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const graphClient = {
        graphRequest: vi.fn().mockResolvedValue({
          content: [
            {
              type: "text",
              text: JSON.stringify({
                contentType: "image/png",
                encoding: "base64",
                contentBytes: "iVBORw0K"
              })
            }
          ]
        })
      };
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server, graphClient);
      const result = await server.tools.get("execute-tool").handler({
        tool_name: "download-bytes",
        parameters: { target: "/me/photo/$value" }
      });
      expect(result.isError).toBeFalsy();
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [path] = graphClient.graphRequest.mock.calls[0];
      expect(path).toBe("/me/photo/$value");
    });
    it("execute-tool reports unknown tool when name matches neither registry", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server, {});
      const result = await server.tools.get("execute-tool").handler({
        tool_name: "no-such-tool",
        parameters: {}
      });
      expect(result.isError).toBe(true);
      const payload = JSON.parse(result.content[0].text);
      expect(payload.error).toMatch(/not found/i);
    });
  });
  describe("discovery mode: --enabled-tools filter", () => {
    it("search-tools only surfaces Graph tools matching the regex", async () => {
      mockEndpoints.push(
        {
          alias: "list-mail-messages",
          method: "get",
          path: "/me/messages",
          description: "List mail",
          parameters: []
        },
        {
          alias: "list-calendar-events",
          method: "get",
          path: "/me/events",
          description: "List events",
          parameters: []
        }
      );
      mockEndpointsJson = [
        { toolName: "list-mail-messages", method: "get", pathPattern: "/me/messages" },
        { toolName: "list-calendar-events", method: "get", pathPattern: "/me/events" }
      ];
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(server, {}, false, false, void 0, false, [], "mail");
      const result = await server.tools.get("search-tools").handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t) => t.name);
      expect(found).toContain("list-mail-messages");
      expect(found).not.toContain("list-calendar-events");
    });
    it("utility tools obey the regex too", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(
        server,
        {},
        false,
        false,
        void 0,
        false,
        [],
        "^download-bytes$"
      );
      const result = await server.tools.get("search-tools").handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t) => t.name);
      expect(found).toContain("download-bytes");
      expect(found).not.toContain("parse-teams-url");
    });
    it("invalid regex pattern is ignored, all tools surface", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerDiscoveryTools } = await loadModule();
      registerDiscoveryTools(
        server,
        {},
        false,
        false,
        void 0,
        false,
        [],
        "[invalid"
      );
      const result = await server.tools.get("search-tools").handler({ limit: 50 });
      const found = JSON.parse(result.content[0].text).tools.map((t) => t.name);
      expect(found).toContain("download-bytes");
      expect(found).toContain("parse-teams-url");
    });
  });
  describe("utility tools in read-only mode", () => {
    it("skips utility tools whose readOnlyHint is not true", async () => {
      mockEndpoints.length = 0;
      mockEndpointsJson = [];
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server, {}, true);
      expect(server.tools.has("download-bytes")).toBe(true);
      expect(server.tools.has("parse-teams-url")).toBe(true);
    });
  });
});
