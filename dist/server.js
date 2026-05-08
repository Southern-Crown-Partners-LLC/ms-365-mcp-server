import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { mcpAuthRouter } from "@modelcontextprotocol/sdk/server/auth/router.js";
import express from "express";
import logger, { enableConsoleLogging } from "./logger.js";
import { registerAuthTools } from "./auth-tools.js";
import { registerGraphTools, registerDiscoveryTools } from "./graph-tools.js";
import { buildMcpServerInstructions } from "./mcp-instructions.js";
import GraphClient from "./graph-client.js";
import { buildScopesFromEndpoints } from "./auth.js";
import { MicrosoftOAuthProvider } from "./oauth-provider.js";
import {
  exchangeCodeForToken,
  microsoftBearerTokenAuthMiddleware,
  refreshAccessToken
} from "./lib/microsoft-auth.js";
import { getSecrets } from "./secrets.js";
import { getCloudEndpoints } from "./cloud-config.js";
import { requestContext } from "./request-context.js";
import crypto from "node:crypto";
import OboClient from "./obo-client.js";
function parseHttpOption(httpOption) {
  if (typeof httpOption === "boolean") {
    return { host: void 0, port: 3e3 };
  }
  const httpString = httpOption.trim();
  if (httpString.includes(":")) {
    const [hostPart, portPart] = httpString.split(":");
    const host = hostPart || void 0;
    const port2 = parseInt(portPart) || 3e3;
    return { host, port: port2 };
  }
  const port = parseInt(httpString) || 3e3;
  return { host: void 0, port };
}
class MicrosoftGraphServer {
  constructor(authManager, options = {}) {
    this.version = "0.0.0";
    this.multiAccount = false;
    this.accountNames = [];
    // Two-leg PKCE: stores client's code_challenge and server's code_verifier, keyed by OAuth state
    this.pkceStore = /* @__PURE__ */ new Map();
    this.authManager = authManager;
    this.options = options;
    this.graphClient = null;
    this.server = null;
    this.secrets = null;
    this.oboClient = null;
  }
  createMcpServer() {
    const server = new McpServer(
      {
        name: "Microsoft365MCP",
        version: this.version
      },
      {
        instructions: buildMcpServerInstructions({
          discovery: Boolean(this.options.discovery),
          orgMode: Boolean(this.options.orgMode),
          readOnly: Boolean(this.options.readOnly),
          multiAccount: this.multiAccount
        })
      }
    );
    const shouldRegisterAuthTools = !this.options.http || this.options.enableAuthTools;
    if (shouldRegisterAuthTools) {
      registerAuthTools(server, this.authManager);
    }
    if (this.options.discovery) {
      registerDiscoveryTools(
        server,
        this.graphClient,
        this.options.readOnly,
        this.options.orgMode,
        this.authManager,
        this.multiAccount,
        this.accountNames,
        this.options.enabledTools
      );
    } else {
      registerGraphTools(
        server,
        this.graphClient,
        this.options.readOnly,
        this.options.enabledTools,
        this.options.orgMode,
        this.authManager,
        this.multiAccount,
        this.accountNames
      );
    }
    return server;
  }
  async initialize(version) {
    this.secrets = await getSecrets();
    this.version = version;
    try {
      this.multiAccount = await this.authManager.isMultiAccount();
      if (this.multiAccount) {
        const accounts = await this.authManager.listAccounts();
        this.accountNames = accounts.map((a) => a.username).filter((u) => !!u);
        logger.info(
          `Multi-account mode detected (${this.accountNames.length} accounts): "account" parameter will be injected into all tool schemas`
        );
      }
    } catch (err) {
      logger.warn(`Failed to detect multi-account mode: ${err.message}`);
    }
    if (this.options.obo) {
      if (!this.options.http) {
        throw new Error("--obo requires --http (On-Behalf-Of flow only works in HTTP mode).");
      }
      if (!this.secrets.clientSecret) {
        throw new Error(
          "--obo requires MS365_MCP_CLIENT_SECRET to be set (confidential client required for On-Behalf-Of flow)."
        );
      }
      this.oboClient = new OboClient(this.secrets);
      logger.info("On-Behalf-Of (OBO) flow enabled");
    }
    const outputFormat = this.options.toon ? "toon" : "json";
    this.graphClient = new GraphClient(this.authManager, this.secrets, outputFormat);
    if (!this.options.http) {
      this.server = this.createMcpServer();
    }
    if (this.options.discovery) {
      logger.info("Discovery mode enabled (experimental) - registering discovery tool only");
    }
  }
  async start() {
    if (this.options.v) {
      enableConsoleLogging();
    }
    logger.info("Microsoft 365 MCP Server starting...");
    logger.info("Secrets Check:", {
      CLIENT_ID: this.secrets?.clientId ? `${this.secrets.clientId.substring(0, 8)}...` : "NOT SET",
      CLIENT_SECRET: this.secrets?.clientSecret ? "SET" : "NOT SET",
      TENANT_ID: this.secrets?.tenantId || "NOT SET",
      NODE_ENV: process.env.NODE_ENV || "NOT SET"
    });
    if (this.options.readOnly) {
      logger.info("Server running in READ-ONLY mode. Write operations are disabled.");
    }
    if (this.options.http) {
      const { host, port } = parseHttpOption(this.options.http);
      const app = express();
      app.set("trust proxy", true);
      app.use(express.json());
      app.use(express.urlencoded({ extended: true }));
      const corsOrigin = process.env.MS365_MCP_CORS_ORIGIN || "http://localhost:3000";
      app.use((req, res, next) => {
        res.header("Access-Control-Allow-Origin", corsOrigin);
        res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
        res.header(
          "Access-Control-Allow-Headers",
          "Origin, X-Requested-With, Content-Type, Accept, Authorization, mcp-protocol-version"
        );
        if (req.method === "OPTIONS") {
          res.sendStatus(200);
          return;
        }
        next();
      });
      const oauthProvider = new MicrosoftOAuthProvider(this.authManager, this.secrets);
      const publicUrlRaw = this.options.publicUrl || process.env.MS365_MCP_PUBLIC_URL || this.options.baseUrl || process.env.MS365_MCP_BASE_URL || null;
      const publicBase = publicUrlRaw ? new URL(publicUrlRaw).href.replace(/\/$/, "") : null;
      app.get("/.well-known/oauth-authorization-server", async (req, res) => {
        const protocol = req.secure ? "https" : "http";
        const requestOrigin = `${protocol}://${req.get("host")}`;
        const browserBase = publicBase ?? requestOrigin;
        const scopes = buildScopesFromEndpoints(this.options.orgMode, this.options.enabledTools);
        const metadata = {
          issuer: browserBase,
          authorization_endpoint: `${browserBase}/authorize`,
          token_endpoint: `${requestOrigin}/token`,
          response_types_supported: ["code"],
          response_modes_supported: ["query"],
          grant_types_supported: ["authorization_code", "refresh_token"],
          token_endpoint_auth_methods_supported: ["none"],
          code_challenge_methods_supported: ["S256"],
          scopes_supported: scopes
        };
        if (this.options.enableDynamicRegistration) {
          metadata.registration_endpoint = `${requestOrigin}/register`;
        }
        res.json(metadata);
      });
      app.get("/.well-known/oauth-protected-resource", async (req, res) => {
        const protocol = req.secure ? "https" : "http";
        const requestOrigin = `${protocol}://${req.get("host")}`;
        const browserBase = publicBase ?? requestOrigin;
        const scopes = this.options.obo ? [`api://${this.secrets.clientId}/access_as_user`] : buildScopesFromEndpoints(this.options.orgMode, this.options.enabledTools);
        res.json({
          resource: `${requestOrigin}/mcp`,
          authorization_servers: [browserBase],
          scopes_supported: scopes,
          bearer_methods_supported: ["header"],
          resource_documentation: browserBase
        });
      });
      if (this.options.enableDynamicRegistration) {
        app.post("/register", async (req, res) => {
          const body = req.body;
          logger.info("Client registration request", { body });
          const clientId = `mcp-client-${Date.now()}`;
          res.status(201).json({
            client_id: clientId,
            client_id_issued_at: Math.floor(Date.now() / 1e3),
            redirect_uris: body.redirect_uris || [],
            grant_types: body.grant_types || ["authorization_code", "refresh_token"],
            response_types: body.response_types || ["code"],
            token_endpoint_auth_method: body.token_endpoint_auth_method || "none",
            client_name: body.client_name || "MCP Client"
          });
        });
      }
      app.get("/authorize", async (req, res) => {
        const url = new URL(req.url, `${req.protocol}://${req.get("host")}`);
        const tenantId = this.secrets?.tenantId || "common";
        const clientId = this.secrets.clientId;
        const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
        const microsoftAuthUrl = new URL(
          `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/authorize`
        );
        const clientCodeChallenge = url.searchParams.get("code_challenge");
        const clientCodeChallengeMethod = url.searchParams.get("code_challenge_method");
        const state = url.searchParams.get("state");
        const allowedParams = [
          "response_type",
          "redirect_uri",
          "scope",
          "state",
          "response_mode",
          "prompt",
          "login_hint",
          "domain_hint"
        ];
        allowedParams.forEach((param) => {
          const value = url.searchParams.get(param);
          if (value) {
            microsoftAuthUrl.searchParams.set(param, value);
          }
        });
        if (clientCodeChallenge && state) {
          const serverCodeVerifier = crypto.randomBytes(32).toString("base64url");
          const serverCodeChallenge = crypto.createHash("sha256").update(serverCodeVerifier).digest("base64url");
          const now = Date.now();
          const maxAge = 10 * 60 * 1e3;
          const maxEntries = 1e3;
          for (const [key, value] of this.pkceStore) {
            if (now - value.createdAt > maxAge) {
              this.pkceStore.delete(key);
            }
          }
          if (this.pkceStore.size >= maxEntries) {
            logger.warn(
              `PKCE store at capacity (${maxEntries} entries) \u2014 rejecting new authorization request`
            );
            res.status(503).json({
              error: "server_busy",
              error_description: "Too many pending authorization requests. Try again later."
            });
            return;
          }
          this.pkceStore.set(state, {
            clientCodeChallenge,
            clientCodeChallengeMethod: clientCodeChallengeMethod || "S256",
            serverCodeVerifier,
            createdAt: Date.now()
          });
          microsoftAuthUrl.searchParams.set("code_challenge", serverCodeChallenge);
          microsoftAuthUrl.searchParams.set("code_challenge_method", "S256");
          logger.info("Two-leg PKCE: stored client challenge, generated server challenge", {
            state: state.substring(0, 8) + "..."
          });
        } else if (clientCodeChallenge) {
          microsoftAuthUrl.searchParams.set("code_challenge", clientCodeChallenge);
          if (clientCodeChallengeMethod) {
            microsoftAuthUrl.searchParams.set("code_challenge_method", clientCodeChallengeMethod);
          }
        }
        microsoftAuthUrl.searchParams.set("client_id", clientId);
        if (!microsoftAuthUrl.searchParams.get("scope")) {
          microsoftAuthUrl.searchParams.set(
            "scope",
            "User.Read Files.Read Mail.Read offline_access"
          );
        } else {
          const scopeValue = microsoftAuthUrl.searchParams.get("scope");
          const scopeList = scopeValue.split(/\s+/).filter(Boolean);
          if (!scopeList.includes("offline_access")) {
            scopeList.push("offline_access");
            microsoftAuthUrl.searchParams.set("scope", scopeList.join(" "));
          }
        }
        res.redirect(microsoftAuthUrl.toString());
      });
      app.post("/token", async (req, res) => {
        try {
          logger.info("Token endpoint called", {
            method: req.method,
            url: req.url,
            contentType: req.get("Content-Type"),
            grant_type: req.body?.grant_type
          });
          const body = req.body;
          if (!body) {
            logger.error("Token endpoint: Request body is undefined");
            res.status(400).json({
              error: "invalid_request",
              error_description: "Request body is required"
            });
            return;
          }
          if (!body.grant_type) {
            logger.error("Token endpoint: grant_type is missing", { body });
            res.status(400).json({
              error: "invalid_request",
              error_description: "grant_type parameter is required"
            });
            return;
          }
          if (body.grant_type === "authorization_code") {
            const tenantId = this.secrets?.tenantId || "common";
            const clientId = this.secrets.clientId;
            const clientSecret = this.secrets?.clientSecret;
            logger.info("Token endpoint: authorization_code exchange", {
              redirect_uri: body.redirect_uri,
              has_code: !!body.code,
              has_code_verifier: !!body.code_verifier,
              clientId,
              tenantId,
              hasClientSecret: !!clientSecret
            });
            let serverCodeVerifier;
            if (body.code_verifier) {
              const clientVerifier = body.code_verifier;
              const clientChallengeComputed = crypto.createHash("sha256").update(clientVerifier).digest("base64url");
              for (const [state, pkceData] of this.pkceStore) {
                if (pkceData.clientCodeChallenge === clientChallengeComputed) {
                  serverCodeVerifier = pkceData.serverCodeVerifier;
                  this.pkceStore.delete(state);
                  logger.info("Two-leg PKCE: matched client verifier, using server verifier", {
                    state: state.substring(0, 8) + "..."
                  });
                  break;
                }
              }
            }
            const result = await exchangeCodeForToken(
              body.code,
              body.redirect_uri,
              clientId,
              clientSecret,
              tenantId,
              serverCodeVerifier || body.code_verifier,
              this.secrets.cloudType
            );
            res.json(result);
          } else if (body.grant_type === "refresh_token") {
            const tenantId = this.secrets?.tenantId || "common";
            const clientId = this.secrets.clientId;
            const clientSecret = this.secrets?.clientSecret;
            if (clientSecret) {
              logger.info("Refresh endpoint: Using confidential client with client_secret");
            } else {
              logger.info("Refresh endpoint: Using public client without client_secret");
            }
            const result = await refreshAccessToken(
              body.refresh_token,
              clientId,
              clientSecret,
              tenantId,
              this.secrets.cloudType
            );
            res.json(result);
          } else {
            res.status(400).json({
              error: "unsupported_grant_type",
              error_description: `Grant type '${body.grant_type}' is not supported`
            });
          }
        } catch (error) {
          logger.error("Token endpoint error:", error);
          res.status(500).json({
            error: "server_error",
            error_description: "Internal server error during token exchange"
          });
        }
      });
      app.use(
        mcpAuthRouter({
          provider: oauthProvider,
          issuerUrl: new URL(publicBase ?? `http://localhost:${port}`)
        })
      );
      app.get(
        "/mcp",
        microsoftBearerTokenAuthMiddleware,
        async (req, res) => {
          const handler = async () => {
            const server = this.createMcpServer();
            const transport = new StreamableHTTPServerTransport({
              sessionIdGenerator: void 0
              // Stateless mode
            });
            res.on("close", () => {
              transport.close();
              server.close();
            });
            await server.connect(transport);
            await transport.handleRequest(req, res, void 0);
          };
          try {
            if (req.microsoftAuth) {
              let accessToken = req.microsoftAuth.accessToken;
              if (this.oboClient) {
                accessToken = await this.oboClient.exchangeToken(accessToken);
              }
              await requestContext.run({ accessToken }, handler);
            } else {
              await handler();
            }
          } catch (error) {
            logger.error("Error handling MCP GET request:", error);
            if (!res.headersSent) {
              res.status(500).json({
                jsonrpc: "2.0",
                error: {
                  code: -32603,
                  message: "Internal server error"
                },
                id: null
              });
            }
          }
        }
      );
      app.post(
        "/mcp",
        microsoftBearerTokenAuthMiddleware,
        async (req, res) => {
          const handler = async () => {
            const server = this.createMcpServer();
            const transport = new StreamableHTTPServerTransport({
              sessionIdGenerator: void 0
              // Stateless mode
            });
            res.on("close", () => {
              transport.close();
              server.close();
            });
            await server.connect(transport);
            await transport.handleRequest(req, res, req.body);
          };
          try {
            if (req.microsoftAuth) {
              let accessToken = req.microsoftAuth.accessToken;
              if (this.oboClient) {
                accessToken = await this.oboClient.exchangeToken(accessToken);
              }
              await requestContext.run({ accessToken }, handler);
            } else {
              await handler();
            }
          } catch (error) {
            logger.error("Error handling MCP POST request:", error);
            if (!res.headersSent) {
              res.status(500).json({
                jsonrpc: "2.0",
                error: {
                  code: -32603,
                  message: "Internal server error"
                },
                id: null
              });
            }
          }
        }
      );
      app.get("/", (req, res) => {
        res.send("Microsoft 365 MCP Server is running");
      });
      if (host) {
        app.listen(port, host, () => {
          logger.info(`Server listening on ${host}:${port}`);
          logger.info(`  - MCP endpoint: http://${host}:${port}/mcp`);
          logger.info(`  - OAuth endpoints: http://${host}:${port}/auth/*`);
          logger.info(
            `  - OAuth discovery: http://${host}:${port}/.well-known/oauth-authorization-server`
          );
        });
      } else {
        app.listen(port, () => {
          logger.info(`Server listening on all interfaces (0.0.0.0:${port})`);
          logger.info(`  - MCP endpoint: http://localhost:${port}/mcp`);
          logger.info(`  - OAuth endpoints: http://localhost:${port}/auth/*`);
          logger.info(
            `  - OAuth discovery: http://localhost:${port}/.well-known/oauth-authorization-server`
          );
        });
      }
    } else {
      const transport = new StdioServerTransport();
      await this.server.connect(transport);
      logger.info("Server connected to stdio transport");
    }
  }
}
var server_default = MicrosoftGraphServer;
export {
  server_default as default
};
