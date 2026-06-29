import { randomUUID } from "crypto";
import logger from "./logger.js";
import { auditLog, getUserIdentityForAudit } from "./audit-log.js";
import {
  getEndpointScopeGroups,
  getMissingAllowedScopesForGroups,
  parseAllowedScopes
} from "./auth.js";
import { api } from "./generated/client.js";
import { api as betaApi } from "./generated/client-beta.js";
const allEndpoints = [...api.endpoints, ...betaApi.endpoints];
import { z } from "zod";
import { readFileSync } from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { TOOL_CATEGORIES } from "./tool-categories.js";
import { getRequestTokens } from "./request-context.js";
import { parseTeamsUrl } from "./lib/teams-url-parser.js";
import { buildBM25Index, scoreQuery, tokenize } from "./lib/bm25.js";
import { describeToolSchema, describeUtilityToolSchema } from "./lib/tool-schema.js";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, "endpoints.json"), "utf8")
);
const TOP_UNSUPPORTED_DELTA_TOOLS = /* @__PURE__ */ new Set([
  "list-calendar-events-delta",
  "list-calendar-view-delta"
]);
function withApiVersionPrefix(description, config) {
  return config?.apiVersion === "beta" ? `[beta] ${description}` : description;
}
function maxTopFromEnv() {
  const raw = process.env.MS365_MCP_MAX_TOP;
  if (raw === void 0 || raw === "") return void 0;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) {
    logger.warn(
      `Ignoring invalid MS365_MCP_MAX_TOP=${JSON.stringify(raw)} (use a positive integer)`
    );
    return void 0;
  }
  return n;
}
function clampTopQueryParam(queryParams) {
  const cap = maxTopFromEnv();
  if (cap === void 0 || queryParams["$top"] === void 0) return;
  const requested = Number.parseInt(queryParams["$top"], 10);
  if (!Number.isFinite(requested) || requested <= cap) return;
  logger.info(`Clamping $top from ${requested} to ${cap} (MS365_MCP_MAX_TOP)`);
  queryParams["$top"] = String(cap);
}
const DEFAULT_MAX_PAGES = 100;
const DEFAULT_MAX_ITEMS = 1e4;
function positiveIntFromEnv(name, defaultValue) {
  const raw = process.env[name];
  if (raw === void 0 || raw === "") return defaultValue;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) {
    logger.warn(`Ignoring invalid ${name}=${JSON.stringify(raw)} (use a positive integer)`);
    return defaultValue;
  }
  return n;
}
function paginationAllowed() {
  const raw = process.env.MS365_MCP_ALLOW_PAGINATION;
  if (raw === void 0 || raw === "") return true;
  return !/^(0|false|no)$/i.test(raw.trim());
}
function formatDisabledToolsForLog(disabledTools) {
  const shown = disabledTools.slice(0, 20).map((tool) => `${tool.toolName} (missing: ${tool.missingScopes.join(", ")})`);
  const suffix = disabledTools.length > shown.length ? `, ... +${disabledTools.length - shown.length} more` : "";
  return `${shown.join("; ")}${suffix}`;
}
async function checkAccountParamInBearerMode(accountParam, authManager) {
  if (!accountParam || !authManager) return null;
  const contextToken = getRequestTokens()?.accessToken;
  if (!contextToken && !authManager.isOAuthModeEnabled()) return null;
  const bearerToken = contextToken ?? await authManager.getToken().catch(() => null) ?? void 0;
  const bearerIdentity = getUserIdentityForAudit(bearerToken);
  if (bearerIdentity && bearerIdentity.toLowerCase() === accountParam.toLowerCase()) return null;
  return `The 'account' parameter is not supported in HTTP/OAuth mode: every request uses the identity of the connecting client's bearer token` + (bearerIdentity ? ` ('${bearerIdentity}')` : "") + `, so account switching is not possible. To act as '${accountParam}', reconnect the MCP client authenticated as that account, or run the server in stdio mode (or HTTP with --trust-proxy-auth) where cached accounts are available.`;
}
const UTILITY_TOOLS = [
  {
    name: "parse-teams-url",
    method: "POST",
    path: "tool:parse-teams-url",
    description: "Converts any Teams meeting URL format (short /meet/, full /meetup-join/, or recap ?threadId=) into a standard joinWebUrl. Use this before list-online-meetings when the user provides a recap or short URL.",
    readOnlyHint: true,
    openWorldHint: false,
    buildSchema: () => ({
      url: z.string().describe("Teams meeting URL in any format")
    }),
    execute: async (params) => {
      const url = params.url;
      if (typeof url !== "string") {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: "url is required." }) }],
          isError: true
        };
      }
      try {
        const joinWebUrl = parseTeamsUrl(url);
        return { content: [{ type: "text", text: joinWebUrl }] };
      } catch (error) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: error.message }) }],
          isError: true
        };
      }
    }
  },
  {
    name: "download-bytes",
    method: "GET",
    path: "tool:download-bytes",
    description: 'Download binary content from Microsoft Graph and return it as base64. Single tool for any binary read: drive file content, mail attachment, profile photo, Teams hosted content, meeting recording. Returns { contentType, encoding: "base64", contentLength, contentBytes }. For large drive/SharePoint file content, prefer get-download-url, which returns a pre-authenticated URL to stream bytes out-of-band instead of base64 through the agent context.',
    readOnlyHint: true,
    openWorldHint: true,
    buildSchema: (ctx) => {
      const schema = {
        target: z.string().describe(
          'Relative Microsoft Graph path starting with "/". Common paths: /drives/{drive-id}/items/{driveItem-id}/content (drive file content); /me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment, list-mail-attachments returns the IDs); /me/photo/$value or /users/{user-id}/photo/$value (profile photo); /chats/{chat-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams chat hosted content, list-chat-message-hosted-contents returns the IDs); /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams channel hosted content). For meeting recordings, use get-meeting-recording-content where available; Microsoft Graph returns authenticated recording bytes, not a pre-authenticated download URL.'
        )
      };
      if (ctx.multiAccount) {
        schema["account"] = z.string().optional().describe(
          "Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts)."
        );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager }) => {
      const target = params.target;
      const accountParam = params.account;
      if (typeof target !== "string" || target.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ error: "target is required and must be a non-empty string." })
            }
          ],
          isError: true
        };
      }
      if (!target.startsWith("/")) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: 'target must be a relative Microsoft Graph path starting with "/", e.g. /me/photo/$value or /drives/{drive-id}/items/{driveItem-id}/content. Absolute URLs are not accepted; if you have an @microsoft.graph.downloadUrl, use the equivalent /content or /$value path instead (Graph 302-redirects to the same bytes).'
              })
            }
          ],
          isError: true
        };
      }
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: "text", text: JSON.stringify({ error: accountModeError }) }],
            isError: true
          };
        }
        let accountAccessToken;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        return await graphClient.graphRequest(target, {
          accessToken: accountAccessToken,
          rawResponse: true
        });
      } catch (error) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: error.message }) }],
          isError: true
        };
      }
    }
  },
  {
    name: "get-download-url",
    method: "GET",
    path: "tool:get-download-url",
    searchKeywords: "download file download drive file download onedrive file sharepoint file download large drive file large sharepoint file large file out-of-band download pre-authenticated url",
    description: "Resolve a short-lived, pre-authenticated download URL for Microsoft Graph binary content that exposes one (drive/SharePoint file content). The returned URL streams the bytes with NO Authorization header, so the client can fetch it straight to disk (e.g. curl) without round-tripping base64 through the agent context. Prefer this over download-bytes for any file above a few KB or any bulk download. Returns { downloadUrl, name?, size?, contentType? }. NOTE: mail file attachments (/messages/{id}/attachments/{id}/$value) and meeting recordings do NOT expose a pre-authenticated URL \u2014 Graph offers no such link for them; use download-bytes for small ones.",
    readOnlyHint: true,
    openWorldHint: true,
    buildSchema: (ctx) => {
      const schema = {
        target: z.string().describe(
          'Relative Microsoft Graph path starting with "/". Either a driveItem content path or the item path itself, e.g. /drives/{drive-id}/items/{driveItem-id}/content, /me/drive/items/{driveItem-id}/content, or /sites/{site-id}/drive/items/{driveItem-id}. A trailing /content is optional and is stripped automatically for drive items. Mail attachment $value paths and meeting recordings are not supported (Graph exposes no pre-authenticated URL for them).'
        )
      };
      if (ctx.multiAccount) {
        schema["account"] = z.string().optional().describe(
          "Account to use when multiple Microsoft accounts are configured. Required when multiple accounts exist (see list-accounts)."
        );
      }
      return schema;
    },
    execute: async (params, { graphClient, authManager }) => {
      const target = params.target;
      const accountParam = params.account;
      if (typeof target !== "string" || target.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ error: "target is required and must be a non-empty string." })
            }
          ],
          isError: true
        };
      }
      if (!target.startsWith("/")) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: 'target must be a relative Microsoft Graph path starting with "/", e.g. /drives/{drive-id}/items/{driveItem-id}/content.'
              })
            }
          ],
          isError: true
        };
      }
      const queryIdx = target.indexOf("?");
      if (queryIdx >= 0) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: "target must not include query parameters. Pass the drive item /content path or item metadata path without $select, $expand, or other query options."
              })
            }
          ],
          isError: true
        };
      }
      const pathPart = target.replace(/\/+$/, "");
      if (/^(\/me|\/users\/[^/]+)\/messages\/[^/]+\/attachments\//.test(pathPart) || /^(\/me|\/users\/[^/]+)\/events\/[^/]+\/attachments\//.test(pathPart) || /^\/groups\/[^/]+\/messages\/[^/]+\/attachments\//.test(pathPart) || /^\/groups\/[^/]+\/events\/[^/]+\/attachments\//.test(pathPart)) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: "Mail and calendar event attachments do not expose a pre-authenticated download URL. Use download-bytes for small attachments."
              })
            }
          ],
          isError: true
        };
      }
      if (/^(\/me|\/users\/[^/]+)\/onlineMeetings\/[^/]+\/recordings\/[^/]+(?:\/content)?$/.test(
        pathPart
      ) || /^\/communications\/calls\/[^/]+\/recordings\/[^/]+(?:\/content)?$/.test(pathPart)) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: "Meeting recordings do not expose a pre-authenticated download URL. Use download-bytes for small recordings or get-meeting-recording-content where available."
              })
            }
          ],
          isError: true
        };
      }
      if (pathPart.endsWith("/$value")) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: "$value byte endpoints do not expose a pre-authenticated download URL. Use download-bytes to read these bytes."
              })
            }
          ],
          isError: true
        };
      }
      const isDriveItemById = /^\/drives\/[^/]+\/items\/[^/]+(?:\/content)?$/.test(pathPart) || /^\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/drive\/items\/[^/]+(?:\/content)?$/.test(
        pathPart
      ) || /^\/(?:groups\/[^/]+|sites\/[^/]+)\/drives\/[^/]+\/items\/[^/]+(?:\/content)?$/.test(
        pathPart
      );
      const isDriveItemByPath = /^\/drives\/[^/]+\/root:\/.+:(?:\/content)?$/.test(pathPart) || /^\/(?:me|users\/[^/]+|groups\/[^/]+|sites\/[^/]+)\/drive\/root:\/.+:(?:\/content)?$/.test(
        pathPart
      ) || /^\/(?:groups\/[^/]+|sites\/[^/]+)\/drives\/[^/]+\/root:\/.+:(?:\/content)?$/.test(
        pathPart
      );
      if (!isDriveItemById && !isDriveItemByPath) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                error: "target must identify a driveItem in OneDrive or SharePoint. Use a drive item metadata path or /content path, such as /drives/{drive-id}/items/{driveItem-id}/content, /me/drive/items/{driveItem-id}, /sites/{site-id}/drive/items/{driveItem-id}, or /me/drive/root:/path/file.ext:/content. Other Graph byte resources must use download-bytes."
              })
            }
          ],
          isError: true
        };
      }
      const isDriveContentEndpoint = /\/items\/[^/]+\/content$/.test(pathPart) || pathPart.endsWith(":/content");
      const itemPath = isDriveContentEndpoint ? pathPart.slice(0, -"/content".length) : pathPart;
      try {
        const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
        if (accountModeError) {
          return {
            content: [{ type: "text", text: JSON.stringify({ error: accountModeError }) }],
            isError: true
          };
        }
        let accountAccessToken;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        const response = await graphClient.graphRequest(itemPath, {
          accessToken: accountAccessToken
        });
        if (response?.isError) {
          return response;
        }
        const text = response?.content?.[0]?.text;
        let item;
        if (typeof text === "string") {
          try {
            item = JSON.parse(text);
          } catch {
            item = void 0;
          }
        }
        const downloadUrl = item?.["@microsoft.graph.downloadUrl"];
        if (!downloadUrl) {
          return {
            content: [
              {
                type: "text",
                text: JSON.stringify({
                  error: "No pre-authenticated download URL is available for this resource. It may not be a drive item, or it exposes bytes only via download-bytes."
                })
              }
            ],
            isError: true
          };
        }
        const file = item?.file;
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({
                downloadUrl,
                name: item?.name,
                size: item?.size,
                contentType: file?.mimeType
              })
            }
          ]
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: JSON.stringify({ error: error.message }) }],
          isError: true
        };
      }
    }
  }
];
function registerUtilityToolWithMcp(server, utility, ctx) {
  server.tool(
    utility.name,
    utility.description,
    utility.buildSchema(ctx),
    {
      title: utility.name,
      readOnlyHint: utility.readOnlyHint ?? true,
      openWorldHint: utility.openWorldHint ?? true
    },
    async (params) => utility.execute(params, ctx)
  );
}
async function executeGraphTool(tool, config, graphClient, params, authManager) {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
  const requestId = randomUUID();
  const startTime = Date.now();
  const upn = getUserIdentityForAudit(getRequestTokens()?.accessToken);
  const httpMethod = tool.method.toUpperCase();
  try {
    const accountParam = params.account;
    const accountModeError = await checkAccountParamInBearerMode(accountParam, authManager);
    if (accountModeError) {
      return {
        content: [{ type: "text", text: JSON.stringify({ error: accountModeError }) }],
        isError: true
      };
    }
    let accountAccessToken;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      try {
        accountAccessToken = await authManager.getTokenForAccount(accountParam);
      } catch (err) {
        return {
          content: [
            {
              type: "text",
              text: JSON.stringify({ error: err.message })
            }
          ],
          isError: true
        };
      }
    }
    const parameterDefinitions = tool.parameters || [];
    let path2 = tool.path;
    const queryParams = {};
    const headers = {};
    let body = null;
    for (const [paramName, paramValue] of Object.entries(params)) {
      if ([
        "account",
        "fetchAllPages",
        "includeHeaders",
        "excludeResponse",
        "timezone",
        "expandExtendedProperties"
      ].includes(paramName)) {
        continue;
      }
      const odataParams = [
        "filter",
        "select",
        "expand",
        "orderby",
        "skip",
        "top",
        "count",
        "search",
        "format"
      ];
      const normalizedParamName = paramName.startsWith("$") ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      const camelCaseParamName = paramName.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
      const paramDef = parameterDefinitions.find(
        (p) => p.name === paramName || p.name === camelCaseParamName || isOdataParam && p.name === normalizedParamName
      );
      if (paramDef) {
        switch (paramDef.type) {
          case "Path": {
            const shouldSkipEncoding = config?.skipEncoding?.includes(paramName) ?? false;
            const encodedValue = shouldSkipEncoding ? paramValue : encodeURIComponent(paramValue).replace(/%3D/g, "=");
            path2 = path2.replace(`{${paramName}}`, encodedValue).replace(`:${paramName}`, encodedValue).replace(`{${camelCaseParamName}}`, encodedValue).replace(`:${camelCaseParamName}`, encodedValue);
            break;
          }
          case "Query":
            if (paramValue !== "" && paramValue != null) {
              queryParams[fixedParamName] = `${paramValue}`;
            }
            break;
          case "Body":
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;
          case "Header":
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === "body") {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      } else if (path2.includes(`:${paramName}`) || path2.includes(`{${paramName}}`) || path2.includes(`:${camelCaseParamName}`) || path2.includes(`{${camelCaseParamName}}`)) {
        const encodedValue = encodeURIComponent(paramValue).replace(/%3D/g, "=");
        path2 = path2.replace(`{${paramName}}`, encodedValue).replace(`:${paramName}`, encodedValue).replace(`{${camelCaseParamName}}`, encodedValue).replace(`:${camelCaseParamName}`, encodedValue);
        logger.info(`Path param fallback: replaced :${camelCaseParamName} with encoded value`);
      } else if (isOdataParam) {
        queryParams[fixedParamName] = `${paramValue}`;
        logger.info(`OData param fallback: forwarded ${fixedParamName}=${paramValue}`);
      }
    }
    if (TOP_UNSUPPORTED_DELTA_TOOLS.has(tool.alias)) {
      delete queryParams["$top"];
    }
    clampTopQueryParam(queryParams);
    const preferValues = [];
    if (config?.supportsTimezone && params.timezone) {
      preferValues.push(`outlook.timezone="${params.timezone}"`);
      logger.info(`Setting timezone preference: outlook.timezone="${params.timezone}"`);
    }
    const bodyFormat = process.env.MS365_MCP_BODY_FORMAT || "text";
    if (bodyFormat !== "html" && tool.method.toUpperCase() === "GET") {
      preferValues.push(`outlook.body-content-type="${bodyFormat}"`);
    }
    if (preferValues.length > 0) {
      headers["Prefer"] = preferValues.join(", ");
    }
    if (config?.supportsExpandExtendedProperties && params.expandExtendedProperties === true) {
      const expandValue = "singleValueExtendedProperties";
      if (queryParams["$expand"]) {
        queryParams["$expand"] += `,${expandValue}`;
      } else {
        queryParams["$expand"] = expandValue;
      }
      logger.info(`Adding $expand=${expandValue} for extended properties`);
    }
    if (config?.contentType) {
      headers["Content-Type"] = config.contentType;
      logger.info(`Setting custom Content-Type: ${config.contentType}`);
    }
    if (config?.acceptType) {
      headers["Accept"] = config.acceptType;
      logger.info(`Setting custom Accept: ${config.acceptType}`);
    }
    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams).map(([key, value]) => `${key}=${encodeURIComponent(value).replace(/%2C/gi, ",")}`).join("&");
      path2 = `${path2}${path2.includes("?") ? "&" : "?"}${queryString}`;
    }
    const options = {
      method: tool.method.toUpperCase(),
      headers
    };
    if (config?.apiVersion) {
      options.apiVersion = config.apiVersion;
    }
    if (options.method !== "GET" && body) {
      if (tool.requestFormat === "binary" && typeof body === "string") {
        options.body = Buffer.from(body, "base64");
        if (!config?.contentType) {
          headers["Content-Type"] = "application/octet-stream";
        }
      } else if (config?.contentType === "text/html") {
        if (typeof body === "string") {
          options.body = body;
        } else if (typeof body === "object" && "content" in body) {
          options.body = body.content;
        } else {
          options.body = String(body);
        }
      } else {
        options.body = typeof body === "string" ? body : JSON.stringify(body);
      }
    }
    const isProbablyMediaContent = tool.errors?.some((error) => error.description === "Retrieved media content") || path2.endsWith("/content");
    if (config?.returnDownloadUrl && path2.endsWith("/content")) {
      path2 = path2.replace(/\/content$/, "");
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }
    if (accountAccessToken) {
      options.accessToken = accountAccessToken;
    }
    const { accessToken: _redacted, ...safeOptions } = options;
    logger.info(
      `Making graph request to ${path2} with options: ${JSON.stringify(safeOptions)}${_redacted ? " [accessToken=REDACTED]" : ""}`
    );
    let response = await graphClient.graphRequest(path2, options);
    const fetchAllPages = params.fetchAllPages === true;
    const paginationEnabled = paginationAllowed();
    if (fetchAllPages && !paginationEnabled) {
      logger.info(
        "fetchAllPages requested but MS365_MCP_ALLOW_PAGINATION is disabled; returning first page only"
      );
    }
    if (fetchAllPages && paginationEnabled && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse["@odata.nextLink"];
        let pageCount = 1;
        const maxPages = positiveIntFromEnv("MS365_MCP_MAX_PAGES", DEFAULT_MAX_PAGES);
        const maxItems = positiveIntFromEnv("MS365_MCP_MAX_ITEMS", DEFAULT_MAX_ITEMS);
        let deltaLink = combinedResponse["@odata.deltaLink"];
        while (nextLink && pageCount < maxPages && allItems.length < maxItems) {
          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);
          const url = new URL(nextLink);
          const nextPath = url.pathname.replace(/^\/(v1\.0|beta)/, "") + url.search;
          const nextOptions = { ...options };
          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse["@odata.nextLink"];
            if (nextJsonResponse["@odata.deltaLink"]) {
              deltaLink = nextJsonResponse["@odata.deltaLink"];
            }
            pageCount++;
          } else {
            break;
          }
        }
        if (pageCount >= maxPages) {
          logger.warn(`Reached maximum page limit (${maxPages}) for pagination`);
        }
        if (allItems.length >= maxItems) {
          logger.warn(
            `Reached maximum item limit (${maxItems}) for pagination \u2014 truncated at ${allItems.length} items`
          );
        }
        combinedResponse.value = allItems;
        if (combinedResponse["@odata.count"]) {
          combinedResponse["@odata.count"] = allItems.length;
        }
        delete combinedResponse["@odata.nextLink"];
        if (deltaLink) {
          combinedResponse["@odata.deltaLink"] = deltaLink;
        }
        response.content[0].text = JSON.stringify(combinedResponse);
        logger.info(
          `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
        );
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }
    }
    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);
      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse["@odata.nextLink"]) {
          logger.info(`Response has pagination nextLink: ${jsonResponse["@odata.nextLink"]}`);
        }
      } catch {
      }
    }
    const content = response.content.map((item) => ({
      type: "text",
      text: item.text
    }));
    auditLog({
      event: "tool.call",
      request_id: requestId,
      user_principal_name: upn,
      tool: tool.alias,
      http_method: httpMethod,
      status: response.isError ? "error" : "success",
      duration_ms: Date.now() - startTime
    });
    return {
      content,
      _meta: response._meta,
      isError: response.isError
    };
  } catch (error) {
    const err = error;
    logger.error(`Error in tool ${tool.alias}: ${error.message}`);
    auditLog({
      event: "tool.call",
      request_id: requestId,
      user_principal_name: upn,
      tool: tool.alias,
      http_method: httpMethod,
      status: "error",
      duration_ms: Date.now() - startTime,
      error_type: err?.name || "Error",
      error_code: err?.status ?? err?.code
    });
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${error.message}`
          })
        }
      ],
      isError: true
    };
  }
}
function registerGraphTools(server, graphClient, readOnly = false, enabledToolsPattern, orgMode = false, authManager, multiAccount = false, accountNames = [], allowedScopesValue) {
  let enabledToolsRegex;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, "i");
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }
  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;
  const allowedScopes = parseAllowedScopes(allowedScopesValue);
  const disabledByAllowedScopes = [];
  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }
    const method = tool.method.toUpperCase();
    if (readOnly && method !== "GET") {
      if (!(method === "POST" && endpointConfig?.readOnly)) {
        logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
        skippedCount++;
        continue;
      }
    }
    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }
    const missingScopes = allowedScopes !== void 0 && !endpointConfig ? ["endpoint scope metadata"] : getMissingAllowedScopesForGroups(
      getEndpointScopeGroups(endpointConfig, orgMode),
      allowedScopes
    );
    if (missingScopes.length > 0) {
      disabledByAllowedScopes.push({ toolName: tool.alias, missingScopes });
      skippedCount++;
      continue;
    }
    const paramSchema = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        paramSchema[param.name] = param.schema || z.any();
      }
    }
    const pathParamMatches = tool.path.matchAll(/:([a-zA-Z]+)/g);
    for (const match of pathParamMatches) {
      const pathParamName = match[1];
      if (!(pathParamName in paramSchema)) {
        paramSchema[pathParamName] = z.string().describe(`Path parameter: ${pathParamName}`);
      }
    }
    if (tool.method.toUpperCase() === "GET" && tool.path.includes("/") && paginationAllowed()) {
      const maxPages = positiveIntFromEnv("MS365_MCP_MAX_PAGES", DEFAULT_MAX_PAGES);
      paramSchema["fetchAllPages"] = z.boolean().describe(
        `Follow @odata.nextLink and merge up to ${maxPages} pages into one response. Can return enormous payloads\u2014only when the user explicitly needs a full export. Prefer a small $top first, then paginate or narrow with $filter/$search.`
      ).optional();
    }
    if (paramSchema["filter"] !== void 0 || paramSchema["$filter"] !== void 0) {
      const key = paramSchema["$filter"] !== void 0 ? "$filter" : "filter";
      paramSchema[key] = z.string().describe(
        "OData filter expression. Add $count=true for advanced filters (flag/flagStatus, contains()). Cannot combine with $search."
      ).optional();
    }
    if (paramSchema["search"] !== void 0 || paramSchema["$search"] !== void 0) {
      const key = paramSchema["$search"] !== void 0 ? "$search" : "search";
      paramSchema[key] = z.string().describe("KQL search query \u2014 wrap value in double quotes. Cannot combine with $filter.").optional();
    }
    if (paramSchema["select"] !== void 0 || paramSchema["$select"] !== void 0) {
      const key = paramSchema["$select"] !== void 0 ? "$select" : "select";
      paramSchema[key] = z.string().describe("Comma-separated fields to return, e.g. id,subject,from,receivedDateTime").optional();
    }
    if (paramSchema["orderby"] !== void 0 || paramSchema["$orderby"] !== void 0) {
      const key = paramSchema["$orderby"] !== void 0 ? "$orderby" : "orderby";
      paramSchema[key] = z.string().describe("Sort expression, e.g. receivedDateTime desc").optional();
    }
    if (TOP_UNSUPPORTED_DELTA_TOOLS.has(tool.alias)) {
      delete paramSchema["top"];
      delete paramSchema["$top"];
    } else if (paramSchema["top"] !== void 0 || paramSchema["$top"] !== void 0) {
      const key = paramSchema["$top"] !== void 0 ? "$top" : "top";
      paramSchema[key] = z.number().describe(
        "Page size (Graph $top). Start small (e.g. 5\u201315) so responses fit the model context; raise only if needed. Use $select to return fewer fields per item. For more rows, use @odata.nextLink from the response instead of a very large $top."
      ).optional();
    }
    if (paramSchema["skip"] !== void 0 || paramSchema["$skip"] !== void 0) {
      const key = paramSchema["$skip"] !== void 0 ? "$skip" : "skip";
      paramSchema[key] = z.number().describe("Items to skip for pagination. Not supported with $search.").optional();
    }
    if (paramSchema["count"] !== void 0 || paramSchema["$count"] !== void 0) {
      const countKey = paramSchema["$count"] !== void 0 ? "$count" : "count";
      paramSchema[countKey] = z.boolean().describe(
        "Set true to enable advanced query mode (ConsistencyLevel: eventual). Required for complex $filter on flag/flagStatus or contains()."
      ).optional();
    }
    if (multiAccount) {
      const accountHint = accountNames.length > 0 ? `Known accounts: ${accountNames.join(", ")}. ` : "";
      paramSchema["account"] = z.string().describe(
        `${accountHint}Microsoft account email to use for this request. Required when multiple accounts are configured. Use the list-accounts tool to discover all currently available accounts.`
      ).optional();
    }
    paramSchema["includeHeaders"] = z.boolean().describe("Include response headers (including ETag) in the response metadata").optional();
    paramSchema["excludeResponse"] = z.boolean().describe("Exclude the full response body and only return success or failure indication").optional();
    if (endpointConfig?.supportsTimezone) {
      paramSchema["timezone"] = z.string().describe(
        'IANA timezone name (e.g., "America/New_York", "Europe/London", "Asia/Tokyo") for calendar event times. If not specified, times are returned in UTC.'
      ).optional();
    }
    if (endpointConfig?.supportsExpandExtendedProperties) {
      paramSchema["expandExtendedProperties"] = z.boolean().describe(
        "When true, expands singleValueExtendedProperties on each event. Use this to retrieve custom extended properties (e.g., sync metadata) stored on calendar events."
      ).optional();
    }
    let toolDescription = withApiVersionPrefix(
      (endpointConfig?.descriptionOverride ?? tool.description) || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`,
      endpointConfig
    );
    if (endpointConfig?.llmTip) {
      toolDescription += `

\u{1F4A1} TIP: ${endpointConfig.llmTip}`;
    }
    const isReadOnlyTool = tool.method.toUpperCase() === "GET" || endpointConfig?.readOnly === true;
    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: isReadOnlyTool,
          destructiveHint: !isReadOnlyTool && ["POST", "PATCH", "DELETE"].includes(tool.method.toUpperCase()),
          openWorldHint: true
          // All tools call Microsoft Graph API
        },
        async (params) => executeGraphTool(tool, endpointConfig, graphClient, params, authManager)
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${error.message}`);
      failedCount++;
    }
  }
  if (multiAccount) {
    logger.info('Multi-account mode: "account" parameter injected into all tool schemas');
  }
  if (disabledByAllowedScopes.length > 0) {
    logger.info(
      `Allowed scopes disabled ${disabledByAllowedScopes.length} Graph tools: ${formatDisabledToolsForLog(disabledByAllowedScopes)}`
    );
  }
  const utilityCtx = {
    graphClient,
    authManager,
    multiAccount,
    accountNames
  };
  for (const utility of UTILITY_TOOLS) {
    if (readOnly && !utility.readOnlyHint) continue;
    if (enabledToolsRegex && !enabledToolsRegex.test(utility.name)) continue;
    try {
      registerUtilityToolWithMcp(server, utility, utilityCtx);
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${utility.name}: ${error.message}`);
      failedCount++;
    }
  }
  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}
function buildToolsRegistry(readOnly, orgMode, enabledToolsRegex, allowedScopesValue, disabledByAllowedScopes = []) {
  const toolsMap = /* @__PURE__ */ new Map();
  const allowedScopes = parseAllowedScopes(allowedScopesValue);
  for (const tool of allEndpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }
    const method = tool.method.toUpperCase();
    if (readOnly && method !== "GET") {
      if (!(method === "POST" && endpointConfig?.readOnly)) {
        continue;
      }
    }
    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      continue;
    }
    const missingScopes = allowedScopes !== void 0 && !endpointConfig ? ["endpoint scope metadata"] : getMissingAllowedScopesForGroups(
      getEndpointScopeGroups(endpointConfig, orgMode),
      allowedScopes
    );
    if (missingScopes.length > 0) {
      disabledByAllowedScopes.push({ toolName: tool.alias, missingScopes });
      continue;
    }
    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }
  return toolsMap;
}
function buildDiscoverySearchIndex(toolsRegistry, utilityTools = []) {
  const TIP_EXCERPT_TOKENS = 12;
  const DESC_CAP_TOKENS = 40;
  const docs = [];
  const nameTokens = /* @__PURE__ */ new Map();
  for (const [name, { tool, config }] of toolsRegistry) {
    const nt = tokenize(name);
    nameTokens.set(name, new Set(nt));
    const pathTokens = tokenize(tool.path);
    const descTokens = tokenize(config?.descriptionOverride ?? tool.description).slice(
      0,
      DESC_CAP_TOKENS
    );
    const tipTokens = tokenize(config?.llmTip).slice(0, TIP_EXCERPT_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...tipTokens,
      ...descTokens
    ];
    docs.push({ id: name, tokens });
  }
  for (const utility of utilityTools) {
    const nt = tokenize(utility.name);
    nameTokens.set(utility.name, new Set(nt));
    const pathTokens = tokenize(utility.path);
    const keywordTokens = tokenize(utility.searchKeywords);
    const descTokens = tokenize(utility.description).slice(0, DESC_CAP_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...keywordTokens,
      ...keywordTokens,
      ...descTokens
    ];
    docs.push({ id: utility.name, tokens });
  }
  return { bm25: buildBM25Index(docs), nameTokens };
}
function scoreDiscoveryQuery(query, index) {
  const queryTokenSet = new Set(tokenize(query));
  if (queryTokenSet.size === 0) return [];
  const ranked = scoreQuery(query, index.bm25);
  const NAME_BONUS_WEIGHT = 2;
  for (const r of ranked) {
    const nt = index.nameTokens.get(r.id);
    if (!nt || nt.size === 0) continue;
    let matchedIdf = 0;
    let matchedCount = 0;
    for (const qt of queryTokenSet) {
      if (nt.has(qt)) {
        matchedCount++;
        matchedIdf += index.bm25.idf.get(qt) ?? 0;
      }
    }
    if (matchedCount === 0) continue;
    const precision = matchedCount / nt.size;
    r.score += precision * matchedIdf * NAME_BONUS_WEIGHT;
  }
  ranked.sort((a, b) => b.score - a.score);
  return ranked;
}
function registerDiscoveryTools(server, graphClient, readOnly = false, orgMode = false, authManager, multiAccount = false, accountNames = [], enabledTools, allowedScopesValue) {
  let enabledToolsRegex;
  if (enabledTools) {
    try {
      enabledToolsRegex = new RegExp(enabledTools, "i");
      logger.info(`Discovery mode: filtering tools with pattern ${enabledTools}`);
    } catch (error) {
      logger.error(
        `Invalid --enabled-tools regex ${JSON.stringify(enabledTools)} \u2014 ignoring filter: ${error.message}`
      );
    }
  }
  const disabledByAllowedScopes = [];
  const toolsRegistry = buildToolsRegistry(
    readOnly,
    orgMode,
    enabledToolsRegex,
    allowedScopesValue,
    disabledByAllowedScopes
  );
  if (disabledByAllowedScopes.length > 0) {
    logger.info(
      `Discovery mode: allowed scopes disabled ${disabledByAllowedScopes.length} Graph tools: ${formatDisabledToolsForLog(disabledByAllowedScopes)}`
    );
  }
  const utilityTools = UTILITY_TOOLS.filter((u) => {
    if (readOnly && !u.readOnlyHint) return false;
    if (enabledToolsRegex && !enabledToolsRegex.test(u.name)) return false;
    return true;
  });
  const searchIndex = buildDiscoverySearchIndex(toolsRegistry, utilityTools);
  const totalCount = toolsRegistry.size + utilityTools.length;
  logger.info(
    `Discovery mode: ${totalCount} tools (${toolsRegistry.size} Graph + ${utilityTools.length} utility)`
  );
  const utilityCtx = {
    graphClient,
    authManager,
    multiAccount,
    accountNames
  };
  const utilityByName = new Map(utilityTools.map((u) => [u.name, u]));
  const categoryNames = Object.keys(TOOL_CATEGORIES).join(", ");
  const toResultEntry = (name) => {
    const entry = toolsRegistry.get(name);
    if (entry) {
      const { tool, config } = entry;
      return {
        name,
        method: tool.method.toUpperCase(),
        path: tool.path,
        description: withApiVersionPrefix(
          (config?.descriptionOverride ?? tool.description) || `${tool.method.toUpperCase()} ${tool.path}`,
          config
        ),
        ...config?.llmTip ? { llmTip: config.llmTip } : {}
      };
    }
    const utility = utilityByName.get(name);
    if (utility) {
      return {
        name: utility.name,
        method: utility.method,
        path: utility.path,
        description: utility.description
      };
    }
    return null;
  };
  server.tool(
    "search-tools",
    `Search through ${totalCount} tools (${toolsRegistry.size} Microsoft Graph API operations + ${utilityTools.length} server utilities like download-bytes). Ranks results by BM25 over tool name, llmTip, description, and path. After picking a tool, call get-tool-schema for parameters, then execute-tool.`,
    {
      query: z.string().describe(
        'Natural-language query. Tokenized and BM25-ranked. E.g. "send email", "download photo", "list unread messages".'
      ).optional(),
      category: z.string().describe(`Optional pre-filter by category: ${categoryNames}`).optional(),
      limit: z.number().describe("Maximum results (default: 10, max: 50)").optional()
    },
    {
      title: "search-tools",
      readOnlyHint: true,
      openWorldHint: true
    },
    async ({ query, category, limit = 10 }) => {
      const maxLimit = Math.min(Math.max(limit, 1), 50);
      const categoryDef = category ? TOOL_CATEGORIES[category] : void 0;
      const categoryFilter = (name) => !categoryDef || categoryDef.pattern.test(name);
      let orderedNames;
      if (query && query.trim().length > 0) {
        const ranked = scoreDiscoveryQuery(query, searchIndex);
        orderedNames = ranked.map((r) => r.id).filter(categoryFilter);
      } else {
        orderedNames = [...toolsRegistry.keys(), ...utilityTools.map((u) => u.name)].filter(
          categoryFilter
        );
      }
      const tools = orderedNames.slice(0, maxLimit).map(toResultEntry).filter(Boolean);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                found: tools.length,
                total: totalCount,
                tools,
                tip: "Call get-tool-schema(tool_name) to see parameters before invoking execute-tool."
              },
              null,
              2
            )
          }
        ]
      };
    }
  );
  server.tool(
    "get-tool-schema",
    "Returns the full parameter schema (name, placement, required, JSON Schema) for a tool discovered via search-tools. Call this before execute-tool so you know what parameters to pass and what enum values are valid.",
    {
      tool_name: z.string().describe('Exact tool name from search-tools (e.g. "send-mail")')
    },
    {
      title: "get-tool-schema",
      readOnlyHint: true,
      openWorldHint: false
    },
    async ({ tool_name }) => {
      const entry = toolsRegistry.get(tool_name);
      if (entry) {
        const schema = describeToolSchema(
          entry.tool,
          entry.config?.llmTip,
          entry.config?.descriptionOverride
        );
        return {
          content: [{ type: "text", text: JSON.stringify(schema, null, 2) }]
        };
      }
      const utility = utilityByName.get(tool_name);
      if (utility) {
        const schema = describeUtilityToolSchema(utility, utilityCtx);
        return {
          content: [{ type: "text", text: JSON.stringify(schema, null, 2) }]
        };
      }
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: `Tool not found: ${tool_name}`,
              tip: "Use search-tools to find available tools."
            })
          }
        ],
        isError: true
      };
    }
  );
  server.tool(
    "execute-tool",
    "Execute a Microsoft Graph API tool by name. Workflow: search-tools \u2192 get-tool-schema \u2192 execute-tool. Call get-tool-schema first for any tool you have not seen before \u2014 passing the wrong shape to parameters will fail validation or return a Graph 400. For list endpoints, prefer modest $top plus $select.",
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z.record(z.any()).describe(
        'Parameters shaped per get-tool-schema. Path/query/header params go at the top level; request bodies go under "body".'
      ).optional()
    },
    {
      title: "execute-tool",
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (toolData) {
        return executeGraphTool(
          toolData.tool,
          toolData.config,
          graphClient,
          parameters,
          authManager
        );
      }
      const utility = utilityByName.get(tool_name);
      if (utility) {
        return utility.execute(parameters, utilityCtx);
      }
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: `Tool not found: ${tool_name}`,
              tip: "Use search-tools to find available tools."
            })
          }
        ],
        isError: true
      };
    }
  );
}
export {
  UTILITY_TOOLS,
  buildDiscoverySearchIndex,
  buildToolsRegistry,
  registerDiscoveryTools,
  registerGraphTools,
  scoreDiscoveryQuery
};
