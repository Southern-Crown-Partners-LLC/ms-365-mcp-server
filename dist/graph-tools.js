import logger from "./logger.js";
import { api } from "./generated/client.js";
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
    description: 'Download binary content from Microsoft Graph and return it as base64. Single tool for any binary read: drive file content, mail attachment, profile photo, Teams hosted content, meeting recording. Returns { contentType, encoding: "base64", contentLength, contentBytes }.',
    readOnlyHint: true,
    openWorldHint: true,
    buildSchema: (ctx) => {
      const schema = {
        target: z.string().describe(
          'Relative Microsoft Graph path starting with "/". Common paths: /drives/{drive-id}/items/{driveItem-id}/content (drive file content); /me/messages/{message-id}/attachments/{attachment-id}/$value (mail attachment, list-mail-attachments returns the IDs); /me/photo/$value or /users/{user-id}/photo/$value (profile photo); /chats/{chat-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams chat hosted content, list-chat-message-hosted-contents returns the IDs); /teams/{team-id}/channels/{channel-id}/messages/{chatMessage-id}/hostedContents/{chatMessageHostedContent-id}/$value (Teams channel hosted content). For meeting recordings (often large), use get-meeting-recording-content which returns a URL for out-of-band download by the client.'
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
        let accountAccessToken;
        if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
          accountAccessToken = await authManager.getTokenForAccount(accountParam);
        }
        return await graphClient.graphRequest(target, { accessToken: accountAccessToken });
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
  try {
    let accountAccessToken;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      const accountParam = params.account;
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
      }
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
    if (fetchAllPages && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse["@odata.nextLink"];
        let pageCount = 1;
        const maxPages = 100;
        const maxItems = 1e4;
        while (nextLink && pageCount < maxPages && allItems.length < maxItems) {
          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);
          const url = new URL(nextLink);
          const nextPath = url.pathname.replace("/v1.0", "") + url.search;
          const nextOptions = { ...options };
          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse["@odata.nextLink"];
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
    return {
      content,
      _meta: response._meta,
      isError: response.isError
    };
  } catch (error) {
    logger.error(`Error in tool ${tool.alias}: ${error.message}`);
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
function registerGraphTools(server, graphClient, readOnly = false, enabledToolsPattern, orgMode = false, authManager, multiAccount = false, accountNames = []) {
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
  for (const tool of api.endpoints) {
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
    if (tool.method.toUpperCase() === "GET" && tool.path.includes("/")) {
      paramSchema["fetchAllPages"] = z.boolean().describe(
        "Follow @odata.nextLink and merge up to 100 pages into one response. Can return enormous payloads\u2014only when the user explicitly needs a full export. Prefer a small $top first, then paginate or narrow with $filter/$search."
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
    if (paramSchema["top"] !== void 0 || paramSchema["$top"] !== void 0) {
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
    let toolDescription = tool.description || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`;
    if (endpointConfig?.llmTip) {
      toolDescription += `

\u{1F4A1} TIP: ${endpointConfig.llmTip}`;
    }
    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: tool.method.toUpperCase() === "GET",
          destructiveHint: ["POST", "PATCH", "DELETE"].includes(tool.method.toUpperCase()),
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
function buildToolsRegistry(readOnly, orgMode, enabledToolsRegex) {
  const toolsMap = /* @__PURE__ */ new Map();
  for (const tool of api.endpoints) {
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
    const descTokens = tokenize(tool.description).slice(0, DESC_CAP_TOKENS);
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
    const descTokens = tokenize(utility.description).slice(0, DESC_CAP_TOKENS);
    const tokens = [...nt, ...nt, ...nt, ...nt, ...nt, ...pathTokens, ...pathTokens, ...descTokens];
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
function registerDiscoveryTools(server, graphClient, readOnly = false, orgMode = false, authManager, multiAccount = false, accountNames = [], enabledTools) {
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
  const toolsRegistry = buildToolsRegistry(readOnly, orgMode, enabledToolsRegex);
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
        description: tool.description || `${tool.method.toUpperCase()} ${tool.path}`,
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
        const schema = describeToolSchema(entry.tool, entry.config?.llmTip);
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
