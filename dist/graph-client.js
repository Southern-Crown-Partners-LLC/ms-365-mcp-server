import logger from "./logger.js";
import { encode as toonEncode } from "@toon-format/toon";
import { getCloudEndpoints } from "./cloud-config.js";
import { getRequestTokens } from "./request-context.js";
function isBinaryContentType(contentType) {
  if (!contentType) return false;
  const lower = contentType.toLowerCase().split(";")[0].trim();
  if (!lower) return false;
  if (lower.startsWith("image/") || lower.startsWith("video/") || lower.startsWith("audio/") || lower.startsWith("font/")) {
    return true;
  }
  if (lower === "application/octet-stream" || lower === "application/pdf") {
    return true;
  }
  if (lower.startsWith("application/zip") || lower.startsWith("application/x-zip")) {
    return true;
  }
  if (lower.startsWith("application/vnd.") || lower.startsWith("application/x-")) {
    if (lower.endsWith("+json") || lower.endsWith("+xml") || lower.endsWith("+text")) {
      return false;
    }
    return true;
  }
  return false;
}
class GraphClient {
  constructor(authManager, secrets, outputFormat = "json") {
    this.outputFormat = "json";
    this.authManager = authManager;
    this.secrets = secrets;
    this.outputFormat = outputFormat;
  }
  async makeRequest(endpoint, options = {}) {
    const contextTokens = getRequestTokens();
    const accessToken = options.accessToken ?? contextTokens?.accessToken ?? await this.authManager.getToken();
    if (!accessToken) {
      throw new Error("No access token available");
    }
    try {
      const response = await this.performRequest(endpoint, accessToken, options);
      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes("scope") || errorText.includes("permission")) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }
      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }
      const contentTypeHeader = response.headers?.get?.("content-type") || "";
      const isBinaryResponse = isBinaryContentType(contentTypeHeader);
      let result;
      if (isBinaryResponse) {
        const buffer = Buffer.from(await response.arrayBuffer());
        result = {
          message: "OK!",
          contentType: contentTypeHeader,
          encoding: "base64",
          contentLength: buffer.byteLength,
          contentBytes: buffer.toString("base64")
        };
      } else {
        const text = await response.text();
        if (text === "") {
          result = { message: "OK!" };
        } else {
          try {
            result = JSON.parse(text);
          } catch {
            result = { message: "OK!", rawResponse: text };
          }
        }
      }
      if (options.includeHeaders) {
        const etag = response.headers.get("ETag") || response.headers.get("etag");
        if (result && typeof result === "object" && !Array.isArray(result)) {
          return {
            ...result,
            _etag: etag || "no-etag-found"
          };
        }
      }
      return result;
    } catch (error) {
      logger.error("Microsoft Graph API request failed:", error);
      throw error;
    }
  }
  async performRequest(endpoint, accessToken, options) {
    const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
    const url = `${cloudEndpoints.graphApi}/v1.0${endpoint}`;
    logger.info(`[GRAPH CLIENT] Final URL being sent to Microsoft: ${url}`);
    const headers = {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...options.headers
    };
    return fetch(url, {
      method: options.method || "GET",
      headers,
      // Node's fetch accepts Buffer/Uint8Array; TS BodyInit doesn't.
      body: options.body
    });
  }
  serializeData(data, outputFormat, pretty = false) {
    if (outputFormat === "toon") {
      try {
        return toonEncode(data);
      } catch (error) {
        logger.warn(`Failed to encode as TOON, falling back to JSON: ${error}`);
        return JSON.stringify(data, null, pretty ? 2 : void 0);
      }
    }
    return JSON.stringify(data, null, pretty ? 2 : void 0);
  }
  async graphRequest(endpoint, options = {}) {
    try {
      logger.info(`Calling ${endpoint} with options: ${JSON.stringify(options)}`);
      const result = await this.makeRequest(endpoint, options);
      return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      return {
        content: [{ type: "text", text: JSON.stringify({ error: error.message }) }],
        isError: true
      };
    }
  }
  formatJsonResponse(data, rawResponse = false, excludeResponse = false) {
    if (excludeResponse) {
      return {
        content: [{ type: "text", text: this.serializeData({ success: true }, this.outputFormat) }]
      };
    }
    if (data && typeof data === "object" && "_headers" in data) {
      const responseData = data;
      const meta = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }
      if (responseData._headers) {
        meta.headers = responseData._headers;
      }
      if (rawResponse) {
        return {
          content: [
            { type: "text", text: this.serializeData(responseData.data, this.outputFormat) }
          ],
          _meta: meta
        };
      }
      if (responseData.data === null || responseData.data === void 0) {
        return {
          content: [
            { type: "text", text: this.serializeData({ success: true }, this.outputFormat) }
          ],
          _meta: meta
        };
      }
      const removeODataProps2 = (obj) => {
        if (typeof obj === "object" && obj !== null) {
          Object.keys(obj).forEach((key) => {
            if (key.startsWith("@odata.") && key !== "@odata.nextLink") {
              delete obj[key];
            } else if (typeof obj[key] === "object") {
              removeODataProps2(obj[key]);
            }
          });
        }
      };
      removeODataProps2(responseData.data);
      return {
        content: [
          { type: "text", text: this.serializeData(responseData.data, this.outputFormat, true) }
        ],
        _meta: meta
      };
    }
    if (rawResponse) {
      return {
        content: [{ type: "text", text: this.serializeData(data, this.outputFormat) }]
      };
    }
    if (data === null || data === void 0) {
      return {
        content: [{ type: "text", text: this.serializeData({ success: true }, this.outputFormat) }]
      };
    }
    const removeODataProps = (obj) => {
      if (typeof obj === "object" && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (key.startsWith("@odata.") && key !== "@odata.nextLink") {
            delete obj[key];
          } else if (typeof obj[key] === "object") {
            removeODataProps(obj[key]);
          }
        });
      }
    };
    removeODataProps(data);
    return {
      content: [{ type: "text", text: this.serializeData(data, this.outputFormat, true) }]
    };
  }
}
var graph_client_default = GraphClient;
export {
  graph_client_default as default,
  isBinaryContentType
};
