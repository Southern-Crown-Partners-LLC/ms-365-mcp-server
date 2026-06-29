import logger from "../logger.js";
import { getCloudEndpoints } from "../cloud-config.js";
function buildResourceMetadataUrl(req, publicUrl) {
  if (publicUrl) {
    const parsed = new URL(publicUrl);
    const path = parsed.pathname.replace(/\/$/, "");
    return `${parsed.origin}/.well-known/oauth-protected-resource${path}`;
  }
  const protocol = req.secure ? "https" : "http";
  const origin = `${protocol}://${req.get("host")}`;
  return `${origin}/.well-known/oauth-protected-resource`;
}
function buildWwwAuthenticate(req, error, description, publicUrl) {
  const resourceMetadata = buildResourceMetadataUrl(req, publicUrl);
  return `Bearer resource_metadata="${resourceMetadata}", error="${error}", error_description="${description}"`;
}
function isJwtExpired(token) {
  const parts = token.split(".");
  if (parts.length !== 3) return false;
  try {
    const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf-8"));
    if (typeof payload.exp !== "number") return false;
    return payload.exp * 1e3 < Date.now();
  } catch {
    return false;
  }
}
const DISCOVERY_METHODS = /* @__PURE__ */ new Set([
  "initialize",
  "notifications/initialized",
  "tools/list",
  "prompts/list",
  "resources/list",
  "ping"
]);
function isDiscoveryRequest(req) {
  if (req.method !== "POST" || !req.body) return false;
  const body = req.body;
  if (Array.isArray(body)) {
    return body.every((item) => DISCOVERY_METHODS.has(item?.method));
  }
  return DISCOVERY_METHODS.has(body?.method);
}
const microsoftBearerTokenAuthMiddleware = (opts = {}) => (req, res, next) => {
  if (opts.trustProxyAuth) {
    next();
    return;
  }
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    if (opts.allowUnauthenticatedDiscovery && isDiscoveryRequest(req)) {
      next();
      return;
    }
    res.status(401).set(
      "WWW-Authenticate",
      buildWwwAuthenticate(
        req,
        "invalid_token",
        "Missing or malformed Authorization header",
        opts.publicUrl
      )
    ).json({
      error: "invalid_token",
      error_description: "Missing or malformed Authorization header"
    });
    return;
  }
  const accessToken = authHeader.substring(7);
  if (isJwtExpired(accessToken)) {
    res.status(401).set(
      "WWW-Authenticate",
      buildWwwAuthenticate(req, "invalid_token", "The access token has expired", opts.publicUrl)
    ).json({ error: "invalid_token", error_description: "The access token has expired" });
    return;
  }
  req.microsoftAuth = { accessToken };
  next();
};
class OAuthUpstreamError extends Error {
  constructor(status, raw, body) {
    const suffix = body.error_description ? ` - ${body.error_description}` : "";
    super(`OAuth upstream error: ${body.error}${suffix}`);
    this.name = "OAuthUpstreamError";
    this.status = status;
    this.body = body;
    this.raw = raw;
  }
}
function parseUpstreamOAuthError(raw) {
  try {
    const json = JSON.parse(raw);
    if (json !== null && typeof json === "object" && typeof json.error === "string") {
      return json;
    }
  } catch {
  }
  return null;
}
function toOAuthErrorResponse(error) {
  if (error instanceof OAuthUpstreamError) {
    const body = {
      error: error.body.error
    };
    if (error.body.error_description) body.error_description = error.body.error_description;
    if (error.body.suberror) body.suberror = error.body.suberror;
    return { status: 400, body };
  }
  return {
    status: 500,
    body: {
      error: "server_error",
      error_description: "Internal server error during token exchange"
    }
  };
}
async function exchangeCodeForToken(code, redirectUri, clientId, clientSecret, tenantId = "common", codeVerifier, cloudType = "global") {
  const cloudEndpoints = getCloudEndpoints(cloudType);
  const params = new URLSearchParams({
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    client_id: clientId
  });
  if (clientSecret) {
    params.append("client_secret", clientSecret);
  }
  if (codeVerifier) {
    params.append("code_verifier", codeVerifier);
  }
  const response = await fetch(`${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: params
  });
  if (!response.ok) {
    const raw = await response.text();
    const parsed = parseUpstreamOAuthError(raw);
    if (parsed) {
      logger.warn(`Token endpoint upstream OAuth error: ${parsed.error}`, {
        status: response.status,
        error: parsed.error,
        suberror: parsed.suberror,
        error_codes: parsed.error_codes,
        correlation_id: parsed.correlation_id
      });
      throw new OAuthUpstreamError(response.status, raw, parsed);
    }
    logger.error(`Failed to exchange code for token: ${raw}`);
    throw new Error(`Failed to exchange code for token: ${raw}`);
  }
  return response.json();
}
async function refreshAccessToken(refreshToken, clientId, clientSecret, tenantId = "common", cloudType = "global") {
  const cloudEndpoints = getCloudEndpoints(cloudType);
  const params = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    client_id: clientId
  });
  if (clientSecret) {
    params.append("client_secret", clientSecret);
  }
  const response = await fetch(`${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: params
  });
  if (!response.ok) {
    const raw = await response.text();
    const parsed = parseUpstreamOAuthError(raw);
    if (parsed) {
      logger.warn(`Token endpoint upstream OAuth error: ${parsed.error}`, {
        status: response.status,
        error: parsed.error,
        suberror: parsed.suberror,
        error_codes: parsed.error_codes,
        correlation_id: parsed.correlation_id
      });
      throw new OAuthUpstreamError(response.status, raw, parsed);
    }
    logger.error(`Failed to refresh token: ${raw}`);
    throw new Error(`Failed to refresh token: ${raw}`);
  }
  return response.json();
}
export {
  OAuthUpstreamError,
  exchangeCodeForToken,
  microsoftBearerTokenAuthMiddleware,
  refreshAccessToken,
  toOAuthErrorResponse
};
