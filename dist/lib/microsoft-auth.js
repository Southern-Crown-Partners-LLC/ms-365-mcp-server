import logger from "../logger.js";
import { getCloudEndpoints } from "../cloud-config.js";
function buildWwwAuthenticate(req, error, description) {
  const protocol = req.secure ? "https" : "http";
  const origin = `${protocol}://${req.get("host")}`;
  const resourceMetadata = `${origin}/.well-known/oauth-protected-resource`;
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
const microsoftBearerTokenAuthMiddleware = (req, res, next) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    res.status(401).set(
      "WWW-Authenticate",
      buildWwwAuthenticate(req, "invalid_token", "Missing or malformed Authorization header")
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
      buildWwwAuthenticate(req, "invalid_token", "The access token has expired")
    ).json({ error: "invalid_token", error_description: "The access token has expired" });
    return;
  }
  req.microsoftAuth = { accessToken };
  next();
};
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
    const error = await response.text();
    logger.error(`Failed to exchange code for token: ${error}`);
    throw new Error(`Failed to exchange code for token: ${error}`);
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
    const error = await response.text();
    logger.error(`Failed to refresh token: ${error}`);
    throw new Error(`Failed to refresh token: ${error}`);
  }
  return response.json();
}
export {
  exchangeCodeForToken,
  microsoftBearerTokenAuthMiddleware,
  refreshAccessToken
};
