const REDACTIONS = [
  // JSON Web Tokens (header.payload.signature) — access_token, id_token, etc.
  {
    pattern: /eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/g,
    replacement: "[REDACTED_JWT]"
  },
  // Authorization: Bearer <token>
  {
    pattern: /(Bearer\s+)[A-Za-z0-9._~+/-]+=*/gi,
    replacement: "$1[REDACTED]"
  },
  // OAuth token fields in query strings or JSON bodies:
  // refresh_token=..., "access_token": "...", client_secret=...
  // The \b keeps composite keys like `statusCode` from matching via a substring.
  {
    pattern: /(["']?\b(?:refresh_token|access_token|id_token|client_secret|assertion)["']?\s*[=:]\s*["']?)[A-Za-z0-9._~+/-]+=*/gi,
    replacement: "$1[REDACTED]"
  },
  // OAuth authorization codes, query-string form only (?code=... / &code=...).
  // `code:` in JSON or prose is an error/status code (ECONNRESET, AADSTS...),
  // which must stay readable for diagnostics.
  {
    pattern: /([?&]code=)[^&\s"']+/gi,
    replacement: "$1[REDACTED]"
  },
  // Email addresses / UPNs
  {
    pattern: /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g,
    replacement: "[REDACTED_EMAIL]"
  }
];
function redactionEnabled() {
  const raw = process.env.MS365_MCP_REDACT_PII;
  return raw === "true" || raw === "1";
}
function redactSensitive(input) {
  let out = input;
  for (const { pattern, replacement } of REDACTIONS) {
    out = out.replace(pattern, replacement);
  }
  return out;
}
export {
  redactSensitive,
  redactionEnabled
};
