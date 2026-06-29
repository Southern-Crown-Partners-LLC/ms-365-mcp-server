const LOOPBACK_HOSTS = /* @__PURE__ */ new Set(["localhost", "127.0.0.1", "[::1]", "::1"]);
function isLoopbackHost(host) {
  return LOOPBACK_HOSTS.has(host.toLowerCase());
}
function parseAllowlist(raw) {
  if (!raw) return null;
  const list = raw.split(",").map((s) => s.trim()).filter(Boolean);
  return list.length > 0 ? list : null;
}
function isAllowedRedirectUri(value, allowlist) {
  if (typeof value !== "string" || value.length === 0) return false;
  let url;
  try {
    url = new URL(value);
  } catch {
    return false;
  }
  if (url.protocol !== "http:" && url.protocol !== "https:") {
    return false;
  }
  if (allowlist && allowlist.length > 0) {
    return allowlist.includes(value);
  }
  if (url.protocol === "http:") {
    return isLoopbackHost(url.hostname);
  }
  return true;
}
export {
  isAllowedRedirectUri,
  isLoopbackHost,
  parseAllowlist
};
