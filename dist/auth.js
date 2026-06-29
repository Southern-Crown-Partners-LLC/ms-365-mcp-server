import { AuthError, PublicClientApplication } from "@azure/msal-node";
import logger from "./logger.js";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import path from "path";
import { getSecrets } from "./secrets.js";
import { getCloudEndpoints, getDefaultClientId } from "./cloud-config.js";
import {
  createTokenCacheStorage,
  DefaultTokenCacheStorage,
  getSelectedAccountPath,
  getTokenCachePath,
  pickNewest,
  unwrapCache,
  wrapCache
} from "./token-cache-storage.js";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, "endpoints.json"), "utf8")
);
const endpoints = {
  default: endpointsData
};
function createMsalConfig(secrets) {
  const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
  return {
    auth: {
      clientId: secrets.clientId || getDefaultClientId(secrets.cloudType),
      authority: `${cloudEndpoints.authority}/${secrets.tenantId || "common"}`
    }
  };
}
function buildDiskCoherencyCachePlugin(storage) {
  return {
    beforeCacheAccess: async (context) => {
      try {
        const cacheRaw = await storage.load("token-cache");
        if (cacheRaw) {
          context.tokenCache.deserialize(unwrapCache(cacheRaw).data);
        }
      } catch (error) {
        logger.error(`Error reloading token cache: ${error.message}`);
        if (storage.failClosed) {
          throw error;
        }
      }
    },
    afterCacheAccess: async (context) => {
      if (!context.cacheHasChanged) {
        return;
      }
      try {
        await storage.save("token-cache", wrapCache(context.tokenCache.serialize()));
      } catch (error) {
        logger.error(`Error saving token cache: ${error.message}`);
        if (storage.failClosed) {
          throw error;
        }
      }
    }
  };
}
const SCOPE_HIERARCHY = {
  "Mail.ReadWrite": ["Mail.Read"],
  "Calendars.ReadWrite": ["Calendars.Read"],
  "Files.ReadWrite": ["Files.Read"],
  "Tasks.ReadWrite": ["Tasks.Read"],
  "Contacts.ReadWrite": ["Contacts.Read"]
};
function parseAllowedScopes(value) {
  if (value === void 0) {
    return void 0;
  }
  return Array.from(new Set(value.trim().split(/\s+/).filter(Boolean)));
}
function getEndpointRequiredScopes(endpoint, includeWorkAccountScopes = false) {
  if (!endpoint) {
    return [];
  }
  const scopes = /* @__PURE__ */ new Set();
  getEndpointScopeGroups(endpoint, includeWorkAccountScopes).forEach(
    (group) => group.forEach((scope) => scopes.add(scope))
  );
  return Array.from(scopes);
}
function toScopeGroups(value) {
  if (!value || value.length === 0) {
    return [];
  }
  return Array.isArray(value[0]) ? value : [value];
}
function getEndpointScopeGroups(endpoint, includeWorkAccountScopes = false) {
  if (!endpoint) {
    return [];
  }
  const groups = [...toScopeGroups(endpoint.scopes)];
  if (includeWorkAccountScopes) {
    groups.push(...toScopeGroups(endpoint.workScopes));
  }
  return groups;
}
function getEndpointLoginScopes(endpoint, includeWorkAccountScopes = false) {
  const groups = getEndpointScopeGroups(endpoint, includeWorkAccountScopes);
  return groups.length > 0 ? groups[0] : [];
}
function getMissingAllowedScopesForGroups(scopeGroups, allowedScopes) {
  if (allowedScopes === void 0 || scopeGroups.length === 0) {
    return [];
  }
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopes));
  let closest;
  for (const group of scopeGroups) {
    const missing = group.filter((scope) => !coveredAllowedScopes.has(scope));
    if (missing.length === 0) {
      return [];
    }
    if (!closest || missing.length < closest.length) {
      closest = missing;
    }
  }
  return closest ?? [];
}
function getEndpointEffectiveLoginScopes(scopeGroups, allowedScopes) {
  if (scopeGroups.length === 0) {
    return [];
  }
  if (allowedScopes === void 0) {
    return scopeGroups[0];
  }
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopes));
  const satisfied = scopeGroups.find(
    (group) => group.every((scope) => coveredAllowedScopes.has(scope))
  );
  return satisfied ?? [];
}
function collapseRedundantScopes(scopes) {
  const scopesSet = new Set(scopes);
  Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
    if (scopesSet.has(higherScope) && lowerScopes.every((scope) => scopesSet.has(scope))) {
      lowerScopes.forEach((scope) => scopesSet.delete(scope));
    }
  });
  return Array.from(scopesSet);
}
function buildScopesFromEndpoints(includeWorkAccountScopes = false, enabledToolsPattern, readOnly = false) {
  const scopesSet = /* @__PURE__ */ new Set();
  let enabledToolsRegex;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, "i");
      logger.info(`Building scopes with tool filter pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(
        `Invalid tool filter regex pattern: ${enabledToolsPattern}. Building scopes without filter.`
      );
    }
  }
  endpoints.default.forEach((endpoint) => {
    if (readOnly && endpoint.method.toUpperCase() !== "GET") {
      if (!(endpoint.method.toUpperCase() === "POST" && endpoint.readOnly)) {
        return;
      }
    }
    if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
      return;
    }
    if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
      return;
    }
    getEndpointLoginScopes(endpoint, includeWorkAccountScopes).forEach(
      (scope) => scopesSet.add(scope)
    );
  });
  const scopes = collapseRedundantScopes(Array.from(scopesSet));
  if (enabledToolsPattern) {
    logger.info(`Built ${scopes.length} scopes for filtered tools: ${scopes.join(", ")}`);
  }
  return scopes;
}
function lowerScopesFor(scope) {
  const lowerScopes = new Set(SCOPE_HIERARCHY[scope] ?? []);
  if (scope.endsWith(".ReadWrite.All")) {
    const readAllScope = scope.replace(/\.ReadWrite\.All$/, ".Read.All");
    const readWriteScope = scope.replace(/\.ReadWrite\.All$/, ".ReadWrite");
    const readScope = scope.replace(/\.ReadWrite\.All$/, ".Read");
    lowerScopes.add(readAllScope);
    lowerScopes.add(readWriteScope);
    lowerScopes.add(readScope);
  } else if (scope.endsWith(".ReadWrite.Shared")) {
    lowerScopes.add(scope.replace(/\.ReadWrite\.Shared$/, ".Read.Shared"));
  } else if (scope.endsWith(".ReadWrite")) {
    lowerScopes.add(scope.replace(/\.ReadWrite$/, ".Read"));
  } else if (scope.endsWith(".Read.All")) {
    lowerScopes.add(scope.replace(/\.Read\.All$/, ".Read"));
  }
  return Array.from(lowerScopes);
}
function addImpliedScopes(scope, scopesSet) {
  for (const lowerScope of lowerScopesFor(scope)) {
    if (!scopesSet.has(lowerScope)) {
      scopesSet.add(lowerScope);
      addImpliedScopes(lowerScope, scopesSet);
    }
  }
}
function collapseScopeHierarchy(scopes) {
  const scopesSet = new Set(scopes);
  for (const scope of scopes) {
    addImpliedScopes(scope, scopesSet);
  }
  return Array.from(scopesSet);
}
function getMissingAllowedScopes(requiredScopes, allowedScopes) {
  if (allowedScopes === void 0) {
    return [];
  }
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopes));
  return requiredScopes.filter((scope) => !coveredAllowedScopes.has(scope));
}
function isScopeUsedByTools(allowedScope, toolScopes) {
  const coveredByAllowedScope = new Set(collapseScopeHierarchy([allowedScope]));
  return toolScopes.some((scope) => coveredByAllowedScope.has(scope));
}
function endpointMatchesNormalToolSurface(endpoint, includeWorkAccountScopes, enabledToolsRegex, readOnly = false) {
  if (readOnly && endpoint.method.toUpperCase() !== "GET") {
    if (!(endpoint.method.toUpperCase() === "POST" && endpoint.readOnly)) {
      return false;
    }
  }
  if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
    return false;
  }
  if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
    return false;
  }
  return true;
}
function buildAllowedScopeDiagnostics(options = {}) {
  const allowedScopes = parseAllowedScopes(options.allowedScopes);
  let enabledToolsRegex;
  if (options.enabledTools) {
    try {
      enabledToolsRegex = new RegExp(options.enabledTools, "i");
    } catch {
      logger.error(
        `Invalid tool filter regex pattern: ${options.enabledTools}. Building diagnostics without filter.`
      );
    }
  }
  const normalToolScopes = /* @__PURE__ */ new Set();
  const effectiveToolScopes = /* @__PURE__ */ new Set();
  const effectiveToolScopesAllGroups = /* @__PURE__ */ new Set();
  const disabledTools = [];
  for (const endpoint of endpoints.default) {
    if (!endpointMatchesNormalToolSurface(
      endpoint,
      Boolean(options.orgMode),
      enabledToolsRegex,
      Boolean(options.readOnly)
    )) {
      continue;
    }
    const scopeGroups = getEndpointScopeGroups(endpoint, Boolean(options.orgMode));
    const loginScopes = getEndpointLoginScopes(endpoint, Boolean(options.orgMode));
    const allScopes = getEndpointRequiredScopes(endpoint, Boolean(options.orgMode));
    loginScopes.forEach((scope) => normalToolScopes.add(scope));
    const missingScopes = getMissingAllowedScopesForGroups(scopeGroups, allowedScopes);
    if (missingScopes.length > 0) {
      disabledTools.push({
        toolName: endpoint.toolName,
        requiredScopes: allScopes.sort((a, b) => a.localeCompare(b)),
        missingScopes: missingScopes.sort((a, b) => a.localeCompare(b))
      });
      continue;
    }
    getEndpointEffectiveLoginScopes(scopeGroups, allowedScopes).forEach(
      (scope) => effectiveToolScopes.add(scope)
    );
    allScopes.forEach((scope) => effectiveToolScopesAllGroups.add(scope));
  }
  const toolPermissions = collapseRedundantScopes(Array.from(normalToolScopes)).sort(
    (a, b) => a.localeCompare(b)
  );
  const effectivePermissions = collapseRedundantScopes(Array.from(effectiveToolScopes)).sort(
    (a, b) => a.localeCompare(b)
  );
  const sortedAllowedScopes = allowedScopes ? [...allowedScopes].sort((a, b) => a.localeCompare(b)) : void 0;
  const missingAllowedScopesForTools = Array.from(
    new Set(disabledTools.flatMap((tool) => tool.missingScopes))
  ).sort((a, b) => a.localeCompare(b));
  const allEffectiveToolScopes = Array.from(effectiveToolScopesAllGroups);
  const extraAllowedScopesNotUsedByTools = sortedAllowedScopes?.filter((scope) => !isScopeUsedByTools(scope, allEffectiveToolScopes)) ?? [];
  return {
    permissions: effectivePermissions,
    toolPermissions,
    effectivePermissions,
    ...sortedAllowedScopes ? { allowedScopes: sortedAllowedScopes } : {},
    disabledTools,
    missingAllowedScopesForTools,
    extraAllowedScopesNotUsedByTools
  };
}
function resolveAuthScopes(options = {}) {
  const toolScopes = buildAllowedScopeDiagnostics(options).effectivePermissions;
  const extraScopes = parseAllowedScopes(options.extraScopes);
  if (!extraScopes || extraScopes.length === 0) {
    return toolScopes;
  }
  return Array.from(/* @__PURE__ */ new Set([...toolScopes, ...extraScopes]));
}
function buildScopeDiagnostics(toolScopes, allowedScopesInput) {
  const toolPermissions = [...toolScopes].sort((a, b) => a.localeCompare(b));
  const coveredAllowedScopes = new Set(collapseScopeHierarchy(allowedScopesInput));
  const missingAllowedScopesForTools = toolPermissions.filter(
    (scope) => !coveredAllowedScopes.has(scope)
  );
  return {
    permissions: toolPermissions.filter((scope) => coveredAllowedScopes.has(scope)),
    toolPermissions,
    effectivePermissions: toolPermissions.filter((scope) => coveredAllowedScopes.has(scope)),
    allowedScopes: [...allowedScopesInput].sort((a, b) => a.localeCompare(b)),
    disabledTools: [],
    missingAllowedScopesForTools,
    extraAllowedScopesNotUsedByTools: [...allowedScopesInput].sort((a, b) => a.localeCompare(b)).filter((scope) => !isScopeUsedByTools(scope, toolPermissions))
  };
}
function describeAuthError(error) {
  if (error instanceof AuthError) {
    const suberror = error.subError ? ` / ${error.subError}` : "";
    return `${error.errorCode}${suberror} (correlationId: ${error.correlationId || "none"}): ${error.errorMessage}`;
  }
  return error.message;
}
const MSA_HOME_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad";
function consumersAuthorityHint(error, account, authority) {
  if (error instanceof AuthError && error.errorCode === "invalid_grant" && account?.tenantId === MSA_HOME_TENANT_ID && (!authority || /\/common\/?$/i.test(authority))) {
    return `This looks like a known issue (June 2026) where Microsoft rejects refresh tokens issued to personal accounts via the default 'common' authority. If this server is used only with personal accounts, set MS365_MCP_TENANT_ID=consumers and re-login with: --login`;
  }
  return null;
}
class AuthManager {
  constructor(config, scopes = [], expectedAccount, storage) {
    logger.info(`And scopes are ${scopes.join(", ")}`, scopes);
    this.scopes = scopes;
    this.storage = storage ?? new DefaultTokenCacheStorage();
    this.config = {
      ...config,
      cache: {
        ...config.cache,
        cachePlugin: buildDiskCoherencyCachePlugin(this.storage)
      }
    };
    this.msalApp = new PublicClientApplication(this.config);
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;
    this.useInteractiveAuth = false;
    this.expectedUsername = this.normalizeExpectedUsername(expectedAccount?.expectedUsername);
    this.expectedHomeAccountId = this.normalizeExpectedHomeAccountId(
      expectedAccount?.expectedHomeAccountId
    );
    const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
    this.oauthToken = oauthTokenFromEnv ?? null;
    this.isOAuthMode = oauthTokenFromEnv != null;
  }
  /**
   * Creates an AuthManager instance with secrets loaded from the configured provider.
   * Uses Key Vault if MS365_MCP_KEYVAULT_URL is set, otherwise environment variables.
   */
  static async create(scopes = [], expectedAccount, options = {}) {
    const secrets = await getSecrets();
    const config = createMsalConfig(secrets);
    const storage = options.storage ?? await createTokenCacheStorage({ allowCommandStorage: false, logProvider: true });
    return new AuthManager(config, scopes, expectedAccount, storage);
  }
  async loadTokenCache() {
    try {
      const cacheRaw = await this.storage.load("token-cache");
      if (cacheRaw) {
        this.msalApp.getTokenCache().deserialize(unwrapCache(cacheRaw).data);
      }
      await this.loadSelectedAccount();
    } catch (error) {
      logger.error(`Error loading token cache: ${error.message}`);
      if (this.storage.failClosed) {
        throw error;
      }
    }
  }
  async loadSelectedAccount() {
    try {
      const selectedAccountRaw = await this.storage.load("selected-account");
      if (selectedAccountRaw) {
        const parsed = JSON.parse(unwrapCache(selectedAccountRaw).data);
        this.selectedAccountId = parsed.accountId;
        logger.info(`Loaded selected account: ${this.selectedAccountId}`);
      }
    } catch (error) {
      logger.error(`Error loading selected account: ${error.message}`);
      if (this.storage.failClosed) {
        throw error;
      }
    }
  }
  async saveSelectedAccount() {
    try {
      const stamped = wrapCache(JSON.stringify({ accountId: this.selectedAccountId }));
      await this.storage.save("selected-account", stamped);
    } catch (error) {
      logger.error(`Error saving selected account: ${error.message}`);
      if (this.storage.failClosed) {
        throw error;
      }
    }
  }
  normalizeExpectedUsername(value) {
    if (value === void 0) {
      return null;
    }
    const trimmed = value.trim();
    if (trimmed === "") {
      throw new Error("Expected Microsoft account username was provided but is empty.");
    }
    return trimmed.toLowerCase();
  }
  normalizeExpectedHomeAccountId(value) {
    if (value === void 0) {
      return null;
    }
    const trimmed = value.trim();
    if (trimmed === "") {
      throw new Error("Expected Microsoft account homeAccountId was provided but is empty.");
    }
    return trimmed;
  }
  hasExpectedAccount() {
    return this.expectedUsername !== null || this.expectedHomeAccountId !== null;
  }
  expectedAccountLabel() {
    const parts = [];
    if (this.expectedUsername) {
      parts.push(`username ${this.expectedUsername}`);
    }
    if (this.expectedHomeAccountId) {
      parts.push(`homeAccountId ${this.expectedHomeAccountId}`);
    }
    return parts.join(" and ");
  }
  describeAccount(account) {
    return account?.username || account?.name || "unknown";
  }
  describeCachedAccounts(accounts) {
    if (accounts.length === 0) {
      return "none";
    }
    return accounts.map((account) => this.describeAccount(account)).join(", ");
  }
  accountMatchesExpected(account) {
    if (!this.hasExpectedAccount() || !account) {
      return !this.hasExpectedAccount();
    }
    if (this.expectedUsername && account.username?.toLowerCase() !== this.expectedUsername) {
      return false;
    }
    if (this.expectedHomeAccountId && account.homeAccountId !== this.expectedHomeAccountId) {
      return false;
    }
    return true;
  }
  buildExpectedAccountMissingError(accounts) {
    return new Error(
      `Expected Microsoft account '${this.expectedAccountLabel()}' not found in token cache. Cached accounts: ${this.describeCachedAccounts(accounts)}. Run --login after configuring the expected account, or use --select-account to recover.`
    );
  }
  resolveExpectedAccountFromAccounts(accounts) {
    if (!this.hasExpectedAccount()) {
      throw new Error("No expected Microsoft account is configured.");
    }
    const usernameMatch = this.expectedUsername ? accounts.find((account) => account.username?.toLowerCase() === this.expectedUsername) : void 0;
    const homeAccountIdMatch = this.expectedHomeAccountId ? accounts.find((account) => account.homeAccountId === this.expectedHomeAccountId) : void 0;
    if (this.expectedUsername && this.expectedHomeAccountId) {
      if (!usernameMatch || !homeAccountIdMatch) {
        throw this.buildExpectedAccountMissingError(accounts);
      }
      if (usernameMatch.homeAccountId !== homeAccountIdMatch.homeAccountId) {
        throw new Error(
          `Expected Microsoft account pins conflict: username ${this.expectedUsername} matched ${this.describeAccount(usernameMatch)}, but homeAccountId ${this.expectedHomeAccountId} matched ${this.describeAccount(homeAccountIdMatch)}.`
        );
      }
      return usernameMatch;
    }
    const expectedAccount = usernameMatch ?? homeAccountIdMatch;
    if (!expectedAccount) {
      throw this.buildExpectedAccountMissingError(accounts);
    }
    return expectedAccount;
  }
  async assertExpectedAccountAvailable() {
    if (!this.hasExpectedAccount()) {
      return;
    }
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    this.resolveExpectedAccountFromAccounts(accounts);
  }
  async rejectUnexpectedLoginAccount(account) {
    if (!this.hasExpectedAccount()) {
      return;
    }
    if (this.accountMatchesExpected(account)) {
      return;
    }
    this.accessToken = null;
    this.tokenExpiry = null;
    if (account) {
      try {
        await this.msalApp.getTokenCache().removeAccount(account);
      } catch (error) {
        logger.error(`Failed to remove unexpected account from cache: ${error.message}`);
        throw new Error(
          `Authenticated Microsoft account '${this.describeAccount(account)}' does not match expected Microsoft account '${this.expectedAccountLabel()}', and it could not be removed from the token cache (${error.message}). Its tokens may remain persisted - run --logout to clear the cache, then re-login.`
        );
      }
      throw new Error(
        `Authenticated Microsoft account '${this.describeAccount(account)}' does not match expected Microsoft account '${this.expectedAccountLabel()}'. Login was not persisted.`
      );
    }
    throw new Error(
      `Microsoft login did not return an account. Expected Microsoft account '${this.expectedAccountLabel()}'. Login was not persisted.`
    );
  }
  async setOAuthToken(token) {
    this.oauthToken = token;
    this.isOAuthMode = true;
  }
  async getToken(forceRefresh = false) {
    if (this.isOAuthMode && this.oauthToken) {
      return this.oauthToken;
    }
    if (this.accessToken && this.tokenExpiry && this.tokenExpiry > Date.now() && !forceRefresh) {
      return this.accessToken;
    }
    const currentAccount = await this.getCurrentAccount();
    if (currentAccount) {
      const silentRequest = {
        account: currentAccount,
        scopes: this.scopes
      };
      try {
        const response = await this.msalApp.acquireTokenSilent(silentRequest);
        this.accessToken = response.accessToken;
        this.tokenExpiry = response.expiresOn ? new Date(response.expiresOn).getTime() : null;
        return this.accessToken;
      } catch (error) {
        const hint = consumersAuthorityHint(error, currentAccount, this.config.auth.authority);
        logger.error(
          `Silent token acquisition failed: ${describeAuthError(error)}${hint ? ` ${hint}` : ""}`
        );
        throw new Error(
          hint ? `Silent token acquisition failed. ${hint}` : "Silent token acquisition failed"
        );
      }
    }
    throw new Error("No valid token found");
  }
  async getCurrentAccount() {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    if (this.hasExpectedAccount()) {
      return this.resolveExpectedAccountFromAccounts(accounts);
    }
    if (accounts.length === 0) {
      return null;
    }
    if (this.selectedAccountId) {
      const selectedAccount = accounts.find(
        (account) => account.homeAccountId === this.selectedAccountId
      );
      if (selectedAccount) {
        return selectedAccount;
      }
      logger.warn(
        `Selected account ${this.selectedAccountId} not found, falling back to first account`
      );
    }
    return accounts[0];
  }
  async acquireTokenByDeviceCode(hack) {
    const deviceCodeRequest = {
      scopes: this.scopes,
      deviceCodeCallback: (response) => {
        const text = ["\n", response.message, "\n"].join("");
        if (hack) {
          hack(text + 'After login run the "verify login" command');
        } else {
          console.log(text);
        }
        logger.info("Device code login initiated");
      }
    };
    try {
      logger.info("Requesting device code...");
      logger.info(`Requesting scopes: ${this.scopes.join(", ")}`);
      const response = await this.msalApp.acquireTokenByDeviceCode(deviceCodeRequest);
      logger.info(`Granted scopes: ${response?.scopes?.join(", ") || "none"}`);
      logger.info("Device code login successful");
      this.accessToken = response?.accessToken || null;
      this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;
      await this.rejectUnexpectedLoginAccount(response?.account);
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }
      return this.accessToken;
    } catch (error) {
      logger.error(`Error in device code flow: ${error.message}`);
      throw error;
    }
  }
  setUseInteractiveAuth(value) {
    this.useInteractiveAuth = value;
  }
  getUseInteractiveAuth() {
    return this.useInteractiveAuth;
  }
  async acquireTokenInteractive(hack) {
    const open = (await import("open")).default;
    const interactiveRequest = {
      scopes: this.scopes,
      openBrowser: async (url) => {
        const message = "Opening browser for Microsoft sign-in...";
        if (hack) {
          hack(message);
        }
        logger.info(message);
        await open(url);
      },
      successTemplate: "<h1>Authentication successful!</h1><p>You can close this window and return to your application.</p>",
      errorTemplate: "<h1>Authentication failed</h1><p>Something went wrong. Please try again.</p>"
    };
    try {
      logger.info("Requesting interactive browser login...");
      logger.info(`Requesting scopes: ${this.scopes.join(", ")}`);
      const response = await this.msalApp.acquireTokenInteractive(interactiveRequest);
      logger.info(`Granted scopes: ${response?.scopes?.join(", ") || "none"}`);
      logger.info("Interactive browser login successful");
      this.accessToken = response?.accessToken || null;
      this.tokenExpiry = response?.expiresOn ? new Date(response.expiresOn).getTime() : null;
      await this.rejectUnexpectedLoginAccount(response?.account);
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }
      return this.accessToken;
    } catch (error) {
      logger.error(`Error in interactive browser flow: ${error.message}`);
      throw error;
    }
  }
  async testLogin() {
    try {
      logger.info("Testing login...");
      const token = await this.getToken();
      if (!token) {
        logger.error("Login test failed - no token received");
        return {
          success: false,
          message: "Login failed - no token received"
        };
      }
      logger.info("Token retrieved successfully, testing Graph API access...");
      try {
        const secrets = await getSecrets();
        const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
        const response = await fetch(`${cloudEndpoints.graphApi}/v1.0/me`, {
          headers: {
            Authorization: `Bearer ${token}`
          }
        });
        if (response.ok) {
          const userData = await response.json();
          logger.info("Graph API user data fetch successful");
          return {
            success: true,
            message: "Login successful",
            userData: {
              displayName: userData.displayName,
              userPrincipalName: userData.userPrincipalName
            }
          };
        } else {
          const errorText = await response.text();
          logger.error(`Graph API user data fetch failed: ${response.status} - ${errorText}`);
          return {
            success: false,
            message: `Login successful but Graph API access failed: ${response.status}`
          };
        }
      } catch (graphError) {
        logger.error(`Error fetching user data: ${graphError.message}`);
        return {
          success: false,
          message: `Login successful but Graph API access failed: ${graphError.message}`
        };
      }
    } catch (error) {
      logger.error(`Login test failed: ${error.message}`);
      return {
        success: false,
        message: `Login failed: ${error.message}`
      };
    }
  }
  async logout() {
    try {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.msalApp.getTokenCache().removeAccount(account);
      }
      this.accessToken = null;
      this.tokenExpiry = null;
      this.selectedAccountId = null;
      await this.storage.delete("token-cache");
      await this.storage.delete("selected-account");
      return true;
    } catch (error) {
      logger.error(`Error during logout: ${error.message}`);
      throw error;
    }
  }
  // Multi-account support methods
  async listAccounts() {
    return await this.msalApp.getTokenCache().getAllAccounts();
  }
  async selectAccount(identifier) {
    const account = await this.resolveAccount(identifier);
    if (this.hasExpectedAccount() && !this.accountMatchesExpected(account)) {
      throw new Error(
        `Account '${identifier}' does not match expected Microsoft account '${this.expectedAccountLabel()}'.`
      );
    }
    this.selectedAccountId = account.homeAccountId;
    await this.saveSelectedAccount();
    this.accessToken = null;
    this.tokenExpiry = null;
    logger.info(`Selected account: ${account.username} (${account.homeAccountId})`);
    return true;
  }
  async removeAccount(identifier) {
    const account = await this.resolveAccount(identifier);
    try {
      await this.msalApp.getTokenCache().removeAccount(account);
      if (this.selectedAccountId === account.homeAccountId) {
        this.selectedAccountId = null;
        await this.saveSelectedAccount();
        this.accessToken = null;
        this.tokenExpiry = null;
      }
      logger.info(`Removed account: ${account.username} (${account.homeAccountId})`);
      return true;
    } catch (error) {
      logger.error(`Failed to remove account ${identifier}: ${error.message}`);
      return false;
    }
  }
  getSelectedAccountId() {
    return this.selectedAccountId;
  }
  /**
   * Returns true if auth is in OAuth/HTTP mode (token supplied via env or setOAuthToken).
   * In this mode, account resolution should be skipped — the request context drives token selection.
   */
  isOAuthModeEnabled() {
    return this.isOAuthMode;
  }
  /**
   * Resolves an account by identifier (email or homeAccountId).
   * Resolution: username match (case-insensitive) → homeAccountId match → throw.
   */
  async resolveAccount(identifier) {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    if (accounts.length === 0) {
      throw new Error("No accounts found. Please login first.");
    }
    const lowerIdentifier = identifier.toLowerCase();
    let account = accounts.find((a) => a.username?.toLowerCase() === lowerIdentifier) ?? null;
    if (!account) {
      account = accounts.find((a) => a.homeAccountId === identifier) ?? null;
    }
    if (!account) {
      const availableAccounts = accounts.map((a) => a.username || a.name || "unknown").join(", ");
      throw new Error(
        `Account '${identifier}' not found. Available accounts: ${availableAccounts}`
      );
    }
    return account;
  }
  /**
   * Returns true if the MSAL cache contains more than one account.
   * Used to decide whether to inject the `account` parameter into tool schemas.
   */
  async isMultiAccount() {
    if (this.hasExpectedAccount()) {
      return false;
    }
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
    return accounts.length > 1;
  }
  /**
   * Acquires a token for a specific account identified by username (email) or homeAccountId,
   * WITHOUT changing the persisted selectedAccountId.
   *
   * Resolution order:
   *  1. Exact match on username (case-insensitive)
   *  2. Exact match on homeAccountId
   *  3. If identifier is empty/undefined AND only 1 account exists → auto-select
   *  4. If identifier is empty/undefined AND multiple accounts → use selectedAccountId or throw
   *
   * @returns The access token string.
   */
  async getTokenForAccount(identifier) {
    if (this.isOAuthMode && this.oauthToken) {
      if (identifier) {
        throw new Error(
          `Cannot switch to account '${identifier}': the server is in OAuth mode and always uses the identity of the supplied bearer token. Account switching requires stdio mode (or HTTP with --trust-proxy-auth).`
        );
      }
      return this.oauthToken;
    }
    let targetAccount = null;
    if (this.hasExpectedAccount()) {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      targetAccount = this.resolveExpectedAccountFromAccounts(accounts);
      if (identifier) {
        const requestedAccount = await this.resolveAccount(identifier);
        if (requestedAccount.homeAccountId !== targetAccount.homeAccountId) {
          throw new Error(
            `Account '${identifier}' does not match expected Microsoft account '${this.expectedAccountLabel()}'.`
          );
        }
      }
    } else if (identifier) {
      targetAccount = await this.resolveAccount(identifier);
    } else {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      if (accounts.length === 0) {
        throw new Error("No accounts found. Please login first.");
      }
      if (accounts.length === 1) {
        targetAccount = accounts[0];
      } else {
        if (this.selectedAccountId) {
          targetAccount = accounts.find((a) => a.homeAccountId === this.selectedAccountId) ?? null;
        }
        if (!targetAccount) {
          const availableAccounts = accounts.map((a) => a.username || a.name || "unknown").join(", ");
          throw new Error(
            `Multiple accounts configured but no 'account' parameter provided and no default selected. Available accounts: ${availableAccounts}. Pass account="<email>" in your tool call or use select-account to set a default.`
          );
        }
      }
    }
    const silentRequest = {
      account: targetAccount,
      scopes: this.scopes
    };
    try {
      const response = await this.msalApp.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (error) {
      const hint = consumersAuthorityHint(error, targetAccount, this.config.auth.authority);
      logger.error(
        `Silent token acquisition failed: ${describeAuthError(error)}${hint ? ` ${hint}` : ""}`
      );
      throw new Error(
        `Failed to acquire token for account '${targetAccount.username || targetAccount.name || "unknown"}'. ` + (hint ?? "The token may have expired. Please re-login with: --login")
      );
    }
  }
}
var auth_default = AuthManager;
export {
  buildAllowedScopeDiagnostics,
  buildDiskCoherencyCachePlugin,
  buildScopeDiagnostics,
  buildScopesFromEndpoints,
  collapseScopeHierarchy,
  consumersAuthorityHint,
  auth_default as default,
  describeAuthError,
  getEndpointRequiredScopes,
  getEndpointScopeGroups,
  getMissingAllowedScopes,
  getMissingAllowedScopesForGroups,
  getSelectedAccountPath,
  getTokenCachePath,
  parseAllowedScopes,
  pickNewest,
  resolveAuthScopes,
  unwrapCache,
  wrapCache
};
