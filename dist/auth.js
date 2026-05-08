import { PublicClientApplication } from "@azure/msal-node";
import logger from "./logger.js";
import fs, { existsSync, readFileSync } from "fs";
import { fileURLToPath } from "url";
import path from "path";
import { getSecrets } from "./secrets.js";
import { getCloudEndpoints, getDefaultClientId } from "./cloud-config.js";
let keytar = null;
async function getKeytar() {
  if (keytar === void 0) {
    return null;
  }
  if (keytar === null) {
    try {
      const mod = await import("keytar");
      keytar = mod.default ?? mod;
      return keytar;
    } catch (error) {
      logger.info("keytar not available, using file-based credential storage");
      keytar = void 0;
      return null;
    }
  }
  return keytar;
}
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, "endpoints.json"), "utf8")
);
const endpoints = {
  default: endpointsData
};
const SERVICE_NAME = "ms-365-mcp-server";
const TOKEN_CACHE_ACCOUNT = "msal-token-cache";
const SELECTED_ACCOUNT_KEY = "selected-account";
const FALLBACK_DIR = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_TOKEN_CACHE_PATH = path.join(FALLBACK_DIR, "..", ".token-cache.json");
const DEFAULT_SELECTED_ACCOUNT_PATH = path.join(FALLBACK_DIR, "..", ".selected-account.json");
function getTokenCachePath() {
  const envPath = process.env.MS365_MCP_TOKEN_CACHE_PATH?.trim();
  return envPath || DEFAULT_TOKEN_CACHE_PATH;
}
function getSelectedAccountPath() {
  const envPath = process.env.MS365_MCP_SELECTED_ACCOUNT_PATH?.trim();
  return envPath || DEFAULT_SELECTED_ACCOUNT_PATH;
}
function ensureParentDir(filePath) {
  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true, mode: 448 });
}
function wrapCache(data) {
  return JSON.stringify({ _cacheEnvelope: true, data, savedAt: Date.now() });
}
function unwrapCache(raw) {
  try {
    const parsed = JSON.parse(raw);
    if (parsed._cacheEnvelope && typeof parsed.data === "string") {
      return { data: parsed.data, savedAt: parsed.savedAt };
    }
  } catch {
  }
  return { data: raw };
}
function pickNewest(keytarRaw, fileRaw) {
  if (!keytarRaw && !fileRaw) return void 0;
  if (keytarRaw && !fileRaw) return unwrapCache(keytarRaw).data;
  if (!keytarRaw && fileRaw) return unwrapCache(fileRaw).data;
  const kt = unwrapCache(keytarRaw);
  const file = unwrapCache(fileRaw);
  if (kt.savedAt === void 0 && file.savedAt === void 0) return kt.data;
  if (kt.savedAt !== void 0 && file.savedAt === void 0) return kt.data;
  if (kt.savedAt === void 0 && file.savedAt !== void 0) return file.data;
  return kt.savedAt >= file.savedAt ? kt.data : file.data;
}
function createMsalConfig(secrets) {
  const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
  return {
    auth: {
      clientId: secrets.clientId || getDefaultClientId(secrets.cloudType),
      authority: `${cloudEndpoints.authority}/${secrets.tenantId || "common"}`
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
function buildScopesFromEndpoints(includeWorkAccountScopes = false, enabledToolsPattern, readOnly = false) {
  const scopesSet = /* @__PURE__ */ new Set();
  let enabledToolsRegex;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, "i");
      logger.info(`Building scopes with tool filter pattern: ${enabledToolsPattern}`);
    } catch (error) {
      logger.error(
        `Invalid tool filter regex pattern: ${enabledToolsPattern}. Building scopes without filter.`
      );
    }
  }
  endpoints.default.forEach((endpoint) => {
    if (readOnly && endpoint.method.toUpperCase() !== "GET") {
      return;
    }
    if (enabledToolsRegex && !enabledToolsRegex.test(endpoint.toolName)) {
      return;
    }
    if (!includeWorkAccountScopes && !endpoint.scopes && endpoint.workScopes) {
      return;
    }
    if (endpoint.scopes && Array.isArray(endpoint.scopes)) {
      endpoint.scopes.forEach((scope) => scopesSet.add(scope));
    }
    if (includeWorkAccountScopes && endpoint.workScopes && Array.isArray(endpoint.workScopes)) {
      endpoint.workScopes.forEach((scope) => scopesSet.add(scope));
    }
  });
  Object.entries(SCOPE_HIERARCHY).forEach(([higherScope, lowerScopes]) => {
    if (scopesSet.has(higherScope) && lowerScopes.every((scope) => scopesSet.has(scope))) {
      lowerScopes.forEach((scope) => scopesSet.delete(scope));
    }
  });
  const scopes = Array.from(scopesSet);
  if (enabledToolsPattern) {
    logger.info(`Built ${scopes.length} scopes for filtered tools: ${scopes.join(", ")}`);
  }
  return scopes;
}
class AuthManager {
  constructor(config, scopes = buildScopesFromEndpoints()) {
    logger.info(`And scopes are ${scopes.join(", ")}`, scopes);
    this.config = config;
    this.scopes = scopes;
    this.msalApp = new PublicClientApplication(this.config);
    this.accessToken = null;
    this.tokenExpiry = null;
    this.selectedAccountId = null;
    this.useInteractiveAuth = false;
    const oauthTokenFromEnv = process.env.MS365_MCP_OAUTH_TOKEN;
    this.oauthToken = oauthTokenFromEnv ?? null;
    this.isOAuthMode = oauthTokenFromEnv != null;
  }
  /**
   * Creates an AuthManager instance with secrets loaded from the configured provider.
   * Uses Key Vault if MS365_MCP_KEYVAULT_URL is set, otherwise environment variables.
   */
  static async create(scopes = buildScopesFromEndpoints()) {
    const secrets = await getSecrets();
    const config = createMsalConfig(secrets);
    return new AuthManager(config, scopes);
  }
  async loadTokenCache() {
    try {
      let keytarRaw;
      try {
        const kt = await getKeytar();
        if (kt) {
          keytarRaw = await kt.getPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT) ?? void 0;
        }
      } catch (keytarError) {
        logger.warn(`Keychain access failed: ${keytarError.message}`);
      }
      let fileRaw;
      const cachePath = getTokenCachePath();
      if (existsSync(cachePath)) {
        fileRaw = readFileSync(cachePath, "utf8");
      }
      const cacheData = pickNewest(keytarRaw, fileRaw);
      if (cacheData) {
        this.msalApp.getTokenCache().deserialize(cacheData);
      }
      await this.loadSelectedAccount();
    } catch (error) {
      logger.error(`Error loading token cache: ${error.message}`);
    }
  }
  async loadSelectedAccount() {
    try {
      let keytarRaw;
      try {
        const kt = await getKeytar();
        if (kt) {
          keytarRaw = await kt.getPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY) ?? void 0;
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain access failed for selected account: ${keytarError.message}`
        );
      }
      let fileRaw;
      const accountPath = getSelectedAccountPath();
      if (existsSync(accountPath)) {
        fileRaw = readFileSync(accountPath, "utf8");
      }
      const selectedAccountData = pickNewest(keytarRaw, fileRaw);
      if (selectedAccountData) {
        const parsed = JSON.parse(selectedAccountData);
        this.selectedAccountId = parsed.accountId;
        logger.info(`Loaded selected account: ${this.selectedAccountId}`);
      }
    } catch (error) {
      logger.error(`Error loading selected account: ${error.message}`);
    }
  }
  async saveTokenCache() {
    try {
      const stamped = wrapCache(this.msalApp.getTokenCache().serialize());
      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.setPassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT, stamped);
        } else {
          const cachePath = getTokenCachePath();
          ensureParentDir(cachePath);
          fs.writeFileSync(cachePath, stamped, { mode: 384 });
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain save failed, falling back to file storage: ${keytarError.message}`
        );
        const cachePath = getTokenCachePath();
        ensureParentDir(cachePath);
        fs.writeFileSync(cachePath, stamped, { mode: 384 });
      }
    } catch (error) {
      logger.error(`Error saving token cache: ${error.message}`);
    }
  }
  async saveSelectedAccount() {
    try {
      const stamped = wrapCache(JSON.stringify({ accountId: this.selectedAccountId }));
      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.setPassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY, stamped);
        } else {
          const accountPath = getSelectedAccountPath();
          ensureParentDir(accountPath);
          fs.writeFileSync(accountPath, stamped, { mode: 384 });
        }
      } catch (keytarError) {
        logger.warn(
          `Keychain save failed for selected account, falling back to file storage: ${keytarError.message}`
        );
        const accountPath = getSelectedAccountPath();
        ensureParentDir(accountPath);
        fs.writeFileSync(accountPath, stamped, { mode: 384 });
      }
    } catch (error) {
      logger.error(`Error saving selected account: ${error.message}`);
    }
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
        await this.saveTokenCache();
        return this.accessToken;
      } catch {
        logger.error("Silent token acquisition failed");
        throw new Error("Silent token acquisition failed");
      }
    }
    throw new Error("No valid token found");
  }
  async getCurrentAccount() {
    const accounts = await this.msalApp.getTokenCache().getAllAccounts();
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
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }
      await this.saveTokenCache();
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
      if (!this.selectedAccountId && response?.account) {
        this.selectedAccountId = response.account.homeAccountId;
        await this.saveSelectedAccount();
        logger.info(`Auto-selected new account: ${response.account.username}`);
      }
      await this.saveTokenCache();
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
      try {
        const kt = await getKeytar();
        if (kt) {
          await kt.deletePassword(SERVICE_NAME, TOKEN_CACHE_ACCOUNT);
          await kt.deletePassword(SERVICE_NAME, SELECTED_ACCOUNT_KEY);
        }
      } catch (keytarError) {
        logger.warn(`Keychain deletion failed: ${keytarError.message}`);
      }
      const cachePath = getTokenCachePath();
      if (fs.existsSync(cachePath)) {
        fs.unlinkSync(cachePath);
      }
      const accountPath = getSelectedAccountPath();
      if (fs.existsSync(accountPath)) {
        fs.unlinkSync(accountPath);
      }
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
      return this.oauthToken;
    }
    let targetAccount = null;
    if (identifier) {
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
      await this.saveTokenCache();
      return response.accessToken;
    } catch {
      throw new Error(
        `Failed to acquire token for account '${targetAccount.username || targetAccount.name || "unknown"}'. The token may have expired. Please re-login with: --login`
      );
    }
  }
}
var auth_default = AuthManager;
export {
  buildScopesFromEndpoints,
  auth_default as default,
  getSelectedAccountPath,
  getTokenCachePath,
  pickNewest,
  unwrapCache,
  wrapCache
};
