import { spawn } from "node:child_process";
import fs, { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import logger from "./logger.js";
const SERVICE_NAME = "ms-365-mcp-server";
const TOKEN_CACHE_ACCOUNT = "msal-token-cache";
const SELECTED_ACCOUNT_KEY = "selected-account";
const AUTH_CACHE_COMMAND_ENV = "MS365_MCP_AUTH_CACHE_COMMAND";
const AUTH_CACHE_COMMAND_TIMEOUT_ENV = "MS365_MCP_AUTH_CACHE_COMMAND_TIMEOUT_MS";
const DEFAULT_AUTH_CACHE_COMMAND_TIMEOUT_MS = 1e4;
const STDERR_LIMIT = 2048;
const COMMAND_KILL_GRACE_MS = 1e3;
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FALLBACK_DIR = __dirname;
const DEFAULT_TOKEN_CACHE_PATH = path.join(FALLBACK_DIR, "..", ".token-cache.json");
const DEFAULT_SELECTED_ACCOUNT_PATH = path.join(FALLBACK_DIR, "..", ".selected-account.json");
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
    } catch {
      logger.info("keytar not available, using file-based credential storage");
      keytar = void 0;
      return null;
    }
  }
  return keytar;
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
  const newest = pickNewestRaw(keytarRaw, fileRaw);
  return newest ? unwrapCache(newest).data : void 0;
}
function pickNewestRaw(keytarRaw, fileRaw) {
  if (!keytarRaw && !fileRaw) return void 0;
  if (keytarRaw && !fileRaw) return keytarRaw;
  if (!keytarRaw && fileRaw) return fileRaw;
  const kt = unwrapCache(keytarRaw);
  const file = unwrapCache(fileRaw);
  if (kt.savedAt === void 0 && file.savedAt === void 0) return keytarRaw;
  if (kt.savedAt !== void 0 && file.savedAt === void 0) return keytarRaw;
  if (kt.savedAt === void 0 && file.savedAt !== void 0) return fileRaw;
  return kt.savedAt >= file.savedAt ? keytarRaw : fileRaw;
}
function getTokenCachePath() {
  const envPath = process.env.MS365_MCP_TOKEN_CACHE_PATH?.trim();
  return envPath || DEFAULT_TOKEN_CACHE_PATH;
}
function getSelectedAccountPath() {
  const envPath = process.env.MS365_MCP_SELECTED_ACCOUNT_PATH?.trim();
  return envPath || DEFAULT_SELECTED_ACCOUNT_PATH;
}
function storageAccountForKey(key) {
  assertValidKey(key);
  return key === "token-cache" ? TOKEN_CACHE_ACCOUNT : SELECTED_ACCOUNT_KEY;
}
function filePathForKey(key) {
  assertValidKey(key);
  return key === "token-cache" ? getTokenCachePath() : getSelectedAccountPath();
}
function assertValidKey(key) {
  if (key !== "token-cache" && key !== "selected-account") {
    throw new Error(`Unknown auth cache storage key: ${String(key)}`);
  }
}
function ensureParentDir(filePath) {
  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true, mode: 448 });
}
function writeFileAtomically(filePath, value) {
  ensureParentDir(filePath);
  const tempPath = path.join(
    path.dirname(filePath),
    `.${path.basename(filePath)}.${process.pid}.${Date.now()}.tmp`
  );
  fs.writeFileSync(tempPath, value, { mode: 384 });
  fs.renameSync(tempPath, filePath);
}
class DefaultTokenCacheStorage {
  constructor() {
    this.description = "default (keytar+file)";
    this.failClosed = false;
  }
  async load(key) {
    assertValidKey(key);
    let keytarRaw;
    try {
      const kt = await getKeytar();
      if (kt) {
        keytarRaw = await kt.getPassword(SERVICE_NAME, storageAccountForKey(key)) ?? void 0;
      }
    } catch (error) {
      logger.warn(`Keychain access failed for ${key}: ${error.message}`);
    }
    let fileRaw;
    const cachePath = filePathForKey(key);
    if (existsSync(cachePath)) {
      fileRaw = readFileSync(cachePath, "utf8");
    }
    return pickNewestRaw(keytarRaw, fileRaw);
  }
  async save(key, value) {
    assertValidKey(key);
    try {
      const kt = await getKeytar();
      if (kt) {
        await kt.setPassword(SERVICE_NAME, storageAccountForKey(key), value);
        return;
      }
    } catch (error) {
      logger.warn(
        `Keychain save failed for ${key}, falling back to file storage: ${error.message}`
      );
    }
    writeFileAtomically(filePathForKey(key), value);
  }
  async delete(key) {
    assertValidKey(key);
    try {
      const kt = await getKeytar();
      if (kt) {
        await kt.deletePassword(SERVICE_NAME, storageAccountForKey(key));
      }
    } catch (error) {
      logger.warn(`Keychain deletion failed for ${key}: ${error.message}`);
    }
    const cachePath = filePathForKey(key);
    try {
      if (fs.existsSync(cachePath)) {
        fs.unlinkSync(cachePath);
      }
    } catch (error) {
      logger.warn(`File deletion failed for ${key}: ${error.message}`);
    }
  }
}
class CommandTokenCacheStorage {
  constructor(commandPath, timeoutMs = DEFAULT_AUTH_CACHE_COMMAND_TIMEOUT_MS, spawnCommand = spawn) {
    this.commandPath = commandPath;
    this.timeoutMs = timeoutMs;
    this.spawnCommand = spawnCommand;
    this.failClosed = true;
    this.description = `command (${path.basename(commandPath)})`;
  }
  async load(key) {
    assertValidKey(key);
    const result = await this.invoke("load", key);
    const trimmed = result.stdout.trim();
    if (trimmed === "") {
      return void 0;
    }
    let parsed;
    try {
      parsed = JSON.parse(trimmed);
    } catch {
      throw new Error(`Auth cache command returned invalid JSON for load ${key}.`);
    }
    if (!parsed || typeof parsed !== "object") {
      throw new Error(`Auth cache command returned invalid JSON shape for load ${key}.`);
    }
    const response = parsed;
    if (response.found === false) {
      return void 0;
    }
    if (response.found === true && typeof response.value === "string") {
      return response.value;
    }
    throw new Error(`Auth cache command returned invalid load response for ${key}.`);
  }
  async save(key, value) {
    assertValidKey(key);
    await this.invoke("save", key, JSON.stringify({ value }));
  }
  async delete(key) {
    assertValidKey(key);
    await this.invoke("delete", key);
  }
  async invoke(operation, key, stdinPayload) {
    let result;
    try {
      result = await runCommand(
        this.commandPath,
        [operation, key],
        stdinPayload,
        this.timeoutMs,
        this.spawnCommand
      );
    } catch (error) {
      throw new Error(
        `Auth cache command failed for ${operation} ${key}: ${error.message}`
      );
    }
    if (result.timedOut) {
      throw new Error(`Auth cache command timed out for ${operation} ${key}.`);
    }
    if (result.exitCode !== 0) {
      throw new Error(
        `Auth cache command failed for ${operation} ${key} (exit ${result.exitCode ?? `signal ${result.signal ?? "unknown"}`})${formatStderr(
          result.stderr
        )}.`
      );
    }
    return result;
  }
}
function formatStderr(stderr) {
  const trimmed = stderr.trim();
  if (!trimmed) {
    return "";
  }
  const truncated = trimmed.length > STDERR_LIMIT ? `${trimmed.slice(0, STDERR_LIMIT)}...` : trimmed;
  return `: ${truncated}`;
}
function runCommand(commandPath, args, stdinPayload, timeoutMs, spawnCommand) {
  return new Promise((resolve, reject) => {
    let child;
    try {
      child = spawnCommand(commandPath, args, { stdio: "pipe", shell: false });
    } catch (error) {
      reject(new Error(`could not be started: ${error.message}`));
      return;
    }
    let stdout = "";
    let stderr = "";
    let timedOut = false;
    let killTimer;
    const timeout = setTimeout(() => {
      timedOut = true;
      child.kill("SIGTERM");
      killTimer = setTimeout(() => {
        child.kill("SIGKILL");
      }, COMMAND_KILL_GRACE_MS);
    }, timeoutMs);
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.stdin.on("error", () => {
    });
    child.once("error", (error) => {
      clearTimeout(timeout);
      if (killTimer) clearTimeout(killTimer);
      reject(new Error(`could not be started: ${error.message}`));
    });
    child.once("close", (exitCode, signal) => {
      clearTimeout(timeout);
      if (killTimer) clearTimeout(killTimer);
      resolve({ exitCode, signal, stdout, stderr, timedOut });
    });
    if (stdinPayload !== void 0) {
      child.stdin.end(stdinPayload, "utf8");
    } else {
      child.stdin.end();
    }
  });
}
function parseTimeoutMs(value) {
  if (value === void 0 || value.trim() === "") {
    return DEFAULT_AUTH_CACHE_COMMAND_TIMEOUT_MS;
  }
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    throw new Error(`${AUTH_CACHE_COMMAND_TIMEOUT_ENV} must be a positive integer.`);
  }
  return parsed;
}
async function assertCommandUsable(commandPath) {
  let stats;
  try {
    stats = await fs.promises.stat(commandPath);
  } catch {
    throw new Error(`${AUTH_CACHE_COMMAND_ENV} points to a path that does not exist.`);
  }
  if (!stats.isFile()) {
    throw new Error(`${AUTH_CACHE_COMMAND_ENV} must point to an executable file.`);
  }
  if (process.platform !== "win32" && (stats.mode & 73) === 0) {
    throw new Error(`${AUTH_CACHE_COMMAND_ENV} must point to an executable file.`);
  }
}
async function createTokenCacheStorage(options = {}) {
  const allowCommandStorage = options.allowCommandStorage ?? true;
  const configuredCommand = process.env[AUTH_CACHE_COMMAND_ENV];
  let storage;
  if (allowCommandStorage && configuredCommand !== void 0) {
    const commandPath = configuredCommand.trim();
    if (commandPath === "") {
      throw new Error(`${AUTH_CACHE_COMMAND_ENV} was provided but is empty.`);
    }
    await assertCommandUsable(commandPath);
    storage = new CommandTokenCacheStorage(
      commandPath,
      parseTimeoutMs(process.env[AUTH_CACHE_COMMAND_TIMEOUT_ENV])
    );
  } else {
    storage = new DefaultTokenCacheStorage();
  }
  if (options.logProvider) {
    logger.info(`Auth cache storage provider: ${storage.description}`);
  }
  return storage;
}
export {
  CommandTokenCacheStorage,
  DefaultTokenCacheStorage,
  createTokenCacheStorage,
  getSelectedAccountPath,
  getTokenCachePath,
  pickNewest,
  unwrapCache,
  wrapCache
};
