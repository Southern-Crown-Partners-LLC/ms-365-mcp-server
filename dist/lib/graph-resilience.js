import logger from "../logger.js";
function loadResilienceConfig() {
  const intEnv = (name, fallback) => {
    const raw = process.env[name];
    if (raw === void 0 || raw === "") return fallback;
    const n = Number.parseInt(raw, 10);
    if (!Number.isFinite(n) || n < 0) {
      logger.warn(`Ignoring invalid ${name}=${JSON.stringify(raw)} (use a non-negative integer)`);
      return fallback;
    }
    return n;
  };
  return {
    maxRetries: intEnv("MS365_MCP_GRAPH_MAX_RETRIES", 3),
    baseBackoffMs: intEnv("MS365_MCP_GRAPH_BASE_BACKOFF_MS", 200),
    maxBackoffMs: intEnv("MS365_MCP_GRAPH_MAX_BACKOFF_MS", 5e3),
    fetchTimeoutMs: intEnv("MS365_MCP_GRAPH_TIMEOUT_MS", 1e5),
    circuitFailureThreshold: intEnv("MS365_MCP_GRAPH_CIRCUIT_THRESHOLD", 5),
    circuitCooldownMs: intEnv("MS365_MCP_GRAPH_CIRCUIT_COOLDOWN_MS", 3e4),
    circuitDisabled: process.env.MS365_MCP_GRAPH_CIRCUIT_DISABLED === "true" || process.env.MS365_MCP_GRAPH_CIRCUIT_DISABLED === "1"
  };
}
class CircuitOpenError extends Error {
  constructor(cooldownMs) {
    super(
      `Graph circuit breaker is open (cooldown ${cooldownMs} ms). Upstream has failed repeatedly; refusing to flood it.`
    );
    this.code = "circuit_open";
    this.name = "CircuitOpenError";
    this.cooldownMs = cooldownMs;
  }
}
class CircuitBreaker {
  constructor(threshold, cooldownMs, disabled, now = () => Date.now()) {
    this.threshold = threshold;
    this.cooldownMs = cooldownMs;
    this.disabled = disabled;
    this.now = now;
    this.failures = 0;
    this.openedAt = null;
  }
  /**
   * @returns the time-remaining (in ms) before the circuit can be probed,
   *          or `null` if the circuit is closed and the call should proceed.
   */
  checkBeforeRequest() {
    if (this.disabled) return null;
    if (this.openedAt === null) return null;
    const elapsed = this.now() - this.openedAt;
    if (elapsed >= this.cooldownMs) {
      return null;
    }
    return this.cooldownMs - elapsed;
  }
  recordSuccess() {
    if (this.failures !== 0 || this.openedAt !== null) {
      logger.info("Graph circuit: success \u2014 closing breaker");
    }
    this.failures = 0;
    this.openedAt = null;
  }
  recordFailure() {
    if (this.disabled) return;
    this.failures += 1;
    if (this.failures >= this.threshold && this.openedAt === null) {
      this.openedAt = this.now();
      logger.warn(
        `Graph circuit: ${this.failures} consecutive failures \u2014 opening breaker for ${this.cooldownMs} ms`
      );
    } else if (this.openedAt !== null) {
      this.openedAt = this.now();
      logger.warn("Graph circuit: probe failed \u2014 extending cooldown");
    }
  }
  /** Exposed for tests / metrics. */
  getState() {
    return {
      failures: this.failures,
      openedAt: this.openedAt,
      open: this.checkBeforeRequest() !== null
    };
  }
}
function parseRetryAfterMs(header) {
  if (!header) return null;
  const trimmed = header.trim();
  if (trimmed === "") return null;
  const asInt = Number.parseInt(trimmed, 10);
  if (Number.isFinite(asInt) && asInt >= 0 && String(asInt) === trimmed) {
    return Math.min(asInt * 1e3, 6e4);
  }
  if (!/[-/:,]| GMT$/i.test(trimmed) && !/\s+\d/.test(trimmed)) {
    return null;
  }
  const dateMs = Date.parse(trimmed);
  if (Number.isFinite(dateMs)) {
    const delta = dateMs - Date.now();
    if (delta <= 0) return 0;
    return Math.min(delta, 6e4);
  }
  return null;
}
function backoffDelayMs(attempt, baseMs, maxMs, rand = Math.random) {
  const exp = Math.min(maxMs, baseMs * 2 ** attempt);
  return Math.floor(rand() * exp);
}
function isRetriableStatus(status) {
  return status === 429 || status === 503 || status === 504;
}
function isMethodIdempotent(method) {
  const m = method.toUpperCase();
  return m === "GET" || m === "HEAD" || m === "PUT" || m === "DELETE" || m === "OPTIONS" || m === "TRACE";
}
function isAbortError(err) {
  return typeof err === "object" && err !== null && "name" in err && err.name === "AbortError";
}
async function fetchWithResilience(url, init, config, breaker, sleep = (ms) => new Promise((r) => setTimeout(r, ms))) {
  const remainingCooldown = breaker.checkBeforeRequest();
  if (remainingCooldown !== null) {
    throw new CircuitOpenError(remainingCooldown);
  }
  const method = (init?.method ?? "GET").toString().toUpperCase();
  const methodIsIdempotent = isMethodIdempotent(method);
  let attempt = 0;
  while (true) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), config.fetchTimeoutMs);
    let response = null;
    let networkError = null;
    try {
      response = await fetch(url, { ...init, signal: controller.signal });
    } catch (err) {
      networkError = err;
    } finally {
      clearTimeout(timer);
    }
    if (response !== null && !isRetriableStatus(response.status)) {
      breaker.recordSuccess();
      return response;
    }
    const is429 = response !== null && response.status === 429;
    const retryAllowedByMethod = methodIsIdempotent || is429;
    const canRetry = attempt < config.maxRetries && retryAllowedByMethod;
    if (!canRetry) {
      breaker.recordFailure();
      if (response !== null) {
        if (!retryAllowedByMethod && attempt === 0) {
          logger.warn(
            `Graph ${method} ${response.status}: not retried (non-idempotent method, side-effect may have landed)`
          );
        }
        return response;
      }
      if (!retryAllowedByMethod && attempt === 0) {
        logger.warn(
          `Graph ${method} network error: not retried (non-idempotent method, side-effect may have landed)`
        );
      }
      throw networkError ?? new Error("Graph fetch failed (unknown error)");
    }
    let delayMs;
    if (response !== null && response.status === 429) {
      const retryAfter = parseRetryAfterMs(response.headers.get("retry-after"));
      delayMs = retryAfter !== null ? retryAfter : backoffDelayMs(attempt, config.baseBackoffMs, config.maxBackoffMs);
    } else {
      delayMs = backoffDelayMs(attempt, config.baseBackoffMs, config.maxBackoffMs);
    }
    const reason = response !== null ? `HTTP ${response.status}` : isAbortError(networkError) ? `timeout (${config.fetchTimeoutMs} ms)` : `network error: ${networkError?.message ?? "unknown"}`;
    logger.warn(
      `Graph retry ${attempt + 1}/${config.maxRetries} after ${reason} \u2014 sleeping ${delayMs} ms`
    );
    if (response !== null) {
      try {
        await response.arrayBuffer();
      } catch {
      }
    }
    breaker.recordFailure();
    attempt += 1;
    await sleep(delayMs);
  }
}
let _sharedBreaker = null;
function getSharedBreaker() {
  if (_sharedBreaker === null) {
    const cfg = loadResilienceConfig();
    _sharedBreaker = new CircuitBreaker(
      cfg.circuitFailureThreshold,
      cfg.circuitCooldownMs,
      cfg.circuitDisabled
    );
  }
  return _sharedBreaker;
}
function __resetSharedBreakerForTests() {
  _sharedBreaker = null;
}
export {
  CircuitBreaker,
  CircuitOpenError,
  __resetSharedBreakerForTests,
  backoffDelayMs,
  fetchWithResilience,
  getSharedBreaker,
  isMethodIdempotent,
  loadResilienceConfig,
  parseRetryAfterMs
};
