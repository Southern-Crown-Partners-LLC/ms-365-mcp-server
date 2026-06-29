import winston from "winston";
import path from "path";
import fs from "fs";
import os from "os";
const logsDir = process.env.MS365_MCP_LOG_DIR || path.join(os.homedir(), ".ms-365-mcp-server", "logs");
const FILE_MODE = 384;
const auditLogPath = path.join(logsDir, "audit.log");
try {
  if (!fs.existsSync(logsDir)) {
    fs.mkdirSync(logsDir, { recursive: true, mode: 448 });
  }
  if (fs.existsSync(auditLogPath)) {
    fs.chmodSync(auditLogPath, FILE_MODE);
  }
} catch {
}
const auditLogger = winston.createLogger({
  level: "info",
  format: winston.format.combine(
    winston.format.timestamp({ format: "YYYY-MM-DDTHH:mm:ss.SSSZ" }),
    winston.format.json()
  ),
  defaultMeta: {
    service: "ms-365-mcp-server",
    stream: "audit"
  },
  transports: [
    new winston.transports.Console({
      // Route audit events to stderr so they don't collide with JSON-RPC on
      // stdout when this server runs in stdio mode. Container platforms
      // (Container Apps, App Service, Docker) capture both stdout and stderr
      // and forward to Log Analytics, so the production audit sink is
      // unaffected. Vitest sets `VITEST=true`; staying silent there avoids
      // polluting unrelated tests that exercise the real graph-tools module.
      stderrLevels: ["info"],
      silent: process.env.SILENT === "true" || process.env.SILENT === "1" || process.env.VITEST === "true"
    }),
    new winston.transports.File({
      filename: auditLogPath,
      options: { flags: "a", mode: FILE_MODE }
    })
  ]
});
function isAuditLogEnabled() {
  return process.env.MS365_MCP_AUDIT_LOG !== "false";
}
function auditLog(evt) {
  if (!isAuditLogEnabled()) return;
  auditLogger.info(evt);
}
function getUserIdentityForAudit(token) {
  if (!token) return void 0;
  try {
    const parts = token.split(".");
    if (parts.length < 2) return void 0;
    let b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padNeeded = (4 - b64.length % 4) % 4;
    b64 = b64 + "=".repeat(padNeeded);
    const payload = JSON.parse(Buffer.from(b64, "base64").toString("utf-8"));
    const candidate = payload.upn || payload.preferred_username || payload.email || payload.sub;
    return typeof candidate === "string" ? candidate : void 0;
  } catch {
    return void 0;
  }
}
const __testing = { auditLogger, auditLogPath };
export {
  __testing,
  auditLog,
  getUserIdentityForAudit,
  isAuditLogEnabled
};
