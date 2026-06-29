import winston from "winston";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";
import os from "os";
import { redactionEnabled, redactSensitive } from "./lib/log-redactor.js";
const redactFormat = winston.format((info) => {
  if (!redactionEnabled()) return info;
  if (typeof info.message === "string") {
    info.message = redactSensitive(info.message);
  }
  return info;
});
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const logsDir = process.env.MS365_MCP_LOG_DIR || path.join(os.homedir(), ".ms-365-mcp-server", "logs");
if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir, { recursive: true, mode: 448 });
} else {
  try {
    fs.chmodSync(logsDir, 448);
  } catch {
  }
}
const FILE_MODE = 384;
function ensureFileMode(filePath) {
  try {
    if (fs.existsSync(filePath)) {
      fs.chmodSync(filePath, FILE_MODE);
    }
  } catch {
  }
}
const errorLogPath = path.join(logsDir, "error.log");
const serverLogPath = path.join(logsDir, "mcp-server.log");
ensureFileMode(errorLogPath);
ensureFileMode(serverLogPath);
const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || "info",
  format: winston.format.combine(
    redactFormat(),
    winston.format.timestamp({
      format: "YYYY-MM-DD HH:mm:ss"
    }),
    winston.format.printf(({ level, message, timestamp }) => {
      return `${timestamp} ${level.toUpperCase()}: ${message}`;
    })
  ),
  transports: [
    new winston.transports.File({
      filename: errorLogPath,
      level: "error",
      options: { flags: "a", mode: FILE_MODE }
    }),
    new winston.transports.File({
      filename: serverLogPath,
      options: { flags: "a", mode: FILE_MODE }
    })
  ]
});
const enableConsoleLogging = () => {
  logger.add(
    new winston.transports.Console({
      format: winston.format.combine(
        redactFormat(),
        winston.format.colorize(),
        winston.format.simple()
      ),
      silent: process.env.SILENT === "true" || process.env.SILENT === "1"
    })
  );
};
var logger_default = logger;
export {
  logger_default as default,
  enableConsoleLogging
};
