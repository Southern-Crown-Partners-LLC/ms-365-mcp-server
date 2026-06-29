function dumpError(reason, depth = 0) {
  if (depth > 5) {
    return { truncated: true };
  }
  if (reason instanceof Error) {
    const dump = {
      name: reason.name,
      constructor: reason.constructor?.name,
      message: reason.message,
      stack: reason.stack
    };
    const properties = {};
    for (const key of Object.getOwnPropertyNames(reason)) {
      if (key === "name" || key === "message" || key === "stack" || key === "cause") continue;
      properties[key] = reason[key];
    }
    if (Object.keys(properties).length > 0) {
      dump.properties = properties;
    }
    if ("cause" in reason && reason.cause !== void 0) {
      dump.cause = dumpError(reason.cause, depth + 1);
    }
    return dump;
  }
  return { type: typeof reason, value: reason };
}
function getActiveResources() {
  const fn = process.getActiveResourcesInfo;
  if (typeof fn !== "function") {
    return "unavailable (node < 17.3)";
  }
  try {
    return fn.call(process);
  } catch (err) {
    return `error: ${err.message}`;
  }
}
export {
  dumpError,
  getActiveResources
};
