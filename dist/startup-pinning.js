const LOCAL_ACCOUNT_COMMANDS = [
  "login",
  "logout",
  "listAccounts",
  "selectAccount",
  "removeAccount",
  "verifyLogin"
];
function getExpectedAccountInertWarning(args, authManager) {
  if (!authManager.hasExpectedAccount()) {
    return null;
  }
  const inertModes = [];
  if (args.http) {
    inertModes.push("--http");
  }
  if (args.obo) {
    inertModes.push("--obo");
  }
  if (authManager.isOAuthModeEnabled()) {
    inertModes.push("MS365_MCP_OAUTH_TOKEN");
  }
  if (inertModes.length === 0) {
    return null;
  }
  return `Warning: expected account pinning is configured, but ${inertModes.join(", ")} uses request-provided tokens for Graph calls. The pin only guards local MSAL auth helpers.`;
}
function shouldAssertExpectedAccountAtStartup(args, authManager) {
  if (!authManager.hasExpectedAccount()) {
    return false;
  }
  if (getExpectedAccountInertWarning(args, authManager)) {
    return false;
  }
  return !LOCAL_ACCOUNT_COMMANDS.some((key) => Boolean(args[key]));
}
function shouldUseLocalAuthStorage(args) {
  if (!args.http) {
    return true;
  }
  if (args.enableAuthTools) {
    return true;
  }
  return LOCAL_ACCOUNT_COMMANDS.some((key) => Boolean(args[key]));
}
export {
  getExpectedAccountInertWarning,
  shouldAssertExpectedAccountAtStartup,
  shouldUseLocalAuthStorage
};
