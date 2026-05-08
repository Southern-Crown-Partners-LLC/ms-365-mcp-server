import { ProxyOAuthServerProvider } from "@modelcontextprotocol/sdk/server/auth/providers/proxyProvider.js";
import logger from "./logger.js";
import { getCloudEndpoints } from "./cloud-config.js";
class MicrosoftOAuthProvider extends ProxyOAuthServerProvider {
  constructor(authManager, secrets) {
    const tenantId = secrets.tenantId || "common";
    const clientId = secrets.clientId;
    const cloudEndpoints = getCloudEndpoints(secrets.cloudType);
    super({
      endpoints: {
        authorizationUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/authorize`,
        tokenUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`,
        revocationUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/logout`
      },
      verifyAccessToken: async (token) => {
        try {
          const response = await fetch(`${cloudEndpoints.graphApi}/v1.0/me`, {
            headers: {
              Authorization: `Bearer ${token}`
            }
          });
          if (response.ok) {
            const userData = await response.json();
            logger.info(`OAuth token verified for user: ${userData.userPrincipalName}`);
            await authManager.setOAuthToken(token);
            return {
              token,
              clientId,
              scopes: []
            };
          } else {
            throw new Error(`Token verification failed: ${response.status}`);
          }
        } catch (error) {
          logger.error(`OAuth token verification error: ${error}`);
          throw error;
        }
      },
      getClient: async (client_id) => {
        return {
          client_id,
          redirect_uris: ["http://localhost:3000/callback"]
        };
      }
    });
    this.authManager = authManager;
  }
}
export {
  MicrosoftOAuthProvider
};
