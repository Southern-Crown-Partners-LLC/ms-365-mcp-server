import logger from "./logger.js";
import { parseCloudType, getDefaultClientId } from "./cloud-config.js";
class EnvironmentSecretsProvider {
  async getSecrets() {
    const cloudType = parseCloudType(process.env.MS365_MCP_CLOUD_TYPE);
    return {
      clientId: process.env.MS365_MCP_CLIENT_ID || getDefaultClientId(cloudType),
      tenantId: process.env.MS365_MCP_TENANT_ID || "common",
      clientSecret: process.env.MS365_MCP_CLIENT_SECRET,
      cloudType
    };
  }
}
class KeyVaultSecretsProvider {
  constructor(vaultUrl) {
    this.vaultUrl = vaultUrl;
  }
  async getSecrets() {
    const { DefaultAzureCredential } = await import("@azure/identity");
    const { SecretClient } = await import("@azure/keyvault-secrets");
    const credential = new DefaultAzureCredential();
    const client = new SecretClient(this.vaultUrl, credential);
    logger.info(`Fetching secrets from Key Vault: ${this.vaultUrl}`);
    const [clientIdSecret, tenantIdSecret, clientSecretResult, cloudTypeResult] = await Promise.all(
      [
        client.getSecret("ms365-mcp-client-id"),
        client.getSecret("ms365-mcp-tenant-id").catch(() => null),
        client.getSecret("ms365-mcp-client-secret").catch(() => null),
        client.getSecret("ms365-mcp-cloud-type").catch(() => null)
      ]
    );
    if (!clientIdSecret.value) {
      throw new Error("Required secret ms365-mcp-client-id not found in Key Vault");
    }
    logger.info("Successfully retrieved secrets from Key Vault");
    return {
      clientId: clientIdSecret.value,
      tenantId: tenantIdSecret?.value || "common",
      clientSecret: clientSecretResult?.value,
      cloudType: parseCloudType(cloudTypeResult?.value)
    };
  }
}
function createSecretsProvider() {
  const vaultUrl = process.env.MS365_MCP_KEYVAULT_URL;
  if (vaultUrl) {
    logger.info("Key Vault URL configured, using Azure Key Vault for secrets");
    return new KeyVaultSecretsProvider(vaultUrl);
  }
  logger.info("Using environment variables for secrets");
  return new EnvironmentSecretsProvider();
}
let cachedSecrets = null;
async function getSecrets() {
  if (cachedSecrets) {
    return cachedSecrets;
  }
  const provider = createSecretsProvider();
  cachedSecrets = await provider.getSecrets();
  return cachedSecrets;
}
function clearSecretsCache() {
  cachedSecrets = null;
}
export {
  clearSecretsCache,
  getSecrets
};
